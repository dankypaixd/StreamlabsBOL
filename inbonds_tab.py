"""Pestana "Inbonds": lee BOL o Commercial Invoice, saca el estado del Ship To y el service level.

No lee el contacto (nombre, correo, telefono). En la hoja DATE pone "INBD" en la columna FK
(Priority Code) de cada BOL. Los PDF pueden ser:
* BOL escaneados (se leen igual que en la pestana de contactos, pero solo BOL, estado y carrier);
* Commercial Invoice (se leen con el mismo lector de la pestana 1, que tambien dice si es BEX o VRP).
"""
from __future__ import annotations

import hashlib
import re
import shutil
from typing import Callable

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image

import bol_generator as bg
import contact_extractor as ce
import contact_review as cr
import destinations
import contacts_tab as ct
import date_export as de
import delivery_ui
import service_levels as sl

XLSX_MIME = ct.XLSX_MIME

TEXT: dict[str, dict[str, str]] = {
    "es": {
        "tab_inbonds": "Inbonds",
        "ib_heading": "Inbonds: estado y service level desde el BOL o la Commercial Invoice",
        "ib_intro": (
            "Lee el BOL o la Commercial Invoice, reconoce el estado del Ship To y calcula el service level y la "
            "fecha de entrega. No lee el contacto. En DATE pone INBD en la columna FK (Priority Code)."
        ),
        "ib_upload": "Sube los PDF (BOL o Commercial Invoice, puedes subir varios)",
        "ib_read": "Leer PDF",
        "ib_reading": "Leyendo {name}...",
        "ib_done": "Se leyeron {count} BOL.",
        "ib_kind": "{name}: {kind}",
        "ib_kind_bol": "BOL escaneado",
        "ib_kind_ci": "Commercial Invoice",
        "ib_nothing": "No encontré ningún BOL en: {items}",
        "ib_error": "No se pudo leer {name}: {error}",
        "ib_changed": "Cambiaste los archivos. Presiona “Leer PDF” para procesarlos.",
        "ib_sec_fail": "🔴 Sin estado o sin BOL: complétalos a mano",
        "ib_sec_doubt": "🟡 Dudosos: revisa contra la captura",
        "ib_sec_ok": "🟢 Sin error",
        "ib_sec_confirmed": "✔ Confirmados",
        "ib_confirm_all_ok": "Confirmar todos los verdes",
        "ib_export_heading": "Excel para importar (INBD)",
        "ib_priority_note": "Cada fila lleva INBD en la columna FK (Priority Code).",
        "ib_download": "Descargar Excel de Inbonds",
        "ib_no_capture": "Sin captura (se leyó el texto del PDF).",
    },
    "en": {
        "tab_inbonds": "Inbonds",
        "ib_heading": "Inbonds: state and service level from the BOL or the Commercial Invoice",
        "ib_intro": (
            "Reads the BOL or the Commercial Invoice, recognizes the Ship To state and calculates the service "
            "level and delivery date. It does not read the contact. In DATE it puts INBD in column FK (Priority Code)."
        ),
        "ib_upload": "Upload the PDFs (BOL or Commercial Invoice, you can upload several)",
        "ib_read": "Read PDFs",
        "ib_reading": "Reading {name}...",
        "ib_done": "{count} BOLs were read.",
        "ib_kind": "{name}: {kind}",
        "ib_kind_bol": "scanned BOL",
        "ib_kind_ci": "Commercial Invoice",
        "ib_nothing": "I did not find any BOL in: {items}",
        "ib_error": "Could not read {name}: {error}",
        "ib_changed": "The files changed. Press “Read PDFs” to process them.",
        "ib_sec_fail": "🔴 Missing state or BOL: complete them by hand",
        "ib_sec_doubt": "🟡 Doubtful: check against the capture",
        "ib_sec_ok": "🟢 No error",
        "ib_sec_confirmed": "✔ Confirmed",
        "ib_confirm_all_ok": "Confirm all green rows",
        "ib_export_heading": "Excel to import (INBD)",
        "ib_priority_note": "Every row carries INBD in column FK (Priority Code).",
        "ib_download": "Download Inbonds Excel",
        "ib_no_capture": "No capture (the PDF text was read).",
    },
}



# ------------------------------------------------------------------ lectura


def pdf_kind(content: bytes) -> str:
    """'ci' si las primeras paginas tienen texto de Commercial Invoice, si no 'bol'."""
    doc = fitz.open(stream=content, filetype="pdf")
    try:
        for index in range(min(3, len(doc))):
            text = doc[index].get_text("text").lower()
            if "commercial invoice" in text or "air waybill" in text or "iec technologies" in text:
                return "ci"
    finally:
        doc.close()
    return "bol"


def state_from_lines(lines: tuple[str, ...] | list[str]) -> tuple[str, str, str]:
    """(estado, zip, como) a partir de las lineas del Ship To de una Commercial Invoice.

    Entiende "DALLAS, TX 75201" y las provincias de Canada ("Mississauga, ON L5N 8E9").
    ``como`` es "code" (escrito), "fixed" (corregido por una confusion de OCR, p. ej. 0/O),
    "zip" (deducido del codigo postal) o "name" (nombre completo, p. ej. Puerto Rico).
    """
    rows = [re.sub(r"\s+", " ", line).strip() for line in lines if line and line.strip()]
    for row in reversed(rows):
        for regex in (ce.STATE_RE, ce.CA_POSTAL_RE):
            match = regex.search(row.upper() if regex is ce.CA_POSTAL_RE else row)
            if not match:
                continue
            code, fixed = ce._fix_state_token(match.group(1))
            if code:
                return code, (match.group(2) if regex is ce.STATE_RE else ""), "fixed" if fixed else "code"
    for row in reversed(rows):
        match = ce.ZIP_RE.search(row)
        if match:
            code = sl.state_from_zip(match.group(1))
            if code:
                return code, match.group(1), "zip"
    state = sl.state_from_name("\n".join(rows))
    return (state, "", "name") if state else ("", "", "")


def ship_to_box(words: list[tuple[str, float, float, float, float]], width: float, height: float):
    """Caja (x0, y0, x1, y1) del bloque SHIP TO de una Commercial Invoice, o None si no se ve.

    ``words`` = (texto, x0, y0, x1, y1). El bloque va desde la etiqueta "SHIP TO" hasta la fila
    "HEART ORDER NO" y, a la derecha, hasta la columna "SOLD TO" / "SHIP VIA".
    """
    def tok(text: str) -> str:
        return re.sub(r"[^a-z]", "", text.lower())

    ordered = sorted(words, key=lambda w: (round(w[2] / 6), w[1]))
    label = None
    for index, word in enumerate(ordered[:-1]):
        nxt = ordered[index + 1]
        same_line = abs(nxt[2] - word[2]) < max(8.0, 0.8 * (word[4] - word[2]))
        if tok(word[0]) == "ship" and tok(nxt[0]) == "to" and same_line and word[1] < 0.5 * width:
            label = word  # si hubiera varios, vale el de mas abajo
    if label is None:
        return None
    sold = [w[1] for w in words if tok(w[0]) in ("sold", "via") and w[1] > label[1] + 0.15 * width]
    right = (min(sold) - 6) if sold else 0.55 * width
    heart = [w[2] for w in words if tok(w[0]) == "heart" and w[2] > label[4]]
    bottom = (min(heart) - 4) if heart else label[2] + 0.30 * height
    return (max(0.0, label[1] - 24), max(0.0, label[2] - 14), min(width, right), min(height, bottom))


def ci_ship_to_capture(content: bytes, page_number: int, dpi: int = 150) -> bytes:
    """JPEG con solo el Ship To de la pagina de la Commercial Invoice (si no lo ubica, la parte de arriba)."""
    import io

    doc = fitz.open(stream=content, filetype="pdf")
    try:
        page = doc[page_number - 1]
        top = fitz.Rect(0, 0, page.rect.width, page.rect.height * 0.55)
        pix = page.get_pixmap(clip=top, dpi=dpi, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        scale = dpi / 72
        native = [w for w in page.get_text("words") if w[1] < top.y1]
        if len(native) >= 15:
            words = [(w[4], w[0] * scale, w[1] * scale, w[2] * scale, w[3] * scale) for w in native]
        else:  # pagina escaneada: se lee con OCR solo la parte de arriba
            data = pytesseract.image_to_data(img, config="--psm 6", output_type=pytesseract.Output.DICT)
            words = []
            for text, left, upper, width, height, conf in zip(
                data["text"], data["left"], data["top"], data["width"], data["height"], data["conf"]
            ):
                try:
                    keep = text.strip() and float(conf) >= 25
                except (TypeError, ValueError):
                    keep = False
                if keep:
                    words.append((text, float(left), float(upper), float(left + width), float(upper + height)))
        box = ship_to_box(words, img.width, img.height)
        if box:
            img = img.crop(tuple(int(v) for v in box))
        if img.width > 900:
            img = img.resize((900, int(img.height * 900 / img.width)), Image.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=85)
        return buffer.getvalue()
    except Exception:
        return b""
    finally:
        doc.close()


def contacts_from_invoice(name: str, content: bytes) -> list[cr.Contact]:
    """Un Contact (solo BOL, estado y carrier) por cada Air Waybill de la Commercial Invoice."""
    contacts: list[cr.Contact] = []
    for record in bg.extract_basic_records(content):
        rule, state = destinations.destination_state(record.ship_to_lines)
        if rule:
            zip_code, how = "", ""  # Mexico / Latinoamerica: Laredo (TX) o Miami (FL), como en la pestana 1
        else:
            state, zip_code, how = state_from_lines(record.ship_to_lines)
        contact = cr.Contact(
            key=record.air_waybill, bol=record.air_waybill, source_file=name, pages=[record.source_page],
            carrier_sources={k: v for k, v in (("pdf", record.source_format), ("file", cr.carrier_from_filename(name))) if v},
            state=state, zip_code=zip_code,
            agreement={"state": 1.0, "bol": 1.0},
        )
        if rule:
            contact.reasons.append(f"destination_{rule}")
        elif how == "zip":
            contact.reasons.append("state_from_zip")
        elif how == "name":
            contact.reasons.append("state_from_name")
        elif how == "fixed":
            contact.reasons.append("state_fixed")
        cr.set_carrier(contact)
        contact.ocr = {"state": state, "bol": contact.bol}
        contacts.append(contact)
    return contacts


def read_files(
    files: list[tuple[str, bytes]], progress: Callable[[str, float], None] | None = None,
) -> tuple[list[cr.Contact], dict[str, str], dict[str, str]]:
    """Lee todos los PDF. Devuelve (contactos, tipo por archivo, errores por archivo)."""
    kinds: dict[str, str] = {}
    errors: dict[str, str] = {}
    contacts: list[cr.Contact] = []
    bol_reads: list[ce.PdfRead] = []
    for index, (name, content) in enumerate(files):
        def tick(done: int, total: int, index: int = index, name: str = name) -> None:
            if progress:
                progress(name, (index + done / max(1, total)) / len(files))

        try:
            kind = pdf_kind(content)
            if kind == "bol":
                pdf = ce.read_pdf(content, name, workers=2, progress=tick, contact=False)
                if any(p.kind == "bol" for p in pdf.pages):
                    kinds[name] = "bol"
                    bol_reads.append(pdf)
                    continue
                kind = "ci"  # no habia BOL: puede ser una Commercial Invoice escaneada
            found = contacts_from_invoice(name, content)
            if found:
                kinds[name] = "ci"
                for contact in found:
                    cr.classify(contact, None, need_contact=False)
                contacts.extend(found)
            else:
                errors[name] = "nothing"
        except Exception as error:  # PDF danado o ilegible
            errors[name] = str(error)
    if bol_reads:
        contacts.extend(cr.build_contacts(bol_reads, None, None, None, need_contact=False))
    contacts.sort(key=lambda c: (ct._numeric(c.bol), c.source_file))
    return contacts, kinds, errors


# ------------------------------------------------------------------ acciones


def _confirm(key: str) -> None:
    job = st.session_state.get("ib_job")
    if not job or key not in job["contacts"]:
        return
    contact: cr.Contact = job["contacts"][key]
    k, ss = ct._wk(key), st.session_state
    contact.state = ss.get(f"ib_state_{k}", contact.state) or ""
    contact.bol = (ss.get(f"ib_bol_{k}", contact.bol) or "").strip()
    if not contact.bol or not sl.normalize_state(contact.state):
        contact.error = "ct_need_state_bol"
        return
    contact.error = ""
    chosen = ss.get(f"ib_carrier_{k}", contact.carrier) or ""
    if chosen in de.SHIPPER_NUMBER:
        contact.carrier = chosen
        contact.reasons = [r for r in contact.reasons if not r.startswith("carrier_")]
    cr.reclassify(contact, None, need_contact=False)
    contact.confirmed = True


def _undo(key: str) -> None:
    job = st.session_state.get("ib_job")
    if job and key in job["contacts"]:
        job["contacts"][key].confirmed = False


def _confirm_all_ok() -> None:
    job = st.session_state.get("ib_job")
    if job:
        for contact in job["contacts"].values():
            if not contact.confirmed and contact.status == "ok":
                contact.confirmed = True


# --------------------------------------------------------------------- tarjeta


def _card(t: Callable[..., str], contact: cr.Contact, ready: str, *, ns: str, job: dict) -> None:
    k = f"{ns}_{ct._wk(contact.key)}"
    key_k = ct._wk(contact.key)
    with st.container(border=True):
        st.markdown(f"**{contact.bol or contact.key}** · {contact.source_file} · p. {', '.join(map(str, contact.pages))}")
        left, right = st.columns([1.15, 1])
        with left:
            if not contact.capture and job["kinds"].get(contact.source_file) == "ci":
                # solo el Ship To; se genera al mostrar la tarjeta (no para los 70+ verdes)
                contact.capture = ci_ship_to_capture(job["files"][contact.source_file], contact.pages[0])
            if contact.capture:
                st.image(contact.capture, width="stretch")
            else:
                st.caption(t("ib_no_capture"))
            for reason in contact.reasons:
                st.caption("• " + ct._reason_text(t, reason))
        with right:
            options = ct._state_options()
            current = sl.normalize_state(contact.state)
            st.session_state.setdefault(f"ib_state_{key_k}", current if current in options else "")
            state = st.selectbox(t("ct_state"), options, format_func=ct._state_label, key=f"ib_state_{key_k}")
            st.session_state.setdefault(f"ib_bol_{key_k}", contact.bol)
            st.text_input(t("ct_bol"), key=f"ib_bol_{key_k}")
            carriers = ["", *de.SHIPPER_NUMBER]
            st.session_state.setdefault(f"ib_carrier_{key_k}", contact.carrier if contact.carrier in carriers else "")
            st.selectbox(
                t("ct_carrier"), carriers, format_func=lambda c: f"{c} ({de.SHIPPER_NUMBER[c]})" if c else "—",
                key=f"ib_carrier_{key_k}",
            )
            st.caption(ct._service_text(t, state, ready))
            if contact.error:
                st.error(t(contact.error))
            if contact.confirmed:
                st.success(t("ct_confirmed_badge"))
                st.button(t("ct_undo"), key=f"{k}_undo", on_click=_undo, args=(contact.key,))
            else:
                st.button(t("ct_confirm"), key=f"{k}_ok", type="primary", on_click=_confirm, args=(contact.key,))


def _table(t: Callable[..., str], contacts: list[cr.Contact], ready: str) -> pd.DataFrame:
    return pd.DataFrame([{
        t("ct_col_bol"): c.bol, t("ct_col_state"): c.state, "Carrier": c.carrier,
        t("ct_col_service"): sl.service_label(c.state),
        t("ct_col_delivery"): sl.delivery_text(ready, c.state) if ready and c.state else "",
    } for c in contacts])


def _status(contact: cr.Contact) -> str:
    return "confirmed" if contact.confirmed else contact.status


# ---------------------------------------------------------------------- render


def render(t: Callable[..., str]) -> None:
    st.subheader(t("ib_heading"))
    st.write(t("ib_intro"))

    if shutil.which("tesseract") is None:
        st.error(t("ct_no_tesseract"))
        return

    pdf_files = st.file_uploader(t("ib_upload"), type=["pdf"], accept_multiple_files=True, key="ib_pdfs")
    ready_input = st.text_input(t("ct_ready"), key="ib_ready", help=t("ct_ready_help"))
    ready = de.format_ready(ready_input) if ready_input.strip() else ""
    ready_invalid = bool(ready_input.strip()) and not ready
    if ready_invalid:
        st.warning(t("ct_ready_invalid"))
    if not pdf_files:
        return

    files = [(f.name, f.getvalue()) for f in pdf_files]
    files_hash = hashlib.sha256(b"".join(hashlib.sha256(c).digest() + n.encode() for n, c in files)).hexdigest()[:16]

    if st.button(t("ib_read"), type="primary", width="stretch", key="ib_read_btn"):
        bar = st.progress(0.0)

        def progress(name: str, fraction: float) -> None:
            bar.progress(min(1.0, fraction), text=t("ib_reading", name=name))

        contacts, kinds, errors = read_files(files, progress)
        bar.empty()
        st.session_state.ib_job = {
            "hash": files_hash, "contacts": {c.key: c for c in contacts}, "kinds": kinds, "errors": errors,
            "files": dict(files),
        }

    job = st.session_state.get("ib_job")
    if not job:
        return
    if job["hash"] != files_hash:
        st.info(t("ib_changed"))
        return

    for name, kind in job["kinds"].items():
        st.caption(t("ib_kind", name=name, kind=t(f"ib_kind_{kind}")))
    nothing = [n for n, e in job["errors"].items() if e == "nothing"]
    if nothing:
        st.warning(t("ib_nothing", items=", ".join(nothing)))
    for name, error in job["errors"].items():
        if error != "nothing":
            st.error(t("ib_error", name=name, error=error))

    contacts = sorted(job["contacts"].values(), key=lambda c: (ct._numeric(c.bol), c.source_file))
    if not contacts:
        return
    st.success(t("ib_done", count=len(contacts)))
    groups = {key: [c for c in contacts if _status(c) == key] for key in ("fail", "doubt", "ok", "confirmed")}
    columns = st.columns(4)
    for column, (label, key) in zip(
        columns, (("ct_metric_ok", "ok"), ("ct_metric_doubt", "doubt"), ("ct_metric_fail", "fail"), ("ct_metric_confirmed", "confirmed"))
    ):
        column.metric(t(label), len(groups[key]))

    for key in ("fail", "doubt"):
        st.markdown(f"##### {t(f'ib_sec_{key}')} ({len(groups[key])})")
        if not groups[key]:
            st.caption(t("ct_none"))
        for contact in ct._paged(t, groups[key], f"ib_{key}"):
            _card(t, contact, ready, ns=f"ib_{key}", job=job)

    st.markdown(f"##### {t('ib_sec_ok')} ({len(groups['ok'])})")
    if groups["ok"]:
        st.dataframe(_table(t, groups["ok"], ready), width="stretch", hide_index=True)
        st.button(t("ib_confirm_all_ok"), key="ib_confirm_all", on_click=_confirm_all_ok)
    else:
        st.caption(t("ct_none"))
    if groups["confirmed"]:
        st.markdown(f"##### {t('ib_sec_confirmed')} ({len(groups['confirmed'])})")
        st.dataframe(_table(t, groups["confirmed"], ready), width="stretch", hide_index=True)

    # ------------------------------------------------------------ exportar
    st.divider()
    st.markdown(f"#### {t('ib_export_heading')}")
    pending_doubt, pending_fail = len(groups["doubt"]), len(groups["fail"])
    if pending_doubt or pending_fail:
        st.warning(t("ct_pending_warn", doubt=pending_doubt, fail=pending_fail))
    if ready_invalid:
        st.info(t("ct_need_ready"))  # un READY mal escrito no se exporta en blanco por error
        return

    by_bol: dict[str, cr.Contact] = {}
    for contact in contacts:
        current = by_bol.get(contact.bol)
        if current is None or (contact.confirmed and not current.confirmed):
            by_bol[contact.bol] = contact
    shipments = [
        de.ShipmentRow(bol=bol, state=c.state, carrier=c.carrier if c.carrier in de.SHIPPER_NUMBER else "")
        for bol, c in sorted(by_bol.items(), key=lambda item: ct._numeric(item[0])) if bol
    ]
    no_state = [s.bol for s in shipments if not sl.normalize_state(s.state)]
    no_carrier = [s.bol for s in shipments if not s.carrier]
    if no_state:
        st.warning(t("ct_no_service_warn", items=", ".join(no_state[:20])))
    if no_carrier:
        st.warning(t("ct_no_carrier_warn", items=", ".join(no_carrier[:20])))

    st.caption(t("ib_priority_note") + " " + t("ct_row_count", count=len(shipments)))
    shipments, default_date, default_time = delivery_ui.render(
        t, shipments, ready, prefix="ib", priority=de.PRIORITY_INBOND,
        extra_columns={"Carrier": [s.carrier for s in shipments]},
    )
    blob, results = de.build_workbook(
        shipments, ready, None, include_data=False, priority=de.PRIORITY_INBOND,
        default_date=default_date, default_time=default_time,
    )

    name_input = st.text_input(t("ct_filename"), key="ib_filename", help=t("ct_filename_help"), placeholder=ready or "DATE_IMPORT")
    file_name = de.safe_filename(name_input, ready)
    st.caption(t("ct_download_name", name=file_name))
    st.download_button(
        t("ib_download"), data=blob, file_name=file_name, mime=XLSX_MIME, width="stretch", key="ib_dl", type="primary",
    )
