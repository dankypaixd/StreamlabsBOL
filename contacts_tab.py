"""Pestana "Leer contactos de BOL": lee Ship To Attention de los BOL, clasifica y exporta DATE/DATA."""
from __future__ import annotations

import hashlib
import re
import shutil
from typing import Callable

import pandas as pd
import streamlit as st

import contact_extractor as ce
import contact_memory
import contact_review as cr
import date_export as de
import delivery_ui
import service_levels as sl

PAGE_SIZE = 8
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

TEXT: dict[str, dict[str, str]] = {
    "es": {
        "tab_contacts": "Leer contactos de BOL",
        "ct_heading": "Contacto, service level y fecha de entrega desde el BOL",
        "ct_intro": (
            "Lee el bloque “Ship To Attention” (nombre, correo y teléfono) y el estado del Ship To de cada "
            "Bill of Lading, calcula el service level y la fecha de entrega, y prepara el Excel para importar."
        ),
        "ct_upload_pdfs": "Sube los PDF de los BOL (puedes subir varios)",
        "ct_upload_data": "Sube los DIMS (.xlsx, .xlsm o .csv)",
        "ct_data_help": (
            "Son las dimensiones que recibes. Sirven para comparar los BOL con los de los PDF, confirmar el carrier "
            "(BEX/VRP), tomar el HAZMAT y armar la hoja DATA del Excel. Pueden traer más o menos columnas: "
            "solo hace falta Barcode."
        ),
        "ct_data_ok": "DIMS leído: {boxes} renglones, {bols} BOL distintos.",
        "ct_data_error": "No se pudo leer el DIMS: {error}",
        "ct_data_header_error": "no encontré una columna llamada Barcode",
        "ct_data_missing_cols": "Este DIMS no trae: {items}. No pasa nada, se trabaja sin eso.",
        "ct_cmp_ok": "✅ Los {count} BOL del DIMS son los mismos que los de los PDF.",
        "ct_cmp_diff": "⚠️ Los BOL del DIMS y de los PDF no son iguales: DIMS {dims}, PDF {pdf}.",
        "ct_cmp_only_dims": "Están en el DIMS pero no en los PDF: {items}",
        "ct_cmp_only_pdf": "Están en los PDF pero no en el DIMS: {items}",
        "ct_cmp_carrier": "Carrier distinto entre PDF, DIMS y nombre del archivo en: {items}",
        "ct_ready": "READY DATE (opcional, MM/DD/YY HHMM)",
        "ct_ready_help": (
            "Ejemplo: 11/19/26 1356. Se usa para calcular la fecha de entrega con el cut-off 0800–1700. "
            "Si lo dejas vacío, esos campos quedan en blanco y puedes poner la entrega a mano."
        ),
        "ct_ready_invalid": "El READY DATE no es válido. Usa el formato 11/19/26 1356.",
        "ct_pod": "POD DATE TIME (opcional)",
        "ct_pod_help": "Ejemplo: 11/19/26 20:54. Se usa para el mensaje ARRIVED AT.",
        "ct_weekday": "Incluir el día de la semana (ej. SATURDAY 11/07/26 10:44)",
        "ct_msg_heading": "Mensajes de cut off (para copiar)",
        "ct_msg_both": "Salida y llegada",
        "ct_msg_left": "Solo salida (LEFT PLANT)",
        "ct_msg_arrived": "Solo llegada (ARRIVED AT)",
        "ct_msg_none": "Escribe el READY DATE (y el POD DATE TIME si lo tienes) para ver los mensajes.",
        "ct_filename": "Nombre del archivo (opcional)",
        "ct_filename_help": "Si lo dejas vacío, el archivo se guarda con el READY DATE.",
        "ct_download_name": "Se guardará como: {name}",
        "ct_carrier": "Carrier",
        "ct_read": "Leer BOL",
        "ct_reading": "Leyendo {name}...",
        "ct_no_tesseract": (
            "Falta Tesseract OCR en el servidor. Agrega un archivo packages.txt con tesseract-ocr y "
            "tesseract-ocr-eng en la raíz del repositorio y vuelve a desplegar."
        ),
        "ct_files_changed": "Los archivos cambiaron. Presiona “Leer BOL” para procesarlos.",
        "ct_done": "Se leyeron {contacts} BOL en {pages} páginas.",
        "ct_skipped": "Páginas omitidas porque no son un BOL: {items}",
        "ct_metric_ok": "🟢 Sin error",
        "ct_metric_doubt": "🟡 Dudosos",
        "ct_metric_fail": "🔴 Incompletos",
        "ct_metric_confirmed": "✔ Confirmados",
        "ct_legend": "Cajas en la captura: azul = nombre, verde = correo, naranja = teléfono, morado = estado, rojo = BOL.",
        "ct_sec_fail": "🔴 Incompletos: faltan datos o el formato no es válido",
        "ct_sec_doubt": "🟡 Dudosos: revisa contra la captura",
        "ct_sec_ok": "🟢 Sin error: las lecturas coinciden y el formato es válido",
        "ct_sec_confirmed": "✔ Confirmados",
        "ct_none": "Ninguno.",
        "ct_page": "Página",
        "ct_name": "Nombre",
        "ct_email": "Correo",
        "ct_phone": "Teléfono (dígitos y +)",
        "ct_state": "Estado",
        "ct_bol": "BOL",
        "ct_read_by_ocr": "Lo que leyó: {name} · {email} · {phone} · {state}",
        "ct_before_memory": "Antes de aplicar lo conocido decía: {items}",
        "ct_service": "Service level: {service} · Entrega: {delivery}",
        "ct_service_none": "Elige el estado para calcular el service level.",
        "ct_confirm": "Confirmar",
        "ct_undo": "Deshacer",
        "ct_confirmed_badge": "✔ Confirmado",
        "ct_need_state_bol": "Para confirmar hace falta el BOL y un estado válido.",
        "ct_confirm_all_ok": "Confirmar todos los verdes",
        "ct_view_one": "Ver captura y editar uno",
        "ct_pick": "Elige un BOL",
        "ct_col_bol": "BOL",
        "ct_col_name": "Nombre",
        "ct_col_email": "Correo",
        "ct_col_phone": "Teléfono",
        "ct_col_state": "Estado",
        "ct_col_service": "Service",
        "ct_col_delivery": "Entrega",
        "ct_col_status": "Estado de lectura",
        "ct_export_heading": "Excel para importar",
        "ct_only_verified": "Solo incluir nombre, correo y teléfono de los verdes y confirmados",
        "ct_only_verified_help": (
            "Los dudosos e incompletos sin confirmar salen con esos tres campos vacíos para que no se "
            "importe un dato sin revisar. Desmárcalo para exportar también lo leído en los dudosos."
        ),
        "ct_pending_warn": "Sin confirmar: {doubt} dudosos y {fail} incompletos.",
        "ct_missing_in_pdf": "BOL del DIMS sin PDF: {items}",
        "ct_extra_in_pdf": "BOL de los PDF que no están en el DIMS (no se incluyen): {items}",
        "ct_no_service_warn": "Sin service level por falta de estado: {items}",
        "ct_no_carrier_warn": "Sin carrier (BEX/VRP), no se puede poner el Shipper Number: {items}",
        "ct_download_date": "Descargar solo la hoja DATE",
        "ct_download_full": "Descargar libro completo (DATE + DATA)",
        "ct_need_ready": "Corrige el READY DATE (o déjalo vacío) para poder descargar.",
        "ct_no_hazmat_note": "Sin DIMS, el HAZMAT (DK) sale en 0 para todos y no se arma la hoja DATA.",
        "ct_files_changed_dims": "Cambiaste los archivos o el DIMS. Presiona “Leer BOL” para procesarlos de nuevo.",
        "ct_memory_error": "No se pudo guardar el aprendizaje ahora; se reintentará la próxima vez.",
        "ct_row_count": "{count} filas en DATE.",
        "r_empty_block": "El bloque Ship To Attention está vacío en el PDF.",
        "r_missing:bol": "No se pudo leer el número de BOL.",
        "r_missing:name": "Falta el nombre.",
        "r_missing:email": "Falta el correo.",
        "r_missing:phone": "Falta el teléfono (no aparece en el PDF o no se leyó).",
        "r_missing:state": "No se encontró el estado del Ship To.",
        "r_invalid:name": "El nombre no parece válido.",
        "r_invalid:email": "El correo no tiene un formato válido.",
        "r_invalid:phone": "El teléfono no tiene un largo razonable.",
        "r_invalid:state": "El estado no está en la lista de service levels.",
        "r_email_by_name": "El correo se eligió entre lecturas parecidas porque coincide con el nombre.",
        "r_email_near_edge": "El correo llega al borde del recuadro: pudo cortarse.",
        "r_ambiguous_glyph": "El correo tiene una I/l que el OCR puede confundir (p. ej. Imco / lmco): compáralo con la captura.",
        "r_name_odd": "El nombre trae números o “/” (tal como está impreso): decide si lo dejas así.",
        "r_state_fixed": "El código de estado se corrigió por una confusión típica del OCR (p. ej. 0/O).",
        "r_state_from_zip": "El estado no estaba escrito: se dedujo del código postal.",
        "r_state_from_name": "El estado no estaba como código: se tomó del nombre escrito.",
        "r_state_zip_mismatch": "El estado no coincide con el código postal.",
        "r_destination_mexico": "Destino en México: se entrega en Laredo, TX (como en la pestaña 1).",
        "r_destination_latam": "Destino en Latinoamérica: se entrega en Miami, FL (como en la pestaña 1).",
        "r_destination_unsure": "Las lecturas no se pusieron de acuerdo sobre si el destino es México/Latinoamérica: revisa la captura.",
        "r_tracking_mismatch": "El BOL del encabezado no coincide con el de la fila Tracking No.",
        "r_bol_snapped": "El BOL se ajustó al más parecido que está en el DIMS.",
        "r_bol_not_in_data": "Este BOL no está en el DIMS.",
        "r_carrier_mismatch": "El carrier no coincide entre el PDF, el DIMS y el nombre del archivo: revisa cuál es.",
        "r_carrier_unknown": "No se pudo saber si es BEX o VRP: elige el carrier para poner el Shipper Number.",
        "r_memory_fix:email": "El correo se corrigió con uno ya confirmado antes.",
        "r_memory_fix:domain": "El dominio del correo se corrigió con uno ya confirmado antes.",
        "r_memory_fix:name": "El nombre se corrigió con el ya confirmado para ese correo.",
        "r_memory_fix:phone": "El teléfono se corrigió con el ya confirmado para ese correo.",
        "r_memory_fix:state": "El estado se corrigió con el ya confirmado para ese código postal.",
        "r_memory_exact": "Coincide exactamente con un contacto confirmado antes.",
        "r_pages_disagree": "Las dos páginas del BOL no coinciden en: {field}.",
        "r_lowagree": "Las lecturas del OCR no coinciden entre sí en: {field}.",
        "f_name": "nombre", "f_email": "correo", "f_phone": "teléfono", "f_state": "estado",
        "f_bol": "BOL", "f_phone_prefix": "prefijo del teléfono",
    },
    "en": {
        "tab_contacts": "Read BOL contacts",
        "ct_heading": "Contact, service level and delivery date from the BOL",
        "ct_intro": (
            "Reads the “Ship To Attention” block (name, email and phone) and the Ship To state of each "
            "Bill of Lading, calculates the service level and delivery date, and prepares the import Excel."
        ),
        "ct_upload_pdfs": "Upload the BOL PDFs (you can upload several)",
        "ct_upload_data": "Upload the DIMS (.xlsx, .xlsm or .csv)",
        "ct_data_help": (
            "The dimensions you receive. They are used to compare the BOLs with the PDFs, confirm the carrier "
            "(BEX/VRP), take HAZMAT and build the DATA sheet. They may have more or fewer columns: "
            "only Barcode is required."
        ),
        "ct_data_ok": "DIMS read: {boxes} rows, {bols} distinct BOLs.",
        "ct_data_error": "Could not read the DIMS: {error}",
        "ct_data_header_error": "I could not find a column named Barcode",
        "ct_data_missing_cols": "This DIMS has no: {items}. That is fine, it works without them.",
        "ct_cmp_ok": "✅ The {count} BOLs in the DIMS are the same as the ones in the PDFs.",
        "ct_cmp_diff": "⚠️ The BOLs in the DIMS and in the PDFs are not the same: DIMS {dims}, PDF {pdf}.",
        "ct_cmp_only_dims": "In the DIMS but not in the PDFs: {items}",
        "ct_cmp_only_pdf": "In the PDFs but not in the DIMS: {items}",
        "ct_cmp_carrier": "Carrier differs between PDF, DIMS and file name in: {items}",
        "ct_ready": "READY DATE (optional, MM/DD/YY HHMM)",
        "ct_ready_help": (
            "Example: 11/19/26 1356. Used to compute the delivery date with the 0800–1700 cut-off. "
            "If you leave it empty, those fields stay blank and you can set the delivery by hand."
        ),
        "ct_ready_invalid": "The READY DATE is not valid. Use the format 11/19/26 1356.",
        "ct_pod": "POD DATE TIME (optional)",
        "ct_pod_help": "Example: 11/19/26 20:54. Used for the ARRIVED AT message.",
        "ct_weekday": "Include the day of the week (e.g. SATURDAY 11/07/26 10:44)",
        "ct_msg_heading": "Cut off messages (to copy)",
        "ct_msg_both": "Left and arrived",
        "ct_msg_left": "Left only (LEFT PLANT)",
        "ct_msg_arrived": "Arrived only (ARRIVED AT)",
        "ct_msg_none": "Enter the READY DATE (and the POD DATE TIME if you have it) to see the messages.",
        "ct_filename": "File name (optional)",
        "ct_filename_help": "If you leave it empty, the file is saved with the READY DATE.",
        "ct_download_name": "It will be saved as: {name}",
        "ct_carrier": "Carrier",
        "ct_read": "Read BOLs",
        "ct_reading": "Reading {name}...",
        "ct_no_tesseract": (
            "Tesseract OCR is missing on the server. Add a packages.txt file with tesseract-ocr and "
            "tesseract-ocr-eng at the repository root and redeploy."
        ),
        "ct_files_changed": "The files changed. Press “Read BOLs” to process them.",
        "ct_done": "Read {contacts} BOLs from {pages} pages.",
        "ct_skipped": "Pages skipped because they are not a BOL: {items}",
        "ct_metric_ok": "🟢 No error",
        "ct_metric_doubt": "🟡 Doubtful",
        "ct_metric_fail": "🔴 Incomplete",
        "ct_metric_confirmed": "✔ Confirmed",
        "ct_legend": "Boxes in the capture: blue = name, green = email, orange = phone, purple = state, red = BOL.",
        "ct_sec_fail": "🔴 Incomplete: data missing or invalid format",
        "ct_sec_doubt": "🟡 Doubtful: check against the capture",
        "ct_sec_ok": "🟢 No error: reads agree and the format is valid",
        "ct_sec_confirmed": "✔ Confirmed",
        "ct_none": "None.",
        "ct_page": "Page",
        "ct_name": "Name",
        "ct_email": "Email",
        "ct_phone": "Phone (digits and +)",
        "ct_state": "State",
        "ct_bol": "BOL",
        "ct_read_by_ocr": "OCR read: {name} · {email} · {phone} · {state}",
        "ct_before_memory": "Before applying known data it said: {items}",
        "ct_service": "Service level: {service} · Delivery: {delivery}",
        "ct_service_none": "Pick the state to calculate the service level.",
        "ct_confirm": "Confirm",
        "ct_undo": "Undo",
        "ct_confirmed_badge": "✔ Confirmed",
        "ct_need_state_bol": "To confirm, the BOL and a valid state are required.",
        "ct_confirm_all_ok": "Confirm all green",
        "ct_view_one": "View capture and edit one",
        "ct_pick": "Pick a BOL",
        "ct_col_bol": "BOL",
        "ct_col_name": "Name",
        "ct_col_email": "Email",
        "ct_col_phone": "Phone",
        "ct_col_state": "State",
        "ct_col_service": "Service",
        "ct_col_delivery": "Delivery",
        "ct_col_status": "Read status",
        "ct_export_heading": "Excel to import",
        "ct_only_verified": "Only include name, email and phone for green and confirmed rows",
        "ct_only_verified_help": (
            "Unconfirmed doubtful and incomplete rows leave those three fields empty so that nothing "
            "unreviewed gets imported. Uncheck to also export what was read for doubtful rows."
        ),
        "ct_pending_warn": "Unconfirmed: {doubt} doubtful and {fail} incomplete.",
        "ct_missing_in_pdf": "BOLs in the DIMS without a PDF: {items}",
        "ct_extra_in_pdf": "BOLs in the PDFs that are not in the DIMS (not included): {items}",
        "ct_no_service_warn": "No service level because the state is missing: {items}",
        "ct_no_carrier_warn": "No carrier (BEX/VRP), the Shipper Number cannot be set: {items}",
        "ct_download_date": "Download the DATE sheet only",
        "ct_download_full": "Download full workbook (DATE + DATA)",
        "ct_need_ready": "Fix the READY DATE (or leave it empty) to enable the downloads.",
        "ct_no_hazmat_note": "Without a DIMS, HAZMAT (DK) is 0 for every row and the DATA sheet is not built.",
        "ct_files_changed_dims": "The files or the DIMS changed. Press “Read BOLs” to process them again.",
        "ct_memory_error": "Could not save what was learned right now; it will be retried next time.",
        "ct_row_count": "{count} rows in DATE.",
        "r_empty_block": "The Ship To Attention block is empty in the PDF.",
        "r_missing:bol": "The BOL number could not be read.",
        "r_missing:name": "Name is missing.",
        "r_missing:email": "Email is missing.",
        "r_missing:phone": "Phone is missing (not in the PDF or not read).",
        "r_missing:state": "The Ship To state was not found.",
        "r_invalid:name": "The name does not look valid.",
        "r_invalid:email": "The email format is not valid.",
        "r_invalid:phone": "The phone length is not reasonable.",
        "r_invalid:state": "The state is not in the service level list.",
        "r_email_by_name": "The email was picked among similar reads because it matches the name.",
        "r_email_near_edge": "The email reaches the edge of the box: it may be cut off.",
        "r_ambiguous_glyph": "The email has an I/l that OCR can confuse (e.g. Imco / lmco): compare it with the capture.",
        "r_name_odd": "The name has digits or “/” (as printed): decide whether to keep it that way.",
        "r_state_fixed": "The state code was fixed for a typical OCR confusion (e.g. 0/O).",
        "r_state_from_zip": "The state was not written: it was deduced from the ZIP code.",
        "r_state_from_name": "The state was not a code: it was taken from the written name.",
        "r_state_zip_mismatch": "The state does not match the ZIP code.",
        "r_destination_mexico": "Destination in Mexico: delivered in Laredo, TX (same as tab 1).",
        "r_destination_latam": "Destination in Latin America: delivered in Miami, FL (same as tab 1).",
        "r_destination_unsure": "The reads disagree on whether the destination is Mexico/Latin America: check the capture.",
        "r_tracking_mismatch": "The header BOL does not match the Tracking No row.",
        "r_bol_snapped": "The BOL was adjusted to the closest one in the DIMS.",
        "r_bol_not_in_data": "This BOL is not in the DIMS.",
        "r_carrier_mismatch": "The carrier differs between the PDF, the DIMS and the file name: check which one it is.",
        "r_carrier_unknown": "Could not tell whether it is BEX or VRP: choose the carrier to set the Shipper Number.",
        "r_memory_fix:email": "The email was corrected with one confirmed before.",
        "r_memory_fix:domain": "The email domain was corrected with one confirmed before.",
        "r_memory_fix:name": "The name was corrected with the one confirmed for that email.",
        "r_memory_fix:phone": "The phone was corrected with the one confirmed for that email.",
        "r_memory_fix:state": "The state was corrected with the one confirmed for that ZIP.",
        "r_memory_exact": "Exactly matches a contact confirmed before.",
        "r_pages_disagree": "The two pages of the BOL disagree on: {field}.",
        "r_lowagree": "The OCR reads do not agree with each other on: {field}.",
        "f_name": "name", "f_email": "email", "f_phone": "phone", "f_state": "state",
        "f_bol": "BOL", "f_phone_prefix": "phone prefix",
    },
}


# ------------------------------------------------------------------ utilidades


@st.cache_resource(show_spinner=False)
def get_memory() -> contact_memory.ContactMemory:
    try:
        secrets = dict(st.secrets)
    except Exception:  # sin secrets.toml
        secrets = {}
    return contact_memory.from_secrets(secrets)


def _wk(key: str) -> str:
    """Clave segura para widgets."""
    return re.sub(r"\W+", "_", key)


def _numeric(value: str) -> tuple[int, str]:
    digits = re.sub(r"\D", "", value)
    return (int(digits) if digits else 10**18, value)


def _reason_text(t: Callable[..., str], reason: str) -> str:
    if ":" in reason and reason.split(":", 1)[0] in ("pages_disagree", "lowagree"):
        kind, field = reason.split(":", 1)
        return t(f"r_{kind}", field=t(f"f_{field}"))
    key = f"r_{reason}"
    try:
        return t(key)
    except KeyError:
        return reason


def _state_options() -> list[str]:
    return [""] + sorted(sl.load_table())


def _state_label(code: str) -> str:
    if not code:
        return "—"
    days = sl.service_days(code)
    return f"{code} ({days}D)" if days is not None else code


def _service_text(t: Callable[..., str], state: str, ready: str) -> str:
    service = sl.service_label(state)
    if not service:
        return t("ct_service_none")
    delivery = sl.delivery_text(ready, state) if ready else "—"
    return t("ct_service", service=service, delivery=delivery or "—")


def _current_status(contact: cr.Contact) -> str:
    return "confirmed" if contact.confirmed else contact.status


def _is_trusted(contact: cr.Contact) -> bool:
    return contact.confirmed or contact.status == "ok"


def _learn(contacts: list[cr.Contact]) -> None:
    memory = get_memory()
    rows = []
    for contact in contacts:
        if contact.learned or not _is_trusted(contact):
            continue
        if not (ce.valid_name(contact.name) and ce.valid_email(contact.email) and ce.valid_phone(contact.phone)):
            continue
        rows.append({"name": contact.name, "email": contact.email, "phone": contact.phone,
                     "state": contact.state, "zip_code": contact.zip_code})
        contact.learned = True
    if rows and memory.learn(rows):
        memory.save_async()


# ------------------------------------------------------------------- acciones


def _confirm(key: str) -> None:
    job = st.session_state.get("ct_job")
    if not job or key not in job["contacts"]:
        return
    contact: cr.Contact = job["contacts"][key]
    k, ss = _wk(key), st.session_state
    contact.name = re.sub(r"\s+", " ", ss.get(f"ct_name_{k}", contact.name) or "").strip()
    contact.email = (ss.get(f"ct_email_{k}", contact.email) or "").strip()
    contact.phone = ce.normalize_phone(ss.get(f"ct_phone_{k}", contact.phone) or "")
    contact.state = ss.get(f"ct_state_{k}", contact.state) or ""
    contact.bol = (ss.get(f"ct_bol_{k}", contact.bol) or "").strip()
    if not contact.bol or not sl.normalize_state(contact.state):
        contact.error = "ct_need_state_bol"
        return
    contact.error = ""
    chosen = ss.get(f"ct_carrier_{k}", contact.carrier) or ""
    if chosen in de.SHIPPER_NUMBER:  # lo eligio una persona: ya no hay duda de carrier
        contact.carrier = chosen
        contact.reasons = [r for r in contact.reasons if not r.startswith("carrier_")]
    cr.reclassify(contact, get_memory())
    contact.confirmed = True
    ss[f"ct_phone_{k}"] = contact.phone
    contact.learned = False
    _learn([contact])


def _undo(key: str) -> None:
    job = st.session_state.get("ct_job")
    if job and key in job["contacts"]:
        job["contacts"][key].confirmed = False


def _confirm_all_ok() -> None:
    job = st.session_state.get("ct_job")
    if not job:
        return
    greens = [c for c in job["contacts"].values() if not c.confirmed and c.status == "ok"]
    for contact in greens:
        contact.confirmed = True
    _learn(greens)


# --------------------------------------------------------------------- tarjeta


def _card(t: Callable[..., str], contact: cr.Contact, ready: str, *, ns: str) -> None:
    k = f"{ns}_{_wk(contact.key)}"
    key_k = _wk(contact.key)  # claves que leen los callbacks
    with st.container(border=True):
        head = f"**{contact.bol or contact.key}** · {contact.source_file} · p. {', '.join(map(str, contact.pages))}"
        st.markdown(head)
        left, right = st.columns([1.15, 1])
        with left:
            if contact.capture:
                st.image(contact.capture, width="stretch")
            st.caption(
                t("ct_read_by_ocr",
                  name=contact.ocr.get("name") or "—", email=contact.ocr.get("email") or "—",
                  phone=contact.ocr.get("phone") or "—", state=contact.ocr.get("state") or "—")
            )
            if contact.fixes:
                st.caption(t("ct_before_memory", items=", ".join(f"{f}: {v}" for f, v in contact.fixes.items())))
            for reason in contact.reasons:
                if reason == "memory_exact":
                    st.caption("✅ " + _reason_text(t, reason))
                else:
                    st.caption("• " + _reason_text(t, reason))
        with right:
            for field, label, initial in (
                ("name", "ct_name", contact.name), ("email", "ct_email", contact.email),
                ("phone", "ct_phone", contact.phone),
            ):
                st.session_state.setdefault(f"ct_{field}_{key_k}", initial)
                st.text_input(t(label), key=f"ct_{field}_{key_k}")
            options = _state_options()
            current = sl.normalize_state(contact.state)
            st.session_state.setdefault(f"ct_state_{key_k}", current if current in options else "")
            state = st.selectbox(t("ct_state"), options, format_func=_state_label, key=f"ct_state_{key_k}")
            st.session_state.setdefault(f"ct_bol_{key_k}", contact.bol)
            st.text_input(t("ct_bol"), key=f"ct_bol_{key_k}")
            carriers = ["", *de.SHIPPER_NUMBER]
            st.session_state.setdefault(f"ct_carrier_{key_k}", contact.carrier if contact.carrier in carriers else "")
            st.selectbox(
                t("ct_carrier"), carriers, format_func=lambda c: f"{c} ({de.SHIPPER_NUMBER[c]})" if c else "—",
                key=f"ct_carrier_{key_k}",
            )
            st.caption(_service_text(t, state, ready))
            if contact.error:
                st.error(t(contact.error))
            if contact.confirmed:
                st.success(t("ct_confirmed_badge"))
                st.button(t("ct_undo"), key=f"{k}_undo", on_click=_undo, args=(contact.key,))
            else:
                st.button(t("ct_confirm"), key=f"{k}_ok", type="primary", on_click=_confirm, args=(contact.key,))


def _paged(t: Callable[..., str], items: list[cr.Contact], group: str) -> list[cr.Contact]:
    if len(items) <= PAGE_SIZE:
        return items
    pages = (len(items) + PAGE_SIZE - 1) // PAGE_SIZE
    page = st.number_input(t("ct_page"), 1, pages, 1, key=f"ct_page_{group}")
    start = (int(page) - 1) * PAGE_SIZE
    return items[start:start + PAGE_SIZE]


def _table(t: Callable[..., str], contacts: list[cr.Contact], ready: str) -> pd.DataFrame:
    rows = []
    for c in contacts:
        rows.append({
            t("ct_col_bol"): c.bol, t("ct_col_name"): c.name, t("ct_col_email"): c.email,
            t("ct_col_phone"): c.phone, t("ct_col_state"): c.state,
            t("ct_col_service"): sl.service_label(c.state),
            t("ct_col_delivery"): sl.delivery_text(ready, c.state) if ready and c.state else "",
        })
    return pd.DataFrame(rows)


# ------------------------------------------------------------------- exportar


def _shipments(job: dict, dims: de.DimsSheet | None, only_verified: bool) -> tuple[list[de.ShipmentRow], dict[str, list[str]]]:
    by_bol: dict[str, cr.Contact] = {}
    for contact in job["contacts"].values():
        current = by_bol.get(contact.bol)
        if current is None or (contact.confirmed and not current.confirmed):
            by_bol[contact.bol] = contact
    bols = dims.barcodes if dims else sorted(by_bol, key=_numeric)
    notes: dict[str, list[str]] = {"missing": [], "extra": [], "no_state": [], "no_carrier": []}
    rows: list[de.ShipmentRow] = []
    for bol in bols:
        contact = by_bol.get(bol)
        dims_carrier = dims.carriers.get(bol, "") if dims else ""
        carrier = (contact.carrier if contact else "") or dims_carrier
        carrier = carrier if carrier in de.SHIPPER_NUMBER else ""
        hazmat = bool(dims and dims.hazmat.get(bol))
        if contact is None:
            notes["missing"].append(bol)
            rows.append(de.ShipmentRow(bol=bol, carrier=carrier, hazmat=hazmat))
            continue
        use = _is_trusted(contact) or not only_verified
        rows.append(de.ShipmentRow(
            bol=bol,
            name=contact.name if use else "", email=contact.email if use else "", phone=contact.phone if use else "",
            state=contact.state, carrier=carrier, hazmat=hazmat,
        ))
        if not sl.normalize_state(contact.state):
            notes["no_state"].append(bol)
    if dims:
        notes["extra"] = sorted((set(by_bol) - set(bols)), key=_numeric)
    notes["no_carrier"] = [r.bol for r in rows if not r.carrier]
    return rows, notes


def _compare_panel(t: Callable[..., str], job: dict, dims: de.DimsSheet) -> None:
    """Compara los BOL del DIMS con los de los PDF y avisa si el carrier no coincide."""
    pdf_bols = {c.bol for c in job["contacts"].values() if c.bol}
    dims_bols = set(dims.barcodes)
    only_dims = sorted(dims_bols - pdf_bols, key=_numeric)
    only_pdf = sorted(pdf_bols - dims_bols, key=_numeric)
    if not only_dims and not only_pdf:
        st.success(t("ct_cmp_ok", count=len(dims_bols)))
    else:
        st.warning(t("ct_cmp_diff", dims=len(dims_bols), pdf=len(pdf_bols)))
        if only_dims:
            st.markdown("- " + t("ct_cmp_only_dims", items=", ".join(only_dims)))
        if only_pdf:
            st.markdown("- " + t("ct_cmp_only_pdf", items=", ".join(only_pdf)))
    mismatched = sorted(
        {c.bol or c.key for c in job["contacts"].values() if "carrier_mismatch" in c.reasons}, key=_numeric
    )
    if mismatched:
        st.warning(t("ct_cmp_carrier", items=", ".join(mismatched)))


def _messages_panel(t: Callable[..., str], ready: str) -> dict[str, str]:
    """POD DATE TIME y los mensajes de cut off que antes salian de DATA!O18:O20."""
    st.markdown(f"##### {t('ct_msg_heading')}")
    pod = st.text_input(t("ct_pod"), key="ct_pod", help=t("ct_pod_help"), placeholder="11/19/26 20:54")
    weekday = st.checkbox(t("ct_weekday"), key="ct_weekday")
    messages = de.cutoff_messages(ready, pod, weekday)
    shown = False
    for key, label in (("both", "ct_msg_both"), ("left", "ct_msg_left"), ("arrived", "ct_msg_arrived")):
        if messages[key]:
            shown = True
            st.caption(t(label))
            st.code(messages[key], language=None)  # el icono de copiar viene incluido
    if not shown:
        st.caption(t("ct_msg_none"))
    return {"pod": re.sub(r"\s+", " ", pod or "").strip(), **messages}


def _on_download(job: dict) -> None:
    _learn(list(job["contacts"].values()))


# ---------------------------------------------------------------------- render


def render(t: Callable[..., str]) -> None:
    st.subheader(t("ct_heading"))
    st.write(t("ct_intro"))

    if shutil.which("tesseract") is None:
        st.error(t("ct_no_tesseract"))
        return

    pdf_files = st.file_uploader(t("ct_upload_pdfs"), type=["pdf"], accept_multiple_files=True, key="ct_pdfs")
    dims_file = st.file_uploader(t("ct_upload_data"), type=["xlsx", "xlsm", "xls", "csv"], key="ct_data",
                                 help=t("ct_data_help"))

    dims: de.DimsSheet | None = None
    data_hash = "nodata"
    if dims_file is not None:
        content = dims_file.getvalue()
        data_hash = hashlib.sha256(content).hexdigest()[:12]
        try:
            dims = de.read_dims(content, dims_file.name)
            st.caption(t("ct_data_ok", boxes=len(dims.rows), bols=len(dims.barcodes)))
            if dims.missing:
                st.caption(t("ct_data_missing_cols", items=", ".join(name.upper() for name in dims.missing)))
        except ValueError as error:
            reason = t("ct_data_header_error") if str(error) == "barcode_header" else str(error)
            st.error(t("ct_data_error", error=reason))
        except Exception as error:  # archivo ilegible
            st.error(t("ct_data_error", error=error))

    ready_input = st.text_input(
        t("ct_ready"), value=(dims.ready if dims else ""), key=f"ct_ready_{data_hash}", help=t("ct_ready_help"),
    )
    ready = de.format_ready(ready_input) if ready_input.strip() else ""
    ready_invalid = bool(ready_input.strip()) and not ready
    if ready_invalid:
        st.warning(t("ct_ready_invalid"))
    extras = _messages_panel(t, ready)

    if not pdf_files:
        return

    files = [(f.name, f.getvalue()) for f in pdf_files]
    files_hash = hashlib.sha256(
        b"".join(hashlib.sha256(c).digest() + n.encode() for n, c in files) + data_hash.encode()
    ).hexdigest()[:16]
    barcodes = set(dims.barcodes) if dims else None

    if st.button(t("ct_read"), type="primary", width="stretch", key="ct_read_btn"):
        memory = get_memory()
        bar = st.progress(0.0)
        pdfs: list[ce.PdfRead] = []
        for index, (name, content) in enumerate(files):
            def progress(done: int, total: int, index: int = index, name: str = name) -> None:
                bar.progress(min(1.0, (index + done / max(1, total)) / len(files)), text=t("ct_reading", name=name))

            pdfs.append(ce.read_pdf(content, name, workers=2, progress=progress))
        bar.empty()
        contacts = cr.build_contacts(pdfs, memory, barcodes, dims.carriers if dims else None)
        contacts.sort(key=lambda c: (_numeric(c.bol), c.source_file))
        st.session_state.ct_job = {
            "hash": files_hash,
            "contacts": {c.key: c for c in contacts},
            "skipped": {p.filename: p.skipped for p in pdfs if p.skipped},
            "pages": sum(p.total_pages for p in pdfs),
        }

    job = st.session_state.get("ct_job")
    if not job:
        return
    if job["hash"] != files_hash:
        st.info(t("ct_files_changed_dims"))
        return

    contacts = sorted(job["contacts"].values(), key=lambda c: (_numeric(c.bol), c.source_file))
    st.success(t("ct_done", contacts=len(contacts), pages=job["pages"]))
    if job["skipped"]:
        st.caption(t("ct_skipped", items="; ".join(f"{n}: {', '.join(map(str, p))}" for n, p in job["skipped"].items())))
    if dims:
        _compare_panel(t, job, dims)
    memory = get_memory()
    if memory.last_error:
        st.caption("⚠️ " + t("ct_memory_error"))

    groups = {key: [c for c in contacts if _current_status(c) == key] for key in ("fail", "doubt", "ok", "confirmed")}
    columns = st.columns(4)
    for column, (label, key) in zip(
        columns, (("ct_metric_ok", "ok"), ("ct_metric_doubt", "doubt"), ("ct_metric_fail", "fail"), ("ct_metric_confirmed", "confirmed"))
    ):
        column.metric(t(label), len(groups[key]))
    st.caption(t("ct_legend"))

    for key, title in (("fail", "ct_sec_fail"), ("doubt", "ct_sec_doubt")):
        st.markdown(f"##### {t(title)} ({len(groups[key])})")
        if not groups[key]:
            st.caption(t("ct_none"))
        for contact in _paged(t, groups[key], key):
            _card(t, contact, ready, ns=key)

    st.markdown(f"##### {t('ct_sec_ok')} ({len(groups['ok'])})")
    if groups["ok"]:
        st.dataframe(_table(t, groups["ok"], ready), width="stretch", hide_index=True)
        st.button(t("ct_confirm_all_ok"), key="ct_confirm_all", on_click=_confirm_all_ok)
        with st.expander(t("ct_view_one")):
            options = [c.key for c in groups["ok"]]
            picked = st.selectbox(t("ct_pick"), options, format_func=lambda k: job["contacts"][k].bol or k, key="ct_pick_ok")
            if picked in job["contacts"] and not job["contacts"][picked].confirmed:
                _card(t, job["contacts"][picked], ready, ns="ok")
    else:
        st.caption(t("ct_none"))

    if groups["confirmed"]:
        st.markdown(f"##### {t('ct_sec_confirmed')} ({len(groups['confirmed'])})")
        st.dataframe(_table(t, groups["confirmed"], ready), width="stretch", hide_index=True)
        with st.expander(t("ct_view_one")):
            options = [c.key for c in groups["confirmed"]]
            picked = st.selectbox(t("ct_pick"), options, format_func=lambda k: job["contacts"][k].bol or k, key="ct_pick_conf")
            if picked in job["contacts"] and job["contacts"][picked].confirmed:
                _card(t, job["contacts"][picked], ready, ns="conf")

    # ------------------------------------------------------------ exportar
    st.divider()
    st.markdown(f"#### {t('ct_export_heading')}")
    only_verified = st.checkbox(t("ct_only_verified"), value=True, help=t("ct_only_verified_help"), key="ct_only_verified")
    pending_doubt = sum(1 for c in groups["doubt"])
    pending_fail = sum(1 for c in groups["fail"])
    if pending_doubt or pending_fail:
        st.warning(t("ct_pending_warn", doubt=pending_doubt, fail=pending_fail))

    if ready_invalid:
        st.info(t("ct_need_ready"))  # un READY mal escrito no se exporta en blanco por error
        return
    shipments, notes = _shipments(job, dims, only_verified)
    if notes["missing"]:
        st.warning(t("ct_missing_in_pdf", items=", ".join(notes["missing"][:20])))
    if notes["extra"]:
        st.warning(t("ct_extra_in_pdf", items=", ".join(notes["extra"][:20])))
    if notes["no_state"]:
        st.warning(t("ct_no_service_warn", items=", ".join(notes["no_state"][:20])))
    if notes["no_carrier"]:
        st.warning(t("ct_no_carrier_warn", items=", ".join(notes["no_carrier"][:20])))
    if dims is None:
        st.caption(t("ct_no_hazmat_note"))

    st.caption(t("ct_row_count", count=len(shipments)))
    shipments, default_date, default_time = delivery_ui.render(
        t, shipments, ready, prefix="ct",
        extra_columns={
            t("ct_col_name"): [r.name for r in shipments],
            t("ct_col_phone"): [r.phone for r in shipments],
            t("ct_col_email"): [r.email for r in shipments],
        },
    )
    overrides = {"default_date": default_date, "default_time": default_time}
    date_bytes, results = de.build_workbook(shipments, ready, dims, include_data=False, **overrides)

    name_input = st.text_input(t("ct_filename"), key="ct_filename", help=t("ct_filename_help"), placeholder=ready or "DATE_IMPORT")
    file_name = de.safe_filename(name_input, ready)
    st.caption(t("ct_download_name", name=file_name))
    if dims is not None:
        full_bytes, _ = de.build_workbook(shipments, ready, dims, include_data=True, extras=extras, **overrides)
        st.download_button(
            t("ct_download_full"), data=full_bytes, file_name=file_name, mime=XLSX_MIME,
            width="stretch", key="ct_dl_full", type="primary", on_click=_on_download, args=(job,),
        )
        date_name = de.safe_filename(f"{file_name[:-5]}_DATE", ready)
    else:
        date_name = file_name
    st.download_button(
        t("ct_download_date"), data=date_bytes, file_name=date_name, mime=XLSX_MIME,
        width="stretch", key="ct_dl_date", on_click=_on_download, args=(job,),
    )
