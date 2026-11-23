"""Comparacion opcional de los BOL de los PDF contra los del DIMS (pestanas "Leer contactos" e Inbonds).

El DIMS ya no aporta datos al Excel: solo sirve para verificar que trae los mismos BOL y la misma cantidad que
los PDF. Es opcional (puede que en ese momento no se tenga) y, si hay diferencias, se pueden aceptar BOL uno por
uno o ignorar la comparacion completa; mientras haya diferencias sin resolver no se habilita la descarga.
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from typing import Callable, Iterable

import streamlit as st

import date_export as de

DIMS_TYPES = ["xlsx", "xlsm", "xls", "csv"]

TEXT: dict[str, dict[str, str]] = {
    "es": {
        "dc_upload": "DIMS (opcional): solo para comparar los BOL con los PDF",
        "dc_upload_help": (
            "Si lo tienes, se comparan los BOL del DIMS contra los de los PDF para verificar que son los mismos y "
            "la misma cantidad. No agrega ni cambia nada en el Excel. Solo hace falta una columna llamada Barcode. "
            "Si no lo tienes ahora, descarga el DATE sin él."
        ),
        "dc_read_ok": "DIMS leído: {boxes} renglones, {bols} BOL distintos.",
        "dc_read_error": "No se pudo leer el DIMS: {error}. Puedes seguir sin él.",
        "dc_header_error": "no encontré una columna llamada Barcode",
        "dc_ok": "✅ Los {count} BOL de los PDF son los mismos que los del DIMS.",
        "dc_diff": "⚠️ Los BOL de los PDF y del DIMS no coinciden: PDF {pdf}, DIMS {dims}.",
        "dc_only_pdf": "Están en los PDF pero no en el DIMS ({count}): {items}",
        "dc_only_dims": "Están en el DIMS pero no en los PDF ({count}): {items}",
        "dc_accept": "Aceptar BOL que no coinciden",
        "dc_accept_help": (
            "Elige los que sí quieres dejar así. Los que están solo en los PDF se quedan en el Excel; los que están "
            "solo en el DIMS no llevan fila porque no hay PDF."
        ),
        "dc_tag_pdf": "solo en PDF",
        "dc_tag_dims": "solo en DIMS",
        "dc_ignore": "Ignorar la comparación",
        "dc_ignore_help": "Permite descargar el DATE aunque los BOL no coincidan.",
        "dc_ignored": "Comparación ignorada: ya puedes descargar el DATE.",
        "dc_accepted_all": "Todas las diferencias están aceptadas: ya puedes descargar el DATE.",
        "dc_pending": "Faltan {count} BOL por aceptar (o marca “Ignorar la comparación”) para descargar el DATE.",
        "dc_pending_download": "Para descargar el DATE, acepta los BOL que no coinciden o ignora la comparación (arriba).",
    },
    "en": {
        "dc_upload": "DIMS (optional): only to compare the BOLs with the PDFs",
        "dc_upload_help": (
            "If you have it, the DIMS BOLs are compared with the PDF BOLs to check they are the same and the same "
            "quantity. It does not add or change anything in the Excel. Only a column named Barcode is required. "
            "If you do not have it yet, download the DATE without it."
        ),
        "dc_read_ok": "DIMS read: {boxes} rows, {bols} distinct BOLs.",
        "dc_read_error": "Could not read the DIMS: {error}. You can go on without it.",
        "dc_header_error": "I could not find a column named Barcode",
        "dc_ok": "✅ The {count} BOLs in the PDFs are the same as the ones in the DIMS.",
        "dc_diff": "⚠️ The BOLs in the PDFs and in the DIMS do not match: PDF {pdf}, DIMS {dims}.",
        "dc_only_pdf": "In the PDFs but not in the DIMS ({count}): {items}",
        "dc_only_dims": "In the DIMS but not in the PDFs ({count}): {items}",
        "dc_accept": "Accept BOLs that do not match",
        "dc_accept_help": (
            "Pick the ones you are fine with. BOLs only in the PDFs stay in the Excel; BOLs only in the DIMS get no "
            "row because there is no PDF."
        ),
        "dc_tag_pdf": "only in PDF",
        "dc_tag_dims": "only in DIMS",
        "dc_ignore": "Ignore the comparison",
        "dc_ignore_help": "Lets you download the DATE even if the BOLs do not match.",
        "dc_ignored": "Comparison ignored: you can download the DATE.",
        "dc_accepted_all": "Every difference is accepted: you can download the DATE.",
        "dc_pending": "{count} BOLs still need to be accepted (or tick “Ignore the comparison”) to download the DATE.",
        "dc_pending_download": "To download the DATE, accept the BOLs that do not match or ignore the comparison (above).",
    },
}


def bol_order(value: str) -> tuple[int, str]:
    digits = re.sub(r"\D", "", value)
    return (int(digits) if digits else 10**18, value)


@dataclass(frozen=True)
class Comparison:
    pdf: list[str]  # BOL unicos de los PDF
    dims: list[str]  # BOL unicos del DIMS
    only_pdf: list[str]
    only_dims: list[str]

    @property
    def matches(self) -> bool:
        return not self.only_pdf and not self.only_dims

    @property
    def differences(self) -> list[str]:
        return sorted(set(self.only_pdf) | set(self.only_dims), key=bol_order)


def compare_bols(pdf_bols: Iterable[str], dims_bols: Iterable[str]) -> Comparison:
    """Compara BOL unicos (sin vacios): mismos BOL y, por lo tanto, misma cantidad."""
    pdf = {b.strip() for b in pdf_bols if b and b.strip()}
    dims = {b.strip() for b in dims_bols if b and b.strip()}
    return Comparison(
        pdf=sorted(pdf, key=bol_order),
        dims=sorted(dims, key=bol_order),
        only_pdf=sorted(pdf - dims, key=bol_order),
        only_dims=sorted(dims - pdf, key=bol_order),
    )


def pending_after(comparison: Comparison, accepted: Iterable[str], ignored: bool) -> list[str]:
    """BOL con diferencia que todavia no se aceptan (vacio si se ignora la comparacion)."""
    if ignored:
        return []
    accepted = set(accepted)
    return [b for b in comparison.differences if b not in accepted]


def uploader(t: Callable[..., str], *, key: str) -> tuple[de.DimsSheet | None, str]:
    """Selector del DIMS (opcional). Devuelve (DIMS leido o None, hash del archivo o 'nodata')."""
    dims_file = st.file_uploader(t("dc_upload"), type=DIMS_TYPES, key=key, help=t("dc_upload_help"))
    if dims_file is None:
        return None, "nodata"
    content = dims_file.getvalue()
    data_hash = hashlib.sha256(content).hexdigest()[:12]
    try:
        dims = de.read_dims(content, dims_file.name)
    except ValueError as error:
        reason = t("dc_header_error") if str(error) == "barcode_header" else str(error)
        st.error(t("dc_read_error", error=reason))
        return None, data_hash
    except Exception as error:  # archivo ilegible
        st.error(t("dc_read_error", error=error))
        return None, data_hash
    st.caption(t("dc_read_ok", boxes=sum(dims.boxes.values()), bols=len(dims.barcodes)))
    return dims, data_hash


def render_comparison(t: Callable[..., str], *, prefix: str, pdf_bols: Iterable[str], dims: de.DimsSheet | None) -> bool:
    """Muestra la comparacion. Devuelve True si se puede descargar (sin DIMS, sin diferencias o ya resueltas)."""
    if dims is None:
        return True
    comparison = compare_bols(pdf_bols, dims.barcodes)
    if comparison.matches:
        st.success(t("dc_ok", count=len(comparison.pdf)))
        return True

    st.warning(t("dc_diff", pdf=len(comparison.pdf), dims=len(comparison.dims)))
    if comparison.only_pdf:
        st.markdown("- " + t("dc_only_pdf", count=len(comparison.only_pdf), items=", ".join(comparison.only_pdf)))
    if comparison.only_dims:
        st.markdown("- " + t("dc_only_dims", count=len(comparison.only_dims), items=", ".join(comparison.only_dims)))

    # La seleccion es por conjunto de diferencias: si cambian los archivos, se empieza de cero.
    signature = hashlib.sha256("|".join(comparison.differences).encode()).hexdigest()[:10]
    ignored = st.checkbox(t("dc_ignore"), key=f"{prefix}_dc_ignore_{signature}", help=t("dc_ignore_help"))
    only_pdf = set(comparison.only_pdf)
    accepted = st.multiselect(
        t("dc_accept"), comparison.differences, key=f"{prefix}_dc_accept_{signature}", help=t("dc_accept_help"),
        format_func=lambda bol: f"{bol} · {t('dc_tag_pdf') if bol in only_pdf else t('dc_tag_dims')}",
        disabled=ignored,
    )
    if ignored:
        st.caption(t("dc_ignored"))
        return True
    pending = pending_after(comparison, accepted, ignored)
    if pending:
        st.info(t("dc_pending", count=len(pending)))
        return False
    st.caption(t("dc_accepted_all"))
    return True
