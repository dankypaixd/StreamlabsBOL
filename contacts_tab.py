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
import dims_compare
import service_levels as sl

PAGE_SIZE = 8
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

TEXT: dict[str, dict[str, str]] = {
    "es": {
        "tab_contacts": "Leer contactos de BOL",
        "ct_heading": "Contacto y service level desde el BOL",
        "ct_intro": (
            "Lee el bloque “Ship To Attention” (nombre, correo y teléfono) y el estado del Ship To de cada "
            "Bill of Lading, calcula el service level y prepara el Excel para importar."
        ),
        "ct_upload_pdfs": "Sube los PDF de los BOL (puedes subir varios)",
        "ct_ready": "READY DATE (opcional, MM/DD/YY HHMM)",
        "ct_ready_help": (
            "Ejemplo: 11/19/26 1356. Va al Excel (Pickup Date, Time y Close Time), se usa para calcular la entrega "
            "que ves en pantalla (cut-off 0800–1700) y para el mensaje LEFT PLANT. Si lo dejas vacío, esos campos quedan en blanco."
        ),
        "ct_ready_invalid": "El READY DATE no es válido. Usa el formato 11/19/26 1356.",
        "ct_push_warn": (
            "⚠️ El READY DATE es a las {time}, después de las 17:00 (fuera del cut-off): "
            "la entrega se está corriendo (push) un día."
        ),
        "ct_pod": "POD DATE TIME (opcional)",
        "ct_pod_help": "Ejemplo: 11/20/26 0800. Solo sirve para armar el mensaje ARRIVED AT que copias; no va al Excel de importar.",
        "ct_weekday": "Con día de la semana",
        "ct_weekday_help": "Agrega el día al mensaje, ej. SATURDAY 11/07/26 10:44.",
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
        "ct_service": "Service level: {service}",
        "ct_service_ready": "Service level: {service} · READY DATE: {date} · {time}",
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
        "ct_col_ready_date": "Fecha ready",
        "ct_col_ready_time": "Hora ready",
        "ct_col_status": "Estado de lectura",
        "ct_export_heading": "Excel para importar",
        "ct_include_attention": "Incluir el Ship To Attention (nombre, correo y teléfono) en la hoja DATE",
        "ct_include_attention_help": "Si lo desmarcas, Consignee Contact, Consignee Phone y Consignee Email salen vacíos.",
        "ct_include_delivery": "Incluir la fecha de entrega en la hoja DATE",
        "ct_include_delivery_help": (
            "Scheduled Delivery Date y Time. La fecha se calcula con el READY DATE y el service level de cada BOL "
            "(días hábiles, cut-off 0800–1700) y la hora sale 1600. Puedes cambiarla abajo, en la tabla."
        ),
        "ct_delivery_all": "Fecha de entrega para todos los BOL (opcional, MM/DD/YY)",
        "ct_delivery_all_help": "Si la escribes, todos los BOL llevan esa fecha en vez de la calculada. Después puedes cambiar alguno en la tabla.",
        "ct_delivery_all_invalid": "La fecha de entrega para todos no es válida. Usa el formato 11/27/26.",
        "ct_delivery_need_ready": "Escribe el READY DATE arriba para que la fecha de entrega se calcule sola (o escríbela a mano).",
        "ct_delivery_hint": "Toca una fecha de la columna «{col}» para cambiarla. Las que cambies se conservan aunque cambies el READY DATE.",
        "ct_delivery_changed": "{count} fecha(s) de entrega cambiadas a mano: {items}",
        "ct_delivery_reset": "Volver a calcular todas las fechas de entrega",
        "ct_delivery_invalid": "Fecha de entrega no válida en: {items}. Usa el formato 11/27/26 (o déjala en blanco).",
        "ct_col_delivery": "Fecha de entrega",
        "ct_col_delivery_help": "MM/DD/YY. Si la dejas en blanco, ese BOL sale sin fecha de entrega.",
        "ct_include_service": "Incluir el service level en la hoja DATE",
        "ct_include_service_help": "Si lo desmarcas, la columna Service sale vacía.",
        "ct_deduped": "BOL repetidos: ya estaban en la lista y se dejó un solo registro de cada uno ({items}).",
        "ct_only_verified": "Solo incluir nombre, correo y teléfono de los verdes y confirmados",
        "ct_only_verified_help": (
            "Los dudosos e incompletos sin confirmar salen con esos tres campos vacíos para que no se "
            "importe un dato sin revisar. Desmárcalo para exportar también lo leído en los dudosos."
        ),
        "ct_pending_warn": "Sin confirmar: {doubt} dudosos y {fail} incompletos.",
        "ct_no_service_warn": "Sin service level por falta de estado: {items}",
        "ct_no_carrier_warn": "Sin carrier (BEX/VRP), no se puede poner el Shipper Number: {items}",
        "ct_download_date": "Descargar la hoja DATE",
        "ct_need_ready": "Corrige el READY DATE (o déjalo vacío) para poder descargar.",
        "ct_files_changed_dims": "Cambiaste los archivos o el DIMS. Presiona “Leer BOL” para procesarlos de nuevo.",
        "ct_memory_error": "No se pudo guardar el aprendizaje ahora; se reintentará la próxima vez.",
        "ct_row_count": "{count} filas en DATE.",
        "r_empty_block": "El bloque Ship To Attention está vacío en el PDF: nombre, correo y teléfono salen en blanco.",
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
        "r_tracking_mismatch_values": "El BOL del encabezado no coincide con la fila Tracking No, que dice {tracking}: puede ser una mala lectura de un dígito en uno de los dos.",
        "r_tracking_other": "La fila Tracking No dice {tracking}, otro número que no se parece al BOL (no cuenta como error).",
        "r_bol_snapped": "El BOL se ajustó al más parecido que está en el DIMS.",
        "r_bol_similar": "Este BOL se parece al {bol}, que ya está en la lista: puede ser una mala lectura de un dígito.",
        "r_bol_alt": "El OCR dudó entre este BOL y el {bol}: revisa el dígito que cambia.",
        "ct_use_bol": "Usar {bol}",
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
        "ct_heading": "Contact and service level from the BOL",
        "ct_intro": (
            "Reads the “Ship To Attention” block (name, email and phone) and the Ship To state of each "
            "Bill of Lading, calculates the service level and prepares the import Excel."
        ),
        "ct_upload_pdfs": "Upload the BOL PDFs (you can upload several)",
        "ct_ready": "READY DATE (optional, MM/DD/YY HHMM)",
        "ct_ready_help": (
            "Example: 11/19/26 1356. Goes to the Excel (Pickup Date, Time and Close Time), is used to compute the delivery "
            "shown on screen (0800–1700 cut-off) and for the LEFT PLANT message. If you leave it empty, those fields stay blank."
        ),
        "ct_ready_invalid": "The READY DATE is not valid. Use the format 11/19/26 1356.",
        "ct_push_warn": (
            "⚠️ The READY DATE is at {time}, after 17:00 (outside the cut-off): "
            "the delivery is being pushed one day."
        ),
        "ct_pod": "POD DATE TIME (optional)",
        "ct_pod_help": "Example: 11/20/26 0800. Only used to build the ARRIVED AT message you copy; it does not go to the import Excel.",
        "ct_weekday": "With weekday",
        "ct_weekday_help": "Adds the day to the message, e.g. SATURDAY 11/07/26 10:44.",
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
        "ct_service": "Service level: {service}",
        "ct_service_ready": "Service level: {service} · READY DATE: {date} · {time}",
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
        "ct_col_ready_date": "Ready date",
        "ct_col_ready_time": "Ready time",
        "ct_col_status": "Read status",
        "ct_export_heading": "Excel to import",
        "ct_include_attention": "Include the Ship To Attention (name, email and phone) in the DATE sheet",
        "ct_include_attention_help": "If unchecked, Consignee Contact, Consignee Phone and Consignee Email are left empty.",
        "ct_include_delivery": "Include the delivery date in the DATE sheet",
        "ct_include_delivery_help": (
            "Scheduled Delivery Date and Time. The date is computed from the READY DATE and each BOL's service level "
            "(business days, cut-off 0800–1700) and the time is 1600. You can change it below, in the table."
        ),
        "ct_delivery_all": "Delivery date for all BOLs (optional, MM/DD/YY)",
        "ct_delivery_all_help": "If you type one, every BOL gets that date instead of the computed one. You can still change individual rows in the table.",
        "ct_delivery_all_invalid": "The delivery date for all is not valid. Use the format 11/27/26.",
        "ct_delivery_need_ready": "Type the READY DATE above so the delivery date is computed automatically (or type it by hand).",
        "ct_delivery_hint": "Tap a date in the \"{col}\" column to change it. Your changes are kept even if you change the READY DATE.",
        "ct_delivery_changed": "{count} delivery date(s) changed by hand: {items}",
        "ct_delivery_reset": "Recalculate all delivery dates",
        "ct_delivery_invalid": "Invalid delivery date in: {items}. Use the format 11/27/26 (or leave it empty).",
        "ct_col_delivery": "Delivery date",
        "ct_col_delivery_help": "MM/DD/YY. If left empty, that BOL goes without a delivery date.",
        "ct_include_service": "Include the service level in the DATE sheet",
        "ct_include_service_help": "If unchecked, the Service column is left empty.",
        "ct_deduped": "Repeated BOLs: they were already in the list, so only one record of each was kept ({items}).",
        "ct_only_verified": "Only include name, email and phone for green and confirmed rows",
        "ct_only_verified_help": (
            "Unconfirmed doubtful and incomplete rows leave those three fields empty so that nothing "
            "unreviewed gets imported. Uncheck to also export what was read for doubtful rows."
        ),
        "ct_pending_warn": "Unconfirmed: {doubt} doubtful and {fail} incomplete.",
        "ct_no_service_warn": "No service level because the state is missing: {items}",
        "ct_no_carrier_warn": "No carrier (BEX/VRP), the Shipper Number cannot be set: {items}",
        "ct_download_date": "Download the DATE sheet",
        "ct_need_ready": "Fix the READY DATE (or leave it empty) to enable the download.",
        "ct_files_changed_dims": "The files or the DIMS changed. Press “Read BOLs” to process them again.",
        "ct_memory_error": "Could not save what was learned right now; it will be retried next time.",
        "ct_row_count": "{count} rows in DATE.",
        "r_empty_block": "The Ship To Attention block is empty in the PDF: name, email and phone are left blank.",
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
        "r_tracking_mismatch_values": "The header BOL does not match the Tracking No row, which says {tracking}: one of the two may have a misread digit.",
        "r_tracking_other": "The Tracking No row says {tracking}, a different number that does not look like the BOL (not counted as an error).",
        "r_bol_snapped": "The BOL was adjusted to the closest one in the DIMS.",
        "r_bol_similar": "This BOL looks like {bol}, which is already in the list: it may be a one-digit misread.",
        "r_bol_alt": "The OCR hesitated between this BOL and {bol}: check the digit that differs.",
        "ct_use_bol": "Use {bol}",
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
    if reason.startswith("bol_similar:"):
        return t("r_bol_similar", bol=reason.split(":", 1)[1])
    if reason.startswith("bol_alt:"):
        return t("r_bol_alt", bol=reason.split(":", 1)[1])
    if reason.startswith("tracking_mismatch:"):
        return t("r_tracking_mismatch_values", tracking=reason.split(":", 1)[1])
    if reason.startswith("tracking_other:"):
        return t("r_tracking_other", tracking=reason.split(":", 1)[1])
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


def ready_parts(ready: str) -> tuple[str, str]:
    """'11/19/26 1356' -> ('11/19/26', '1356'); ('', '') si no hay READY DATE."""
    parts = (ready or "").split()
    return (parts[0], parts[1]) if len(parts) == 2 else ("", "")


def ready_columns(t: Callable[..., str], ready: str) -> dict[str, str]:
    """Las dos columnas de las tablas con la fecha y la hora del READY DATE (ninguna si no hay READY DATE)."""
    date, hour = ready_parts(ready)
    return {t("ct_col_ready_date"): date, t("ct_col_ready_time"): hour} if date else {}


def _service_text(t: Callable[..., str], state: str, ready: str) -> str:
    service = sl.service_label(state)
    if not service:
        return t("ct_service_none")
    date, hour = ready_parts(ready)
    if date:
        return t("ct_service_ready", service=service, date=date, time=hour)
    return t("ct_service", service=service)


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
    removed = cr.merge_duplicate_bol(job["contacts"], contact)  # un BOL corregido a uno que ya estaba: queda uno solo
    if removed:
        job.setdefault("deduped", []).extend(removed)


def use_bol(widget_key: str, bol: str) -> None:
    """Pone en el campo BOL de la tarjeta el que se sugiere (el usuario sigue teniendo que confirmar)."""
    st.session_state[widget_key] = bol


def similar_bol_buttons(t: Callable[..., str], contact: cr.Contact, *, prefix: str, key_k: str, k: str) -> None:
    """Un boton "Usar {BOL}" por cada BOL sugerido (el otro valor que dudo el OCR o uno parecido de la lista)."""
    shown: set[str] = set()
    for reason in contact.reasons:
        if reason.startswith(("bol_similar:", "bol_alt:")):
            bol = reason.split(":", 1)[1]
            if bol in shown:
                continue
            shown.add(bol)
            st.button(
                t("ct_use_bol", bol=bol), key=f"{k}_use_{bol}", on_click=use_bol, args=(f"{prefix}_bol_{key_k}", bol),
            )


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
            similar_bol_buttons(t, contact, prefix="ct", key_k=key_k, k=k)
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
            **ready_columns(t, ready),
        })
    return pd.DataFrame(rows)


# ------------------------------------------------------------------- exportar


def _shipments(job: dict, only_verified: bool) -> tuple[list[de.ShipmentRow], dict[str, list[str]]]:
    """Una fila por BOL de los PDF (el DIMS ya no decide que filas lleva el Excel)."""
    by_bol: dict[str, cr.Contact] = {}
    for contact in job["contacts"].values():
        current = by_bol.get(contact.bol)
        if current is None or (contact.confirmed and not current.confirmed):
            by_bol[contact.bol] = contact
    notes: dict[str, list[str]] = {"no_state": [], "no_carrier": []}
    rows: list[de.ShipmentRow] = []
    for bol in sorted(by_bol, key=_numeric):
        if not bol:
            continue  # un PDF sin BOL legible no tiene fila (sale en rojo para corregirlo)
        contact = by_bol[bol]
        carrier = contact.carrier if contact.carrier in de.SHIPPER_NUMBER else ""
        use = _is_trusted(contact) or not only_verified
        rows.append(de.ShipmentRow(
            bol=bol,
            name=contact.name if use else "", email=contact.email if use else "", phone=contact.phone if use else "",
            state=contact.state, carrier=carrier, consignee=contact.consignee,
        ))
        if not sl.normalize_state(contact.state):
            notes["no_state"].append(bol)
    notes["no_carrier"] = [r.bol for r in rows if not r.carrier]
    return rows, notes


def push_warning(t: Callable[..., str], ready: str) -> None:
    """Aviso cuando el READY DATE cae despues de las 17:00: el dia 0 se corre un dia habil."""
    if sl.is_after_cutoff(ready):
        st.warning(t("ct_push_warn", time=f"{sl.parse_ready(ready):%H:%M}"))


def ready_block(
    t: Callable[..., str], container, *, prefix: str, ready_key: str, default: str = "",
) -> tuple[str, bool, dict[str, str]]:
    """Arriba de la pestana: READY DATE y POD DATE TIME en una fila y, debajo, los mensajes para copiar.

    Los mensajes cambian con lo que se escribe. Sin READY DATE (y sin POD) no se ve ninguno. El POD solo sirve
    para armar el mensaje ARRIVED AT. Devuelve (READY normalizado, si esta mal escrito, extras para el Excel).
    """
    with container:
        c_ready, c_pod, c_day = st.columns([3, 3, 2])
        ready_input = c_ready.text_input(t("ct_ready"), value=default, key=ready_key, help=t("ct_ready_help"))
        pod = c_pod.text_input(t("ct_pod"), key=f"{prefix}_pod", help=t("ct_pod_help"), placeholder="11/20/26 0800")
        weekday = c_day.checkbox(t("ct_weekday"), key=f"{prefix}_weekday", help=t("ct_weekday_help"))
        ready = de.format_ready(ready_input) if ready_input.strip() else ""
        invalid = bool(ready_input.strip()) and not ready
        if invalid:
            st.warning(t("ct_ready_invalid"))
        push_warning(t, ready)
        messages = de.cutoff_messages(ready, pod, weekday)
        shown = [messages[key] for key in ("both", "left", "arrived") if messages[key]]
        if shown:  # uno al lado del otro para que ocupen poco; el icono de copiar viene incluido en cada uno
            for column, text in zip(st.columns(len(shown)), shown):
                column.code(text, language=None, wrap_lines=True)
    return ready, invalid, {"pod": re.sub(r"\s+", " ", pod or "").strip(), **messages}


# ---------------------------------------------------------- fecha de entrega


def _reset_delivery(prefix: str) -> None:
    st.session_state[f"{prefix}_delivery_over"] = {}
    st.session_state[f"{prefix}_delivery_all"] = ""
    st.session_state[f"{prefix}_delivery_gen"] = st.session_state.get(f"{prefix}_delivery_gen", 0) + 1  # editor nuevo


def delivery_inputs(
    t: Callable[..., str], *, prefix: str, ready: str, shipments: list[de.ShipmentRow],
) -> tuple[dict[str, str], bool]:
    """Fecha de entrega para todos y boton de recalcular. Deja en cada fila `delivery` (calculada o la cambiada).

    Devuelve (fecha base de cada BOL, hay error). La base es la fecha para todos, si se escribio, o la calculada
    (READY DATE + service level); lo que la persona cambia en la tabla se guarda aparte y gana sobre la base.
    """
    over: dict[str, str] = st.session_state.setdefault(f"{prefix}_delivery_over", {})
    c_all, c_reset = st.columns([3, 2], vertical_alignment="bottom")
    all_text = c_all.text_input(
        t("ct_delivery_all"), key=f"{prefix}_delivery_all", help=t("ct_delivery_all_help"), placeholder="Ej. 11/27/26",
    )
    c_reset.button(t("ct_delivery_reset"), key=f"{prefix}_delivery_reset", on_click=_reset_delivery, args=(prefix,))
    all_date = de.format_delivery(all_text, ready) if all_text.strip() else ""
    error = bool(all_text.strip()) and not all_date
    if error:
        st.warning(t("ct_delivery_all_invalid"))
    if not ready and not all_date:
        st.caption(t("ct_delivery_need_ready"))
    base = {row.bol: (all_date or de.auto_delivery(row, ready)) for row in shipments}
    for row in shipments:
        row.delivery = over.get(row.bol, base[row.bol])
    return base, error


def delivery_table(
    t: Callable[..., str], *, prefix: str, ready: str, shipments: list[de.ShipmentRow], base: dict[str, str],
    frame: pd.DataFrame, position: int,
) -> tuple[bool, bool]:
    """Vista previa editable: la columna de fecha de entrega se puede cambiar celda por celda.

    Aplica lo que se cambie a `shipments` y lo recuerda por BOL. Devuelve (hubo cambios, hay fechas no validas).
    """
    column = t("ct_col_delivery")
    shown = pd.DataFrame(frame).copy()
    shown.insert(position, column, [row.delivery or "" for row in shipments])
    # La llave del editor NO depende de lo que la persona cambia (cambiarla a media edicion perderia el siguiente
    # cambio); si cambia cualquier otra cosa que mueva las filas o las fechas base, o se recalcula todo, es otro editor.
    basis = "|".join([
        ready, str(st.session_state.get(f"{prefix}_delivery_gen", 0)), *(f"{r.bol}={base.get(r.bol, '')}" for r in shipments),
    ])
    key = f"{prefix}_delivery_editor_{hashlib.sha1(basis.encode()).hexdigest()[:10]}"
    st.session_state[f"{prefix}_delivery_editor_key"] = key
    edited = st.data_editor(
        shown, key=key, hide_index=True, width="stretch", num_rows="fixed",
        disabled=[c for c in shown.columns if c != column],
        column_config={column: st.column_config.TextColumn(column, help=t("ct_col_delivery_help"), max_chars=10)},
    )
    over: dict[str, str] = st.session_state.setdefault(f"{prefix}_delivery_over", {})
    changed, invalid = False, []
    for row, value in zip(shipments, edited[column]):
        text = "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()
        fixed = de.format_delivery(text, ready)
        if text and not fixed:
            invalid.append(row.bol)
        final = fixed if fixed or not text else text
        if final != (row.delivery or ""):
            changed = True
        row.delivery = final
        if final != base.get(row.bol, ""):
            over[row.bol] = final
        else:
            over.pop(row.bol, None)
    manual = [row.bol for row in shipments if row.bol in over]
    if manual:
        st.caption(t("ct_delivery_changed", count=len(manual), items=", ".join(manual[:20])))
    st.caption(t("ct_delivery_hint", col=column))
    if invalid:
        st.warning(t("ct_delivery_invalid", items=", ".join(invalid[:20])))
    return changed, bool(invalid)


def _on_download(job: dict) -> None:
    _learn(list(job["contacts"].values()))


# ---------------------------------------------------------------------- render


def render(t: Callable[..., str]) -> None:
    st.subheader(t("ct_heading"))
    st.write(t("ct_intro"))

    if shutil.which("tesseract") is None:
        st.error(t("ct_no_tesseract"))
        return

    top = st.container()  # READY DATE, POD y mensajes arriba de todo
    pdf_files = st.file_uploader(t("ct_upload_pdfs"), type=["pdf"], accept_multiple_files=True, key="ct_pdfs")
    dims, data_hash = dims_compare.uploader(t, key="ct_data")  # opcional: solo compara los BOL

    ready, ready_invalid, _ = ready_block(t, top, prefix="ct", ready_key="ct_ready")

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
        contacts = cr.build_contacts(pdfs, memory, barcodes)
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
    if job.get("deduped"):
        st.info(t("ct_deduped", items=", ".join(dict.fromkeys(job["deduped"]))))
    if job["skipped"]:
        st.caption(t("ct_skipped", items="; ".join(f"{n}: {', '.join(map(str, p))}" for n, p in job["skipped"].items())))
    compared = dims_compare.render_comparison(
        t, prefix="ct", pdf_bols=[c.bol for c in job["contacts"].values()], dims=dims,
    )
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
    include_attention = st.checkbox(
        t("ct_include_attention"), value=True, help=t("ct_include_attention_help"), key="ct_include_attention",
    )
    only_verified = st.checkbox(
        t("ct_only_verified"), value=True, help=t("ct_only_verified_help"), key="ct_only_verified",
        disabled=not include_attention,
    )
    include_service = st.checkbox(
        t("ct_include_service"), value=True, help=t("ct_include_service_help"), key="ct_include_service",
    )
    include_delivery = st.checkbox(
        t("ct_include_delivery"), value=False, help=t("ct_include_delivery_help"), key="ct_include_delivery",
    )
    pending_doubt = sum(1 for c in groups["doubt"])
    pending_fail = sum(1 for c in groups["fail"])
    if pending_doubt or pending_fail:
        st.warning(t("ct_pending_warn", doubt=pending_doubt, fail=pending_fail))

    if ready_invalid:
        st.info(t("ct_need_ready"))  # un READY mal escrito no se exporta en blanco por error
        return
    shipments, notes = _shipments(job, only_verified)
    if notes["no_state"] and include_service:  # sin service level no hace falta el estado
        st.warning(t("ct_no_service_warn", items=", ".join(notes["no_state"][:20])))
    if notes["no_carrier"]:
        st.warning(t("ct_no_carrier_warn", items=", ".join(notes["no_carrier"][:20])))

    delivery_base, delivery_error = ({}, False)
    if include_delivery:
        delivery_base, delivery_error = delivery_inputs(t, prefix="ct", ready=ready, shipments=shipments)

    def build() -> tuple[bytes, list[de.DateResult]]:
        return de.build_workbook(
            shipments, ready, include_attention=include_attention, include_service=include_service,
            include_delivery=include_delivery,
        )

    date_bytes, results = build()
    preview = pd.DataFrame([{
        t("ct_col_bol"): r.values["B"],
        **({t("ct_col_name"): r.values["AB"], t("ct_col_phone"): r.values["AC"], t("ct_col_email"): r.values["DU"]}
           if include_attention else {}),
        **({t("ct_col_service"): r.service} if include_service else {}),
        **ready_columns(t, ready), "Consignee No.": r.values["T"],
    } for r in results])
    st.caption(t("ct_row_count", count=len(results)))
    if include_delivery:
        changed, bad = delivery_table(  # la fecha va junto al BOL para que se vea en la pantalla del telefono
            t, prefix="ct", ready=ready, shipments=shipments, base=delivery_base, frame=preview, position=1,
        )
        delivery_error = delivery_error or bad
        if changed:
            date_bytes, results = build()
    else:
        st.dataframe(preview, width="stretch", hide_index=True)

    name_input = st.text_input(t("ct_filename"), key="ct_filename", help=t("ct_filename_help"), placeholder=ready or "DATE_IMPORT")
    file_name = de.safe_filename(name_input, ready)
    st.caption(t("ct_download_name", name=file_name))
    if not compared:
        st.info(t("dc_pending_download"))
    st.download_button(
        t("ct_download_date"), data=date_bytes, file_name=file_name, mime=XLSX_MIME,
        width="stretch", key="ct_dl_date", type="primary", on_click=_on_download, args=(job,),
        disabled=not compared or delivery_error,
    )
