"""Fecha y hora de entrega a mano, compartido por las pestanas "Leer contactos de BOL" e Inbonds.

Por defecto la entrega se calcula (service level + READY DATE, hora 1600). Aqui se puede:

* elegir una fecha y/o una hora para *todos* los BOL con una sola seleccion;
* cambiar la fecha u hora de cualquier BOL en la tabla (lo que se escribe en la fila manda sobre lo de todos);
* borrar una celda para dejar ese campo vacio.

Orden de prioridad: fila > para todos > calculado.
"""
from __future__ import annotations

import dataclasses
import hashlib
from typing import Callable, Sequence

import pandas as pd
import streamlit as st

import date_export as de

TEXT: dict[str, dict[str, str]] = {
    "es": {
        "dl_heading": "Fecha y hora de entrega",
        "dl_intro": (
            "Opcional. Si no eliges nada, la entrega se calcula con el service level y el READY DATE "
            "(hora 1600). Elige una fecha u hora para aplicarla a todos los BOL y, si quieres, cambia "
            "cualquiera en la tabla."
        ),
        "dl_date": "Fecha de entrega para todos",
        "dl_date_help": "Déjala vacía para usar la calculada. Formato de salida: MM/DD/YY.",
        "dl_time": "Hora de entrega para todos",
        "dl_time_help": "Déjala vacía para usar 1600. Formato de salida: HHMM (24 horas).",
        "dl_no_ready": (
            "Sin READY DATE: Pickup Date, Pickup Time, Pickup Close Time y la entrega calculada quedan en blanco. "
            "Puedes poner la entrega a mano."
        ),
        "dl_table_help": (
            "Puedes editar Entrega (MM/DD/YY) y Hora (HHMM) de cada BOL. Si borras una celda, ese campo queda vacío."
        ),
        "dl_col_time": "Hora",
        "dl_invalid": "No se entendió y se ignoró (usa MM/DD/YY y HHMM): {items}",
        "dl_changed": "{count} BOL con entrega o hora cambiada a mano.",
    },
    "en": {
        "dl_heading": "Delivery date and time",
        "dl_intro": (
            "Optional. If you pick nothing, delivery is computed from the service level and the READY DATE "
            "(time 1600). Pick a date or time to apply it to every BOL and, if you want, change any of them "
            "in the table."
        ),
        "dl_date": "Delivery date for all",
        "dl_date_help": "Leave empty to use the computed one. Output format: MM/DD/YY.",
        "dl_time": "Delivery time for all",
        "dl_time_help": "Leave empty to use 1600. Output format: HHMM (24-hour).",
        "dl_no_ready": (
            "No READY DATE: Pickup Date, Pickup Time, Pickup Close Time and the computed delivery stay blank. "
            "You can set the delivery by hand."
        ),
        "dl_table_help": (
            "You can edit Delivery (MM/DD/YY) and Time (HHMM) for each BOL. If you clear a cell, that field stays empty."
        ),
        "dl_col_time": "Time",
        "dl_invalid": "Not understood and ignored (use MM/DD/YY and HHMM): {items}",
        "dl_changed": "{count} BOLs with delivery or time changed by hand.",
    },
}


def _text(value: object) -> str:
    """Celda del editor -> texto ('' si esta vacia o es NaN/None)."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def apply_row_edits(
    shipments: Sequence[de.ShipmentRow],
    base: Sequence[de.DateResult],
    edited_dates: Sequence[object],
    edited_times: Sequence[object],
) -> tuple[list[de.ShipmentRow], list[str]]:
    """Convierte las celdas editadas en cambios por BOL.

    Una celda cuenta como cambio solo si dice algo distinto de lo que ya salia (`base`). Si se borra,
    el campo queda vacio (de.BLANK). Lo que no se entiende se ignora y se devuelve en la lista de avisos.
    """
    rows: list[de.ShipmentRow] = []
    invalid: list[str] = []
    for shipment, result, raw_date, raw_time in zip(shipments, base, edited_dates, edited_times):
        changes: dict[str, str] = {}
        for field_name, raw, normalize, current in (
            ("delivery_date", raw_date, de.normalize_delivery_date, result.delivery),
            ("delivery_time", raw_time, de.normalize_time, result.time),
        ):
            text = _text(raw)
            if text == current:
                continue
            if not text:
                changes[field_name] = de.BLANK
                continue
            value = normalize(text)
            if not value:
                invalid.append(f"{shipment.bol}: {text}")
            elif value != current:
                changes[field_name] = value
        rows.append(dataclasses.replace(shipment, **changes) if changes else shipment)
    return rows, invalid


def render(
    t: Callable[..., str],
    shipments: list[de.ShipmentRow],
    ready: str,
    *,
    prefix: str,
    priority: str = "",
    extra_columns: dict[str, list[str]] | None = None,
) -> tuple[list[de.ShipmentRow], str, str]:
    """Pinta la seleccion global y la tabla editable.

    Devuelve (BOL con sus cambios a mano, fecha para todos, hora para todos); se pasan tal cual a
    `date_export.build_workbook(rows, ready, ..., default_date=..., default_time=...)`.
    """
    st.markdown(f"##### {t('dl_heading')}")
    st.caption(t("dl_intro"))
    if not ready:
        st.info(t("dl_no_ready"))
    left, right = st.columns(2)
    picked_date = left.date_input(t("dl_date"), value=None, format="MM/DD/YYYY", key=f"{prefix}_deliv_date", help=t("dl_date_help"))
    picked_time = right.time_input(t("dl_time"), value=None, step=60, key=f"{prefix}_deliv_time", help=t("dl_time_help"))
    default_date = de.normalize_delivery_date(picked_date)
    default_time = de.normalize_time(picked_time)

    base = [de.build_date_row(row, ready, priority, default_date, default_time) for row in shipments]
    if not shipments:
        return shipments, default_date, default_time

    extra = extra_columns or {}
    frame = pd.DataFrame({
        t("ct_col_bol"): [r.bol for r in base],
        **{name: values for name, values in extra.items()},
        t("ct_col_service"): [r.service for r in base],
        t("ct_col_delivery"): [r.delivery for r in base],
        t("dl_col_time"): [r.time for r in base],
    })
    # La tabla se reinicia cuando cambia lo que se calcula (otra fecha/hora para todos, otro READY, otros BOL).
    signature = hashlib.sha256("|".join(f"{r.bol},{r.service},{r.delivery},{r.time}" for r in base).encode()).hexdigest()[:10]
    st.caption(t("dl_table_help"))
    edited = st.data_editor(
        frame,
        key=f"{prefix}_deliv_rows_{signature}",
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        disabled=[c for c in frame.columns if c not in (t("ct_col_delivery"), t("dl_col_time"))],
        column_config={
            t("ct_col_delivery"): st.column_config.TextColumn(t("ct_col_delivery"), max_chars=10, help="MM/DD/YY"),
            t("dl_col_time"): st.column_config.TextColumn(t("dl_col_time"), max_chars=5, help="HHMM"),
        },
    )
    rows, invalid = apply_row_edits(shipments, base, edited[t("ct_col_delivery")].tolist(), edited[t("dl_col_time")].tolist())
    if invalid:
        st.warning(t("dl_invalid", items="; ".join(invalid[:10])))
    changed = sum(1 for before, after in zip(shipments, rows) if before is not after)
    if changed:
        st.caption(t("dl_changed", count=changed))
    return rows, default_date, default_time
