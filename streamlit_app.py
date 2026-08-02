from __future__ import annotations

from dataclasses import replace
from hashlib import sha256
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st

from bol_generator import BolExtractionError, BolRecord, extract_unique_records, generate_bol_pdf
from dimension_converter import (
    DimensionConversionError,
    OutputRow,
    convert_dimension_file,
    create_output_workbook,
)


BASE_DIR = Path(__file__).parent
BOL_TEMPLATE = BASE_DIR / "assets" / "BOL_TEMPLATE.pdf"
DIMENSIONS_TEMPLATE = BASE_DIR / "assets" / "DIMENSIONES_TEMPLATE.xlsx"

st.set_page_config(
    page_title="BOL y dimensiones",
    page_icon="📦",
    layout="wide",
)

# Small visual adjustments for the dark layout and wide editable tables.
st.markdown(
    """
    <style>
    .block-container {max-width: 1600px; padding-top: 1.5rem;}
    div[data-testid="stDataEditor"] {border: 1px solid rgba(255,255,255,.12); border-radius: 10px;}
    div[data-testid="stFileUploader"] {border-radius: 10px;}
    </style>
    """,
    unsafe_allow_html=True,
)


def _numeric_key(value: str) -> tuple[int, int | str, str]:
    text = str(value).strip()
    if text.isdigit():
        return (0, int(text), text)
    return (1, text, text)


def _address_text(lines: tuple[str, ...] | list[str]) -> str:
    return "\n".join(str(line).strip() for line in lines if str(line).strip())


def _address_lines(value: Any) -> tuple[str, ...]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ()
    text = str(value).replace(" | ", "\n")
    return tuple(line.strip() for line in text.splitlines() if line.strip())


def _bol_dataframe(records: list[BolRecord]) -> pd.DataFrame:
    ordered = sorted(records, key=lambda record: _numeric_key(record.air_waybill))
    return pd.DataFrame(
        [
            {
                "Generar": True,
                "BOL": record.air_waybill,
                "ENTRY": record.heart_order,
                "PACKING ID": record.packing_id,
                "DIRECCIÓN ORIGINAL": _address_text(
                    record.original_ship_to_lines or record.ship_to_lines
                ),
                "DIRECCIÓN PARA EL BOL": _address_text(record.ship_to_lines),
            }
            for record in ordered
        ]
    )


def _selected_bol_records(
    edited: pd.DataFrame,
    records: list[BolRecord],
) -> tuple[list[BolRecord], list[str]]:
    by_bol = {record.air_waybill: record for record in records}
    selected: list[BolRecord] = []
    invalid: list[str] = []

    for _, row in edited.iterrows():
        if not bool(row.get("Generar", False)):
            continue
        bol = str(row.get("BOL", "")).strip()
        record = by_bol.get(bol)
        if record is None:
            continue
        destination = _address_lines(row.get("DIRECCIÓN PARA EL BOL"))
        if not destination:
            invalid.append(bol)
            continue
        selected.append(replace(record, ship_to_lines=destination))

    selected.sort(key=lambda record: _numeric_key(record.air_waybill))
    return selected, invalid


def _dimension_dataframe(rows: list[OutputRow]) -> pd.DataFrame:
    ordered_rows = sorted(
        rows,
        key=lambda row: (_numeric_key(row.barcode), _numeric_key(row.trailer)),
    )
    return pd.DataFrame(
        [
            {
                "Generar": True,
                "CARRIER": row.carrier,
                "TRAILER": row.trailer,
                "Barcode": row.barcode,
                "Length": row.length,
                "Width": row.width,
                "Height": row.height,
                "Actual Weight": row.weight_lb,
                "BOX": row.boxes,
                "RACK": row.rack,
                "HAZMAT": row.hazmat,
                "Scanned Time": row.scanned_time,
                "TIME IN": row.time_in,
                "TIME OUT": row.time_out,
            }
            for row in ordered_rows
        ]
    )


def _text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _integer(value: Any, field: str, row_number: int) -> int:
    if value is None or pd.isna(value) or str(value).strip() == "":
        return 0
    try:
        return int(round(float(str(value).replace(",", "."))))
    except (TypeError, ValueError) as error:
        raise ValueError(f"Fila {row_number}: {field} debe ser un número entero.") from error


def _edited_dimension_rows(edited: pd.DataFrame) -> list[OutputRow]:
    rows: list[OutputRow] = []
    for index, row in edited.iterrows():
        if not bool(row.get("Generar", False)):
            continue
        row_number = int(index) + 2 if isinstance(index, int) else len(rows) + 2
        key_values = [row.get("CARRIER"), row.get("TRAILER"), row.get("Barcode")]
        if all(_text(value) == "" for value in key_values):
            continue
        rows.append(
            OutputRow(
                carrier=_text(row.get("CARRIER")),
                trailer=_text(row.get("TRAILER")),
                barcode=_text(row.get("Barcode")),
                length=_integer(row.get("Length"), "Length", row_number),
                width=_integer(row.get("Width"), "Width", row_number),
                height=_integer(row.get("Height"), "Height", row_number),
                weight_lb=_integer(row.get("Actual Weight"), "Actual Weight", row_number),
                boxes=_integer(row.get("BOX"), "BOX", row_number),
                rack=_text(row.get("RACK")),
                hazmat=_text(row.get("HAZMAT")),
                scanned_time=_text(row.get("Scanned Time")),
                time_in=_text(row.get("TIME IN")),
                time_out=_text(row.get("TIME OUT")),
            )
        )
    rows.sort(key=lambda record: (_numeric_key(record.barcode), _numeric_key(record.trailer)))
    return rows


st.title("Generador de BOL y convertidor de dimensiones")

bol_tab, dimensions_tab = st.tabs(["Generar BOL", "Convertir dimensiones"])


with bol_tab:
    st.subheader("Commercial Invoice a Bill of Lading")
    st.write(
        "Acepta Commercial Invoices VRP y BEX, incluso si el PDF es escaneado. "
        "Genera un solo BOL por Air Waybill."
    )

    ci_file = st.file_uploader(
        "Sube la Commercial Invoice",
        type=["pdf"],
        key="ci_pdf",
    )

    if ci_file is not None:
        ci_bytes = ci_file.getvalue()
        ci_hash = sha256(ci_bytes).hexdigest()[:16]
        records_key = f"records_{ci_hash}"
        table_key = f"bol_table_{ci_hash}"
        editor_key = f"bol_editor_{ci_hash}"

        if st.button("Analizar Commercial Invoice", type="primary", use_container_width=True):
            try:
                with st.spinner("Leyendo el PDF y buscando Air Waybills..."):
                    records = extract_unique_records(ci_bytes)
                    records = sorted(records, key=lambda record: _numeric_key(record.air_waybill))
                    st.session_state[records_key] = records
                    st.session_state[table_key] = _bol_dataframe(records)
                    st.session_state.pop(editor_key, None)
                st.success(f"Se reconocieron {len(records)} BOL.")
            except BolExtractionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(f"No se pudo analizar el PDF: {error}")

        records: list[BolRecord] = st.session_state.get(records_key, [])
        if records:
            if table_key not in st.session_state:
                st.session_state[table_key] = _bol_dataframe(records)

            st.markdown("#### BOL reconocidos")
            st.caption(
                "La tabla está ordenada numéricamente por BOL. Puedes editar la dirección final "
                "y elegir cuáles se incluirán en el PDF."
            )

            all_col, none_col, count_col = st.columns([1, 1, 2])
            with all_col:
                if st.button("Marcar todos", use_container_width=True, key=f"all_{ci_hash}"):
                    table = st.session_state[table_key].copy()
                    table["Generar"] = True
                    st.session_state[table_key] = table
                    st.session_state.pop(editor_key, None)
                    st.rerun()
            with none_col:
                if st.button("Desmarcar todos", use_container_width=True, key=f"none_{ci_hash}"):
                    table = st.session_state[table_key].copy()
                    table["Generar"] = False
                    st.session_state[table_key] = table
                    st.session_state.pop(editor_key, None)
                    st.rerun()

            edited_bol_table = st.data_editor(
                st.session_state[table_key],
                key=editor_key,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                height=560,
                row_height=82,
                disabled=["BOL", "ENTRY", "PACKING ID", "DIRECCIÓN ORIGINAL"],
                column_config={
                    "Generar": st.column_config.CheckboxColumn("GENERAR", width="small"),
                    "BOL": st.column_config.TextColumn("BOL", width="small"),
                    "ENTRY": st.column_config.TextColumn("ENTRY", width="small"),
                    "PACKING ID": st.column_config.TextColumn("PACKING ID", width="medium"),
                    "DIRECCIÓN ORIGINAL": st.column_config.TextColumn(
                        "DIRECCIÓN ORIGINAL", width="large"
                    ),
                    "DIRECCIÓN PARA EL BOL": st.column_config.TextColumn(
                        "DIRECCIÓN PARA EL BOL", width="large"
                    ),
                },
            )
            st.session_state[table_key] = edited_bol_table.copy()

            selected_records, invalid_addresses = _selected_bol_records(edited_bol_table, records)
            count_col.info(f"Seleccionados: {len(selected_records)} de {len(records)}")

            if invalid_addresses:
                st.warning(
                    "Estos BOL están marcados pero no tienen dirección final: "
                    + ", ".join(invalid_addresses)
                )

            if st.button(
                "Generar BOL seleccionados",
                type="primary",
                use_container_width=True,
                disabled=not selected_records or bool(invalid_addresses),
            ):
                try:
                    with st.spinner("Generando el PDF..."):
                        output_pdf = generate_bol_pdf(selected_records, BOL_TEMPLATE)
                    st.success(f"PDF generado con {len(selected_records)} BOL.")
                    st.download_button(
                        "Descargar BOLS_GENERADOS.pdf",
                        data=output_pdf,
                        file_name="BOLS_GENERADOS.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception as error:
                    st.error(f"No se pudo generar el PDF: {error}")


with dimensions_tab:
    st.subheader("Archivo VRP/BEX a plantilla de dimensiones")
    st.write(
        "Agrupa cada tracking, suma cajas, pallets y pesos, convierte kg a lb enteras "
        "y crea una fila por pallet con dimensiones 48 x 40 x 40."
    )

    dimension_file = st.file_uploader(
        "Sube el archivo de reparto o dimensiones",
        type=["xlsx", "xls"],
        key="dimension_file",
    )

    if dimension_file is not None:
        dimension_bytes = dimension_file.getvalue()
        dimension_hash = sha256(dimension_bytes).hexdigest()[:16]
        groups_key = f"dimension_groups_{dimension_hash}"
        table_key = f"dimension_table_{dimension_hash}"
        original_key = f"dimension_original_{dimension_hash}"
        editor_key = f"dimension_editor_{dimension_hash}"

        if st.button("Analizar y preparar tabla", type="primary", use_container_width=True):
            try:
                with st.spinner("Leyendo trackings y preparando la tabla editable..."):
                    _, groups, output_rows = convert_dimension_file(
                        dimension_bytes,
                        dimension_file.name,
                        DIMENSIONS_TEMPLATE,
                    )
                    table = _dimension_dataframe(output_rows)
                    st.session_state[groups_key] = groups
                    st.session_state[table_key] = table
                    st.session_state[original_key] = table.copy(deep=True)
                    st.session_state.pop(editor_key, None)
                source_types = ", ".join(sorted({group.source_type for group in groups}))
                st.success(
                    f"Formato detectado: {source_types}. "
                    f"Trackings: {len(groups)}. Pallets de salida: {len(output_rows)}."
                )
            except DimensionConversionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(f"No se pudo convertir el archivo: {error}")

        if table_key in st.session_state:
            st.markdown("#### Tabla final editable")
            st.caption(
                "Todas las filas aparecen marcadas al generarse. Puedes editar cualquier celda, "
                "desmarcar las filas que no quieras, borrar filas o agregar nuevas. "
                "El Excel incluirá únicamente las filas marcadas y las ordenará por Waybill."
            )

            all_col, none_col, reset_col, summary_col = st.columns([1, 1, 1, 3])
            with all_col:
                if st.button(
                    "Marcar todos",
                    use_container_width=True,
                    key=f"dimension_all_{dimension_hash}",
                ):
                    table = st.session_state[table_key].copy(deep=True)
                    table["Generar"] = True
                    st.session_state[table_key] = table
                    st.session_state.pop(editor_key, None)
                    st.rerun()
            with none_col:
                if st.button(
                    "Desmarcar todos",
                    use_container_width=True,
                    key=f"dimension_none_{dimension_hash}",
                ):
                    table = st.session_state[table_key].copy(deep=True)
                    table["Generar"] = False
                    st.session_state[table_key] = table
                    st.session_state.pop(editor_key, None)
                    st.rerun()
            with reset_col:
                if st.button(
                    "Restaurar tabla",
                    use_container_width=True,
                    key=f"dimension_reset_{dimension_hash}",
                ):
                    st.session_state[table_key] = st.session_state[original_key].copy(deep=True)
                    st.session_state.pop(editor_key, None)
                    st.rerun()

            edited_dimension_table = st.data_editor(
                st.session_state[table_key],
                key=editor_key,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                height=600,
                row_height=36,
                column_config={
                    "Generar": st.column_config.CheckboxColumn(
                        "GENERAR",
                        width="small",
                        default=True,
                        help="Desmarca una fila para excluirla del Excel final.",
                    ),
                    "CARRIER": st.column_config.TextColumn("CARRIER", width="small"),
                    "TRAILER": st.column_config.TextColumn("TRAILER", width="small"),
                    "Barcode": st.column_config.TextColumn("Barcode", width="small"),
                    "Length": st.column_config.NumberColumn("Length", step=1, format="%d"),
                    "Width": st.column_config.NumberColumn("Width", step=1, format="%d"),
                    "Height": st.column_config.NumberColumn("Height", step=1, format="%d"),
                    "Actual Weight": st.column_config.NumberColumn(
                        "Actual Weight", step=1, format="%d"
                    ),
                    "BOX": st.column_config.NumberColumn("BOX", step=1, format="%d"),
                    "RACK": st.column_config.TextColumn("RACK", width="small"),
                    "HAZMAT": st.column_config.TextColumn("HAZMAT", width="small"),
                    "Scanned Time": st.column_config.TextColumn(
                        "Scanned Time", width="medium"
                    ),
                    "TIME IN": st.column_config.TextColumn("TIME IN", width="small"),
                    "TIME OUT": st.column_config.TextColumn("TIME OUT", width="small"),
                },
            )
            st.session_state[table_key] = edited_dimension_table.copy()

            try:
                edited_rows = _edited_dimension_rows(edited_dimension_table)
                summary_col.info(f"Filas seleccionadas para exportar: {len(edited_rows)}")
                export_key = f"dimension_export_{dimension_hash}"
                table_fingerprint = sha256(
                    edited_dimension_table.fillna("").to_csv(index=False).encode("utf-8")
                ).hexdigest()

                if not edited_rows:
                    st.warning("La tabla no tiene filas válidas para exportar.")
                else:
                    if st.button(
                        "Preparar Excel con los cambios",
                        type="primary",
                        use_container_width=True,
                    ):
                        st.session_state[export_key] = (
                            table_fingerprint,
                            create_output_workbook(edited_rows, DIMENSIONS_TEMPLATE),
                        )
                        st.success("Excel preparado con los valores actuales de la tabla.")

                    prepared = st.session_state.get(export_key)
                    if prepared and prepared[0] == table_fingerprint:
                        st.download_button(
                            "Descargar DIMENSIONES_CONVERTIDAS.xlsx",
                            data=prepared[1],
                            file_name="DIMENSIONES_CONVERTIDAS.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            on_click="ignore",
                        )
                    elif prepared:
                        st.info(
                            "La tabla cambió después de preparar el archivo. "
                            "Presiona nuevamente “Preparar Excel con los cambios”."
                        )
            except Exception as error:
                st.error(f"Revisa los valores editados: {error}")
