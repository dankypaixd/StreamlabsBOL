from __future__ import annotations

from dataclasses import replace
from hashlib import sha256
from pathlib import Path
import re
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
    page_title="BOL / Dimensions",
    page_icon="📦",
    layout="wide",
)


TRANSLATIONS: dict[str, dict[str, str]] = {
    "es": {
        "app_title": "Generador de BOL y convertidor de dimensiones",
        "switch_language": "🌐 English",
        "switch_dark": "🌙 Modo oscuro",
        "switch_light": "☀️ Modo claro",
        "tab_bol": "Generar BOL",
        "tab_dimensions": "Convertir dimensiones",
        "bol_heading": "Commercial Invoice a Bill of Lading",
        "bol_intro": (
            "Acepta Commercial Invoices VRP y BEX, incluso si el PDF es escaneado. "
            "Genera un solo BOL por Air Waybill."
        ),
        "upload_ci": "Sube la Commercial Invoice",
        "analyze_ci": "Analizar Commercial Invoice",
        "reading_ci": "Leyendo el PDF y buscando Air Waybills...",
        "recognized_bols": "Se reconocieron {count} BOL.",
        "ci_error": "No se pudo analizar el PDF: {error}",
        "bol_table_heading": "BOL reconocidos",
        "bol_table_caption": (
            "La tabla está ordenada numéricamente por BOL. Puedes editar la dirección final "
            "y elegir cuáles se incluirán en el PDF."
        ),
        "select_all": "Marcar todos",
        "select_none": "Desmarcar todos",
        "selected_count": "Seleccionados: {selected} de {total}",
        "missing_address": "Estos BOL están marcados pero no tienen dirección final: {items}",
        "filename_optional": "Nombre del archivo (opcional)",
        "filename_help_pdf": (
            "Escribe el nombre del PDF. Si lo dejas vacío se usará BOLS_GENERADOS.pdf."
        ),
        "generate_selected_bols": "Generar BOL seleccionados",
        "generating_pdf": "Generando el PDF...",
        "pdf_generated": "PDF generado con {count} BOL.",
        "download_file": "Descargar {filename}",
        "pdf_error": "No se pudo generar el PDF: {error}",
        "dimensions_heading": "Archivo VRP/BEX a plantilla de dimensiones",
        "dimensions_intro": (
            "Agrupa cada tracking, suma cajas, pallets y pesos, convierte kg a lb enteras "
            "y crea una fila por pallet con dimensiones 48 x 40 x 40."
        ),
        "upload_dimensions": "Sube el archivo de reparto o dimensiones",
        "analyze_dimensions": "Analizar y preparar tabla",
        "reading_dimensions": "Leyendo trackings y preparando la tabla editable...",
        "format_detected": (
            "Formato detectado: {types}. Trackings: {trackings}. "
            "Pallets de salida: {pallets}."
        ),
        "dimensions_error": "No se pudo convertir el archivo: {error}",
        "editable_table": "Tabla final editable",
        "editable_caption": (
            "Todas las filas aparecen marcadas al generarse. Puedes editar cualquier celda, "
            "desmarcar las filas que no quieras, borrar filas o agregar nuevas. "
            "El Excel incluirá únicamente las filas marcadas y las ordenará por Waybill."
        ),
        "reset_table": "Restaurar tabla",
        "selected_rows": "Filas seleccionadas para exportar: {count}",
        "no_rows": "La tabla no tiene filas válidas para exportar.",
        "filename_help_xlsx": (
            "Escribe el nombre del Excel. Si lo dejas vacío se usará "
            "DIMENSIONES_CONVERTIDAS.xlsx."
        ),
        "prepare_excel": "Preparar Excel con los cambios",
        "excel_prepared": "Excel preparado con los valores actuales de la tabla.",
        "table_changed": (
            "La tabla cambió después de preparar el archivo. "
            "Presiona nuevamente “Preparar Excel con los cambios”."
        ),
        "edited_values_error": "Revisa los valores editados: {error}",
        "row_integer_error": "Fila {row}: {field} debe ser un número entero.",
        "generate_column": "GENERAR",
        "original_address": "DIRECCIÓN ORIGINAL",
        "bol_address": "DIRECCIÓN PARA EL BOL",
        "exclude_row_help": "Desmarca una fila para excluirla del Excel final.",
    },
    "en": {
        "app_title": "BOL generator and dimensions converter",
        "switch_language": "🌐 Español",
        "switch_dark": "🌙 Dark mode",
        "switch_light": "☀️ Light mode",
        "tab_bol": "Generate BOL",
        "tab_dimensions": "Convert dimensions",
        "bol_heading": "Commercial Invoice to Bill of Lading",
        "bol_intro": (
            "Accepts VRP and BEX Commercial Invoices, including scanned PDFs. "
            "It creates one BOL per Air Waybill."
        ),
        "upload_ci": "Upload the Commercial Invoice",
        "analyze_ci": "Analyze Commercial Invoice",
        "reading_ci": "Reading the PDF and finding Air Waybills...",
        "recognized_bols": "{count} BOL records were recognized.",
        "ci_error": "The PDF could not be analyzed: {error}",
        "bol_table_heading": "Recognized BOL records",
        "bol_table_caption": (
            "The table is sorted numerically by BOL. You can edit the final address "
            "and choose which records will be included in the PDF."
        ),
        "select_all": "Select all",
        "select_none": "Clear all",
        "selected_count": "Selected: {selected} of {total}",
        "missing_address": "These selected BOL records do not have a final address: {items}",
        "filename_optional": "File name (optional)",
        "filename_help_pdf": (
            "Enter the PDF name. Leave it blank to use BOLS_GENERADOS.pdf."
        ),
        "generate_selected_bols": "Generate selected BOL records",
        "generating_pdf": "Generating the PDF...",
        "pdf_generated": "PDF generated with {count} BOL records.",
        "download_file": "Download {filename}",
        "pdf_error": "The PDF could not be generated: {error}",
        "dimensions_heading": "VRP/BEX file to dimensions template",
        "dimensions_intro": (
            "Groups each tracking number, totals boxes, pallets and weights, converts kg to "
            "whole pounds, and creates one row per pallet with 48 x 40 x 40 dimensions."
        ),
        "upload_dimensions": "Upload the allocation or dimensions file",
        "analyze_dimensions": "Analyze and prepare table",
        "reading_dimensions": "Reading tracking numbers and preparing the editable table...",
        "format_detected": (
            "Detected format: {types}. Tracking numbers: {trackings}. "
            "Output pallets: {pallets}."
        ),
        "dimensions_error": "The file could not be converted: {error}",
        "editable_table": "Editable final table",
        "editable_caption": (
            "Every row is selected when the table is created. You can edit any cell, clear "
            "rows you do not want, delete rows, or add new ones. The Excel file will include "
            "only selected rows and sort them by Waybill."
        ),
        "reset_table": "Reset table",
        "selected_rows": "Rows selected for export: {count}",
        "no_rows": "The table has no valid rows to export.",
        "filename_help_xlsx": (
            "Enter the Excel file name. Leave it blank to use "
            "DIMENSIONES_CONVERTIDAS.xlsx."
        ),
        "prepare_excel": "Prepare Excel with changes",
        "excel_prepared": "Excel prepared with the table's current values.",
        "table_changed": (
            "The table changed after the file was prepared. "
            "Press “Prepare Excel with changes” again."
        ),
        "edited_values_error": "Check the edited values: {error}",
        "row_integer_error": "Row {row}: {field} must be a whole number.",
        "generate_column": "GENERATE",
        "original_address": "ORIGINAL ADDRESS",
        "bol_address": "BOL ADDRESS",
        "exclude_row_help": "Clear a row to exclude it from the final Excel file.",
    },
}


if "language" not in st.session_state:
    st.session_state.language = "es"
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "dark"


def t(key: str, **values: Any) -> str:
    template = TRANSLATIONS[st.session_state.language][key]
    return template.format(**values) if values else template


def _apply_theme(mode: str) -> None:
    if mode == "dark":
        palette = {
            "bg": "#0E1117",
            "surface": "#171C24",
            "surface2": "#202733",
            "text": "#FAFAFA",
            "muted": "#B8C0CC",
            "border": "rgba(255,255,255,.16)",
            "input": "#141A22",
            "shadow": "rgba(0,0,0,.28)",
        }
    else:
        palette = {
            "bg": "#F7F9FC",
            "surface": "#FFFFFF",
            "surface2": "#EEF2F7",
            "text": "#17202A",
            "muted": "#5B6470",
            "border": "rgba(23,32,42,.18)",
            "input": "#FFFFFF",
            "shadow": "rgba(23,32,42,.10)",
        }

    st.markdown(
        f"""
        <style>
        :root {{
            --background-color: {palette['bg']};
            --secondary-background-color: {palette['surface']};
            --text-color: {palette['text']};
        }}
        html, body, .stApp, [data-testid="stAppViewContainer"] {{
            background: {palette['bg']} !important;
            color: {palette['text']} !important;
        }}
        [data-testid="stHeader"], [data-testid="stToolbar"] {{
            background: {palette['bg']} !important;
        }}
        .block-container {{
            max-width: 1600px;
            padding-top: 1.1rem;
            padding-bottom: 2rem;
        }}
        h1, h2, h3, h4, h5, h6, p, label, span, div {{
            color: {palette['text']};
        }}
        .stCaption, [data-testid="stCaptionContainer"] p {{
            color: {palette['muted']} !important;
        }}
        div[data-testid="stDataEditor"], div[data-testid="stDataFrame"],
        div[data-testid="stFileUploader"], div[data-testid="stAlert"] {{
            border: 1px solid {palette['border']} !important;
            border-radius: 10px !important;
            box-shadow: 0 2px 12px {palette['shadow']};
        }}
        div[data-baseweb="input"] > div,
        div[data-baseweb="textarea"] > div,
        textarea, input {{
            background: {palette['input']} !important;
            color: {palette['text']} !important;
            border-color: {palette['border']} !important;
        }}
        div[data-baseweb="tab-list"] {{
            background: {palette['surface']} !important;
            border-radius: 10px;
            padding: .2rem .4rem;
        }}
        button[kind="secondary"], button[kind="tertiary"] {{
            background: {palette['surface']} !important;
            color: {palette['text']} !important;
            border-color: {palette['border']} !important;
        }}
        [data-testid="stDownloadButton"] button {{
            border-radius: 8px;
        }}
        hr {{ border-color: {palette['border']} !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def _toggle_language() -> None:
    st.session_state.language = "en" if st.session_state.language == "es" else "es"


def _toggle_theme() -> None:
    st.session_state.theme_mode = "light" if st.session_state.theme_mode == "dark" else "dark"


def _numeric_key(value: str) -> tuple[int, int | str, str]:
    text = str(value).strip()
    if text.isdigit():
        return (0, int(text), text)
    return (1, text, text)


def _safe_filename(value: str, default_stem: str, extension: str) -> str:
    text = str(value or "").strip()
    if not text:
        return f"{default_stem}{extension}"
    text = Path(text).name
    text = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", text).strip(" .")
    if not text:
        return f"{default_stem}{extension}"
    if text.lower().endswith(extension.lower()):
        return text
    return f"{text}{extension}"


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
        raise ValueError(t("row_integer_error", row=row_number, field=field)) from error


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


# Global controls: both affect the entire application.
control_spacer, language_col, theme_col = st.columns([6, 1, 1])
with language_col:
    st.button(
        t("switch_language"),
        key="language_toggle",
        use_container_width=True,
        on_click=_toggle_language,
    )
with theme_col:
    theme_label = t("switch_light") if st.session_state.theme_mode == "dark" else t("switch_dark")
    st.button(
        theme_label,
        key="theme_toggle",
        use_container_width=True,
        on_click=_toggle_theme,
    )

_apply_theme(st.session_state.theme_mode)

st.title(t("app_title"))

bol_tab, dimensions_tab = st.tabs([t("tab_bol"), t("tab_dimensions")])


with bol_tab:
    st.subheader(t("bol_heading"))
    st.write(t("bol_intro"))

    ci_file = st.file_uploader(
        t("upload_ci"),
        type=["pdf"],
        key="ci_pdf",
    )

    if ci_file is not None:
        ci_bytes = ci_file.getvalue()
        ci_hash = sha256(ci_bytes).hexdigest()[:16]
        records_key = f"records_{ci_hash}"
        table_key = f"bol_table_{ci_hash}"
        editor_key = f"bol_editor_{ci_hash}"

        if st.button(
            t("analyze_ci"),
            type="primary",
            use_container_width=True,
            key=f"analyze_ci_{ci_hash}",
        ):
            try:
                with st.spinner(t("reading_ci")):
                    records = extract_unique_records(ci_bytes)
                    records = sorted(records, key=lambda record: _numeric_key(record.air_waybill))
                    st.session_state[records_key] = records
                    st.session_state[table_key] = _bol_dataframe(records)
                    st.session_state.pop(editor_key, None)
                st.success(t("recognized_bols", count=len(records)))
            except BolExtractionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(t("ci_error", error=error))

        records: list[BolRecord] = st.session_state.get(records_key, [])
        if records:
            if table_key not in st.session_state:
                st.session_state[table_key] = _bol_dataframe(records)

            st.markdown(f"#### {t('bol_table_heading')}")
            st.caption(t("bol_table_caption"))

            all_col, none_col, count_col = st.columns([1, 1, 2])
            with all_col:
                if st.button(
                    t("select_all"),
                    use_container_width=True,
                    key=f"all_{ci_hash}",
                ):
                    table = st.session_state[table_key].copy()
                    table["Generar"] = True
                    st.session_state[table_key] = table
                    st.session_state.pop(editor_key, None)
                    st.rerun()
            with none_col:
                if st.button(
                    t("select_none"),
                    use_container_width=True,
                    key=f"none_{ci_hash}",
                ):
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
                    "Generar": st.column_config.CheckboxColumn(
                        t("generate_column"), width="small"
                    ),
                    "BOL": st.column_config.TextColumn("BOL", width="small"),
                    "ENTRY": st.column_config.TextColumn("ENTRY", width="small"),
                    "PACKING ID": st.column_config.TextColumn("PACKING ID", width="medium"),
                    "DIRECCIÓN ORIGINAL": st.column_config.TextColumn(
                        t("original_address"), width="large"
                    ),
                    "DIRECCIÓN PARA EL BOL": st.column_config.TextColumn(
                        t("bol_address"), width="large"
                    ),
                },
            )
            st.session_state[table_key] = edited_bol_table.copy()

            selected_records, invalid_addresses = _selected_bol_records(edited_bol_table, records)
            count_col.info(
                t("selected_count", selected=len(selected_records), total=len(records))
            )

            if invalid_addresses:
                st.warning(t("missing_address", items=", ".join(invalid_addresses)))

            bol_filename_input = st.text_input(
                t("filename_optional"),
                key=f"bol_filename_{ci_hash}",
                placeholder="BOLS_GENERADOS",
                help=t("filename_help_pdf"),
            )
            bol_filename = _safe_filename(
                bol_filename_input,
                default_stem="BOLS_GENERADOS",
                extension=".pdf",
            )

            if st.button(
                t("generate_selected_bols"),
                type="primary",
                use_container_width=True,
                disabled=not selected_records or bool(invalid_addresses),
                key=f"generate_bol_{ci_hash}",
            ):
                try:
                    with st.spinner(t("generating_pdf")):
                        output_pdf = generate_bol_pdf(selected_records, BOL_TEMPLATE)
                    st.session_state[f"bol_output_{ci_hash}"] = output_pdf
                    st.success(t("pdf_generated", count=len(selected_records)))
                except Exception as error:
                    st.error(t("pdf_error", error=error))

            output_pdf = st.session_state.get(f"bol_output_{ci_hash}")
            if output_pdf:
                st.download_button(
                    t("download_file", filename=bol_filename),
                    data=output_pdf,
                    file_name=bol_filename,
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"download_bol_{ci_hash}",
                )


with dimensions_tab:
    st.subheader(t("dimensions_heading"))
    st.write(t("dimensions_intro"))

    dimension_file = st.file_uploader(
        t("upload_dimensions"),
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

        if st.button(
            t("analyze_dimensions"),
            type="primary",
            use_container_width=True,
            key=f"analyze_dimensions_{dimension_hash}",
        ):
            try:
                with st.spinner(t("reading_dimensions")):
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
                    t(
                        "format_detected",
                        types=source_types,
                        trackings=len(groups),
                        pallets=len(output_rows),
                    )
                )
            except DimensionConversionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(t("dimensions_error", error=error))

        if table_key in st.session_state:
            st.markdown(f"#### {t('editable_table')}")
            st.caption(t("editable_caption"))

            all_col, none_col, reset_col, summary_col = st.columns([1, 1, 1, 3])
            with all_col:
                if st.button(
                    t("select_all"),
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
                    t("select_none"),
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
                    t("reset_table"),
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
                        t("generate_column"),
                        width="small",
                        default=True,
                        help=t("exclude_row_help"),
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
                summary_col.info(t("selected_rows", count=len(edited_rows)))
                export_key = f"dimension_export_{dimension_hash}"
                table_fingerprint = sha256(
                    edited_dimension_table.fillna("").to_csv(index=False).encode("utf-8")
                ).hexdigest()

                dimensions_filename_input = st.text_input(
                    t("filename_optional"),
                    key=f"dimensions_filename_{dimension_hash}",
                    placeholder="DIMENSIONES_CONVERTIDAS",
                    help=t("filename_help_xlsx"),
                )
                dimensions_filename = _safe_filename(
                    dimensions_filename_input,
                    default_stem="DIMENSIONES_CONVERTIDAS",
                    extension=".xlsx",
                )

                if not edited_rows:
                    st.warning(t("no_rows"))
                else:
                    if st.button(
                        t("prepare_excel"),
                        type="primary",
                        use_container_width=True,
                        key=f"prepare_excel_{dimension_hash}",
                    ):
                        st.session_state[export_key] = (
                            table_fingerprint,
                            create_output_workbook(edited_rows, DIMENSIONS_TEMPLATE),
                        )
                        st.success(t("excel_prepared"))

                    prepared = st.session_state.get(export_key)
                    if prepared and prepared[0] == table_fingerprint:
                        st.download_button(
                            t("download_file", filename=dimensions_filename),
                            data=prepared[1],
                            file_name=dimensions_filename,
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            use_container_width=True,
                            key=f"download_dimensions_{dimension_hash}",
                            on_click="ignore",
                        )
                    elif prepared:
                        st.info(t("table_changed"))
            except Exception as error:
                st.error(t("edited_values_error", error=error))
