from __future__ import annotations

from hashlib import sha256
from pathlib import Path

import streamlit as st

from bol_generator import BolExtractionError, BolRecord, extract_unique_records, generate_bol_pdf
from dimension_converter import DimensionConversionError, convert_dimension_file


BASE_DIR = Path(__file__).parent
BOL_TEMPLATE = BASE_DIR / "assets" / "BOL_TEMPLATE.pdf"
DIMENSIONS_TEMPLATE = BASE_DIR / "assets" / "DIMENSIONES_TEMPLATE.xlsx"

st.set_page_config(
    page_title="BOL y dimensiones",
    page_icon="📦",
    layout="wide",
)

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

        if st.button("Analizar Commercial Invoice", type="primary", use_container_width=True):
            try:
                with st.spinner("Leyendo el PDF y buscando Air Waybills..."):
                    st.session_state[records_key] = extract_unique_records(ci_bytes)
                for record in st.session_state[records_key]:
                    st.session_state[f"select_{ci_hash}_{record.air_waybill}"] = True
                st.success(f"Se reconocieron {len(st.session_state[records_key])} BOL.")
            except BolExtractionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(f"No se pudo analizar el PDF: {error}")

        records: list[BolRecord] = st.session_state.get(records_key, [])
        if records:
            st.markdown("#### BOL reconocidos")
            select_col, clear_col, count_col = st.columns([1, 1, 2])
            with select_col:
                if st.button("Marcar todos", use_container_width=True):
                    for record in records:
                        st.session_state[f"select_{ci_hash}_{record.air_waybill}"] = True
                    st.rerun()
            with clear_col:
                if st.button("Desmarcar todos", use_container_width=True):
                    for record in records:
                        st.session_state[f"select_{ci_hash}_{record.air_waybill}"] = False
                    st.rerun()

            selected_records: list[BolRecord] = []
            with st.container(height=520, border=True):
                for record in records:
                    check_col, info_col = st.columns([0.6, 8])
                    selection_key = f"select_{ci_hash}_{record.air_waybill}"
                    with check_col:
                        selected = st.checkbox(
                            "Seleccionar",
                            key=selection_key,
                            label_visibility="collapsed",
                        )
                    with info_col:
                        st.markdown(
                            f"**{record.air_waybill}** · {record.source_format} · "
                            f"Packing ID: `{record.packing_id}` · Heart Order: `{record.heart_order}`"
                        )
                        st.caption(
                            f"Ship To del BOL: {record.display_ship_to} · Regla: {record.destination_rule}"
                        )
                    if selected:
                        selected_records.append(record)

            count_col.info(f"Seleccionados: {len(selected_records)} de {len(records)}")

            if st.button(
                "Generar BOL seleccionados",
                type="primary",
                use_container_width=True,
                disabled=not selected_records,
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
        if st.button("Convertir dimensiones", type="primary", use_container_width=True):
            try:
                with st.spinner("Leyendo trackings y generando la plantilla..."):
                    workbook_bytes, groups, output_rows = convert_dimension_file(
                        dimension_file.getvalue(),
                        dimension_file.name,
                        DIMENSIONS_TEMPLATE,
                    )

                source_types = ", ".join(sorted({group.source_type for group in groups}))
                st.success(
                    f"Formato detectado: {source_types}. "
                    f"Trackings: {len(groups)}. Pallets de salida: {len(output_rows)}."
                )

                preview = [
                    {
                        "CARRIER": group.source_type,
                        "TRAILER": group.trailer,
                        "Barcode": group.barcode,
                        "Pallets": group.pallets,
                        "Boxes": group.boxes,
                        "Peso kg": float(group.weight_kg),
                        "Peso lb": group.weight_lb,
                    }
                    for group in groups
                ]
                st.dataframe(preview, use_container_width=True, hide_index=True)

                st.download_button(
                    "Descargar DIMENSIONES_CONVERTIDAS.xlsx",
                    data=workbook_bytes,
                    file_name="DIMENSIONES_CONVERTIDAS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except DimensionConversionError as error:
                st.error(str(error))
            except Exception as error:
                st.error(f"No se pudo convertir el archivo: {error}")
