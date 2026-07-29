from pathlib import Path

import streamlit as st

from bol_generator import generate_bol_pdf


st.set_page_config(page_title="Generador de BOL", page_icon="📄", layout="centered")

st.title("Generador de Bill of Lading")
st.write("Sube el PDF de Commercial Invoices y genera un BOL por cada Air Waybill diferente.")

uploaded = st.file_uploader("Commercial Invoice en PDF", type=["pdf"])

if uploaded is not None and st.button("Generar BOL", type="primary", use_container_width=True):
    try:
        with st.spinner("Leyendo Commercial Invoices y generando BOLs..."):
            template_path = Path(__file__).parent / "assets" / "BOL_TEMPLATE.pdf"
            output_pdf, records = generate_bol_pdf(uploaded.getvalue(), template_path)

        st.success(f"Se generaron {len(records)} BOL sin duplicar Air Waybills.")
        st.download_button(
            "Descargar BOLs",
            data=output_pdf,
            file_name="BOLS_GENERADOS.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
    except Exception as error:
        st.error(str(error))
