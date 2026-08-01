# Generador de BOL y convertidor de dimensiones

Aplicación Streamlit para:

1. Leer Commercial Invoices VRP y BEX y generar un BOL por Air Waybill.
2. Revisar los BOL detectados en una tabla ordenada numéricamente.
3. Seleccionar cuáles BOL generar y editar la dirección que aparecerá en el PDF.
4. Leer archivos de dimensiones VRP/BEX, distribuir cajas y pesos por pallet y exportar la plantilla final.
5. Editar cualquier celda de la tabla de dimensiones, agregar filas o eliminar filas antes de exportar.

## Archivos principales

- `streamlit_app.py`: interfaz de la aplicación.
- `bol_generator.py`: lectura de CI y generación de BOL.
- `dimension_converter.py`: lectura de archivos VRP/BEX y creación del Excel.
- `assets/BOL_TEMPLATE.pdf`: plantilla de BOL.
- `assets/DIMENSIONES_TEMPLATE.xlsx`: plantilla de dimensiones.
- `.streamlit/config.toml`: tema oscuro y configuración de carga.

## Despliegue

En Streamlit Community Cloud usa:

- Branch: `main`
- Main file path: `streamlit_app.py`

Al subir cambios al repositorio, reinicia la aplicación si Streamlit no los muestra automáticamente.
