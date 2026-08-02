# Generador de BOL y convertidor de dimensiones

Aplicación Streamlit para:

1. Leer Commercial Invoices VRP y BEX y generar un BOL por Air Waybill.
2. Revisar los BOL detectados en una tabla ordenada numéricamente.
3. Seleccionar cuáles BOL generar y editar la dirección que aparecerá en el PDF.
4. Leer archivos de dimensiones VRP/BEX, distribuir cajas y pesos por pallet y exportar la plantilla final.
5. Editar cualquier celda de la tabla de dimensiones, seleccionar o desmarcar filas y exportar solo las elegidas.

## Archivos principales

- `streamlit_app.py`: interfaz de la aplicación.
- `bol_generator.py`: lectura de CI y generación de BOL.
- `dimension_converter.py`: lectura de archivos VRP/BEX y creación del Excel.
- `assets/BOL_TEMPLATE.pdf`: plantilla de BOL.
- `assets/DIMENSIONES_TEMPLATE.xlsx`: plantilla de dimensiones.
- `.streamlit/config.toml`: configuración base y límite de carga.

## Despliegue

En Streamlit Community Cloud usa:

- Branch: `main`
- Main file path: `streamlit_app.py`

Al subir cambios al repositorio, reinicia la aplicación si Streamlit no los muestra automáticamente.

## Cambios V6

- Botón global para cambiar toda la interfaz entre español e inglés.
- Botón global para cambiar toda la página entre modo oscuro y modo claro.
- Campo opcional para personalizar el nombre del PDF de BOL.
- Campo opcional para personalizar el nombre del Excel de dimensiones.
- Si el nombre queda vacío, se conservan `BOLS_GENERADOS.pdf` y `DIMENSIONES_CONVERTIDAS.xlsx`.
- Se conservan la selección de filas, la edición de tablas y el orden numérico de Waybills.
