# Generador de BOL y convertidor de dimensiones

Aplicación Streamlit para:

1. Leer Commercial Invoices VRP y BEX y generar un BOL por Air Waybill.
2. Revisar los BOL detectados en una tabla ordenada numéricamente.
3. Seleccionar cuáles BOL generar y editar la dirección que aparecerá en el PDF.
4. Leer archivos de dimensiones VRP/BEX, distribuir cajas y pesos por pallet y exportar la plantilla final.
5. Editar cualquier celda de la tabla de dimensiones, seleccionar o desmarcar filas y exportar solo las elegidas.
6. **Leer contactos de BOL (V8):** leer el bloque “Ship To Attention” (nombre, correo, teléfono) y el estado del
   Ship To de cada Bill of Lading, calcular el service level y la fecha de entrega, y exportar el Excel de importación.

## Archivos principales

- `streamlit_app.py`: interfaz de la aplicación.
- `bol_generator.py`: lectura de CI y generación de BOL.
- `dimension_converter.py`: lectura de archivos VRP/BEX y creación del Excel.
- `contacts_tab.py`: pestaña “Leer contactos de BOL” (interfaz, capturas, confirmación, descargas).
- `contact_extractor.py`: OCR del Ship To Attention y del estado (varias lecturas con votación, orientación y enderezado).
- `contact_review.py`: une las páginas de un BOL, aplica lo aprendido y clasifica en verde / amarillo / rojo.
- `contact_memory.py`: memoria de contactos confirmados (cifrada, guardada en una rama de GitHub).
- `service_levels.py`: service level por estado y fecha de entrega con cut-off 0800–1700 (sin fines de semana).
- `date_export.py`: genera las hojas DATE (198 columnas del template) y DATA.
- `assets/SERVICE_LEVELS.csv`: lista estado → días (editable). `assets/DATE_HEADERS.json`: encabezados del template.
- `packages.txt`: paquetes del sistema para Streamlit Cloud (Tesseract OCR).
- `assets/BOL_TEMPLATE.pdf`: plantilla de BOL.
- `assets/DIMENSIONES_TEMPLATE.xlsx`: plantilla de dimensiones.
- `.streamlit/config.toml`: configuración base y límite de carga.

## Despliegue

En Streamlit Community Cloud usa:

- Branch: `main`
- Main file path: `streamlit_app.py`

Al subir cambios al repositorio, reinicia la aplicación si Streamlit no los muestra automáticamente.

## Cambios V7

- Botones globales visibles debajo del título para cambiar toda la interfaz entre español e inglés y alternar modo oscuro/claro.
- Indicador visible de “Versión 7” para confirmar que Streamlit cargó la actualización correcta.
- Campo opcional para personalizar el nombre del PDF de BOL.
- Campo opcional para personalizar el nombre del Excel de dimensiones.
- Si el nombre queda vacío, se conservan `BOLS_GENERADOS.pdf` y `DIMENSIONES_CONVERTIDAS.xlsx`.
- Se conservan la selección de filas, la edición de tablas y el orden numérico de Waybills.


## Cambios V8: pestaña “Leer contactos de BOL”

**Qué hace.** Subes los PDF de los BOL (se reconocen por “Bill of Lading #” arriba a la derecha). Por cada BOL lee
nombre, correo y teléfono del bloque *Ship To Attention* y el estado del *Ship To*. Cada BOL muestra una captura con
lo leído marcado por colores (azul nombre, verde correo, naranja teléfono, morado estado, rojo BOL) y, al lado, los
campos tal como quedaron escritos, que se pueden editar.

**Clasificación.**

- 🟢 *Sin error*: nombre, correo y teléfono presentes y válidos, y todas las lecturas del OCR coinciden entre sí.
- 🟡 *Dudoso*: está todo, pero alguna lectura no coincide, el estado se dedujo del ZIP, el correo se corrigió con
  lo aprendido, las dos páginas del BOL difieren, etc. Cada motivo se explica en la tarjeta.
- 🔴 *Incompleto*: falta algo (bloque vacío en el PDF, sin teléfono…) o el formato no es válido.

Un teléfono no tiene formato fijo: se aceptan dígitos y el signo `+`.

**Service level y fecha de entrega.** Salen del estado del Ship To con `assets/SERVICE_LEVELS.csv` (`3D`, `4D`…).
La fecha se mide desde el READY DATE (`MM/DD/YY HHMM`, celda N2 de DATA): dentro del cut-off 0800–1700 el mismo día
es el día 0; después de las 17:00 el día 0 es el siguiente día hábil; los fines de semana no cuentan. Se escribe `MM/DD/YY`.
Supuestos que no estaban en la especificación (ajustables en `service_levels.py`): antes de las 08:00 el día 0
es el mismo día; un READY DATE en fin de semana arranca el lunes; no se descuentan feriados (`HOLIDAYS`).

**Excel de salida.** Con tu DATA subida (o la salida de “Convertir dimensiones”) el libro trae las hojas `DATE` y `DATA`
(Q y R rellenos por BOL). La hoja `DATE` tiene las 198 columnas del template; se llenan A, B, E, Q, R, S, AB, AC, AF,
AG, AK, CO, DJ, DK y DU con los mismos valores que calculan tus fórmulas. También se puede bajar solo `DATE`.
Por defecto, los dudosos e incompletos sin confirmar salen con nombre, correo y teléfono vacíos (casilla en pantalla).

**Memoria.** Lo confirmado (y los verdes al descargar) se aprende en segundo plano: correo → nombre y teléfonos,
dominios, y código postal → estado. Sirve para corregir confusiones típicas del OCR (l/i/1, o/0, rn/m) contra algo
ya verificado y para dejar en verde un contacto idéntico a uno confirmado. Una corrección aproximada siempre queda
en amarillo. No se muestra en pantalla.

Se guarda en la rama `bol-memory` del repositorio (así no redespliega la app), **cifrada** porque son datos de clientes.
Configura en *Settings → Secrets* de Streamlit Cloud:

```toml
GITHUB_TOKEN = "github_pat_..."   # fine-grained, solo este repositorio, permiso Contents: Read and write
# Opcionales:
# MEMORY_KEY = "una-frase-larga"  # llave de cifrado; si falta se deriva del token
# GITHUB_REPO = "dankypaixd/StreamlabsBOL"
# MEMORY_BRANCH = "bol-memory"
```

Sin `GITHUB_TOKEN` la memoria usa un archivo local (`.bol_memory/`) que Streamlit Cloud borra al reiniciar.

**Paquetes del sistema.** El OCR necesita Tesseract: en la raíz del repositorio debe existir `packages.txt`
(con `tesseract-ocr`, `tesseract-ocr-eng`, `tesseract-ocr-osd`). Elimina el archivo mal nombrado `packag2es.txt`.

**Pruebas.** `python -m pytest tests` (las que usan OCR se omiten si Tesseract no está instalado; no se incluyen
documentos reales: los PDF de prueba se generan con datos inventados en `tests/synthetic.py`).
