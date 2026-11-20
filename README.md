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
- `date_export.py`: lee el DIMS (por nombre de columna) y genera las hojas DATE (198 columnas del template) y DATA, los mensajes de cut off y el nombre del archivo.
- `inbonds_tab.py`: pestaña “Inbonds” (BOL o Commercial Invoice → estado, service level y `INBD` en la columna FK).
- `assets/SERVICE_LEVELS.csv`: lista estado → días (editable). `assets/DATE_HEADERS.json`: encabezados del template.
- `packages.txt`: paquetes del sistema para Streamlit Cloud (Tesseract OCR).
- `assets/BOL_TEMPLATE.pdf`: plantilla de BOL.
- `assets/DIMENSIONES_TEMPLATE.xlsx`: plantilla de dimensiones.
- `.streamlit/config.toml`: temas claro y oscuro y límite de carga.
- `.github/workflows/keepalive.yml` y `.github/keepalive.py`: visita la app cada 6 horas para que Streamlit Cloud no la duerma.

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
La fecha se mide desde el READY DATE (`MM/DD/YY HHMM`, que se escribe en la pantalla): dentro del cut-off 0800–1700 el mismo día
es el día 0; después de las 17:00 el día 0 es el siguiente día hábil; los fines de semana no cuentan. Se escribe `MM/DD/YY`.
Supuestos que no estaban en la especificación (ajustables en `service_levels.py`): antes de las 08:00 el día 0
es el mismo día; un READY DATE en fin de semana arranca el lunes; no se descuentan feriados (`HOLIDAYS`).

**Excel de salida.** Con el DIMS subido el libro trae las hojas `DATE` y `DATA` (la app arma DATA con el orden del
template; Q y R rellenos por BOL; el TRACKER, si viene, va en la columna S). La hoja `DATE` tiene las 198 columnas del template; se llenan A, B, E, Q, R, S, AB, AC, AF,
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

## Cambios V9.1

**México y otros países (como en la pestaña 1).** Si el Ship To de un BOL o de una Commercial Invoice es de México, el
estado se toma de la dirección de entrega de la pestaña 1 (Laredo, **TX**); si es de Latinoamérica, de Miami (**FL**).
Puerto Rico, Canadá y EE. UU. no cambian. Se reutiliza la misma función (`apply_destination_rule`) y las mismas
direcciones, así que si cambian en la pestaña 1 cambian aquí. Una diferencia a propósito: si el Ship To termina en
`US`, nunca se toma como extranjero (así “Panama City, FL” o “Mexico, MO” siguen siendo EE. UU.). El BOL queda con una
nota informativa (no se pone en amarillo); solo si las lecturas del OCR no coinciden sobre el país sale el aviso
`destination_unsure`. El módulo es `destinations.py`.

**READY DATE opcional.** Si se deja vacío, `Pickup Date`, `Pickup Time`, `Pickup Close Time` y la entrega calculada
quedan en blanco. Un READY DATE mal escrito sí bloquea la descarga (para no exportar en blanco por un error de dedo).

**Consignee Number (`T` en DATE).** Si el Ship To es la dirección de Laredo (11905 Conly Road, Laredo TX 78045) se
escribe `147229`; si es la de Miami (12430 NW 25th Street, Miami FL 33182) se escribe `210475`. Cuenta tanto cuando la
dirección se convierte (país México → Laredo, Latinoamérica → Miami) como cuando ya viene escrita así en el BOL o en la
Commercial Invoice. En cualquier otra dirección queda vacío. Aplica a las pestañas de contactos e Inbonds y se ve en la
vista previa (`Consignee No.`). El módulo es `destinations.py`.

**Mensajes de cut off y aviso de push (pestañas de contactos e Inbonds).** Debajo del READY DATE hay un campo
`POD DATE TIME` donde se escribe, por ejemplo, `11/20/26 0800`. Con eso la app arma los mensajes que antes salían de la
hoja DATA, listos para copiar: `ARRIVED AT 11/20/26 0800` (del POD), `LEFT PLANT 11/20/26 1456` (del READY DATE) y los
dos juntos. Si el READY DATE es después de las 17:00 (fuera del cut-off 0800–1700) sale un aviso de que la entrega se
está corriendo (push) un día. No cambia nada del Excel: la fecha y hora de entrega siguen saliendo del cálculo.

## Cambios V9

**DIMS en lugar de DATA.** Ya no hace falta pegar nada en la hoja DATA. En la pestaña de contactos se sube el DIMS
que se recibe (`.xlsx`, `.xlsm`, `.xls` o `.csv`). Se lee **por nombre de columna**, así que no importa el orden ni que
falten `HAZMAT`, `TRACKER` o `RACK` (o que vengan vacías): solo se exige `Barcode`. Con él la app:

- compara los BOL del DIMS con los de los PDF y avisa cuáles faltan de un lado y del otro;
- confirma el carrier de cada BOL;
- toma el HAZMAT (`DK`) y arma la hoja DATA.

**Carrier y Shipper Number (`E`).** Se identifica igual que la pestaña 1, por el contenido del PDF: el Ship From
“IEC TECHNOLOGIES” es BEX (`149570`) y “Foxconn” es VRP (`144463`). Se confirma con la columna `CARRIER` del DIMS y con
el nombre del archivo; si no coinciden gana la mayoría (en empate, el PDF) y el BOL queda en amarillo con el aviso. Si
no se puede saber, se elige a mano en la tarjeta.

**Nombre del archivo.** Campo opcional; si se deja vacío el archivo se guarda con el READY DATE (`11-19-26 1356.xlsx`).

**Mensajes de cut off.** Debajo del READY DATE hay un campo `POD DATE TIME` y tres mensajes listos para copiar, los mismos
que armaba `DATA!O18:O20`: `LEFT PLANT … ARRIVED AT …`, `LEFT PLANT …` y `ARRIVED AT …` (con casilla opcional para
agregar el día de la semana, p. ej. `SATURDAY`). También salen en la hoja DATA del Excel (N2, N3:N4 y N18:O20).

**Pestaña Inbonds.** Lee BOL escaneados o Commercial Invoice (el mismo lector de la pestaña 1), reconoce el estado del
Ship To y calcula service level y fecha de entrega. No lee el contacto. En DATE escribe `INBD` en la columna `FK`
(Priority Code) de cada BOL.

**Temas claro y oscuro.** El tema sigue el del dispositivo (se cambia en ⋮ → Settings → Theme). Se quitó el botón de
modo claro/oscuro propio y los colores forzados, que dejaban texto blanco sobre fondos claros.

**Que la app no se duerma.** Streamlit Community Cloud duerme las apps sin visitas (unas 12 horas). El workflow
`keepalive.yml` abre la app con un navegador cada 6 horas y, si estaba dormida, la despierta. Para activarlo:

1. En GitHub: *Settings → Secrets and variables → Actions → Variables → New repository variable*, nombre `APP_URL`,
   valor la dirección de tu app (`https://....streamlit.app`). Alternativa: poner la dirección en `.github/app_url.txt`.
2. En la pestaña *Actions* abre “Mantener la app despierta” y pulsa *Run workflow* para probarlo.

Es una ayuda, no una garantía: Streamlit puede cambiar sus reglas, y GitHub pausa los workflows programados si el
repositorio pasa 60 días sin actividad.
