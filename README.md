# Generador de BOL y convertidor de dimensiones

Aplicación Streamlit para:

1. Leer Commercial Invoices VRP y BEX y generar un BOL por Air Waybill.
2. Revisar los BOL detectados en una tabla ordenada numéricamente.
3. Seleccionar cuáles BOL generar y editar la dirección que aparecerá en el PDF.
4. Leer archivos de dimensiones VRP/BEX, distribuir cajas y pesos por pallet y exportar la plantilla final.
5. Editar cualquier celda de la tabla de dimensiones, seleccionar o desmarcar filas y exportar solo las elegidas.
6. **Leer contactos de BOL (V8):** leer el bloque “Ship To Attention” (nombre, correo, teléfono) y el estado del
   Ship To de cada Bill of Lading, calcular el service level y exportar el Excel de importación.

## Archivos principales

- `streamlit_app.py`: interfaz de la aplicación.
- `bol_generator.py`: lectura de CI y generación de BOL.
- `dimension_converter.py`: lectura de archivos VRP/BEX y creación del Excel.
- `contacts_tab.py`: pestaña “Leer contactos de BOL” (interfaz, capturas, confirmación, descargas).
- `contact_extractor.py`: OCR del Ship To Attention y del estado (varias lecturas con votación, orientación y enderezado).
- `contact_review.py`: une las páginas de un BOL, aplica lo aprendido y clasifica en verde / amarillo / rojo.
- `contact_memory.py`: memoria de contactos confirmados (cifrada, guardada en una rama de GitHub).
- `service_levels.py`: service level por estado y fecha de entrega con cut-off 0800–1700 (sin fines de semana).
- `date_export.py`: genera la hoja DATE (198 columnas del template) **sin formato**, lee la columna Barcode del DIMS, arma los mensajes de cut off y el nombre del archivo.
- `dims_compare.py`: comparación opcional de los BOL de los PDF contra los del DIMS (contactos e Inbonds), con aceptar / ignorar.
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
- 🔴 *Incompleto*: falta algo (sin teléfono, sin estado…) o el formato no es válido.

Un bloque *Ship To Attention* **vacío en el PDF no es error**: queda en 🟢 y el Excel lleva nombre, correo y teléfono en
blanco (si el estado del Ship To se dedujo del ZIP, sigue en 🟡 por eso). En cuanto se escribe algún dato a mano, deja de
ser un bloque vacío y se exige el resto.

Un teléfono no tiene formato fijo: se aceptan dígitos y el signo `+`.

**Service level.** Sale del estado del Ship To con `assets/SERVICE_LEVELS.csv` (`3D`, `4D`…). El READY DATE
(`MM/DD/YY HHMM`, que se escribe en la pantalla) va al Excel como Pickup Date / Time / Close Time. La fecha de entrega
(READY DATE + service level, cut-off 0800–1700) ya no se muestra ni se escribe en ningún Excel; `service_levels.py` la
sigue calculando. Supuestos que no estaban en la especificación (ajustables ahí): después de las 17:00 el día 0 es el
siguiente día hábil; antes de las 08:00 el día 0 es el mismo día; un READY DATE en fin de semana arranca el lunes; no se
descuentan feriados (`HOLIDAYS`).

**Excel de salida.** Solo se descarga la hoja `DATE` (ya no hay hoja `DATA` ni libro completo). Tiene las 198 columnas del
template; se llenan A, B, E, Q, R, S (READY DATE: Pickup Date, Time y Close Time), T (Consignee Number), AB, AC, AK, CO, DJ, DK y DU
con los mismos valores que calculan tus fórmulas. **La fecha y hora de entrega (AF y AG) quedan vacías**; los encabezados se
conservan para no mover columnas. Por defecto, los dudosos e incompletos sin confirmar salen con nombre, correo y teléfono
vacíos (casilla en pantalla).

Dos casillas arriba de la descarga deciden qué lleva la hoja: **Incluir el Ship To Attention** (nombre, correo y teléfono:
`AB`, `AC` y `DU`) e **Incluir el service level** (`AK`); en Inbonds solo existe la del service level. Ambas vienen
marcadas; al desmarcarlas esas columnas salen vacías y la tabla de vista previa las oculta.

**Sin BOL repetidos.** Si al revisar corriges un BOL (p. ej. una mala lectura `98463300` en lugar de `98163300`) y ese
número ya estaba en la lista, queda un solo registro: el que ya estaba confirmado o, si no, el que acabas de corregir (se
suman sus páginas). Se avisa con un mensaje arriba de los resultados y el conteo de BOL baja.

**Lectura del BOL reforzada.** El número de BOL se lee de dos lugares de la página (el encabezado y la fila *Tracking No*).
Si los dos coinciden en una lectura rápida, listo; si no, cada recorte se vuelve a leer con 16 versiones (4 escalas × 4
tratamientos de la imagen) y se vota el número completo. Pasa a 🟢 solo si el número es fiable: las dos fuentes coinciden o una
lo lee con claridad (≥ 70 % de las lecturas). Además:

- *Dígito dividido.* Si la mitad de las lecturas dicen `96433794` y la otra mitad `98433794`, aunque encabezado y fila
  coincidan el BOL queda en 🟡 y la tarjeta ofrece el botón **Usar 98433794** (el 4↔8, 1↔4↔7, 6↔8… se lee igual en los dos
  lugares). Una coincidencia exacta con la memoria ya no esconde un BOL dudoso.
- *Páginas “1 OF 2” / “2 OF 2”.* Se lee “Page: n OF m” del encabezado; una segunda página consecutiva cuyo BOL difiere en uno o
  dos dígitos y viene dudosa se une a la primera en vez de contarse como otro BOL.
- *Parecido a otro BOL de la lista.* Un BOL dudoso que difiere en un solo dígito de otro ya leído muestra “Usar ese BOL”;
  nunca se corrige solo (los BOL consecutivos sí pueden diferir en un dígito).

**Nombre y correo en varias líneas.** Si el nombre o el correo se escribieron en renglones distintos, se vuelven a armar:
el `.com` (o `com`) que quedó abajo, el salto justo después de la `@`, la `@` al inicio de la línea, el punto leído como espacio
o dos puntos (`fakehp : com`) y las manchas del escaneo (`|`). Solo se une si el resultado es un correo válido, para no tragarse el
teléfono. Si aun así falta la terminación, se vuelve a leer la zona de la derecha y de abajo y se acepta una terminación conocida
si coinciden 2 de 3 lecturas. Las palabras sueltas de baja confianza (`a`, `seh:`) y los símbolos aislados se quitan del nombre,
y las confusiones `i/l/I` del correo se desempatan con el nombre (queda en 🟡).

**El archivo sale sin formato, para importarlo tal cual.** Antes había que copiar la hoja y pegarla como “solo valores” en otra.
El motivo: el archivo traía el primer renglón inmovilizado, anchos en las 198 columnas y los textos guardados “en línea”
(no como los guarda Excel). Ahora se escribe con XlsxWriter sin anchos, sin paneles, sin estilos y con textos compartidos; las
celdas sin dato quedan realmente vacías (no con un texto en blanco). El BOL numérico va como número y lo demás como texto
(`R` conserva su espacio inicial, como `RIGHT(N2,5)`).

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

**READY DATE opcional.** Si se deja vacío, `Pickup Date`, `Pickup Time` y `Pickup Close Time` quedan en blanco (y las columnas
de fecha y hora ready de las tablas se ocultan). Un READY DATE mal escrito sí bloquea la descarga, para no exportar en blanco por un error de dedo.

**Consignee Number (`T` en DATE).** Si el Ship To es la dirección de Laredo (11905 Conly Road, Laredo TX 78045) se
escribe `147229`; si es la de Miami (12430 NW 25th Street, Miami FL 33182) se escribe `210475`. Cuenta tanto cuando la
dirección se convierte (país México → Laredo, Latinoamérica → Miami) como cuando ya viene escrita así en el BOL o en la
Commercial Invoice. En cualquier otra dirección queda vacío. Aplica a las pestañas de contactos e Inbonds y se ve en la
vista previa (`Consignee No.`). El módulo es `destinations.py`.

**Arriba de cada pestaña (contactos e Inbonds): READY DATE, POD y mensajes para copiar.** READY DATE y `POD DATE TIME`
van en una fila al inicio. Con ellos la app arma, uno al lado del otro y listos para copiar con su icono:
`LEFT PLANT 11/20/26 1456 ARRIVED AT 11/20/26 0800`, `LEFT PLANT 11/20/26 1456` y `ARRIVED AT 11/20/26 0800`. Cambian con
lo que se escribe; sin READY DATE (ni POD) no se ve ninguno. El POD solo sirve para armar el mensaje ARRIVED AT. Si el
READY DATE es después de las 17:00 (fuera del cut-off 0800–1700) sale un aviso de que la entrega se está corriendo (push)
un día. En todas las tablas y tarjetas (la fecha de entrega ya no se muestra) salen las columnas **Fecha ready** y **Hora
ready** del READY DATE.

## Cambios V9

**DIMS opcional, solo para comparar.** En las pestañas de contactos e Inbonds se puede subir el DIMS que se recibe (`.xlsx`,
`.xlsm`, `.xls` o `.csv`). Se lee **por nombre de columna**: solo se exige `Barcode`, en cualquier orden. Su única función es
verificar que trae **los mismos BOL y la misma cantidad** que los PDF; no agrega ni cambia nada en el Excel (ni carrier, ni
HAZMAT, ni READY DATE). Es opcional, porque puede que en ese momento no se tenga: sin DIMS se descarga el DATE directamente.

Si hay diferencias se listan por lado (*solo en PDF* / *solo en DIMS*) y la descarga queda bloqueada hasta que hagas una de
estas dos cosas: **aceptar** los BOL que no coinciden (uno por uno, en la lista) o marcar **Ignorar la comparación**. Los BOL
que están solo en los PDF se quedan en el Excel; los que están solo en el DIMS no llevan fila porque no hay PDF. Un DIMS
ilegible o sin `Barcode` solo muestra el error y no bloquea nada. Como el Excel ya no toma el HAZMAT del DIMS, `DK` sale en `0`.

**Carrier y Shipper Number (`E`).** Se identifica igual que la pestaña 1, por el contenido del PDF: el Ship From
“IEC TECHNOLOGIES” es BEX (`149570`) y “Foxconn” es VRP (`144463`). Se confirma con
el nombre del archivo; si no coinciden gana el PDF y el BOL queda en amarillo con el aviso. Si
no se puede saber, se elige a mano en la tarjeta.

**Nombre del archivo.** Campo opcional; si se deja vacío el archivo se guarda con el READY DATE (`11-19-26 1356.xlsx`).

**Mensajes de cut off.** Debajo del READY DATE hay un campo `POD DATE TIME` y tres mensajes listos para copiar, los mismos
que armaba `DATA!O18:O20`: `LEFT PLANT … ARRIVED AT …`, `LEFT PLANT …` y `ARRIVED AT …` (con casilla opcional para
agregar el día de la semana, p. ej. `SATURDAY`). Son solo para copiar y pegar: no van al Excel.

**Pestaña Inbonds.** Lee BOL escaneados o Commercial Invoice (el mismo lector de la pestaña 1), reconoce el estado del
Ship To y calcula el service level. No lee el contacto. En DATE escribe `INBD` en la columna `FK`
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
