# Generador de BOL y dimensiones

Aplicacion Streamlit con dos apartados:

## 1. Generar BOL

- Lee Commercial Invoices VRP y BEX.
- Primero utiliza el texto del PDF y aplica OCR si la pagina es una imagen.
- Extrae Air Waybill, Packing ID, Heart Order y Ship To.
- Ignora el Order Item que aparece junto a Ship To en las CI de BEX.
- Deduplica paginas: un BOL por Air Waybill.
- Permite marcar todos, desmarcar todos o elegir BOL individualmente.
- Fecha y hora del BOL en horario de Costa Rica.
- Reglas de destino:
  - Mexico -> Corporativo Galvan S.C, Laredo.
  - Otros paises de Latinoamerica -> US_MIAMI DSV Inc. - MIA.
  - Puerto Rico se conserva sin cambios.
  - Los destinos fuera de Latinoamerica se conservan.

## 2. Convertir dimensiones

- Acepta `.xlsx` de VRP y `.xls` de BEX.
- Agrupa por trailer + tracking.
- Suma pallets, cajas y peso en kg.
- Convierte el peso total a libras enteras.
- Distribuye cajas y libras enteras entre los pallets sin alterar los totales.
- Crea una fila por pallet.
- Todas las dimensiones se exportan como `48 x 40 x 40`.
- Deja RACK, HAZMAT, TIME IN y TIME OUT vacios.

## Publicar en Streamlit Community Cloud

1. Descomprime el ZIP.
2. En GitHub abre el repositorio.
3. Usa **Add file -> Upload files**.
4. Arrastra el contenido de la carpeta, no el ZIP.
5. Confirma que `streamlit_app.py` quede en la raiz y que existan:

```text
assets/BOL_TEMPLATE.pdf
assets/DIMENSIONES_TEMPLATE.xlsx
.streamlit/config.toml
bol_generator.py
dimension_converter.py
packages.txt
requirements.txt
streamlit_app.py
```

6. Haz **Commit changes**.
7. En Streamlit Community Cloud despliega:

```text
Repository: usuario/repositorio
Branch: main
Main file path: streamlit_app.py
```

Cuando actualices archivos en GitHub, Streamlit volvera a desplegar la aplicacion automaticamente.
