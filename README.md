# Generador de BOL

Aplicación de Streamlit que:

- Lee un PDF de Commercial Invoices.
- Usa texto nativo y OCR como respaldo.
- Genera un solo BOL por cada Air Waybill diferente.
- Coloca únicamente Ship To, Air Waybill, Packing ID, Order Item y Date & Time.

## Ejecutar localmente

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

Para que funcione el OCR también debe estar instalado Tesseract OCR en Windows. Los PDF con texto copiable funcionan sin OCR.

## Publicar en Streamlit Community Cloud

1. Subir esta carpeta a un repositorio de GitHub.
2. Entrar a Streamlit Community Cloud e iniciar sesión con GitHub.
3. Crear una aplicación nueva.
4. Seleccionar el repositorio y `streamlit_app.py`.
5. Presionar **Deploy**.

Community Cloud instalará las librerías de `requirements.txt` y Tesseract desde `packages.txt`.
