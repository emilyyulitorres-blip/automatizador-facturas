# Automatizador de Facturas

Aplicación web en Streamlit para cargar facturas PDF o imágenes, extraer datos principales y generar un Excel.

## Ejecutar localmente

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Para OCR en Windows instala Tesseract OCR. La app detecta automáticamente esta ruta si existe:

`C:\Program Files\Tesseract-OCR\tesseract.exe`

## Publicar en Streamlit Community Cloud

1. Sube estos archivos a un repositorio de GitHub.
2. En Streamlit Community Cloud crea una nueva app.
3. Selecciona el repositorio, rama `main` y archivo principal `app.py`.
4. Mantén `requirements.txt` y `packages.txt` en la raíz del repositorio.

`packages.txt` instala Tesseract OCR y español en la nube.

## Archivos importantes

- `app.py`: interfaz principal.
- `modules/readers.py`: lectura de PDF e imágenes.
- `modules/extractor.py`: extracción de datos.
- `modules/excel_manager.py`: descarga y guardado en Excel.
