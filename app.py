import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO
from PIL import Image
import pytesseract

st.set_page_config(page_title="Automatizador de Facturas", layout="wide")

st.markdown("""
# ☀️ Automatizador de Facturas

Carga facturas PDF o imágenes, revisa los datos extraídos y genera un Excel listo para descargar.

---
""")

def buscar(patron, texto):
    m = re.search(patron, texto, re.IGNORECASE)
    return m.group(1).strip() if m else ""

def leer_pdf(file):
    texto = ""
    data = file.read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            texto += page.get_text("text") + "\n"
    return texto

def leer_imagen(file):
    img = Image.open(file)
    return pytesseract.image_to_string(img, lang="spa")

def tipo_documento(texto):
    t = texto.upper()
    if "PROFORMA" in t:
        return "NO VÁLIDA - PROFORMA"
    if "FACTURA" in t:
        return "FACTURA"
    return "REVISAR - NO IDENTIFICADO"

def extraer(texto, archivo):
    tipo = tipo_documento(texto)

    fecha = buscar(r"Fecha Emisi[oó]n[:\s]*([0-9]{2}/[0-9]{2}/[0-9]{4})", texto)
    if not fecha:
        fecha = buscar(r"([0-9]{2}/[0-9]{2}/[0-9]{4})", texto)

    factura = buscar(r"([0-9]{3}-[0-9]{3}-[0-9]{9})", texto)
    ruc = buscar(r"R\.?U\.?C\.?[:\s]*([0-9]{13})", texto)

    subtotal = buscar(r"SUBTOTAL SIN IMPUESTOS\s*([0-9]+[\.,][0-9]{2})", texto)
    if not subtotal:
        subtotal = buscar(r"SUBTOTAL\s*[0-9]*%?\s*([0-9]+[\.,][0-9]{2})", texto)

    iva = buscar(r"IVA\s*[0-9]*%?\s*([0-9]+[\.,][0-9]{2})", texto)

    total = buscar(r"VALOR TOTAL\s*([0-9]+[\.,][0-9]{2})", texto)
    if not total:
        total = buscar(r"TOTAL\s*([0-9]+[\.,][0-9]{2})", texto)

    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    proveedor = ""
    for linea in lineas[:15]:
        if len(linea) > 5 and "RUC" not in linea.upper() and "FACTURA" not in linea.upper():
            proveedor = linea
            break

    estado = "OK" if tipo == "FACTURA" else "REVISAR"

    return {
        "Estado": estado,
        "Tipo documento": tipo,
        "Archivo": archivo,
        "Fecha": fecha,
        "Factura": factura,
        "Proveedor": proveedor,
        "RUC Emisor": ruc,
        "Subtotal": subtotal,
        "IVA": iva,
        "Total": total,
        "Proyecto": "",
        "Consumo": "",
        "Categoría": "",
        "Observación": "" if estado == "OK" else "Documento para revisión"
    }

archivos = st.file_uploader(
    "Sube facturas PDF o imágenes",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if archivos:
    resultados = []

    for archivo in archivos:
        try:
            nombre = archivo.name.lower()

            if nombre.endswith(".pdf"):
                texto = leer_pdf(archivo)
            else:
                texto = leer_imagen(archivo)

            resultados.append(extraer(texto, archivo.name))

        except Exception as e:
            resultados.append({
                "Estado": "ERROR",
                "Tipo documento": "ERROR",
                "Archivo": archivo.name,
                "Fecha": "",
                "Factura": "",
                "Proveedor": "",
                "RUC Emisor": "",
                "Subtotal": "",
                "IVA": "",
                "Total": "",
                "Proyecto": "",
                "Consumo": "",
                "Categoría": "",
                "Observación": str(e)
            })

    df = pd.DataFrame(resultados)

    ok = len(df[df["Estado"] == "OK"])
    revisar = len(df[df["Estado"] == "REVISAR"])
    error = len(df[df["Estado"] == "ERROR"])

    st.markdown("### 📊 Resumen del procesamiento")

c1, c2, c3 = st.columns(3)

with c1:
    st.success(f"✅ Facturas OK: {ok}")

with c2:
    st.warning(f"⚠️ Para revisar: {revisar}")

with c3:
    st.error(f"❌ Errores: {error}")

    st.subheader("Revisa y corrige antes de descargar")
    df_editado = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_editado.to_excel(writer, index=False, sheet_name="Facturas")

    st.download_button(
        "Descargar Excel",
        data=output.getvalue(),
        file_name="facturas_extraidas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
