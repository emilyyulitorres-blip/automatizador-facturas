import os
import re
import unicodedata
from io import BytesIO
from pathlib import Path

import fitz
import pandas as pd
import pytesseract
import streamlit as st
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

st.set_page_config(
    page_title="Automatizador de Facturas",
    page_icon="📄",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent

# Busca el logo con cualquiera de estos nombres.
LOGO_CANDIDATES = [
    BASE_DIR / "logo_solarteam.png",
    BASE_DIR / "LOGO SOLAR TEAM.png",
    BASE_DIR / "logo.png",
]

# Ruta local de Tesseract en Windows.
if os.name == "nt":
    tesseract_windows = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if tesseract_windows.exists():
        pytesseract.pytesseract.tesseract_cmd = str(tesseract_windows)

# -------------------------------------------------------------------
# DISEÑO
# -------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {
        max-width: 1550px;
        padding-top: 1.4rem;
        padding-bottom: 2rem;
    }
    .main-title {
        font-size: 2.7rem;
        line-height: 1.05;
        font-weight: 800;
        color: #172B4D;
        margin-bottom: 0.5rem;
    }
    .subtitle {
        font-size: 1.05rem;
        color: #475569;
    }
    .section-title {
        font-size: 1.55rem;
        font-weight: 750;
        color: #172B4D;
        margin-top: 1.6rem;
        margin-bottom: 0.8rem;
    }
    .summary-card {
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 20px 24px;
        min-height: 125px;
        background: white;
        box-shadow: 0 3px 12px rgba(15,23,42,.05);
    }
    .summary-label {
        color: #334155;
        font-size: 1rem;
        margin-bottom: 14px;
    }
    .summary-value {
        font-size: 2.1rem;
        font-weight: 800;
    }
    .ok-value { color: #16A34A; }
    .review-value { color: #F59E0B; }
    .error-value { color: #E11D48; }
    .footer {
        margin-top: 2rem;
        padding: 20px 24px;
        border-radius: 14px;
        background: #F8FAFC;
        color: #64748B;
        font-size: .9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------
# SINÓNIMOS
# -------------------------------------------------------------------
ETIQUETAS = {
    "subtotal": [
        "SUBTOTAL SIN IMPUESTOS",
        "SUBTOTAL GRAVADO",
        "SUBTOTAL 15%",
        "SUBTOTAL 12%",
        "BASE IMPONIBLE",
        "BASE GRAVADA",
        "SUBTOTAL",
    ],
    "iva": [
        "IVA 15%",
        "IVA 12%",
        "IMPUESTO IVA",
        "TOTAL IVA",
        "IMPUESTO AL VALOR AGREGADO",
        "IVA",
    ],
    "total": [
        "VALOR TOTAL",
        "TOTAL A PAGAR",
        "IMPORTE TOTAL",
        "MONTO TOTAL",
        "TOTAL FACTURA",
        "TOTAL",
    ],
}

PATRON_FACTURA = re.compile(r"\b\d{3}-\d{3}-\d{9}\b")
PATRON_RUC = re.compile(r"\b\d{13}\b")
PATRON_FECHA = re.compile(
    r"\b(?:\d{2}[/-]\d{2}[/-]\d{4}|\d{4}[/-]\d{2}[/-]\d{2})\b"
)
PATRON_MONTO = re.compile(
    r"(?<!\d)(?:\$?\s*)"
    r"(\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{2})(?!\d)"
)

RUC_SOLARTEAM = "1793069479001"


def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper().replace("\xa0", " ")
    texto = re.sub(r"[ \t]+", " ", texto)
    return texto.strip()


def lineas_limpias(texto: str) -> list[str]:
    return [
        re.sub(r"\s+", " ", linea).strip()
        for linea in texto.splitlines()
        if linea.strip()
    ]


def monto_normalizado(valor: str) -> str:
    if not valor:
        return ""

    valor = valor.replace("$", "").replace(" ", "")

    if "," in valor and "." in valor:
        if valor.rfind(",") > valor.rfind("."):
            valor = valor.replace(".", "").replace(",", ".")
        else:
            valor = valor.replace(",", "")
    else:
        valor = valor.replace(",", ".")

    try:
        return f"{float(valor):.2f}"
    except ValueError:
        return valor


def montos_en_linea(linea: str) -> list[str]:
    return [monto_normalizado(valor) for valor in PATRON_MONTO.findall(linea)]


def buscar_monto(
    lineas: list[str],
    etiquetas: list[str],
    excluir: list[str] | None = None,
) -> str:
    excluir = excluir or []
    resultados = []

    for indice, linea in enumerate(lineas):
        linea_normalizada = normalizar(linea)

        if not any(normalizar(etiqueta) in linea_normalizada for etiqueta in etiquetas):
            continue

        if any(normalizar(palabra) in linea_normalizada for palabra in excluir):
            continue

        valores = montos_en_linea(linea)
        if valores:
            resultados.append((indice, valores[-1]))
            continue

        # Algunos PDF colocan la etiqueta y el valor en líneas separadas.
        for siguiente in range(indice + 1, min(indice + 4, len(lineas))):
            valores_siguientes = montos_en_linea(lineas[siguiente])
            if valores_siguientes:
                resultados.append((siguiente, valores_siguientes[-1]))
                break

    if not resultados:
        return ""

    # El valor correcto suele aparecer al final del resumen tributario.
    return sorted(resultados, key=lambda item: item[0])[-1][1]


def detectar_tipo(texto: str) -> str:
    texto_normalizado = normalizar(texto)
    texto_compacto = re.sub(r"[^A-Z]", "", texto_normalizado)

    if "PROFORMA" in texto_normalizado or "PROFORMA" in texto_compacto:
        return "NO VÁLIDA - PROFORMA"

    if "COTIZACION" in texto_normalizado:
        return "NO VÁLIDA - COTIZACIÓN"

    if "NOTA DE CREDITO" in texto_normalizado:
        return "NO VÁLIDA - NOTA DE CRÉDITO"

    if "NOTA DE VENTA" in texto_normalizado:
        return "REVISAR - NOTA DE VENTA"

    if "FACTURA" in texto_normalizado or "FACTURA" in texto_compacto:
        if (
            "NUMERO DE AUTORIZACION" in texto_normalizado
            or "CLAVE DE ACCESO" in texto_normalizado
        ):
            return "FACTURA ELECTRÓNICA"
        return "FACTURA"

    return "REVISAR - NO IDENTIFICADO"


def extraer_factura(texto: str) -> str:
    coincidencias = PATRON_FACTURA.findall(texto)
    return coincidencias[0] if coincidencias else ""


def extraer_fecha(texto: str, lineas: list[str]) -> str:
    # Se prioriza la fecha de emisión.
    for indice, linea in enumerate(lineas):
        linea_normalizada = normalizar(linea)

        if (
            "FECHA EMISION" in linea_normalizada
            or "FECHA DE EMISION" in linea_normalizada
            or linea_normalizada == "FECHA"
            or linea_normalizada.startswith("FECHA ")
        ):
            fecha = PATRON_FECHA.search(linea)
            if fecha:
                return fecha.group(0)

            for siguiente in range(indice + 1, min(indice + 4, len(lineas))):
                fecha = PATRON_FECHA.search(lineas[siguiente])
                if fecha:
                    return fecha.group(0)

    fechas = PATRON_FECHA.findall(texto)
    return fechas[-1] if fechas else ""


def extraer_ruc_emisor(texto: str) -> str:
    rucs = PATRON_RUC.findall(texto)

    if not rucs:
        return ""

    # Evita tomar el RUC del cliente Solarteam como RUC del proveedor.
    for ruc in rucs:
        if ruc != RUC_SOLARTEAM:
            return ruc

    return rucs[0]


def es_linea_proveedor_valida(linea: str) -> bool:
    n = normalizar(linea)

    etiquetas_descartadas = [
        "RUC",
        "FACTURA",
        "NUMERO DE AUTORIZACION",
        "CLAVE DE ACCESO",
        "FECHA",
        "AMBIENTE",
        "EMISION",
        "PRODUCCION",
        "NORMAL",
        "DIRECCION",
        "OBLIGADO",
        "CONTRIBUYENTE",
        "RAZON SOCIAL",
        "IDENTIFICACION",
        "SOLARTEAM",
        "SOLAR TEAM",
        "SUBTOTAL",
        "VALOR TOTAL",
    ]

    return (
        len(linea.strip()) >= 7
        and not any(etiqueta in n for etiqueta in etiquetas_descartadas)
        and not PATRON_RUC.search(linea)
        and not PATRON_FACTURA.search(linea)
        and not PATRON_FECHA.search(linea)
        and bool(re.search(r"[A-ZÁÉÍÓÚÑ]{3,}", linea.upper()))
    )


def extraer_proveedor(lineas: list[str], ruc_emisor: str) -> str:
    # Primero busca el nombre inmediatamente después del RUC del emisor.
    for indice, linea in enumerate(lineas):
        if ruc_emisor and ruc_emisor in linea:
            for siguiente in range(indice + 1, min(indice + 9, len(lineas))):
                candidato = lineas[siguiente].strip()
                if es_linea_proveedor_valida(candidato):
                    return candidato

    # Segunda estrategia: buscar en las primeras líneas.
    for linea in lineas[:45]:
        if es_linea_proveedor_valida(linea):
            return linea.strip()

    return ""


def extraer_cliente(texto: str) -> str:
    n = normalizar(texto)

    if "SOLARTEAM SAS" in n:
        return "SOLARTEAM SAS"

    if "SOLAR TEAM S.A.S" in n:
        return "SOLAR TEAM S.A.S"

    return ""


def validar(datos: dict) -> tuple[str, str]:
    if datos["Tipo documento"].startswith("NO VÁLIDA"):
        return "REVISAR", "Documento no registrable como factura"

    obligatorios = [
        "Fecha",
        "Factura",
        "Proveedor",
        "RUC Emisor",
        "Subtotal",
        "IVA",
        "Total",
    ]

    faltantes = [
        campo
        for campo in obligatorios
        if not str(datos.get(campo, "")).strip()
    ]

    if faltantes:
        return "REVISAR", "Faltan: " + ", ".join(faltantes)

    try:
        subtotal = float(datos["Subtotal"])
        iva = float(datos["IVA"])
        total = float(datos["Total"])

        if abs((subtotal + iva) - total) > 0.10:
            return "REVISAR", "Subtotal + IVA no coincide con el total"
    except ValueError:
        return "REVISAR", "Los valores monetarios requieren revisión"

    return "OK", ""


def extraer_datos(texto: str, nombre_archivo: str) -> dict:
    lineas = lineas_limpias(texto)
    tipo = detectar_tipo(texto)
    ruc_emisor = extraer_ruc_emisor(texto)

    datos = {
        "Estado": "",
        "Tipo documento": tipo,
        "Archivo": nombre_archivo,
        "Fecha": extraer_fecha(texto, lineas),
        "Factura": extraer_factura(texto),
        "Proveedor": extraer_proveedor(lineas, ruc_emisor),
        "RUC Emisor": ruc_emisor,
        "Cliente": extraer_cliente(texto),
        "RUC Cliente": RUC_SOLARTEAM if RUC_SOLARTEAM in texto else "",
        "Subtotal": buscar_monto(
            lineas,
            ETIQUETAS["subtotal"],
            excluir=["NO OBJETO", "EXENTO", "DESCUENTO"],
        ),
        "IVA": buscar_monto(
            lineas,
            ETIQUETAS["iva"],
            excluir=["INCLUYE IVA"],
        ),
        "Total": buscar_monto(
            lineas,
            ETIQUETAS["total"],
            excluir=["SIN SUBSIDIO", "DESCUENTO", "SUBTOTAL"],
        ),
        "Proyecto": "",
        "Consumo": "",
        "Categoría": "",
        "Observación": "",
    }

    estado, observacion = validar(datos)
    datos["Estado"] = estado
    datos["Observación"] = observacion
    return datos


def preparar_imagen(imagen: Image.Image) -> Image.Image:
    imagen = ImageOps.exif_transpose(imagen).convert("L")
    imagen = ImageOps.autocontrast(imagen)
    imagen = ImageEnhance.Contrast(imagen).enhance(1.8)
    imagen = imagen.filter(ImageFilter.SHARPEN)
    return imagen


def puntaje_texto_ocr(texto: str) -> int:
    n = normalizar(texto)
    palabras = ["FACTURA", "RUC", "TOTAL", "IVA", "FECHA", "SUBTOTAL"]
    return sum(n.count(palabra) for palabra in palabras) + len(texto) // 200


def aplicar_ocr(imagen: Image.Image) -> str:
    imagen = preparar_imagen(imagen)
    mejores_textos = []

    # Prueba varias orientaciones para fotografías giradas.
    for angulo in [0, 90, 270, 180]:
        rotada = imagen.rotate(angulo, expand=True)
        texto = pytesseract.image_to_string(
            rotada,
            lang="spa",
            config="--oem 3 --psm 6",
        )
        mejores_textos.append((puntaje_texto_ocr(texto), texto))

    return max(mejores_textos, key=lambda item: item[0])[1]


def leer_imagen(archivo) -> str:
    imagen = Image.open(BytesIO(archivo.getvalue()))
    return aplicar_ocr(imagen)


def leer_pdf(archivo) -> str:
    texto = ""
    datos = archivo.getvalue()

    with fitz.open(stream=datos, filetype="pdf") as documento:
        for pagina in documento:
            texto_pagina = pagina.get_text("text")
            texto += texto_pagina + "\n"

            # PDF escaneado o sin texto digital.
            if len(texto_pagina.strip()) < 40:
                pixmap = pagina.get_pixmap(
                    matrix=fitz.Matrix(2.4, 2.4),
                    alpha=False,
                )
                imagen = Image.frombytes(
                    "RGB",
                    [pixmap.width, pixmap.height],
                    pixmap.samples,
                )
                texto += aplicar_ocr(imagen) + "\n"

    return texto


def crear_excel(df: pd.DataFrame) -> bytes:
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Facturas")
        hoja = writer.sheets["Facturas"]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        for columna in hoja.columns:
            ancho = min(
                max(len(str(celda.value or "")) for celda in columna) + 2,
                45,
            )
            hoja.column_dimensions[columna[0].column_letter].width = ancho

    return salida.getvalue()


# -------------------------------------------------------------------
# INTERFAZ
# -------------------------------------------------------------------
col_logo, col_titulo = st.columns(
    [1.25, 5.75],
    vertical_alignment="center",
)

with col_logo:
    logo = next(
        (ruta for ruta in LOGO_CANDIDATES if ruta.exists()),
        None,
    )
    if logo:
        st.image(str(logo), use_container_width=True)

with col_titulo:
    st.markdown(
        '<div class="main-title">Automatizador de Facturas</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="subtitle">
            Carga facturas PDF o imágenes, revisa la información extraída
            y genera un Excel listo para descargar.
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()
st.markdown(
    '<div class="section-title">📂 Sube tus facturas</div>',
    unsafe_allow_html=True,
)

archivos = st.file_uploader(
    "Sube facturas PDF o imágenes",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if archivos:
    resultados = []
    barra = st.progress(0, text="Preparando documentos...")

    for indice, archivo in enumerate(archivos, start=1):
        try:
            nombre = archivo.name.lower()

            if nombre.endswith(".pdf"):
                texto = leer_pdf(archivo)
            else:
                texto = leer_imagen(archivo)

            resultados.append(extraer_datos(texto, archivo.name))

        except Exception as error:
            resultados.append(
                {
                    "Estado": "ERROR",
                    "Tipo documento": "ERROR",
                    "Archivo": archivo.name,
                    "Fecha": "",
                    "Factura": "",
                    "Proveedor": "",
                    "RUC Emisor": "",
                    "Cliente": "",
                    "RUC Cliente": "",
                    "Subtotal": "",
                    "IVA": "",
                    "Total": "",
                    "Proyecto": "",
                    "Consumo": "",
                    "Categoría": "",
                    "Observación": str(error),
                }
            )

        barra.progress(
            indice / len(archivos),
            text=f"Procesando {indice} de {len(archivos)} documentos...",
        )

    barra.empty()
    df = pd.DataFrame(resultados)

    cantidad_ok = int((df["Estado"] == "OK").sum())
    cantidad_revisar = int((df["Estado"] == "REVISAR").sum())
    cantidad_error = int((df["Estado"] == "ERROR").sum())

    st.markdown(
        '<div class="section-title">📊 Resumen del procesamiento</div>',
        unsafe_allow_html=True,
    )

    columna_ok, columna_revisar, columna_error = st.columns(3)

    with columna_ok:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">✅ Facturas OK</div>
                <div class="summary-value ok-value">{cantidad_ok}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with columna_revisar:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">⚠️ Para revisar</div>
                <div class="summary-value review-value">{cantidad_revisar}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with columna_error:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">❌ Errores</div>
                <div class="summary-value error-value">{cantidad_error}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="section-title">📋 Revisa y corrige antes de descargar</div>',
        unsafe_allow_html=True,
    )

    df_editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        height=min(620, 130 + len(df) * 40),
        disabled=["Archivo"],
        column_config={
            "Estado": st.column_config.SelectboxColumn(
                "Estado",
                options=["OK", "REVISAR", "ERROR"],
                width="small",
            ),
            "Tipo documento": st.column_config.TextColumn(width="medium"),
            "Archivo": st.column_config.TextColumn(width="large"),
            "Fecha": st.column_config.TextColumn(width="small"),
            "Factura": st.column_config.TextColumn(width="medium"),
            "Proveedor": st.column_config.TextColumn(width="large"),
            "RUC Emisor": st.column_config.TextColumn(width="medium"),
            "Cliente": st.column_config.TextColumn(width="medium"),
            "RUC Cliente": st.column_config.TextColumn(width="medium"),
            "Subtotal": st.column_config.TextColumn(width="small"),
            "IVA": st.column_config.TextColumn(width="small"),
            "Total": st.column_config.TextColumn(width="small"),
            "Observación": st.column_config.TextColumn(width="large"),
        },
    )

    st.download_button(
        "📥 Descargar Excel",
        data=crear_excel(df_editado),
        file_name="facturas_extraidas.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

else:
    st.info(
        "Carga uno o varios archivos PDF, JPG, JPEG o PNG para comenzar."
    )

st.markdown(
    """
    <div class="footer">
        <b>Solar Team</b><br>
        Automatización de información de facturas.
    </div>
    """,
    unsafe_allow_html=True,
)
