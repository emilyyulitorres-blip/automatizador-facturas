
import os
import re
import unicodedata
from datetime import datetime
from io import BytesIO
from pathlib import Path

import fitz
import pandas as pd
import pytesseract
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


# =========================================================
# CONFIGURACIÓN
# =========================================================
st.set_page_config(
    page_title="Automatizador de Facturas",
    page_icon="📄",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
RUC_SOLARTEAM = "1793069479001"

LOGO_CANDIDATES = [
    BASE_DIR / "logo.png",
    BASE_DIR / "logo_solarteam.png",
    BASE_DIR / "LOGO SOLAR TEAM.png",
]

if os.name == "nt":
    ruta_tesseract = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if ruta_tesseract.exists():
        pytesseract.pytesseract.tesseract_cmd = str(ruta_tesseract)


# =========================================================
# ESTILOS
# =========================================================
st.markdown(
    """
    <style>
    .block-container {
        max-width: 1550px;
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }

    .main-title {
        font-size: 2.65rem;
        line-height: 1.05;
        font-weight: 800;
        color: #172B4D;
        margin-bottom: 0.45rem;
    }

    .subtitle {
        font-size: 1.02rem;
        color: #475569;
    }

    .section-title {
        font-size: 1.5rem;
        font-weight: 750;
        color: #172B4D;
        margin-top: 1.5rem;
        margin-bottom: 0.8rem;
    }

    .summary-card {
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 22px;
        min-height: 115px;
        background: #FFFFFF;
        box-shadow: 0 3px 12px rgba(15, 23, 42, 0.05);
    }

    .summary-label {
        font-size: 0.95rem;
        color: #475569;
        margin-bottom: 12px;
    }

    .summary-value {
        font-size: 2rem;
        font-weight: 800;
    }

    .ok-value { color: #16A34A; }
    .review-value { color: #F59E0B; }
    .error-value { color: #E11D48; }
    .total-value { color: #172B4D; }

    .footer {
        margin-top: 2rem;
        padding: 18px 22px;
        border-radius: 14px;
        background: #F8FAFC;
        color: #64748B;
        font-size: 0.9rem;
    }

    [data-testid="stFileUploader"] {
        border: 1px dashed #94A3B8;
        border-radius: 14px;
        padding: 0.4rem;
        background: #F8FAFC;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# PATRONES
# =========================================================
PATRON_FACTURA = re.compile(r"\b\d{3}-\d{3}-\d{9}\b")
PATRON_RUC = re.compile(r"\b\d{13}\b")
PATRON_FECHA = re.compile(
    r"\b(?:\d{2}[/-]\d{2}[/-]\d{4}|\d{4}[/-]\d{2}[/-]\d{2})\b"
)
PATRON_MONTO = re.compile(
    r"(?<!\d)(\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+[.,]\d{2})(?!\d)"
)


# =========================================================
# UTILIDADES
# =========================================================
def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )
    texto = texto.upper().replace("\xa0", " ")
    texto = re.sub(r"\s+", " ", texto)
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
        return ""


def es_monto_valido(valor: str) -> bool:
    try:
        float(valor)
        return True
    except (TypeError, ValueError):
        return False


# =========================================================
# IDENTIFICACIÓN DEL DOCUMENTO
# =========================================================
def detectar_tipo(texto: str) -> str:
    n = normalizar(texto)
    compacto = re.sub(r"[^A-Z]", "", n)

    if "PROFORMA" in n or "PROFORMA" in compacto:
        return "NO VÁLIDA - PROFORMA"

    if "COTIZACION" in n:
        return "NO VÁLIDA - COTIZACIÓN"

    if "NOTA DE CREDITO" in n:
        return "NO VÁLIDA - NOTA DE CRÉDITO"

    if "NOTA DE VENTA" in n:
        return "REVISAR - NOTA DE VENTA"

    if "FACTURA" in n or "FACTURA" in compacto:
        if "NUMERO DE AUTORIZACION" in n or "CLAVE DE ACCESO" in n:
            return "FACTURA ELECTRÓNICA"
        return "FACTURA"

    return "REVISAR - NO IDENTIFICADO"


def extraer_factura(texto: str) -> str:
    coincidencias = PATRON_FACTURA.findall(texto)
    return coincidencias[0] if coincidencias else ""


def extraer_fecha(texto: str) -> str:
    lineas = lineas_limpias(texto)

    for indice, linea in enumerate(lineas):
        n = normalizar(linea)

        if (
            "FECHA EMISION" in n
            or "FECHA DE EMISION" in n
            or n == "FECHA"
            or n.startswith("FECHA ")
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

    for ruc in rucs:
        if ruc != RUC_SOLARTEAM:
            return ruc

    return rucs[0] if rucs else ""


def extraer_cliente(texto: str) -> str:
    n = normalizar(texto)

    if "SOLARTEAM SAS" in n:
        return "SOLARTEAM SAS"

    if "SOLAR TEAM S.A.S" in n:
        return "SOLAR TEAM S.A.S"

    return ""


def es_proveedor_valido(linea: str) -> bool:
    n = normalizar(linea)

    excluir = [
        "RUC", "FACTURA", "NUMERO DE AUTORIZACION", "CLAVE DE ACCESO",
        "FECHA", "AMBIENTE", "EMISION", "PRODUCCION", "NORMAL",
        "DIRECCION", "OBLIGADO", "CONTRIBUYENTE", "RAZON SOCIAL",
        "IDENTIFICACION", "SOLARTEAM", "SOLAR TEAM", "SUBTOTAL",
        "TOTAL", "CODIGO", "DESCRIPCION", "CANTIDAD", "PRECIO",
    ]

    return (
        len(linea.strip()) >= 7
        and not any(palabra in n for palabra in excluir)
        and not PATRON_RUC.search(linea)
        and not PATRON_FACTURA.search(linea)
        and not PATRON_FECHA.search(linea)
        and bool(re.search(r"[A-ZÁÉÍÓÚÑ]{3,}", linea.upper()))
    )


def extraer_proveedor(texto: str, ruc_emisor: str) -> str:
    lineas = lineas_limpias(texto)

    for indice, linea in enumerate(lineas):
        if ruc_emisor and ruc_emisor in linea:
            for siguiente in range(indice + 1, min(indice + 10, len(lineas))):
                candidato = lineas[siguiente].strip()
                if es_proveedor_valido(candidato):
                    return candidato

    for linea in lineas[:45]:
        if es_proveedor_valido(linea):
            return linea.strip()

    return ""


# =========================================================
# EXTRACCIÓN DE MONTOS
# =========================================================
def extraer_monto_desde_bloques(
    bloques: list[tuple],
    etiquetas: list[str],
    excluir: list[str] | None = None,
) -> str:
    excluir = excluir or []

    for etiqueta in etiquetas:
        etiqueta_n = normalizar(etiqueta)

        for bloque in bloques:
            texto_bloque = bloque[4] if len(bloque) > 4 else ""
            bloque_n = normalizar(texto_bloque)

            if etiqueta_n not in bloque_n:
                continue

            if any(normalizar(palabra) in bloque_n for palabra in excluir):
                continue

            montos = PATRON_MONTO.findall(texto_bloque)

            if montos:
                return monto_normalizado(montos[0])

    return ""


def buscar_monto_en_texto(texto: str, patrones: list[str]) -> str:
    n = normalizar(texto)

    for patron in patrones:
        coincidencia = re.search(patron, n, re.IGNORECASE)
        if coincidencia:
            return monto_normalizado(coincidencia.group(1))

    return ""


def extraer_subtotal(texto: str, bloques: list[tuple]) -> str:
    valor = extraer_monto_desde_bloques(
        bloques,
        [
            "SUBTOTAL SIN IMPUESTOS",
            "SUBTOTAL 15%",
            "SUBTOTAL 12%",
            "BASE IMPONIBLE",
            "BASE GRAVADA",
        ],
        excluir=["NO OBJETO", "EXENTO", "DESCUENTO"],
    )

    if valor:
        return valor

    return buscar_monto_en_texto(
        texto,
        [
            r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*[:$]?\s*([0-9.,]+\d{2})",
            r"SUBTOTAL\s+15\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"SUBTOTAL\s+12\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"BASE\s+IMPONIBLE\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


def extraer_iva(texto: str, bloques: list[tuple]) -> str:
    valor = extraer_monto_desde_bloques(
        bloques,
        ["IVA 15%", "IVA 12%", "TOTAL IVA", "IMPUESTO IVA"],
        excluir=["SUBTOTAL", "NO OBJETO", "EXENTO", "INCLUYE IVA"],
    )

    if valor:
        return valor

    return buscar_monto_en_texto(
        texto,
        [
            r"\bIVA\s+15\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"\bIVA\s+12\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"TOTAL\s+IVA\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


def extraer_total(texto: str, bloques: list[tuple]) -> str:
    valor = extraer_monto_desde_bloques(
        bloques,
        [
            "VALOR TOTAL",
            "TOTAL A PAGAR",
            "IMPORTE TOTAL",
            "MONTO TOTAL",
            "TOTAL FACTURA",
        ],
        excluir=["SIN SUBSIDIO", "DESCUENTO", "SUBTOTAL", "AHORRO"],
    )

    if valor:
        return valor

    return buscar_monto_en_texto(
        texto,
        [
            r"VALOR\s+TOTAL(?!\s+SIN\s+SUBSIDIO)\s*[:$]?\s*([0-9.,]+\d{2})",
            r"TOTAL\s+A\s+PAGAR\s*[:$]?\s*([0-9.,]+\d{2})",
            r"IMPORTE\s+TOTAL\s*[:$]?\s*([0-9.,]+\d{2})",
            r"MONTO\s+TOTAL\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


# =========================================================
# VALIDACIÓN, CONFIANZA Y DUPLICADOS
# =========================================================
def calcular_confianza(datos: dict) -> int:
    campos = [
        "Fecha", "Factura", "Proveedor", "RUC Emisor",
        "Subtotal", "IVA", "Total",
    ]

    completos = sum(
        1
        for campo in campos
        if str(datos.get(campo, "")).strip()
    )

    confianza = int((completos / len(campos)) * 100)

    if datos["Tipo documento"].startswith("NO VÁLIDA"):
        confianza = min(confianza, 40)

    if datos.get("Observación"):
        confianza = max(0, confianza - 10)

    return confianza


def validar(datos: dict) -> tuple[str, str]:
    if datos["Tipo documento"].startswith("NO VÁLIDA"):
        return "REVISAR", "Documento no registrable como factura"

    obligatorios = [
        "Fecha", "Factura", "Proveedor", "RUC Emisor",
        "Subtotal", "IVA", "Total",
    ]

    faltantes = [
        campo
        for campo in obligatorios
        if not str(datos.get(campo, "")).strip()
    ]

    if faltantes:
        return "REVISAR", "Faltan: " + ", ".join(faltantes)

    if len(datos["RUC Emisor"]) != 13:
        return "REVISAR", "RUC emisor inválido"

    try:
        subtotal = float(datos["Subtotal"])
        iva = float(datos["IVA"])
        total = float(datos["Total"])

        if abs((subtotal + iva) - total) > 0.10:
            return "REVISAR", "Subtotal + IVA no coincide con el total"

    except ValueError:
        return "REVISAR", "Los valores monetarios requieren revisión"

    return "OK", ""


def extraer_datos(
    texto: str,
    bloques: list[tuple],
    nombre_archivo: str,
) -> dict:
    tipo = detectar_tipo(texto)
    ruc_emisor = extraer_ruc_emisor(texto)

    datos = {
        "Estado": "",
        "Confianza": 0,
        "Tipo documento": tipo,
        "Archivo": nombre_archivo,
        "Fecha": extraer_fecha(texto),
        "Factura": extraer_factura(texto),
        "Proveedor": extraer_proveedor(texto, ruc_emisor),
        "RUC Emisor": ruc_emisor,
        "Cliente": extraer_cliente(texto),
        "RUC Cliente": RUC_SOLARTEAM if RUC_SOLARTEAM in texto else "",
        "Subtotal": extraer_subtotal(texto, bloques),
        "IVA": extraer_iva(texto, bloques),
        "Total": extraer_total(texto, bloques),
        "Proyecto": "",
        "Consumo": "",
        "Categoría": "",
        "Observación": "",
    }

    estado, observacion = validar(datos)
    datos["Estado"] = estado
    datos["Observación"] = observacion
    datos["Confianza"] = calcular_confianza(datos)

    return datos


def marcar_duplicados(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "Factura" not in df.columns:
        return df

    duplicados = (
        df["Factura"].astype(str).str.strip().ne("")
        & df["Factura"].astype(str).duplicated(keep=False)
    )

    for indice in df.index[duplicados]:
        df.at[indice, "Estado"] = "REVISAR"
        observacion = str(df.at[indice, "Observación"] or "").strip()
        mensaje = "Factura duplicada en la carga"
        df.at[indice, "Observación"] = (
            f"{observacion}; {mensaje}".strip("; ")
        )
        df.at[indice, "Confianza"] = min(
            int(df.at[indice, "Confianza"]),
            70,
        )

    return df


# =========================================================
# OCR Y LECTURA
# =========================================================
def preparar_imagen(imagen: Image.Image) -> Image.Image:
    imagen = ImageOps.exif_transpose(imagen).convert("L")
    imagen = ImageOps.autocontrast(imagen)
    imagen = ImageEnhance.Contrast(imagen).enhance(1.8)
    imagen = imagen.filter(ImageFilter.SHARPEN)
    return imagen


def puntaje_ocr(texto: str) -> int:
    n = normalizar(texto)
    palabras = ["FACTURA", "RUC", "TOTAL", "IVA", "FECHA", "SUBTOTAL"]
    return sum(n.count(palabra) for palabra in palabras) + len(texto) // 200


def aplicar_ocr(imagen: Image.Image) -> str:
    imagen = preparar_imagen(imagen)
    resultados = []

    for angulo in [0, 90, 270, 180]:
        rotada = imagen.rotate(angulo, expand=True)
        texto = pytesseract.image_to_string(
            rotada,
            lang="spa",
            config="--oem 3 --psm 6",
        )
        resultados.append((puntaje_ocr(texto), texto))

    return max(resultados, key=lambda resultado: resultado[0])[1]


def leer_imagen(archivo) -> tuple[str, list[tuple]]:
    imagen = Image.open(BytesIO(archivo.getvalue()))
    texto = aplicar_ocr(imagen)
    return texto, []


def leer_pdf(archivo) -> tuple[str, list[tuple]]:
    texto_total = ""
    bloques_totales = []
    datos = archivo.getvalue()

    with fitz.open(stream=datos, filetype="pdf") as documento:
        for pagina in documento:
            texto_pagina = pagina.get_text("text")
            bloques_pagina = pagina.get_text("blocks")

            texto_total += texto_pagina + "\n"
            bloques_totales.extend(bloques_pagina)

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

                texto_total += aplicar_ocr(imagen) + "\n"

    return texto_total, bloques_totales


def vista_previa_archivo(archivo):
    nombre = archivo.name.lower()

    if nombre.endswith((".jpg", ".jpeg", ".png")):
        imagen = Image.open(BytesIO(archivo.getvalue()))
        st.image(imagen, use_container_width=True)
        return

    if nombre.endswith(".pdf"):
        with fitz.open(stream=archivo.getvalue(), filetype="pdf") as documento:
            pagina = documento[0]
            pixmap = pagina.get_pixmap(
                matrix=fitz.Matrix(1.4, 1.4),
                alpha=False,
            )
            imagen = Image.frombytes(
                "RGB",
                [pixmap.width, pixmap.height],
                pixmap.samples,
            )
            st.image(imagen, use_container_width=True)


# =========================================================
# EXCEL PROFESIONAL
# =========================================================
def crear_excel(df: pd.DataFrame) -> bytes:
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Facturas",
        )

    salida.seek(0)
    libro = load_workbook(salida)
    hoja = libro["Facturas"]

    color_azul = "172B4D"
    color_verde = "1F9D68"

    for celda in hoja[1]:
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor=color_azul,
        )
        celda.font = Font(
            color="FFFFFF",
            bold=True,
        )
        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    hoja.row_dimensions[1].height = 24

    for columna in hoja.columns:
        ancho = min(
            max(len(str(celda.value or "")) for celda in columna) + 2,
            45,
        )
        hoja.column_dimensions[columna[0].column_letter].width = ancho

    if hoja.max_row >= 2 and hoja.max_column >= 1:
        referencia = f"A1:{hoja.cell(hoja.max_row, hoja.max_column).coordinate}"
        tabla = Table(
            displayName="TablaFacturas",
            ref=referencia,
        )
        estilo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        tabla.tableStyleInfo = estilo
        hoja.add_table(tabla)

    salida_final = BytesIO()
    libro.save(salida_final)
    return salida_final.getvalue()


# =========================================================
# ENCABEZADO
# =========================================================
col_logo, col_titulo = st.columns(
    [1.2, 5.8],
    vertical_alignment="center",
)

with col_logo:
    logo = next(
        (ruta for ruta in LOGO_CANDIDATES if ruta.exists()),
        None,
    )

    if logo:
        st.image(
            str(logo),
            use_container_width=True,
        )

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


# =========================================================
# CARGA DE ARCHIVOS
# =========================================================
st.markdown(
    '<div class="section-title">📂 Sube tus facturas</div>',
    unsafe_allow_html=True,
)

archivos = st.file_uploader(
    "Arrastra aquí tus archivos o selecciónalos",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    help="Formatos permitidos: PDF, JPG, JPEG y PNG",
)


# =========================================================
# PROCESAMIENTO
# =========================================================
if archivos:
    resultados = []
    barra = st.progress(0, text="Preparando documentos...")

    for indice, archivo in enumerate(archivos, start=1):
        try:
            nombre = archivo.name.lower()

            if nombre.endswith(".pdf"):
                texto, bloques = leer_pdf(archivo)
            else:
                texto, bloques = leer_imagen(archivo)

            resultados.append(
                extraer_datos(
                    texto,
                    bloques,
                    archivo.name,
                )
            )

        except Exception as error:
            resultados.append(
                {
                    "Estado": "ERROR",
                    "Confianza": 0,
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
    df = marcar_duplicados(df)

    cantidad_ok = int((df["Estado"] == "OK").sum())
    cantidad_revisar = int((df["Estado"] == "REVISAR").sum())
    cantidad_error = int((df["Estado"] == "ERROR").sum())

    total_facturado = 0.0
    for valor in df["Total"].tolist():
        if es_monto_valido(valor):
            total_facturado += float(valor)

    st.markdown(
        '<div class="section-title">📊 Resumen del procesamiento</div>',
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">✅ Facturas OK</div>
                <div class="summary-value ok-value">{cantidad_ok}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">⚠️ Para revisar</div>
                <div class="summary-value review-value">{cantidad_revisar}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">❌ Errores</div>
                <div class="summary-value error-value">{cantidad_error}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c4:
        st.markdown(
            f"""
            <div class="summary-card">
                <div class="summary-label">💰 Total detectado</div>
                <div class="summary-value total-value">${total_facturado:,.2f}</div>
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
        height=min(640, 140 + len(df) * 42),
        disabled=["Archivo"],
        column_config={
            "Estado": st.column_config.SelectboxColumn(
                "Estado",
                options=["OK", "REVISAR", "ERROR"],
                width="small",
            ),
            "Confianza": st.column_config.ProgressColumn(
                "Confianza",
                min_value=0,
                max_value=100,
                format="%d%%",
                width="small",
            ),
            "Archivo": st.column_config.TextColumn(width="large"),
            "Factura": st.column_config.TextColumn(width="medium"),
            "Proveedor": st.column_config.TextColumn(width="large"),
            "Observación": st.column_config.TextColumn(width="large"),
        },
    )

    nombre_excel = (
        "FACTURAS_"
        + datetime.now().strftime("%Y-%m-%d_%H-%M")
        + ".xlsx"
    )

    st.download_button(
        "📥 Descargar Excel",
        data=crear_excel(df_editado),
        file_name=nombre_excel,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("👁️ Vista previa de los documentos", expanded=False):
        archivo_vista = st.selectbox(
            "Selecciona un archivo",
            options=[archivo.name for archivo in archivos],
        )

        archivo_seleccionado = next(
            archivo
            for archivo in archivos
            if archivo.name == archivo_vista
        )

        vista_previa_archivo(archivo_seleccionado)

else:
    st.info(
        "Carga uno o varios archivos PDF, JPG, JPEG o PNG para comenzar."
    )


st.markdown(
    """
    <div class="footer">
        <b>Solar Team</b><br>
        Automatización de información de facturas · Versión 1.0
    </div>
    """,
    unsafe_allow_html=True,
)
