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


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Automatizador de Facturas",
    page_icon="📄",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
RUC_SOLARTEAM = "1793069479001"

LOGO_CANDIDATES = [
    BASE_DIR / "logo_solarteam.png",
    BASE_DIR / "LOGO SOLAR TEAM.png",
    BASE_DIR / "logo.png",
]

# Ruta local de Tesseract en Windows.
# En Streamlit Cloud se instala mediante packages.txt.
if os.name == "nt":
    tesseract_windows = Path(
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

    if tesseract_windows.exists():
        pytesseract.pytesseract.tesseract_cmd = str(
            tesseract_windows
        )


# =========================================================
# ESTILOS
# =========================================================
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

    .footer {
        margin-top: 2rem;
        padding: 20px 24px;
        border-radius: 14px;
        background: #F8FAFC;
        color: #64748B;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# PATRONES GENERALES
# =========================================================
PATRON_FACTURA = re.compile(
    r"\b\d{3}-\d{3}-\d{9}\b"
)

PATRON_RUC = re.compile(
    r"\b\d{13}\b"
)

PATRON_FECHA = re.compile(
    r"\b(?:"
    r"\d{2}[/-]\d{2}[/-]\d{4}"
    r"|"
    r"\d{4}[/-]\d{2}[/-]\d{2}"
    r")\b"
)


# =========================================================
# NORMALIZACIÓN
# =========================================================
def normalizar(texto: str) -> str:
    texto = unicodedata.normalize(
        "NFKD",
        texto or "",
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = texto.upper()
    texto = texto.replace("\xa0", " ")
    texto = re.sub(r"[ \t]+", " ", texto)
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

    valor = valor.replace("$", "")
    valor = valor.replace(" ", "")

    if "," in valor and "." in valor:
        if valor.rfind(",") > valor.rfind("."):
            valor = valor.replace(".", "")
            valor = valor.replace(",", ".")
        else:
            valor = valor.replace(",", "")
    else:
        valor = valor.replace(",", ".")

    try:
        return f"{float(valor):.2f}"

    except ValueError:
        return valor


# =========================================================
# EXTRACCIÓN DE MONTOS
# =========================================================
def buscar_monto_en_texto(
    texto: str,
    patrones: list[str],
) -> str:
    texto_busqueda = normalizar(texto)

    for patron in patrones:
        coincidencia = re.search(
            patron,
            texto_busqueda,
            re.IGNORECASE,
        )

        if coincidencia:
            return monto_normalizado(
                coincidencia.group(1)
            )

    return ""


def extraer_subtotal(texto: str) -> str:
    return buscar_monto_en_texto(
        texto,
        [
            r"SUBTOTAL\s+SIN\s+IMPUESTOS\s*[:$]?\s*([0-9.,]+\d{2})",
            r"SUBTOTAL\s+15\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"SUBTOTAL\s+12\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"BASE\s+IMPONIBLE\s*[:$]?\s*([0-9.,]+\d{2})",
            r"BASE\s+GRAVADA\s*[:$]?\s*([0-9.,]+\d{2})",
            r"\bSUBTOTAL\b"
            r"(?!\s+(?:NO\s+OBJETO|EXENTO|DESCUENTO))"
            r"\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


def extraer_iva(texto: str) -> str:
    return buscar_monto_en_texto(
        texto,
        [
            r"\bIVA\s+15\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"\bIVA\s+12\s*%\s*[:$]?\s*([0-9.,]+\d{2})",
            r"TOTAL\s+IVA\s*[:$]?\s*([0-9.,]+\d{2})",
            r"IMPUESTO\s+IVA\s*[:$]?\s*([0-9.,]+\d{2})",
            r"IMPUESTO\s+AL\s+VALOR\s+AGREGADO"
            r"\s*[:$]?\s*([0-9.,]+\d{2})",
            r"\bIVA\b"
            r"(?!\s+(?:CUANDO|INCLUIDO|INCLUYE))"
            r"\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


def extraer_total(texto: str) -> str:
    return buscar_monto_en_texto(
        texto,
        [
            r"VALOR\s+TOTAL"
            r"(?!\s+SIN\s+SUBSIDIO)"
            r"\s*[:$]?\s*([0-9.,]+\d{2})",
            r"TOTAL\s+A\s+PAGAR\s*[:$]?\s*([0-9.,]+\d{2})",
            r"IMPORTE\s+TOTAL\s*[:$]?\s*([0-9.,]+\d{2})",
            r"MONTO\s+TOTAL\s*[:$]?\s*([0-9.,]+\d{2})",
            r"TOTAL\s+FACTURA\s*[:$]?\s*([0-9.,]+\d{2})",
            r"\bTOTAL\b"
            r"(?!\s+(?:DESCUENTO|SIN\s+SUBSIDIO|NO\s+OBJETO|EXENTO))"
            r"\s*[:$]?\s*([0-9.,]+\d{2})",
        ],
    )


# =========================================================
# TIPO DE DOCUMENTO
# =========================================================
def detectar_tipo(texto: str) -> str:
    texto_normalizado = normalizar(texto)

    texto_compacto = re.sub(
        r"[^A-Z]",
        "",
        texto_normalizado,
    )

    if (
        "PROFORMA" in texto_normalizado
        or "PROFORMA" in texto_compacto
    ):
        return "NO VÁLIDA - PROFORMA"

    if "COTIZACION" in texto_normalizado:
        return "NO VÁLIDA - COTIZACIÓN"

    if "NOTA DE CREDITO" in texto_normalizado:
        return "NO VÁLIDA - NOTA DE CRÉDITO"

    if "NOTA DE VENTA" in texto_normalizado:
        return "REVISAR - NOTA DE VENTA"

    if (
        "FACTURA" in texto_normalizado
        or "FACTURA" in texto_compacto
    ):
        if (
            "NUMERO DE AUTORIZACION"
            in texto_normalizado
            or "CLAVE DE ACCESO"
            in texto_normalizado
        ):
            return "FACTURA ELECTRÓNICA"

        return "FACTURA"

    return "REVISAR - NO IDENTIFICADO"


# =========================================================
# CAMPOS PRINCIPALES
# =========================================================
def extraer_factura(texto: str) -> str:
    coincidencias = PATRON_FACTURA.findall(texto)

    if coincidencias:
        return coincidencias[0]

    return ""


def extraer_fecha(
    texto: str,
    lineas: list[str],
) -> str:
    for indice, linea in enumerate(lineas):
        linea_normalizada = normalizar(linea)

        es_fecha_emision = (
            "FECHA EMISION" in linea_normalizada
            or "FECHA DE EMISION" in linea_normalizada
            or linea_normalizada == "FECHA"
            or linea_normalizada.startswith("FECHA ")
        )

        if not es_fecha_emision:
            continue

        fecha = PATRON_FECHA.search(linea)

        if fecha:
            return fecha.group(0)

        for siguiente in range(
            indice + 1,
            min(indice + 4, len(lineas)),
        ):
            fecha = PATRON_FECHA.search(
                lineas[siguiente]
            )

            if fecha:
                return fecha.group(0)

    fechas = PATRON_FECHA.findall(texto)

    if fechas:
        return fechas[-1]

    return ""


def extraer_ruc_emisor(texto: str) -> str:
    rucs = PATRON_RUC.findall(texto)

    if not rucs:
        return ""

    for ruc in rucs:
        if ruc != RUC_SOLARTEAM:
            return ruc

    return rucs[0]


def extraer_cliente(texto: str) -> str:
    texto_normalizado = normalizar(texto)

    if "SOLARTEAM SAS" in texto_normalizado:
        return "SOLARTEAM SAS"

    if "SOLAR TEAM S.A.S" in texto_normalizado:
        return "SOLAR TEAM S.A.S"

    return ""


def es_linea_proveedor_valida(
    linea: str,
) -> bool:
    linea_normalizada = normalizar(linea)

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
        "CODIGO",
        "DESCRIPCION",
        "CANTIDAD",
        "PRECIO",
    ]

    tiene_etiqueta_descartada = any(
        etiqueta in linea_normalizada
        for etiqueta in etiquetas_descartadas
    )

    return (
        len(linea.strip()) >= 7
        and not tiene_etiqueta_descartada
        and not PATRON_RUC.search(linea)
        and not PATRON_FACTURA.search(linea)
        and not PATRON_FECHA.search(linea)
        and bool(
            re.search(
                r"[A-ZÁÉÍÓÚÑ]{3,}",
                linea.upper(),
            )
        )
    )


def extraer_proveedor(
    lineas: list[str],
    ruc_emisor: str,
) -> str:
    for indice, linea in enumerate(lineas):
        if not ruc_emisor:
            continue

        if ruc_emisor not in linea:
            continue

        for siguiente in range(
            indice + 1,
            min(indice + 10, len(lineas)),
        ):
            candidato = lineas[siguiente].strip()

            if es_linea_proveedor_valida(
                candidato
            ):
                return candidato

    for linea in lineas[:45]:
        if es_linea_proveedor_valida(linea):
            return linea.strip()

    return ""


# =========================================================
# VALIDACIÓN
# =========================================================
def validar(
    datos: dict,
) -> tuple[str, str]:
    if datos["Tipo documento"].startswith(
        "NO VÁLIDA"
    ):
        return (
            "REVISAR",
            "Documento no registrable como factura",
        )

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
        if not str(
            datos.get(campo, "")
        ).strip()
    ]

    if faltantes:
        return (
            "REVISAR",
            "Faltan: " + ", ".join(faltantes),
        )

    try:
        subtotal = float(datos["Subtotal"])
        iva = float(datos["IVA"])
        total = float(datos["Total"])

        diferencia = abs(
            (subtotal + iva) - total
        )

        if diferencia > 0.10:
            return (
                "REVISAR",
                "Subtotal + IVA no coincide con el total",
            )

    except ValueError:
        return (
            "REVISAR",
            "Los valores monetarios requieren revisión",
        )

    return "OK", ""


def extraer_datos(
    texto: str,
    nombre_archivo: str,
) -> dict:
    lineas = lineas_limpias(texto)
    tipo = detectar_tipo(texto)
    ruc_emisor = extraer_ruc_emisor(texto)

    datos = {
        "Estado": "",
        "Tipo documento": tipo,
        "Archivo": nombre_archivo,
        "Fecha": extraer_fecha(
            texto,
            lineas,
        ),
        "Factura": extraer_factura(texto),
        "Proveedor": extraer_proveedor(
            lineas,
            ruc_emisor,
        ),
        "RUC Emisor": ruc_emisor,
        "Cliente": extraer_cliente(texto),
        "RUC Cliente": (
            RUC_SOLARTEAM
            if RUC_SOLARTEAM in texto
            else ""
        ),
        "Subtotal": extraer_subtotal(texto),
        "IVA": extraer_iva(texto),
        "Total": extraer_total(texto),
        "Proyecto": "",
        "Consumo": "",
        "Categoría": "",
        "Observación": "",
    }

    estado, observacion = validar(datos)

    datos["Estado"] = estado
    datos["Observación"] = observacion

    return datos


# =========================================================
# OCR Y LECTURA
# =========================================================
def preparar_imagen(
    imagen: Image.Image,
) -> Image.Image:
    imagen = ImageOps.exif_transpose(imagen)
    imagen = imagen.convert("L")
    imagen = ImageOps.autocontrast(imagen)

    imagen = ImageEnhance.Contrast(
        imagen
    ).enhance(1.8)

    imagen = imagen.filter(
        ImageFilter.SHARPEN
    )

    return imagen


def puntaje_texto_ocr(
    texto: str,
) -> int:
    texto_normalizado = normalizar(texto)

    palabras = [
        "FACTURA",
        "RUC",
        "TOTAL",
        "IVA",
        "FECHA",
        "SUBTOTAL",
    ]

    puntaje_palabras = sum(
        texto_normalizado.count(palabra)
        for palabra in palabras
    )

    puntaje_longitud = len(texto) // 200

    return puntaje_palabras + puntaje_longitud


def aplicar_ocr(
    imagen: Image.Image,
) -> str:
    imagen = preparar_imagen(imagen)
    resultados = []

    for angulo in [0, 90, 270, 180]:
        rotada = imagen.rotate(
            angulo,
            expand=True,
        )

        texto = pytesseract.image_to_string(
            rotada,
            lang="spa",
            config="--oem 3 --psm 6",
        )

        puntaje = puntaje_texto_ocr(texto)

        resultados.append(
            (puntaje, texto)
        )

    mejor_resultado = max(
        resultados,
        key=lambda item: item[0],
    )

    return mejor_resultado[1]


def leer_imagen(archivo) -> str:
    imagen = Image.open(
        BytesIO(
            archivo.getvalue()
        )
    )

    return aplicar_ocr(imagen)


def leer_pdf(archivo) -> str:
    texto = ""
    datos = archivo.getvalue()

    with fitz.open(
        stream=datos,
        filetype="pdf",
    ) as documento:

        for pagina in documento:
            texto_pagina = pagina.get_text(
                "text"
            )

            texto += texto_pagina + "\n"

            if len(texto_pagina.strip()) < 40:
                pixmap = pagina.get_pixmap(
                    matrix=fitz.Matrix(
                        2.4,
                        2.4,
                    ),
                    alpha=False,
                )

                imagen = Image.frombytes(
                    "RGB",
                    [
                        pixmap.width,
                        pixmap.height,
                    ],
                    pixmap.samples,
                )

                texto += aplicar_ocr(imagen) + "\n"

    return texto


# =========================================================
# EXCEL
# =========================================================
def crear_excel(
    df: pd.DataFrame,
) -> bytes:
    salida = BytesIO()

    with pd.ExcelWriter(
        salida,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Facturas",
        )

        hoja = writer.sheets["Facturas"]

        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        for columna in hoja.columns:
            ancho = min(
                max(
                    len(
                        str(
                            celda.value or ""
                        )
                    )
                    for celda in columna
                ) + 2,
                45,
            )

            hoja.column_dimensions[
                columna[0].column_letter
            ].width = ancho

    return salida.getvalue()


# =========================================================
# ENCABEZADO
# =========================================================
col_logo, col_titulo = st.columns(
    [1.25, 5.75],
    vertical_alignment="center",
)

with col_logo:
    logo = next(
        (
            ruta
            for ruta in LOGO_CANDIDATES
            if ruta.exists()
        ),
        None,
    )

    if logo:
        st.image(
            str(logo),
            use_container_width=True,
        )

with col_titulo:
    st.markdown(
        """
        <div class="main-title">
            Automatizador de Facturas
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="subtitle">
            Carga facturas PDF o imágenes,
            revisa la información extraída
            y genera un Excel listo para descargar.
        </div>
        """,
        unsafe_allow_html=True,
    )

st.divider()

st.markdown(
    """
    <div class="section-title">
        📂 Sube tus facturas
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# CARGA Y PROCESAMIENTO
# =========================================================
archivos = st.file_uploader(
    "Sube facturas PDF o imágenes",
    type=[
        "pdf",
        "jpg",
        "jpeg",
        "png",
    ],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if archivos:
    resultados = []

    barra = st.progress(
        0,
        text="Preparando documentos...",
    )

    for indice, archivo in enumerate(
        archivos,
        start=1,
    ):
        try:
            nombre = archivo.name.lower()

            if nombre.endswith(".pdf"):
                texto = leer_pdf(archivo)
            else:
                texto = leer_imagen(archivo)

            resultados.append(
                extraer_datos(
                    texto,
                    archivo.name,
                )
            )

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
            text=(
                f"Procesando {indice} "
                f"de {len(archivos)} documentos..."
            ),
        )

    barra.empty()

    df = pd.DataFrame(resultados)

    cantidad_ok = int(
        (df["Estado"] == "OK").sum()
    )

    cantidad_revisar = int(
        (df["Estado"] == "REVISAR").sum()
    )

    cantidad_error = int(
        (df["Estado"] == "ERROR").sum()
    )

    st.markdown(
        """
        <div class="section-title">
            📊 Resumen del procesamiento
        </div>
        """,
        unsafe_allow_html=True,
    )

    columna_ok, columna_revisar, columna_error = st.columns(
        3
    )

    with columna_ok:
        st.metric(
            label="✅ Facturas OK",
            value=cantidad_ok,
        )

    with columna_revisar:
        st.metric(
            label="⚠️ Para revisar",
            value=cantidad_revisar,
        )

    with columna_error:
        st.metric(
            label="❌ Errores",
            value=cantidad_error,
        )

    st.markdown(
        """
        <div class="section-title">
            📋 Revisa y corrige antes de descargar
        </div>
        """,
        unsafe_allow_html=True,
    )

    df_editado = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        height=min(
            620,
            130 + len(df) * 40,
        ),
        disabled=[
            "Archivo",
        ],
        column_config={
            "Estado": st.column_config.SelectboxColumn(
                "Estado",
                options=[
                    "OK",
                    "REVISAR",
                    "ERROR",
                ],
                width="small",
            ),
            "Tipo documento": st.column_config.TextColumn(
                width="medium",
            ),
            "Archivo": st.column_config.TextColumn(
                width="large",
            ),
            "Fecha": st.column_config.TextColumn(
                width="small",
            ),
            "Factura": st.column_config.TextColumn(
                width="medium",
            ),
            "Proveedor": st.column_config.TextColumn(
                width="large",
            ),
            "RUC Emisor": st.column_config.TextColumn(
                width="medium",
            ),
            "Cliente": st.column_config.TextColumn(
                width="medium",
            ),
            "RUC Cliente": st.column_config.TextColumn(
                width="medium",
            ),
            "Subtotal": st.column_config.TextColumn(
                width="small",
            ),
            "IVA": st.column_config.TextColumn(
                width="small",
            ),
            "Total": st.column_config.TextColumn(
                width="small",
            ),
            "Observación": st.column_config.TextColumn(
                width="large",
            ),
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


# =========================================================
# PIE DE PÁGINA
# =========================================================
st.markdown(
    """
    <div class="footer">
        <b>Solar Team</b><br>
        Automatización de información de facturas.
    </div>
    """,
    unsafe_allow_html=True,
)
