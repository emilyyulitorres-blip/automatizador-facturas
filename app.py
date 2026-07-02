import pandas as pd
import streamlit as st

from modules.excel_manager import create_excel_download, save_to_database
from modules.extractor import extract_data
from modules.readers import read_file

st.set_page_config(
    page_title="Automatizador de Facturas",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.6rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        color: #667085;
        font-size: 1.05rem;
        margin-bottom: 1.2rem;
    }
    .status-ok {
        background:#ECFDF3;
        color:#027A48;
        padding:6px 10px;
        border-radius:20px;
        font-weight:700;
    }
    .status-review {
        background:#FFFAEB;
        color:#B54708;
        padding:6px 10px;
        border-radius:20px;
        font-weight:700;
    }
    .status-error {
        background:#FEF3F2;
        color:#B42318;
        padding:6px 10px;
        border-radius:20px;
        font-weight:700;
    }
    .small-note {
        color:#667085;
        font-size:0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">📄 Automatizador de Facturas</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Sube PDF o imágenes, revisa los datos extraídos y descarga/guarda el Excel.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("⚙️ Configuración")
    st.write("Campos manuales para completar antes de guardar:")
    proyecto_default = st.text_input("Proyecto predeterminado", value="")
    consumo_default = st.text_input("Consumo predeterminado", value="")
    categoria_default = st.text_input("Categoría predeterminada", value="")

    st.divider()
    guardar_proformas = st.checkbox("Permitir guardar proformas", value=False)
    mostrar_texto = st.checkbox("Mostrar texto extraído para diagnóstico", value=False)

    st.divider()
    st.markdown("**Estados**")
    st.markdown("🟢 OK: factura detectada")
    st.markdown("🟡 REVISAR: proforma o datos incompletos")
    st.markdown("🔴 ERROR: no se pudo leer")

files = st.file_uploader(
    "Arrastra aquí tus facturas PDF o imágenes",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if not files:
    c1, c2, c3 = st.columns(3)
    c1.info("1. Sube PDF o imágenes")
    c2.info("2. Revisa/corrige la tabla")
    c3.info("3. Descarga o guarda Excel")
    st.stop()

rows = []
debug_texts = {}
progress = st.progress(0, text="Preparando lectura...")

for i, file in enumerate(files, start=1):
    progress.progress(i / len(files), text=f"Procesando {i} de {len(files)}: {file.name}")
    try:
        text, read_note = read_file(file)
        debug_texts[file.name] = text
        row = extract_data(text, file.name)

        if proyecto_default:
            row["Proyecto"] = proyecto_default
        if consumo_default:
            row["Consumo"] = consumo_default
        if categoria_default:
            row["Categoría"] = categoria_default

        if read_note:
            row["Observación"] = (row.get("Observación", "") + " | " + read_note).strip(" |")

        if not text.strip():
            row["Estado"] = "REVISAR"
            row["Tipo documento"] = "SIN TEXTO"
            row["Observación"] = "No se extrajo texto. Revisar calidad del archivo o instalar OCR."

        if row["Tipo documento"] == "NO VÁLIDA - PROFORMA" and not guardar_proformas:
            row["Estado"] = "REVISAR"
            row["Observación"] = "Proforma detectada. No registrar como factura final sin validación."

        rows.append(row)

    except Exception as e:
        rows.append({
            "Estado": "ERROR",
            "Tipo documento": "ERROR",
            "Archivo": file.name,
            "Fecha": "",
            "Factura": "",
            "Proveedor": "",
            "RUC Emisor": "",
            "Subtotal": "",
            "IVA": "",
            "Total": "",
            "Proyecto": proyecto_default,
            "Consumo": consumo_default,
            "Categoría": categoria_default,
            "Observación": str(e),
            "Fecha registro": "",
        })

progress.empty()

df = pd.DataFrame(rows)

# Ensure stable column order
columns = [
    "Estado", "Tipo documento", "Archivo", "Fecha", "Factura", "Proveedor", "RUC Emisor",
    "Subtotal", "IVA", "Total", "Proyecto", "Consumo", "Categoría", "Observación", "Fecha registro"
]
for col in columns:
    if col not in df.columns:
        df[col] = ""
df = df[columns]

ok_count = int((df["Estado"] == "OK").sum())
review_count = int((df["Estado"] == "REVISAR").sum())
error_count = int((df["Estado"] == "ERROR").sum())

m1, m2, m3, m4 = st.columns(4)
m1.metric("Archivos procesados", len(df))
m2.metric("OK", ok_count)
m3.metric("Por revisar", review_count)
m4.metric("Errores", error_count)

st.subheader("📋 Revisa y corrige antes de descargar")
st.caption("Puedes editar cualquier celda antes de generar el Excel.")

edited = st.data_editor(
    df,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Estado": st.column_config.SelectboxColumn("Estado", options=["OK", "REVISAR", "ERROR"]),
        "Tipo documento": st.column_config.TextColumn("Tipo documento"),
        "Subtotal": st.column_config.TextColumn("Subtotal"),
        "IVA": st.column_config.TextColumn("IVA"),
        "Total": st.column_config.TextColumn("Total"),
        "Observación": st.column_config.TextColumn("Observación", width="large"),
    },
)

col_a, col_b, col_c = st.columns([1, 1, 2])

with col_a:
    excel_bytes = create_excel_download(edited.copy())
    st.download_button(
        "⬇️ Descargar Excel",
        data=excel_bytes,
        file_name="facturas_extraidas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_b:
    if st.button("💾 Guardar en base local", use_container_width=True):
        path = save_to_database(edited.copy())
        st.success(f"Guardado en {path}")

with col_c:
    st.markdown('<span class="small-note">Para la demo, usa Descargar Excel. La base local sirve cuando la app corre en una PC/servidor propio.</span>', unsafe_allow_html=True)

if mostrar_texto:
    st.subheader("🔎 Texto extraído")
    for name, text in debug_texts.items():
        with st.expander(name):
            st.text_area("Texto", text[:8000], height=250)
