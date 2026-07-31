from datetime import datetime

import pandas as pd
import streamlit as st

from modules import (
    aplicar_memoria,
    cargar_memoria,
    guardar_facturas,
    obtener_supabase,
)
from modules.excel_export import crear_excel
from modules.extractor import extraer_datos, marcar_duplicados
from modules.readers import leer_imagen, leer_pdf, vista_previa
from modules.ui import aplicar_estilos, mostrar_encabezado, pie_pagina, titulo_seccion


st.set_page_config(
    page_title="Automatizador de Facturas",
    page_icon="📄",
    layout="wide",
)

aplicar_estilos()
mostrar_encabezado()

supabase = obtener_supabase()
memoria = cargar_memoria(supabase)

if supabase is None:
    st.warning(
        "La extracción y el Excel funcionan, pero no se pudo conectar "
        "con Supabase. Revisa los Secrets de Streamlit."
    )

titulo_seccion("📂 Sube tus facturas")

archivos = st.file_uploader(
    "Arrastra aquí tus archivos o selecciónalos",
    type=["pdf", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
    help="Formatos permitidos: PDF, JPG, JPEG y PNG",
)

if archivos:
    resultados = []
    barra = st.progress(0, text="Preparando documentos...")

    for indice, archivo in enumerate(archivos, start=1):
        try:
            if archivo.name.lower().endswith(".pdf"):
                texto, bloques = leer_pdf(archivo)
            else:
                texto, bloques = leer_imagen(archivo)

            datos = extraer_datos(texto, bloques, archivo.name)
            resultados.append(aplicar_memoria(datos, memoria))

        except Exception as error:
            resultados.append({
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
            })

        barra.progress(
            indice / len(archivos),
            text=f"Procesando {indice} de {len(archivos)} documentos...",
        )

    barra.empty()

    df = marcar_duplicados(pd.DataFrame(resultados))
    df_original = df.copy(deep=True)

    titulo_seccion("📊 Resumen del procesamiento")
    c1, c2, c3, c4 = st.columns(4)

    c1.metric("✅ Facturas OK", int((df["Estado"] == "OK").sum()))
    c2.metric("⚠️ Para revisar", int((df["Estado"] == "REVISAR").sum()))
    c3.metric("❌ Errores", int((df["Estado"] == "ERROR").sum()))

    total_detectado = pd.to_numeric(df["Total"], errors="coerce").fillna(0).sum()
    c4.metric("💰 Total detectado", f"${total_detectado:,.2f}")

    titulo_seccion("📋 Revisa y corrige antes de descargar")

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
            "Proyecto": st.column_config.TextColumn(
                "Proyecto",
                width="medium",
                help=(
                    "Escribe el proyecto o déjalo vacío. "
                    "Los proyectos nuevos se crean al guardar."
                ),
            ),
            "Observación": st.column_config.TextColumn(width="large"),
        },
    )

    nombre_excel = "FACTURAS_" + datetime.now().strftime("%Y-%m-%d_%H-%M") + ".xlsx"

    col_excel, col_guardar = st.columns(2)

    with col_excel:
        st.download_button(
            "📥 Descargar Excel",
            data=crear_excel(df_editado),
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_guardar:
        guardar = st.button(
            "💾 Guardar en base de datos",
            type="primary",
            use_container_width=True,
            disabled=(supabase is None),
        )

    if guardar and supabase is not None:
        with st.spinner("Guardando facturas y correcciones..."):
            guardadas, duplicadas, errores = guardar_facturas(
                supabase,
                df_original,
                df_editado,
            )

        if guardadas:
            st.success(f"Se guardaron {guardadas} factura(s) correctamente.")
        if duplicadas:
            st.warning(f"{duplicadas} factura(s) ya existían y no se duplicaron.")
        if errores:
            st.error(
                "No se guardaron algunos registros:\n\n"
                + "\n".join(f"- {error}" for error in errores)
            )

    with st.expander("👁️ Vista previa de los documentos"):
        nombre_vista = st.selectbox(
            "Selecciona un archivo",
            [archivo.name for archivo in archivos],
        )
        seleccionado = next(
            archivo for archivo in archivos if archivo.name == nombre_vista
        )
        st.image(vista_previa(seleccionado), use_container_width=True)

else:
    st.info("Carga uno o varios archivos PDF, JPG, JPEG o PNG para comenzar.")

pie_pagina()
