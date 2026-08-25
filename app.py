import datetime
import io
import uuid
import numpy as np
import pandas as pd
import streamlit as st

# Verificación de librerías para exportar a Excel sin romper la app
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Configuración de la página
st.set_page_config(page_title="Sistema de Control y Análisis de Obra", layout="wide", page_icon="🏗️")

st.title("🏗️ Sistema de Control y Análisis de Obra")
st.caption("Visualización, consolidado de archivos y resguardo de información de ingeniería civil.")

# -----------------------------------------------------------------------------
# MODOS DE OPERACIÓN
# -----------------------------------------------------------------------------
modo_analisis = st.radio(
    "📌 Selecciona el modo de trabajo:",
    ["📄 Analizar un solo archivo", "📂 Analizar y consolidar varios archivos"],
    horizontal=True
)

lista_dfs = []

if modo_analisis == "📄 Analizar un solo archivo":
    origen_fuente = st.radio("Fuente del archivo:", ["Subir Archivo (CSV / Excel)", "📷 Tomar Foto (Cámara / Captura)"], horizontal=True)
    
    if origen_fuente == "Subir Archivo (CSV / Excel)":
        archivo_subido = st.file_uploader("Selecciona tu archivo", type=["csv", "xlsx", "xls"], key="single_file")
        if archivo_subido:
            try:
                if archivo_subido.name.endswith('.csv'):
                    df = pd.read_csv(archivo_subido)
                else:
                    df = pd.read_excel(archivo_subido)
                lista_dfs.append(df)
            except Exception as e:
                st.error(f"Error al leer el archivo: {e}")
                
    else:
        foto_captura = st.camera_input("Toma una captura del documento o ticket")
        if foto_captura:
            st.warning("⚠️ Captura recibida. (El reconocimiento óptico de texto OCR requeriría un modelo adicional). Puedes adjuntar manualmente los datos o cargar el archivo digital.")

else:
    archivos_multiples = st.file_uploader(
        "Selecciona todos los archivos CSV o Excel a consolidar (ej. Reportes Semanales)", 
        type=["csv", "xlsx", "xls"], 
        accept_multiple_files=True,
        key="multi_files"
    )
    if archivos_multiples:
        for arch in archivos_multiples:
            try:
                if arch.name.endswith('.csv'):
                    df_temp = pd.read_csv(arch)
                else:
                    df_temp = pd.read_excel(arch)
                df_temp['Origen_Archivo'] = arch.name
                lista_dfs.append(df_temp)
            except Exception as e:
                st.error(f"Error al leer el archivo {arch.name}: {e}")

# -----------------------------------------------------------------------------
# PROCESAMIENTO Y DASHBOARD
# -----------------------------------------------------------------------------
if lista_dfs:
    # Combinar dataframes si es consolidado
    df_consolidado = pd.concat(lista_dfs, ignore_index=True)
    
    st.divider()
    st.subheader("📌 Métricas Generales del Dataset")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Registros", len(df_consolidado))
    m2.metric("Total de Columnas", len(df_consolidado.columns))
    
    col_numericas = df_consolidado.select_dtypes(include=[np.number]).columns.tolist()
    col_categoricas = df_consolidado.select_dtypes(include=['object', 'category']).columns.tolist()
    
    m3.metric("Columnas Numéricas", len(col_numericas))

    # -------------------------------------------------------------------------
    # SECCIÓN DE FILTROS DINÁMICOS
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("🔍 Filtros de Consulta")
    df_filtrado = df_consolidado.copy()
    
    if col_categoricas:
        col_filtro = st.selectbox("Selecciona columna para filtrar:", ["(Sin filtro)"] + col_categoricas)
        if col_filtro != "(Sin filtro)":
            opciones = df_consolidado[col_filtro].dropna().unique().tolist()
            seleccion = st.multiselect(f"Filtrar por {col_filtro}:", opciones)
            if seleccion:
                df_filtrado = df_consolidado[df_consolidado[col_filtro].isin(seleccion)]

    # -------------------------------------------------------------------------
    # TABLA DE DATOS
    # -------------------------------------------------------------------------
    st.subheader("📋 Tabla de Datos Procesados")
    st.dataframe(df_filtrado, use_container_width=True)

    # -------------------------------------------------------------------------
    # SECCIÓN DE DASHBOARD Y GRÁFICAS AUTOMÁTICAS
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📊 Dashboard de Gráficas de Rendimiento y Costos")
    
    if len(col_numericas) > 0:
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Gráfico de Barras", "🍕 Gráfico de Pastel", "📈 Líneas / Tendencia", "📍 Dispersión"])
        
        with tab1:
            st.markdown("**Comparativa por Categorías (Barras)**")
            c_eje_x = st.selectbox("Selecciona Eje X (Categoría):", col_categoricas if col_categoricas else df_filtrado.columns, key="bar_x")
            c_eje_y = st.selectbox("Selecciona Eje Y (Valor Numérico):", col_numericas, key="bar_y")
            if c_eje_x and c_eje_y:
                df_grouped = df_filtrado.groupby(c_eje_x)[c_eje_y].sum().reset_index()
                st.bar_chart(df_grouped.set_index(c_eje_x))

        with tab2:
            st.markdown("**Distribución Porcentual (Pie / Pastel)**")
            c_pie_cat = st.selectbox("Selecciona Categoría:", col_categoricas if col_categoricas else df_filtrado.columns, key="pie_cat")
            c_pie_val = st.selectbox("Selecciona Valor a Sumar:", col_numericas, key="pie_val")
            if c_pie_cat and c_pie_val:
                df_pie = df_filtrado.groupby(c_pie_cat)[c_pie_val].sum()
                st.bar_chart(df_pie)

        with tab3:
            st.markdown("**Tendencia / Seguimiento Cronológico**")
            c_line_y = st.selectbox("Selecciona Variable para Tendencia:", col_numericas, key="line_y")
            if c_line_y:
                st.line_chart(df_filtrado[c_line_y])

        with tab4:
            st.markdown("**Relación entre Variables**")
            if len(col_numericas) >= 2:
                c_scat_x = st.selectbox("Eje X:", col_numericas, index=0, key="scat_x")
                c_scat_y = st.selectbox("Eje Y:", col_numericas, index=1 if len(col_numericas)>1 else 0, key="scat_y")
                st.scatter_chart(df_filtrado[[c_scat_x, c_scat_y]])
            else:
                st.info("Se requieren al menos 2 columnas numéricas para el gráfico de dispersión.")
    else:
        st.warning("No se detectaron columnas numéricas para generar gráficos automáticos.")

    # -------------------------------------------------------------------------
    # RESGUARDO Y EXPORTACIÓN DE DATOS CON ID ÚNICO
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📥 Resguardo de Información y Exportación")
    
    # Generar identificador único y estampilla de tiempo
    id_unico = str(uuid.uuid4())[:8].upper()
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_base = f"Reporte_Obra_ID_{id_unico}_{timestamp_str}"
    
    st.write(f"🏷️ **ID Único de Consulta / Registro:** `{id_unico}`")
    st.write(f"🕒 **Fecha y Hora de Generación:** `{timestamp_str}`")

    exp_col1, exp_col2 = st.columns(2)

    # Exportación CSV
    csv_bytes = df_filtrado.to_csv(index=False).encode('utf-8')
    exp_col1.download_button(
        label="📄 Descargar Consulta (CSV)",
        data=csv_bytes,
        file_name=f"{nombre_base}.csv",
        mime="text/csv",
        use_container_width=True
    )

    # Exportación Excel
    if EXCEL_AVAILABLE:
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Reporte_Consolidado')
        
        exp_col2.download_button(
            label="📊 Descargar Consulta (Excel .xlsx)",
            data=buffer_excel.getvalue(),
            file_name=f"{nombre_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        exp_col2.warning("⚠️ Módulo 'openpyxl' en instalación. Usa la descarga CSV temporalmente.")

else:
    st.info("👋 Por favor sube uno o varios archivos (CSV/Excel) para iniciar la visualización y análisis.")
