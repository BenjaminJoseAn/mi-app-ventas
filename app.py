import datetime
import io
import uuid
import numpy as np
import pandas as pd
import streamlit as st

# Intentar importar librerías avanzadas
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# Configuración inicial del panel
st.set_page_config(page_title="Sistema Ejecutivo de Control de Obra", layout="wide", page_icon="🏗️")

st.markdown("""
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏗️ Sistema Ejecutivo de Control y Reportes de Obra")
st.caption("Consolidación multi-archivo, análisis gráfico interactivo y exportación ejecutiva en PDF.")

# -----------------------------------------------------------------------------
# MODOS DE TRABAJO
# -----------------------------------------------------------------------------
modo_trabajo = st.radio(
    "📌 Selecciona la modalidad de análisis:",
    ["📄 Archivo Único", "📂 Consolidado Semanal / Mensual (Varios Archivos)"],
    horizontal=True
)

lista_dfs = []

if modo_trabajo == "📄 Archivo Único":
    origen = st.radio("Fuente de los datos:", ["📁 Cargar Archivo (CSV / Excel)", "📷 Captura con Cámara (Dispositivo Móvil)"], horizontal=True)
    if origen == "📁 Cargar Archivo (CSV / Excel)":
        archivo = st.file_uploader("Selecciona el archivo de obra", type=["csv", "xlsx", "xls"])
        if archivo:
            try:
                df = pd.read_csv(archivo) if archivo.name.endswith('.csv') else pd.read_excel(archivo)
                lista_dfs.append(df)
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
    else:
        foto = st.camera_input("Toma una captura limpia de la lista o ticket")
        if foto:
            st.info("ℹ️ Captura almacenada correctamente. Para procesamiento estructurado automático se recomienda cargar el formato digital CSV/Excel.")

else:
    archivos = st.file_uploader(
        "Sube los archivos CSV o Excel a consolidar (ej. semanas o frentes de obra distintos):",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True
    )
    if archivos:
        for a in archivos:
            try:
                d = pd.read_csv(a) if a.name.endswith('.csv') else pd.read_excel(a)
                d['Archivo_Origen'] = a.name
                lista_dfs.append(d)
            except Exception as e:
                st.error(f"Error en {a.name}: {e}")

# -----------------------------------------------------------------------------
# PROCESAMIENTO GENERAL DE DATOS
# -----------------------------------------------------------------------------
if lista_dfs:
    df_base = pd.concat(lista_dfs, ignore_index=True)

    # Identificar columnas numéricas y de texto/categoría estrictamente
    col_num = df_base.select_dtypes(include=[np.number]).columns.tolist()
    col_cat = df_base.select_dtypes(include=['object', 'category']).columns.tolist()

    # Identificar columnas con formato tipo Fecha
    col_fechas = [c for c in df_base.columns if 'fecha' in c.lower() or 'date' in c.lower()]

    st.divider()
    
    # Métricas Principales en tarjetas
    st.subheader("📊 Resumen Ejecutivo")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Total de Registros", len(df_base))
    with m2:
        st.metric("Total de Columnas", len(df_base.columns))
    with m3:
        st.metric("Variables Numéricas", len(col_num))
    with m4:
        st.metric("Categorías Detectadas", len(col_cat))

    # -------------------------------------------------------------------------
    # FILTRADO INTELIGENTE
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("🔍 Filtro de Control de Datos")
    df_filtrado = df_base.copy()

    col_filtro = st.selectbox("Selecciona la categoría o fecha por la que deseas filtrar:", ["(Sin Filtro)"] + col_cat)
    
    if col_filtro != "(Sin Filtro)":
        opciones_validas = sorted([str(x) for x in df_base[col_filtro].dropna().unique().tolist()])
        seleccion = st.multiselect(f"Selecciona valores específicos de '{col_filtro}':", opciones_validas)
        if seleccion:
            df_filtrado = df_base[df_base[col_filtro].astype(str).isin(seleccion)]

    # Vista previa de datos
    st.subheader("📋 Tabla de Registros Procesados")
    st.dataframe(df_filtrado, use_container_width=True)

    # -------------------------------------------------------------------------
    # DASHBOARD GRÁFICO (CON PASTEL CORRECTO Y FILTROS LÓGICOS)
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📈 Dashboard Visual e Indicadores")

    if len(col_num) > 0 and len(col_cat) > 0:
        tab_bar, tab_pie, tab_line, tab_scat = st.tabs(["📊 Gráfico de Barras", "🍕 Gráfico de Pastel", "📈 Líneas / Tendencia", "📍 Dispersión"])

        with tab_bar:
            st.markdown("**Comparativa de Sumas por Categoría (Barras)**")
            ej_x = st.selectbox("Categoría a comparar (Eje X):", col_cat, key="bar_x_clean")
            ej_y = st.selectbox("Monto o Cantidad (Eje Y):", col_num, key="bar_y_clean")
            if ej_x and ej_y:
                df_bar = df_filtrado.groupby(ej_x)[ej_y].sum().reset_index()
                st.bar_chart(df_bar.set_index(ej_x))

        with tab_pie:
            st.markdown("**Distribución Porcentual (Pastel)**")
            pie_c = st.selectbox("Categoría para el Pastel:", col_cat, key="pie_c_clean")
            pie_v = st.selectbox("Monto / Valor a representar:", col_num, key="pie_v_clean")
            if pie_c and pie_v:
                df_pie = df_filtrado.groupby(pie_c)[pie_v].sum().reset_index()
                # Gráfico interactivo tipo pastel usando Vega-Lite (Altair nativo)
                chart_pie = {
                    "mark": {"type": "arc", "tooltip": True},
                    "encoding": {
                        "theta": {"field": pie_v, "type": "quantitative"},
                        "color": {"field": pie_c, "type": "nominal"}
                    },
                    "data": {"values": df_pie.to_dict(orient="records")}
                }
                st.vega_lite_chart(chart_pie, use_container_width=True)

        with tab_line:
            st.markdown("**Comportamiento o Tendencia**")
            line_var = st.selectbox("Selecciona variable a evaluar en el tiempo:", col_num, key="line_clean")
            if line_var:
                st.line_chart(df_filtrado[line_var])

        with tab_scat:
            st.markdown("**Relación entre 2 Variables Numéricas**")
            if len(col_num) >= 2:
                scat_x = st.selectbox("Variable X:", col_num, index=0, key="scat_x_c")
                scat_y = st.selectbox("Variable Y:", col_num, index=1 if len(col_num) > 1 else 0, key="scat_y_c")
                st.scatter_chart(df_filtrado[[scat_x, scat_y]])
            else:
                st.info("Se necesitan al menos 2 columnas numéricas para este gráfico.")
    else:
        st.warning("Se requieren columnas de texto/categorías y valores numéricos para construir el panel gráfico.")

    # -------------------------------------------------------------------------
    # EXPEDICIÓN Y RESGUARDO EN PDF Y FORMATOS DIGITALES
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📥 Generación de Reportes e Identificador de Control")

    id_reporte = str(uuid.uuid4())[:8].upper()
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nombre_archivo_base = f"Reporte_Obra_ID_{id_reporte}"

    st.write(f"🔑 **ID de Auditoría / Control:** `{id_reporte}`")
    st.write(f"🕒 **Fecha de Emisión:** `{fecha_actual}`")

    c_csv, c_excel, c_pdf = st.columns(3)

    # Exportación CSV
    csv_bytes = df_filtrado.to_csv(index=False).encode('utf-8')
    c_csv.download_button("📄 Exportar CSV", data=csv_bytes, file_name=f"{nombre_archivo_base}.csv", mime="text/csv", use_container_width=True)

    # Exportación Excel
    if EXCEL_AVAILABLE:
        buf_xl = io.BytesIO()
        with pd.ExcelWriter(buf_xl, engine='openpyxl') as w:
            df_filtrado.to_excel(w, index=False, sheet_name='Reporte')
        c_excel.download_button("📊 Exportar Excel (.xlsx)", data=buf_xl.getvalue(), file_name=f"{nombre_archivo_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        c_excel.warning("Instalando 'openpyxl'...")

    # Exportación PDF Ejecutiva
    if PDF_AVAILABLE:
        buf_pdf = io.BytesIO()
        doc = SimpleDocTemplate(buf_pdf, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        styles = getSampleStyleSheet()

        # Encabezado del PDF
        story.append(Paragraph(f"<b>REPORTE EJECUTIVO DE CONTROL DE OBRA</b>", styles['Title']))
        story.append(Paragraph(f"<b>ID Único:</b> {id_reporte} | <b>Fecha:</b> {fecha_actual}", styles['Normal']))
        story.append(Spacer(1, 15))

        # Métricas principales en el PDF
        story.append(Paragraph(f"<b>Resumen:</b> Total de registros: {len(df_filtrado)}", styles['Heading2']))
        story.append(Spacer(1, 10))

        # Tabla (primeras 20 filas para mantener claridad)
        datos_tabla = [list(df_filtrado.columns)]
        for _, fila in df_filtrado.head(20).iterrows():
            datos_tabla.append([str(val) for val in fila.values])

        t = Table(datos_tabla, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        story.append(t)

        doc.build(story)

        c_pdf.download_button(
            "📕 Exportar Reporte Ejecutivo (PDF)",
            data=buf_pdf.getvalue(),
            file_name=f"{nombre_archivo_base}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    else:
        c_pdf.warning("Instalando 'reportlab' para PDF...")

else:
    st.info("👋 Sube uno o varios archivos CSV / Excel para desplegar el sistema de análisis y reportes.")
