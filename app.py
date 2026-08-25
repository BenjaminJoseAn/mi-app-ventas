import datetime
import io
import uuid
import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Verificación e importación de librerías para exportar
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# Lectura de PDF previa
try:
    import pypdf
    PDF_READ_AVAILABLE = True
except ImportError:
    PDF_READ_AVAILABLE = False


# -----------------------------------------------------------------------------
# FUNCIÓN AUXILIAR PARA PARSEAR PDFs GENERADOS PREVIAMENTE (CORREGIDA)
# -----------------------------------------------------------------------------
def cargar_datos_desde_pdf(archivo_pdf):
    """
    Extrae la información tabular de los reportes PDF exportados previamente
    de forma dinámica sin fallar por diferencias en el número de columnas.
    """
    reader = pypdf.PdfReader(archivo_pdf)
    texto_completo = ""
    
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            texto_completo += txt + "\n"

    lineas = [line.strip() for line in texto_completo.split('\n') if line.strip()]
    
    filas_datos = []
    max_cols = 0
    
    for line in lineas:
        # Omitir títulos/encabezados genéricos del reporte PDF para evitar ruido
        if "REPORTE EJECUTIVO DE CONTROL DE OBRA" in line or "GRÁFICOS E INDICADORES" in line:
            continue
            
        parts = line.split()
        if len(parts) >= 1:
            filas_datos.append(parts)
            if len(parts) > max_cols:
                max_cols = len(parts)
            
    if filas_datos and max_cols > 0:
        # Rellenar filas cortas con cadenas vacías para normalizar dimensiones
        filas_normalizadas = [
            fila + [""] * (max_cols - len(fila)) for fila in filas_datos
        ]
        
        # Asignar nombres genéricos a las columnas
        headers = [f"Columna_{i+1}" for i in range(max_cols)]
        df_pdf = pd.DataFrame(filas_normalizadas, columns=headers)
        
        # Usar la primera fila como encabezado si sus elementos no se repiten
        if len(df_pdf) > 1:
            posible_header = [str(x) for x in df_pdf.iloc[0].values]
            if len(posible_header) == len(set(posible_header)):
                df_pdf.columns = posible_header
                df_pdf = df_pdf.iloc[1:].reset_index(drop=True)
    else:
        df_pdf = pd.DataFrame({'Contenido_PDF': [texto_completo]})
        
    df_pdf['Archivo_Origen'] = archivo_pdf.name
    return df_pdf


# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ENCABEZADO
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Sistema Ejecutivo de Control de Obra", layout="wide", page_icon="🏗️")

st.title("🏗️ Sistema Ejecutivo de Control y Reportes de Obra")
st.caption("Consolidación multi-archivo (CSV, Excel, PDF), seguimiento periódico (días/semanas) y generación de reportes.")

# -----------------------------------------------------------------------------
# MODOS DE TRABAJO
# -----------------------------------------------------------------------------
modo_trabajo = st.radio(
    "📌 Selecciona la modalidad de análisis:",
    ["📄 Archivo Único / PDF Histórico", "📂 Consolidado Periódico (Varios Archivos / PDFs Semanales)"],
    horizontal=True
)

lista_dfs = []

def Cargar_Archivo_Generico(file_obj):
    nombre = file_obj.name.lower()
    if nombre.endswith('.csv'):
        d = pd.read_csv(file_obj)
    elif nombre.endswith(('.xlsx', '.xls')):
        d = pd.read_excel(file_obj)
    elif nombre.endswith('.pdf'):
        if PDF_READ_AVAILABLE:
            d = cargar_datos_desde_pdf(file_obj)
        else:
            st.error("Instala 'pypdf' para habilitar la lectura de archivos PDF.")
            return None
    else:
        return None
    
    d['Archivo_Origen'] = file_obj.name
    return d

if modo_trabajo == "📄 Archivo Único / PDF Histórico":
    origen = st.radio("Fuente de los datos:", ["📁 Cargar Archivo (CSV / Excel / PDF)", "📷 Captura con Cámara"], horizontal=True)
    if origen == "📁 Cargar Archivo (CSV / Excel / PDF)":
        archivo = st.file_uploader("Selecciona el archivo de obra o reporte PDF previo", type=["csv", "xlsx", "xls", "pdf"])
        if archivo:
            try:
                df_cargado = Cargar_Archivo_Generico(archivo)
                if df_cargado is not None:
                    lista_dfs.append(df_cargado)
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
    else:
        foto = st.camera_input("Toma una captura limpia de la lista o ticket")
        if foto:
            st.info("ℹ️ Captura almacenada correctamente.")

else:
    archivos = st.file_uploader(
        "Sube los archivos o reportes PDF semanales/diarios a consolidar:",
        type=["csv", "xlsx", "xls", "pdf"],
        accept_multiple_files=True
    )
    if archivos:
        for a in archivos:
            try:
                df_cargado = Cargar_Archivo_Generico(a)
                if df_cargado is not None:
                    lista_dfs.append(df_cargado)
            except Exception as e:
                st.error(f"Error en {a.name}: {e}")

# -----------------------------------------------------------------------------
# PROCESAMIENTO GENERAL DE DATOS
# -----------------------------------------------------------------------------
if lista_dfs:
    df_base = pd.concat(lista_dfs, ignore_index=True)

    # Conversión numérica compatible con Pandas 2.0+
    for col in df_base.columns:
        converted = pd.to_numeric(df_base[col], errors='coerce')
        if not converted.isna().all():
            df_base[col] = converted

    col_num = df_base.select_dtypes(include=[np.number]).columns.tolist()
    col_todas = df_base.columns.tolist()

    st.divider()
    
    # Resumen Ejecutivo
    st.subheader("📊 Resumen Ejecutivo y Seguimiento Acumulado")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total de Registros", len(df_base))
    m2.metric("Total de Columnas", len(df_base.columns))
    m3.metric("Archivos / PDFs Cargados", len(lista_dfs))

    # -------------------------------------------------------------------------
    # FILTRADO DINÁMICO Y SEGUIMIENTO PERIÓDICO (DÍAS / SEMANAS)
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("🔍 Filtro de Control de Datos y Períodos")
    df_filtrado = df_base.copy()

    col_filtro, col_periodo = st.columns(2)
    
    with col_filtro:
        c_filtro = st.selectbox("Selecciona columna para filtrar:", ["(Sin Filtro)"] + col_todas)
        if c_filtro != "(Sin Filtro)":
            opciones_validas = sorted([str(x) for x in df_base[c_filtro].dropna().unique().tolist()])
            seleccion = st.multiselect(f"Filtrar por '{c_filtro}':", opciones_validas)
            if seleccion:
                df_filtrado = df_base[df_base[c_filtro].astype(str).isin(seleccion)]

    with col_periodo:
        col_fecha = st.selectbox("Columna de Fecha (Opcional para seguimiento semanal/diario):", ["(Sin Fecha)"] + col_todas)
        if col_fecha != "(Sin Fecha)":
            try:
                df_filtrado[col_fecha] = pd.to_datetime(df_filtrado[col_fecha], errors='coerce')
                min_f = df_filtrado[col_fecha].min()
                max_f = df_filtrado[col_fecha].max()
                
                if pd.notna(min_f) and pd.notna(max_f):
                    rango_fechas = st.date_input("Rango de seguimiento (Días / Semanas):", [min_f, max_f])
                    if len(rango_fechas) == 2:
                        df_filtrado = df_filtrado[
                            (df_filtrado[col_fecha].dt.date >= rango_fechas[0]) & 
                            (df_filtrado[col_fecha].dt.date <= rango_fechas[1])
                        ]
            except Exception as e:
                st.caption(f"No se pudo formatear la columna como fecha: {e}")

    # Tabla interactiva de datos
    st.subheader("📋 Tabla de Registros Consolidados")
    st.dataframe(df_filtrado, use_container_width=True)

    # -------------------------------------------------------------------------
    # DASHBOARD GRÁFICO CON MATPLOTLIB (PARA STREAMLIT Y PDF)
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📈 Dashboard Visual e Indicadores de Obra")

    opciones_y = col_num if len(col_num) > 0 else col_todas

    tab_bar, tab_pie, tab_line = st.tabs(["📊 Gráfico de Barras", "🍕 Gráfico de Pastel", "📈 Líneas / Tendencia Semanal"])

    fig_bar, fig_pie, fig_line = None, None, None

    with tab_bar:
        st.markdown("**Comparativa por Categoría (Barras)**")
        ej_x = st.selectbox("Eje X (Categoría):", col_todas, index=0, key="bar_x_clean")
        ej_y = st.selectbox("Eje Y (Suma):", opciones_y, index=0, key="bar_y_clean")
        if ej_x and ej_y:
            try:
                df_bar = df_filtrado.groupby(ej_x)[ej_y].sum().reset_index().dropna()
                fig_bar, ax_b = plt.subplots(figsize=(7, 3.5))
                ax_b.bar(df_bar[ej_x].astype(str), df_bar[ej_y], color='#1f77b4')
                ax_b.set_title(f"Suma de {ej_y} por {ej_x}")
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                st.pyplot(fig_bar)
            except Exception:
                st.warning("Selecciona columnas válidas para el gráfico de barras.")

    with tab_pie:
        st.markdown("**Distribución Porcentual (Pastel)**")
        pie_c = st.selectbox("Categoría:", col_todas, index=0, key="pie_c_clean")
        pie_v = st.selectbox("Valor a representar:", opciones_y, index=0, key="pie_v_clean")
        if pie_c and pie_v:
            try:
                df_pie = df_filtrado.groupby(pie_c)[pie_v].sum().reset_index().dropna()
                fig_pie, ax_p = plt.subplots(figsize=(6, 3.5))
                ax_p.pie(df_pie[pie_v], labels=df_pie[pie_c].astype(str), autopct='%1.1f%%', startangle=90)
                ax_p.axis('equal')
                ax_p.set_title(f"Distribución de {pie_v} por {pie_c}")
                plt.tight_layout()
                st.pyplot(fig_pie)
            except Exception:
                st.warning("Selecciona una columna numérica para el pastel.")

    with tab_line:
        st.markdown("**Comportamiento o Tendencia Temporal (Días / Semanas)**")
        line_var = st.selectbox("Variable para líneas:", opciones_y, key="line_clean")
        if line_var:
            try:
                fig_line, ax_l = plt.subplots(figsize=(7, 3.5))
                ax_l.plot(df_filtrado[line_var].values, marker='o', color='#2ca02c')
                ax_l.set_title(f"Tendencia de {line_var}")
                plt.tight_layout()
                st.pyplot(fig_line)
            except Exception:
                st.warning("Selecciona una variable numérica.")

    # -------------------------------------------------------------------------
    # GENERACIÓN DE REPORTES Y PDF MULTI-TABLA
    # -------------------------------------------------------------------------
    st.divider()
    st.subheader("📥 Generación de Nuevo Reporte Consolidado")

    id_reporte = str(uuid.uuid4())[:8].upper()
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    nombre_archivo_base = f"Reporte_Obra_Consolidado_ID_{id_reporte}"

    st.write(f"🔑 **ID de Auditoría / Control de este corte:** `{id_reporte}`")
    st.write(f"🕒 **Fecha de Emisión:** `{fecha_actual}`")

    c_csv, c_excel, c_pdf = st.columns(3)

    # Exportación CSV
    csv_bytes = df_filtrado.to_csv(index=False).encode('utf-8')
    c_csv.download_button("📄 Exportar CSV", data=csv_bytes, file_name=f"{nombre_archivo_base}.csv", mime="text/csv", use_container_width=True)

    # Exportación Excel
    if EXCEL_AVAILABLE:
        buf_xl = io.BytesIO()
        with pd.ExcelWriter(buf_xl, engine='openpyxl') as w:
            for origen_nombre, df_sub in df_filtrado.groupby('Archivo_Origen'):
                sheet_name = str(origen_nombre)[:30].replace('/', '_').replace('\\', '_')
                df_sub.dropna(how='all', axis=1).to_excel(w, index=False, sheet_name=sheet_name)
        c_excel.download_button("📊 Exportar Excel (.xlsx)", data=buf_xl.getvalue(), file_name=f"{nombre_archivo_base}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # Exportación PDF
    if PDF_AVAILABLE:
        buf_pdf = io.BytesIO()
        doc = SimpleDocTemplate(buf_pdf, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
        story = []
        styles = getSampleStyleSheet()

        # Encabezado principal
        story.append(Paragraph(f"<b>REPORTE EJECUTIVO DE CONTROL DE OBRA</b>", styles['Title']))
        story.append(Paragraph(f"<b>ID Único:</b> {id_reporte} | <b>Fecha:</b> {fecha_actual}", styles['Normal']))
        story.append(Spacer(1, 15))

        # Insertar Gráficos Generados
        story.append(Paragraph("<b>GRÁFICOS E INDICADORES CLAVE</b>", styles['Heading2']))
        story.append(Spacer(1, 5))

        for fig_temp in [fig_bar, fig_pie, fig_line]:
            if fig_temp is not None:
                img_buf = io.BytesIO()
                fig_temp.savefig(img_buf, format='png', dpi=150)
                img_buf.seek(0)
                story.append(RLImage(img_buf, width=450, height=225))
                story.append(Spacer(1, 10))

        story.append(PageBreak())

        # Tablas por Archivo Fuente
        story.append(Paragraph("<b>DETALLE DE TABLAS POR ARCHIVO / PDF FUENTE</b>", styles['Heading1']))
        story.append(Spacer(1, 10))

        for origen_nombre, df_grupo in df_filtrado.groupby('Archivo_Origen'):
            df_limpio = df_grupo.dropna(how='all', axis=1)

            story.append(Paragraph(f"📄 <b>Fuente: {origen_nombre}</b> (Registros: {len(df_limpio)})", styles['Heading2']))
            story.append(Spacer(1, 5))

            datos_tabla = [list(df_limpio.columns)]
            for _, fila in df_limpio.head(15).iterrows():
                datos_tabla.append([str(val) if pd.notna(val) else "" for val in fila.values])

            t = Table(datos_tabla, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ]))
            story.append(t)
            story.append(Spacer(1, 15))

        doc.build(story)

        c_pdf.download_button(
            "📕 Exportar Reporte Ejecutivo Consolidado (PDF)",
            data=buf_pdf.getvalue(),
            file_name=f"{nombre_archivo_base}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
