import io
import pandas as pd
import streamlit as st

# Configuración responsive para dispositivos móviles y escritorio
st.set_page_config(page_title="Gestor dinámico de datos", layout="wide")

st.title("🏗️ Visor y Analizador de Datos")
st.write("Carga cualquier archivo CSV para analizar, filtrar y exportar resultados.")

# Carga dinámica del archivo
uploaded_file = st.file_uploader("Sube tu archivo CSV", type=["csv"])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        
        st.success("¡Archivo cargado correctamente!")
        
        # 1. Métricas generales del archivo
        st.subheader("📌 Resumen del archivo")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Registros", len(df))
        c2.metric("Total de Columnas", len(df.columns))
        
        col_num = df.select_dtypes(include=['float64', 'int64']).columns
        c3.metric("Columnas Numéricas", len(col_num))

        # 2. Sistema de Filtro Dinámico
        st.subheader("🔍 Filtro de información")
        columna_filtro = st.selectbox("Selecciona la columna por la que deseas filtrar:", ["(Sin filtro)"] + list(df.columns))
        
        df_filtrado = df.copy()
        
        if columna_filtro != "(Sin filtro)":
            valores_unicos = df[columna_filtro].dropna().unique().tolist()
            seleccion = st.multiselect(f"Selecciona opciones de '{columna_filtro}':", valores_unicos)
            
            if seleccion:
                df_filtrado = df[df[columna_filtro].isin(seleccion)]

        # 3. Vista previa de la tabla filtrada
        st.subheader("📋 Datos procesados")
        st.dataframe(df_filtrado, use_container_width=True)

        # 4. Sección de Descargas (Resguardo de información)
        st.subheader("📥 Exportar consulta / Resguardo")
        st.info("Descarga los datos mostrados en pantalla para conservar los registros en tu dispositivo.")
        
        col_exp1, col_exp2 = st.columns(2)

        # Exportar en formato CSV
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8')
        col_exp1.download_button(
            label="📄 Descargar Consulta (CSV)",
            data=csv_data,
            file_name="resguardo_consulta.csv",
            mime="text/csv",
            use_container_width=True
        )

        # Exportar en formato Excel (.xlsx)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Consulta')
        
        col_exp2.download_button(
            label="📊 Descargar Consulta (Excel)",
            data=buffer.getvalue(),
            file_name="resguardo_consulta.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # 5. Visualización gráfica automática
        if len(col_num) > 0:
            st.subheader("📈 Gráficos automáticos")
            col_grafico = st.selectbox("Selecciona columna numérica para visualizar:", col_num)
            if col_grafico:
                st.line_chart(df_filtrado[col_grafico])

    except Exception as e:
        st.error(f"Error al procesar el archivo CSV: {e}")
else:
    st.info("Por favor, sube un archivo CSV para generar las tablas y herramientas de descarga.")
