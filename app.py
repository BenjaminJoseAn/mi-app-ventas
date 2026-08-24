import streamlit as st
import pandas as pd

# Configuración de la página para móviles
st.set_page_config(page_title="Mi App de Ventas", layout="wide")

st.title("📊 Panel de Análisis")
st.write("Bienvenido a la aplicación de procesamiento de datos.")

# Cargar archivo CSV o Excel desde el teléfono/PC
uploaded_file = st.file_uploader("Sube tu archivo de datos (CSV)", type=["csv"])

if uploaded_file is not None:
    # Usando Pandas para leer los datos
    df = pd.read_csv(uploaded_file)
    
    st.subheader("Vista previa de los datos")
    st.dataframe(df.head(), use_container_width=True)
    
    # Resumen rápido
    st.subheader("Métricas clave")
    col1, col2 = st.columns(2)
    col1.metric("Total Filas", len(df))
    col2.metric("Columnas", len(df.columns))

    # Ejemplo de gráfico simple con Pandas y Streamlit
    st.subheader("Visualización")
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    
    if len(numeric_cols) > 0:
        columna_numeric = st.selectbox("Selecciona columna para graficar", numeric_cols)
        if columna_numeric:
            st.line_chart(df[columna_numeric])
    else:
        st.warning("El archivo CSV no contiene columnas numéricas para graficar.")
else:
    st.info("Por favor, sube un archivo CSV para empezar.")