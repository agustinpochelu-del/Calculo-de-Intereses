import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Completar Datos", page_icon="📊")

st.title("📊 Completador de Datos para Reportes")
st.write("Subí tu archivo de Excel, elegí la columna que tiene los espacios en blanco y la app los va a rellenar copiando el dato de arriba.")

# 1. Subir el archivo
archivo_subido = st.file_uploader("Subí tu archivo Excel acá", type=["xlsx", "xls"])

if archivo_subido is not None:
    # Leer el archivo Excel
    df = pd.read_excel(archivo_subido)
    
    st.subheader("Vista previa de tu archivo original")
    st.dataframe(df.head(10)) # Muestra las primeras 10 filas

    # 2. Elegir la columna a procesar
    columnas = df.columns.tolist()
    columna_elegida = st.selectbox("¿Qué columna necesitás completar hacia abajo?", columnas)

    # 3. Botón para ejecutar el proceso
    if st.button("Completar Datos"):
        # Hacemos una copia para no alterar el original en memoria
        df_procesado = df.copy()
        
        # Esta es la magia: ffill() rellena los vacíos con el último valor válido
        df_procesado[columna_elegida] = df_procesado[columna_elegida].ffill()

        st.success("¡Datos completados con éxito!")
        
        st.subheader("Vista previa del resultado")
        st.dataframe(df_procesado.head(10))

        # 4. Preparar el archivo para descargar
        buffer = io.BytesIO()
        # Usamos openpyxl como motor para escribir el Excel
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_procesado.to_excel(writer, index=False, sheet_name='Datos_Completos')
        
        # Botón de descarga
        st.download_button(
            label="📥 Descargar Excel Listo",
            data=buffer.getvalue(),
            file_name="reporte_completado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )