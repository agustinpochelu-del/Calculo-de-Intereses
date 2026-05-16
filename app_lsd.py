import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 1: Prueba de Registro 01 (Cabecera)")

# --- BARRA LATERAL (Solo los datos del Registro 01) ---
st.sidebar.header("Datos para Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, placeholder="Ej: 30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, placeholder="Ej: 202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

# --- ÁREA PRINCIPAL ---
archivo_liq = st.file_uploader("Subir Excel/CSV de 'Conceptos y Totales'", type=["xls", "xlsx", "csv"])

# --- PROCESAMIENTO (Solo Registro 01) ---
if st.button("Generar TXT (Solo Registro 01)", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None:
        st.warning("⚠️ Faltan cargar datos o subir el archivo.")
    else:
        try:
            # 1. Leemos el archivo
            nombre_liq = archivo_liq.name.lower()
            if nombre_liq.endswith('.xlsx') or nombre_liq.endswith('.xls'):
                df_liq = pd.read_excel(archivo_liq)
            else:
                df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')

            # 2. Buscamos la columna CUIL para contar la cantidad de empleados
            df_liq.columns = df_liq.columns.str.strip()
            col_cuil = next((col for col in df_liq.columns if 'CUIL' in col.upper().replace('.', '')), None)
            
            # SI FALLA, TE ESCUPE LAS COLUMNAS EXACTAS QUE LEYÓ
            if not col_cuil:
                columnas_leidas = ', '.join([str(c) for c in df_liq.columns[:15]])
                st.error(f"❌ No encontré la columna CUIL.\n\nPython leyó estas columnas en la primera fila: {columnas_leidas}")
                st.stop()
            
            # Contamos cuántos CUILs únicos hay
            cantidad_empleados = df_liq[col_cuil].nunique()

            # 3. ARMADO DEL REGISTRO 01
            cuit_limpio = limpiar_cuit(cuit_empresa_input)
            tipo_liq_letra = tipo_liq[0]
            nro_liq_form = str(nro_liq).zfill(5)
            dias_base_form = str(dias_base).zfill(2)
            cant_empleados_form = str(cantidad_empleados).zfill(6)
            
            # Formato AFIP: 01 + CUIT(11) + SJ + Periodo(6) + Tipo(1) + Nro(5) + Dias(2) + CantEmpleados(6)
            registro_01 = f"01{cuit_limpio}SJ{periodo}{tipo_liq_letra}{nro_liq_form}{dias_base_form}{cant_empleados_form}"
            
            st.success("✅ ¡Registro 01 generado con éxito!")
            st.code(registro_01)
            
            st.download_button(
                label="⬇️ Descargar TXT de prueba",
                data=registro_01,
                file_name=f"LSD_R01_{periodo}.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
