import streamlit as st
import pandas as pd

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 1: Registro 01 (Cabecera)")

# --- BARRA LATERAL ---
st.sidebar.header("Datos para Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

# --- ÁREA PRINCIPAL ---
archivo_liq = st.file_uploader("Subir archivo 'Conceptos y Totales'", type=["csv", "xls", "xlsx"])

if st.button("Generar TXT (Solo Registro 01)", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None:
        st.warning("⚠️ Faltan cargar datos o subir el archivo.")
    else:
        try:
            # LECTURA DIRECTA
            nombre_archivo = archivo_liq.name.lower()
            
            if nombre_archivo.endswith('.csv'):
                # Leemos tu CSV separado por comas (como el que me acabas de subir)
                df_liq = pd.read_csv(archivo_liq, encoding='utf-8', sep=',')
            else:
                # Por si algún día subís un Excel
                df_liq = pd.read_excel(archivo_liq)

            # Limpiamos espacios fantasmas en los títulos
            df_liq.columns = df_liq.columns.str.strip()

            # Buscamos la columna (sea C.U.I.L. o CUIL)
            col_cuil = None
            for col in df_liq.columns:
                if 'CUIL' in col.upper().replace('.', ''):
                    col_cuil = col
                    break

            if not col_cuil:
                st.error(f"❌ No encontré el CUIL. Títulos leídos: {', '.join(df_liq.columns[:5])}")
                st.stop()

            # Contamos los empleados únicos
            cantidad_empleados = df_liq[col_cuil].dropna().nunique()

            # ARMADO DEL REGISTRO 01
            cuit_limpio = limpiar_cuit(cuit_empresa_input)
            tipo_liq_letra = tipo_liq[0]
            nro_liq_form = str(nro_liq).zfill(5)
            dias_base_form = str(dias_base).zfill(2)
            cant_empleados_form = str(cantidad_empleados).zfill(6)
            
            # Formato AFIP: 01 + CUIT(11) + SJ + Periodo(6) + Tipo(1) + Nro(5) + Dias(2) + CantEmpleados(6)
            registro_01 = f"01{cuit_limpio}SJ{periodo}{tipo_liq_letra}{nro_liq_form}{dias_base_form}{cant_empleados_form}"
            
            st.success(f"✅ ¡Éxito! Se detectaron automáticamente **{cantidad_empleados} empleados únicos**.")
            st.code(registro_01)
            
        except Exception as e:
            st.error(f"Error técnico: {e}")
