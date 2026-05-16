import streamlit as st
import pandas as pd

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 1: Registro 01 con Escáner de Pestañas")

# --- BARRA LATERAL ---
st.sidebar.header("Datos para Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

# --- ÁREA PRINCIPAL ---
archivo_liq = st.file_uploader("Subir archivo Excel o CSV", type=["csv", "xls", "xlsx"])

if st.button("Generar TXT (Solo Registro 01)", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None:
        st.warning("⚠️ Faltan cargar datos o subir el archivo.")
    else:
        try:
            nombre_archivo = archivo_liq.name.lower()
            df_liq = None
            col_cuil = None
            
            # 1. LECTURA INTELIGENTE (Busca en TODAS las pestañas del Excel)
            if nombre_archivo.endswith('.xlsx') or nombre_archivo.endswith('.xls'):
                # Leemos todas las pestañas del Excel
                diccionario_hojas = pd.read_excel(archivo_liq, sheet_name=None)
                
                # Buscamos en qué hoja está el CUIL
                for nombre_hoja, df_hoja in diccionario_hojas.items():
                    df_hoja.columns = df_hoja.columns.astype(str).str.strip()
                    for col in df_hoja.columns:
                        if 'CUIL' in col.upper().replace('.', ''):
                            df_liq = df_hoja
                            col_cuil = col
                            st.info(f"🔍 ¡Pestaña correcta detectada! Se extrajeron los empleados de la hoja: **'{nombre_hoja}'**")
                            break
                    if df_liq is not None:
                        break # Ya la encontró, salimos del ciclo de búsqueda
                        
                if df_liq is None:
                    st.error("❌ Revisé TODAS las pestañas del Excel y no encontré el C.U.I.L. en ninguna.")
                    st.stop()
                    
            else:
                # Si es CSV, leemos directo e intentamos con separador coma o punto y coma
                try:
                    df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=',')
                    if len(df_liq.columns) < 5:
                        archivo_liq.seek(0)
                        df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                except:
                    archivo_liq.seek(0)
                    df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                    
                df_liq.columns = df_liq.columns.astype(str).str.strip()
                for col in df_liq.columns:
                    if 'CUIL' in col.upper().replace('.', ''):
                        col_cuil = col
                        break
                        
                if not col_cuil:
                    st.error(f"❌ No encontré el C.U.I.L. Títulos leídos: {', '.join(df_liq.columns[:5])}")
                    st.stop()

            # 2. CONTAMOS LOS EMPLEADOS
            cantidad_empleados = df_liq[col_cuil].dropna().nunique()

            # 3. ARMADO DEL REGISTRO 01
            cuit_limpio = limpiar_cuit(cuit_empresa_input)
            tipo_liq_letra = tipo_liq[0]
            nro_liq_form = str(nro_liq).zfill(5)
            dias_base_form = str(dias_base).zfill(2)
            cant_empleados_form = str(cantidad_empleados).zfill(6)
            
            registro_01 = f"01{cuit_limpio}SJ{periodo}{tipo_liq_letra}{nro_liq_form}{dias_base_form}{cant_empleados_form}"
            
            st.success(f"✅ ¡Éxito! Se detectaron **{cantidad_empleados} empleados únicos**.")
            st.code(registro_01)
            
        except Exception as e:
            st.error(f"Error técnico: {e}")
