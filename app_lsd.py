import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

# --- MOTOR DE BÚSQUEDA ---
def cargar_dataframe_con_busqueda(archivo_subido, palabra_clave):
    nombre = archivo_subido.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        hojas = pd.read_excel(archivo_subido, sheet_name=None, dtype=str)
        for nombre_hoja, df in hojas.items():
            df.columns = df.columns.astype(str).str.strip()
            for col in df.columns:
                if palabra_clave in col.upper().replace('.', ''):
                    return df, col
    else:
        try:
            df = pd.read_csv(archivo_subido, encoding='utf-8', sep=',', dtype=str)
            if len(df.columns) < 3:
                archivo_subido.seek(0)
                df = pd.read_csv(archivo_subido, encoding='latin1', sep=';', dtype=str)
        except:
            archivo_subido.seek(0)
            df = pd.read_csv(archivo_subido, encoding='latin1', sep=';', dtype=str)
            
        df.columns = df.columns.astype(str).str.strip()
        for col in df.columns:
            if palabra_clave in col.upper().replace('.', ''):
                return df, col
    return None, None

# ==========================================
# INTERFAZ DE USUARIO
# ==========================================
st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 2: Corrección Estricta de Registro 02")

# --- BARRA LATERAL ---
st.sidebar.header("1. Parámetros Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

st.sidebar.divider()

st.sidebar.header("2. Fechas Registro 02")
fecha_pago_defecto = datetime.today()
try:
    if len(periodo) == 6:
        año = int(periodo[:4])
        mes = int(periodo[4:])
        fecha_pago_defecto = datetime(año, mes, 1) - relativedelta(months=1)
        fecha_pago_defecto = fecha_pago_defecto.replace(day=5)
except:
    pass

fecha_pago = st.sidebar.date_input("Fecha de Pago", value=fecha_pago_defecto)
fecha_rubrica = st.sidebar.date_input("Fecha de Rúbrica", value=datetime.today())

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("**A. Sábana de Liquidación**")
    archivo_liq = st.file_uploader("Subir archivo 'Conceptos y Totales'", type=["csv", "xls", "xlsx"], key="liq")

with col2:
    st.markdown("**B. Maestro de Empleados**")
    archivo_emp = st.file_uploader("Subir base histórica de Empleados", type=["csv", "xls", "xlsx"], key="emp")

# --- MOTOR DE PROCESAMIENTO ---
if st.button("Procesar y Validar Estructuras", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None or archivo_emp is None:
        st.warning("⚠️ Faltan completar datos o subir archivos.")
    else:
        try:
            # 1. Leer Sábana
            df_liq, col_cuil_liq = cargar_dataframe_con_busqueda(archivo_liq, 'CUIL')
            if df_liq is None:
                st.error("❌ No se encontró la columna de CUIL en la liquidación.")
                st.stop()
                
            df_liq.columns = df_liq.columns.str.upper()
            col_cuil_liq = col_cuil_liq.upper()
            
            empleados_unicos = df_liq[col_cuil_liq].dropna().unique()
            cantidad_empleados = len(empleados_unicos)

            # Red de Seguridad para Dependencia
            col_lugar_liq = None
            for col in df_liq.columns:
                if 'LUGAR DE TRABAJO' in col or 'DEPENDENCIA' in col or 'SUCURSAL' in col:
                    col_lugar_liq = col
                    break
            
            fallback_lugar = {}
            if col_lugar_liq:
                for _, row in df_liq.iterrows():
                    c = limpiar_cuit(row[col_cuil_liq])
                    if c not in fallback_lugar:
                        lugar = str(row[col_lugar_liq]).strip()
                        if lugar.lower() != 'nan' and lugar:
                            fallback_lugar[c] = lugar

            # 2. Leer Maestro de Empleados
            df_emp, col_cuil_emp = cargar_dataframe_con_busqueda(archivo_emp, 'CUIL')
            if df_emp is None:
                st.error("❌ No se encontró la columna de CUIL en el Maestro.")
                st.stop()

            df_emp.columns = df_emp.columns.str.upper()
            col_cuil_emp_upper = col_cuil_emp.upper()

            # MAPEO FIEL AL EXCEL
            maestro_dict = {}
            for _, row in df_emp.iterrows():
                cuil_m = limpiar_cuit(row[col_cuil_emp_upper])
                
                # Legajo: Leído tal cual viene
                legajo_val = str(row.get('LEGAJO', '')).replace('.0', '')
                if legajo_val.lower() == 'nan': legajo_val = ''
                
                dep_val = str(row.get('DEPENDENCIA DE REVISTA', row.get('DEPENDENCIA', ''))).strip()
                if dep_val.lower() == 'nan': dep_val = ''
                
                cbu_val = str(row.get('CBU', '')).strip().replace('-', '').replace(' ', '')
                if cbu_val.lower() == 'nan': cbu_val = ''
                    
                # AHORA SÍ LEO TU COLUMNA DE FORMA DE PAGO
                fp_val = str(row.get('FORMA DE PAGO', '')).replace('.0', '').strip()
                if fp_val.lower() == 'nan': fp_val = ''

                maestro_dict[cuil_m] = {
                    'legajo': legajo_val,
                    'dependencia': dep_val,
                    'cbu': cbu_val,
                    'forma_pago': fp_val
                }

            # 3. ARMADO DE LÍNEAS
            lineas_txt = []
            
            cuit_l = limpiar_cuit(cuit_empresa_input)
            t_liq = tipo_liq[0]
            r01 = f"01{cuit_l}SJ{periodo}{t_liq}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
            lineas_txt.append(r01)

            f_pago_txt = fecha_pago.strftime("%Y%m%d")
            f_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")

            desglose_mostrar = [] 
            
            # REGISTRO 02
            for cuil in empleados_unicos:
                cuil_l = limpiar_cuit(cuil)
                emp_data = maestro_dict.get(cuil_l, {})

                # Dependencia (con Fallback)
                dep_raw = emp_data.get('dependencia', '')
                if not dep_raw: 
                    dep_raw = fallback_lugar.get(cuil_l, 'ADMINISTRACION')
                dependencia_fixed = dep_raw.ljust(50)[:50]

                # Legajo Fiel (Alineado a la derecha como tu sistema)
                legajo_raw = emp_data.get('legajo', '')
                if len(legajo_raw) == 10:
                    legajo_fixed = legajo_raw
                else:
                    legajo_fixed = legajo_raw.strip().rjust(10, ' ')
                    if not legajo_fixed.strip(): 
                        legajo_fixed = '0'.rjust(10, ' ')

                # CBU
                cbu_fixed = emp_data.get('cbu', '')
                if len(cbu_fixed) == 22:
                    cbu_txt = cbu_fixed
                else:
                    cbu_txt = ' ' * 22 

                # FORMA DE PAGO (Prioridad total a lo que dice tu Excel)
                fp_excel = emp_data.get('forma_pago', '')
                if fp_excel in ['1', '2', '3', '4']:
                    forma_pago = fp_excel
                else:
                    # Solo calcula automático si la celda estaba vacía
                    forma_pago = '3' if len(cbu_txt.strip()) == 22 else '1'

                dias_tope_fixed = str(dias_base).zfill(3)

                r02 = f"02{cuil_l}{legajo_fixed}{dependencia_fixed}{cbu_txt}{dias_tope_fixed}{f_pago_txt}{f_rubrica_txt}{forma_pago}"
                lineas_txt.append(r02)
                
                desglose_mostrar.append({
                    "CUIL": cuil_l,
                    "Legajo": f"'{legajo_fixed}'",
                    "Dependencia": f"'{dependencia_fixed}'",
                    "F. Pago Excel": fp_excel if fp_excel else "[Calculado]",
                    "Largo Línea": len(r02)
                })

            # 4. RESULTADO
            texto_final = "\n".join(lineas_txt)
            st.success("✅ ¡Registros procesados respetando tu Excel!")
            
            st.markdown("### 🔍 Tabla de Control (Registro 02)")
            st.dataframe(pd.DataFrame(desglose_mostrar), use_container_width=True)

            with st.expander("👀 Ver Archivo"):
                st.code(texto_final)

            st.download_button(
                label="⬇️ Descargar TXT Validado (01 + 02)",
                data=texto_final,
                file_name=f"LSD_Validado_{periodo}.txt",
                mime="text/plain",
                type="primary"
            )

        except Exception as e:
            st.error(f"Error en procesamiento: {e}")
