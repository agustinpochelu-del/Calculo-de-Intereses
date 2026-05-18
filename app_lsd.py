import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    c = str(cuit).strip()
    if c.endswith('.0'): c = c[:-2]
    return c.replace("-", "").replace(" ", "").replace(".", "").zfill(11)

def clean_column_name(col):
    """Elimina acentos, puntos y espacios para estandarizar los títulos del Excel/CSV."""
    if pd.isna(col): return ""
    text = str(col).strip().upper()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text.replace('.', '')

# --- MOTOR DE LECTURA INTELIGENTE ---
def cargar_dataframe_inteligente(archivo_subido, palabra_clave_columna):
    nombre = archivo_subido.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        hojas = pd.read_excel(archivo_subido, sheet_name=None, dtype=str)
        for _, df in hojas.items():
            # Limpiamos los títulos de esta pestaña
            df.columns = [clean_column_name(c) for c in df.columns]
            if any(palabra_clave_columna in c for c in df.columns):
                return df
    else:
        try:
            df = pd.read_csv(archivo_subido, encoding='utf-8', sep=',', dtype=str)
            if len(df.columns) < 3:
                archivo_subido.seek(0)
                df = pd.read_csv(archivo_subido, encoding='latin1', sep=';', dtype=str)
        except:
            archivo_subido.seek(0)
            df = pd.read_csv(archivo_subido, encoding='latin1', sep=';', dtype=str)
            
        df.columns = [clean_column_name(c) for c in df.columns]
        if any(palabra_clave_columna in c for c in df.columns):
            return df
    return None

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 2: Extracción Directa desde Sábana Mensual")

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

# --- CARGA DE INSUMOS ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📊 Archivo A: Sábana de Liquidación")
    archivo_liq = st.file_uploader("Subir 'Conceptos y Totales'", type=["csv", "xls", "xlsx"], key="liq")

with col2:
    st.markdown("#### 🗂️ Archivo B: Maestro de Empleados")
    archivo_emp = st.file_uploader("Subir 'Listado Empleados'", type=["csv", "xls", "xlsx"], key="emp")

st.divider()

# --- PROCESAMIENTO ---
if st.button("Procesar y Generar Registro 02", type="primary"):
    if archivo_liq is None or archivo_emp is None:
        st.warning("⚠️ Por favor, subí ambos archivos para realizar el cruce.")
    else:
        with st.spinner("Procesando datos con extracción directa..."):
            try:
                # 1. LEER SÁBANA (Buscamos por el CUIL estandarizado)
                df_liq = cargar_dataframe_inteligente(archivo_liq, 'CUIL')
                if df_liq is None:
                    st.error("❌ No encontré la columna CUIL en el archivo de Liquidación.")
                    st.stop()
                
                # Identificar columnas normalizadas en la Sábana
                col_cuil_liq = next((c for c in df_liq.columns if 'CUIL' in c), None)
                col_legajo_liq = next((c for c in df_liq.columns if 'LEGAJO' in c), None)
                col_lugar_liq = next((c for c in df_liq.columns if 'LUGAR DE TRABAJO' in c), None)

                df_liq_validos = df_liq.dropna(subset=[col_cuil_liq]).copy()
                empleados_mes = df_liq_validos[col_cuil_liq].unique()
                cantidad_empleados = len(empleados_mes)

                # 2. LEER MAESTRO DE EMPLEADOS
                df_emp = cargar_dataframe_inteligente(archivo_emp, 'CUIL')
                if df_emp is None:
                    st.error("❌ No encontré la columna CUIL en el Maestro de Empleados.")
                    st.stop()
                
                col_cuil_emp = next((c for c in df_emp.columns if 'CUIL' in c), None)
                col_legajo_emp = next((c for c in df_emp.columns if 'LEGAJO' in c), None)
                col_dep_emp = next((c for c in df_emp.columns if 'DEPENDENCIA' in c or 'REVISTA' in c), None)
                col_cbu_emp = next((c for c in df_emp.columns if 'CBU' in c), None)
                col_fp_emp = next((c for c in df_emp.columns if 'FORMA' in c or 'PAGO' in c), None)

                # Mapear el Maestro para búsqueda rápida (CBU y Forma de Pago)
                maestro_por_cuil = {}
                maestro_por_legajo = {}

                for _, row in df_emp.iterrows():
                    cuil_m = limpiar_cuit(row[col_cuil_emp]) if col_cuil_emp else ""
                    legajo_m = str(row.get(col_legajo_emp, '')).replace('.0', '').strip() if col_legajo_emp else ""
                    dep_m = str(row.get(col_dep_emp, '')).strip() if col_dep_emp else ""
                    cbu_m = str(row.get(col_cbu_emp, '')).strip().replace('-', '').replace(' ', '') if col_cbu_emp else ""
                    fp_m = str(row.get(col_fp_emp, '')).replace('.0', '').strip() if col_fp_emp else ""

                    if legajo_m.lower() == 'nan': legajo_m = ""
                    if dep_m.lower() == 'nan': dep_m = ""
                    if cbu_m.lower() == 'nan': cbu_m = ""
                    if fp_m.lower() == 'nan': fp_m = ""

                    dict_datos = {'cuil': cuil_m, 'legajo': legajo_m, 'dependencia_maestro': dep_m, 'cbu': cbu_m, 'forma_pago': fp_m}
                    if cuil_m: maestro_por_cuil[cuil_m] = dict_datos
                    if legajo_m: maestro_por_legajo[legajo_m] = dict_datos

                # 3. CONSTRUCCIÓN DEL ARCHIVO TXT
                lineas_txt = []
                
                # Registro 01
                cuit_l = limpiar_cuit(cuit_empresa_input)
                t_liq = tipo_liq[0]
                r01 = f"01{cuit_l}SJ{periodo}{t_liq}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
                lineas_txt.append(r01)

                f_pago_txt = fecha_pago.strftime("%Y%m%d")
                f_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")

                tabla_control = []

                # Registro 02
                for cuil_raw in empleados_mes:
                    cuil_l = limpiar_cuit(cuil_raw)
                    
                    # Extraemos los datos DIRECTOS de la Sábana de este mes
                    df_individual_liq = df_liq_validos[df_liq_validos[col_cuil_liq] == cuil_raw]
                    row_individual_liq = df_individual_liq.iloc[0]
                    
                    legajo_sabana = str(row_individual_liq.get(col_legajo_liq, '')).replace('.0', '').strip()
                    if legajo_sabana.lower() == 'nan': legajo_sabana = ""
                        
                    lugar_sabana = str(row_individual_liq.get(col_lugar_liq, '')).strip()
                    if lugar_sabana.lower() == 'nan': lugar_sabana = ""

                    # Buscamos datos de pago en el Maestro
                    emp_data = maestro_por_cuil.get(cuil_l)
                    if not emp_data and legajo_sabana:
                        emp_data = maestro_por_legajo.get(legajo_sabana)
                    if not emp_data:
                        emp_data = {}

                    # ASIGNACIÓN DE LEGAJO (Sábana manda)
                    legajo_final = legajo_sabana if legajo_sabana else emp_data.get('legajo', '0')
                    legajo_fixed = legajo_final.rjust(10, ' ')

                    # ASIGNACIÓN DE DEPENDENCIA (Sábana manda)
                    dep_final = lugar_sabana if lugar_sabana else emp_data.get('dependencia_maestro', 'ADMINISTRACION')
                    dependencia_fixed = dep_final.ljust(50, ' ')[:50]

                    # ASIGNACIÓN DE CBU
                    cbu_final = emp_data.get('cbu', '')
                    cbu_fixed = cbu_final if len(cbu_final) == 22 else (' ' * 22)

                    # ASIGNACIÓN DE FORMA DE PAGO
                    fp_final = emp_data.get('forma_pago', '')
                    if fp_final not in ['1', '2', '3', '4']:
                        fp_final = '3' if len(cbu_fixed.strip()) == 22 else '1'

                    dias_tope_fixed = str(dias_base).zfill(3)

                    # Ensamblado estructural final Registro 02
                    r02 = f"02{cuil_l}{legajo_fixed}{dependencia_fixed}{cbu_fixed}{dias_tope_fixed}{f_pago_txt}{f_rubrica_txt}{fp_final}"
                    lineas_txt.append(r02)

                    tabla_control.append({
                        "CUIL": cuil_l,
                        "Legajo": f"'{legajo_fixed}'",
                        "Lugar de Trabajo (Sábana)": dependencia_fixed.strip(),
                        "CBU": cbu_fixed if cbu_fixed.strip() else "[Efectivo]",
                        "F. Pago": fp_final,
                        "Largo": len(r02)
                    })

                texto_final = "\n".join(lineas_txt)
                st.success("✅ ¡Cruce estructural realizado con éxito!")
                
                st.markdown("### 🔍 Tabla de Auditoría (Registro 02)")
                st.dataframe(pd.DataFrame(tabla_control), use_container_width=True)

                with st.expander("👀 Ver Estructura del Archivo .txt"):
                    st.code(texto_final)

                st.download_button(
                    label="⬇️ Descargar TXT (01 + 02)",
                    data=texto_final,
                    file_name=f"LSD_R01_R02_{periodo}.txt",
                    mime="text/plain",
                    type="primary"
                )

            except Exception as e:
                st.error(f"Error técnico en el proceso: {e}")
