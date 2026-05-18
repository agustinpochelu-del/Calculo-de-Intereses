import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    c = str(cuit).strip()
    if c.endswith('.0'): c = c[:-2]
    return c.replace("-", "").replace(" ", "").replace(".", "").zfill(11)

# --- MOTOR DE LECTURA INTELIGENTE ---
def cargar_dataframe_inteligente(archivo_subido, palabra_clave_columna):
    nombre = archivo_subido.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        hojas = pd.read_excel(archivo_subido, sheet_name=None, dtype=str)
        for _, df in hojas.items():
            df.columns = df.columns.astype(str).str.strip().str.upper()
            for col in df.columns:
                if palabra_clave_columna in col.replace('.', ''):
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
            
        df.columns = df.columns.astype(str).str.strip().str.upper()
        for col in df.columns:
            if palabra_clave_columna in col.replace('.', ''):
                return df
    return None

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 2: Cruce Seguro y Formateo de Registro 02")

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
        st.warning("⚠️ Por favor, tenés que subir ambos archivos para realizar el cruce.")
    else:
        with st.spinner("Cruzando datos y validando longitudes..."):
            try:
                # 1. LEER Y DETECTAR COLUMNAS DE LA SÁBANA
                df_liq = cargar_dataframe_inteligente(archivo_liq, 'CUIL')
                if df_liq is None:
                    st.error("❌ No encontré la columna CUIL en el archivo de Liquidación.")
                    st.stop()
                
                # PARCHE APLICADO: IGNORA PUNTOS EN LA BÚSQUEDA
                col_cuil_liq = next((col for col in df_liq.columns if 'CUIL' in col.replace('.', '')), None)
                col_legajo_liq = next((col for col in df_liq.columns if 'LEGAJO' in col.replace('.', '')), None)
                col_lugar_liq = next((col for col in df_liq.columns if 'LUGAR' in col.replace('.', '') or 'SUCURSAL' in col.replace('.', '') or 'DEPEN' in col.replace('.', '')), None)

                if not col_cuil_liq:
                    st.error("❌ Error interno: Falló la asignación de la columna CUIL de Liquidación.")
                    st.stop()

                df_liq_validos = df_liq.dropna(subset=[col_cuil_liq]).copy()
                empleados_mes = df_liq_validos[col_cuil_liq].unique()
                cantidad_empleados = len(empleados_mes)

                backup_lugar = {}
                if col_lugar_liq:
                    for _, r in df_liq_validos.iterrows():
                        c_limp = limpiar_cuit(r[col_cuil_liq])
                        lug = str(r[col_lugar_liq]).strip()
                        if lug and lug.lower() != 'nan':
                            backup_lugar[c_limp] = lug

                # 2. LEER Y DETECTAR COLUMNAS DEL MAESTRO
                df_emp = cargar_dataframe_inteligente(archivo_emp, 'CUIL')
                if df_emp is None:
                    st.error("❌ No encontré la columna CUIL en el Maestro de Empleados.")
                    st.stop()
                
                # PARCHE APLICADO: IGNORA PUNTOS EN EL MAESTRO TAMBIÉN
                col_cuil_emp = next((col for col in df_emp.columns if 'CUIL' in col.replace('.', '')), None)
                col_legajo_emp = next((col for col in df_emp.columns if 'LEGAJO' in col.replace('.', '')), None)
                col_dep_emp = next((col for col in df_emp.columns if 'DEPEN' in col.replace('.', '') or 'REVISTA' in col.replace('.', '')), None)
                col_cbu_emp = next((col for col in df_emp.columns if 'CBU' in col.replace('.', '')), None)
                col_fp_emp = next((col for col in df_emp.columns if 'FORMA' in col.replace('.', '') or 'PAGO' in col.replace('.', '')), None)

                maestro_por_cuil = {}
                maestro_por_legajo = {}

                for _, row in df_emp.iterrows():
                    cuil_m = limpiar_cuit(row[col_cuil_emp]) if col_cuil_emp else ""
                    
                    legajo_m = str(row[col_legajo_emp]).replace('.0', '').strip() if col_legajo_emp else ""
                    if legajo_m.lower() == 'nan': legajo_m = ""
                    
                    dep_m = str(row[col_dep_emp]).strip() if col_dep_emp else ""
                    if dep_m.lower() == 'nan': dep_m = ""
                    
                    cbu_m = str(row[col_cbu_emp]).strip().replace('-', '').replace(' ', '') if col_cbu_emp else ""
                    if cbu_m.lower() == 'nan': cbu_m = ""
                    
                    fp_m = str(row[col_fp_emp]).replace('.0', '').strip() if col_fp_emp else ""
                    if fp_m.lower() == 'nan': fp_m = ""

                    dict_datos = {
                        'cuil': cuil_m, 'legajo': legajo_m, 'dependencia': dep_m, 'cbu': cbu_m, 'forma_pago': fp_m
                    }
                    
                    if cuil_m: maestro_por_cuil[cuil_m] = dict_datos
                    if legajo_m: maestro_por_legajo[legajo_m] = dict_datos

                # 3. CONSTRUCCIÓN DEL ARCHIVO TXT
                lineas_txt = []
                
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
                    
                    row_individual_liq = df_liq_validos[df_liq_validos[col_cuil_liq] == cuil_raw].iloc[0]
                    legajo_backup = str(row_individual_liq[col_legajo_liq]).replace('.0', '').strip() if col_legajo_liq else ""
                    if legajo_backup.lower() == 'nan': legajo_backup = ""

                    # DOBLE MOTOR
                    emp_data = maestro_por_cuil.get(cuil_l)
                    if not emp_data and legajo_backup:
                        emp_data = maestro_por_legajo.get(legajo_backup)
                    if not emp_data:
                        emp_data = {}

                    legajo_final = emp_data.get('legajo', legajo_backup)
                    if not legajo_final: legajo_final = "0"
                    legajo_fixed = legajo_final.rjust(10, ' ')

                    dep_final = emp_data.get('dependencia', backup_lugar.get(cuil_l, 'ADMINISTRACION'))
                    if not dep_final: dep_final = "ADMINISTRACION"
                    dependencia_fixed = dep_final.ljust(50, ' ')[:50]

                    cbu_final = emp_data.get('cbu', '')
                    cbu_fixed = cbu_final if len(cbu_final) == 22 else (' ' * 22)

                    fp_final = emp_data.get('forma_pago', '')
                    if fp_final not in ['1', '2', '3', '4']:
                        fp_final = '3' if len(cbu_fixed.strip()) == 22 else '1'

                    dias_tope_fixed = str(dias_base).zfill(3)

                    r02 = f"02{cuil_l}{legajo_fixed}{dependencia_fixed}{cbu_fixed}{dias_tope_fixed}{f_pago_txt}{f_rubrica_txt}{fp_final}"
                    lineas_txt.append(r02)

                    tabla_control.append({
                        "CUIL": cuil_l,
                        "Legajo": f"'{legajo_fixed}'",
                        "Dependencia": f"'{dependencia_fixed}'",
                        "CBU": cbu_fixed if cbu_fixed.strip() else "[Espacio en Blanco]",
                        "F. Pago": fp_final,
                        "Largo": len(r02)
                    })

                texto_final = "\n".join(lineas_txt)
                st.success("✅ ¡Cruce realizado con éxito!")
                
                st.markdown("### 🔍 Tabla de Control Técnico")
                st.dataframe(pd.DataFrame(tabla_control), use_container_width=True)

                with st.expander("👀 Ver Estructura del Archivo"):
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
