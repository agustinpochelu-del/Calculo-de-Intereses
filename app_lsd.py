import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 2: Registro 01 (Cabecera) + Registro 02 (Empleados)")

# --- BARRA LATERAL ---
st.sidebar.header("1. Datos para Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

st.sidebar.divider()

st.sidebar.header("2. Fechas para Registro 02")
# Lógica Matemática de Fecha de Pago (Día 5 del mes anterior)
fecha_pago_defecto = datetime.today()
try:
    if len(periodo) == 6:
        año = int(periodo[:4])
        mes = int(periodo[4:])
        fecha_pago_defecto = datetime(año, mes, 1) - relativedelta(months=1)
        fecha_pago_defecto = fecha_pago_defecto.replace(day=5)
except:
    pass

fecha_pago = st.sidebar.date_input("Fecha de Pago (Se autocalcula)", value=fecha_pago_defecto)
fecha_rubrica = st.sidebar.date_input("Fecha de Rúbrica", value=datetime.today())

# --- ÁREA PRINCIPAL ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("**A. Sábana de Liquidación**")
    archivo_liq = st.file_uploader("Sube el archivo 'Conceptos y Totales'", type=["csv", "xls", "xlsx"])

with col2:
    st.markdown("**B. Maestro de Empleados**")
    archivo_emp = st.file_uploader("Sube la base con Legajo, CBU, etc.", type=["csv", "xls", "xlsx"])

if st.button("Generar TXT (Registros 01 y 02)", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None or archivo_emp is None:
        st.warning("⚠️ Faltan cargar datos o subir alguno de los dos archivos.")
    else:
        try:
            # --- 1. PROCESAR LIQUIDACIÓN (Igual que antes) ---
            if archivo_liq.name.lower().endswith('.csv'):
                df_liq = pd.read_csv(archivo_liq, encoding='utf-8', sep=',')
            else:
                df_liq = pd.read_excel(archivo_liq)

            df_liq.columns = df_liq.columns.str.strip()
            col_cuil_liq = next((col for col in df_liq.columns if 'CUIL' in col.upper().replace('.', '')), None)

            if not col_cuil_liq:
                st.error("❌ No encontré el C.U.I.L. en la liquidación.")
                st.stop()

            # Obtenemos la lista de empleados ÚNICOS que cobraron este mes
            empleados_unicos = df_liq[col_cuil_liq].dropna().unique()
            cantidad_empleados = len(empleados_unicos)

            # --- 2. PROCESAR MAESTRO DE EMPLEADOS ---
            if archivo_emp.name.lower().endswith('.csv'):
                df_emp = pd.read_csv(archivo_emp, encoding='utf-8', sep=',') # Probamos con coma
                if len(df_emp.columns) < 3:
                    archivo_emp.seek(0)
                    df_emp = pd.read_csv(archivo_emp, encoding='latin1', sep=';')
            else:
                df_emp = pd.read_excel(archivo_emp)

            df_emp.columns = df_emp.columns.str.strip()
            col_cuil_emp = next((col for col in df_emp.columns if 'CUIL' in col.upper().replace('.', '')), None)

            if not col_cuil_emp:
                st.error(f"❌ No encontré el C.U.I.L. en el Maestro de Empleados. Títulos leídos: {', '.join(df_emp.columns[:5])}")
                st.stop()

            # Creamos un diccionario rápido para buscar los datos del empleado por su CUIL
            maestro_dict = {}
            for _, row in df_emp.iterrows():
                cuil_limpio = limpiar_cuit(row[col_cuil_emp])
                maestro_dict[cuil_limpio] = {
                    # Si no encuentra la columna, devuelve vacío. Adaptalo a los nombres exactos de tus columnas.
                    'legajo': str(row.get('Legajo', row.get('Número de legajo', ''))).strip(),
                    'dependencia': str(row.get('Dependencia', '')).strip(),
                    'cbu': str(row.get('CBU', '')).strip().replace('-', '').replace(' ', '')
                }

            # --- 3. ARMADO DEL TXT ---
            lineas_txt = []

            # REGISTRO 01
            cuit_limpio = limpiar_cuit(cuit_empresa_input)
            tipo_liq_letra = tipo_liq[0]
            registro_01 = f"01{cuit_limpio}SJ{periodo}{tipo_liq_letra}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
            lineas_txt.append(registro_01)

            # Formato de fechas para TXT (AAAAMMDD)
            fecha_pago_txt = fecha_pago.strftime("%Y%m%d")
            fecha_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")

            # REGISTRO 02 (Uno por empleado)
            for cuil in empleados_unicos:
                cuil_limpio = limpiar_cuit(cuil)
                datos = maestro_dict.get(cuil_limpio, {})

                # Formateos exactos s/ manual AFIP (Rellenar con espacios a la derecha)
                legajo_txt = datos.get('legajo', '').ljust(10)[:10]
                dependencia_txt = datos.get('dependencia', '').ljust(50)[:50]
                cbu_raw = datos.get('cbu', '')

                # Lógica Inteligente de Pago
                if len(cbu_raw) == 22:
                    forma_pago_txt = '3'
                    cbu_txt = cbu_raw
                else:
                    forma_pago_txt = '1'
                    cbu_txt = ' ' * 22 # 22 espacios en blanco si no hay CBU
                
                dias_tope_txt = str(dias_base).zfill(3)

                # Ensamblado: 02 + CUIL(11) + Legajo(10) + Dep(50) + CBU(22) + DiasTope(3) + F.Pago(8) + F.Rubrica(8) + FormaPago(1)
                registro_02 = f"02{cuil_limpio}{legajo_txt}{dependencia_txt}{cbu_txt}{dias_tope_txt}{fecha_pago_txt}{fecha_rubrica_txt}{forma_pago_txt}"
                lineas_txt.append(registro_02)

            # --- RESULTADO ---
            texto_final = "\n".join(lineas_txt)
            
            st.success(f"✅ ¡Éxito! Se generó la cabecera y {cantidad_empleados} registros de empleados.")
            
            with st.expander("👀 Ver previsualización del TXT"):
                st.code(texto_final)
                
            st.download_button(
                label="⬇️ Descargar TXT (Registros 01 y 02)",
                data=texto_final,
                file_name=f"LSD_R01_R02_{periodo}.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Error técnico: {e}")
