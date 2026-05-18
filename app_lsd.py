import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

# --- FUNCIONES DE BASE DE DATOS LOCAL ---
DB_EMPLEADOS = "base_empleados.json"

def cargar_db():
    if os.path.exists(DB_EMPLEADOS):
        with open(DB_EMPLEADOS, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def guardar_db(datos):
    with open(DB_EMPLEADOS, "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=4)

def limpiar_cuit(cuit):
    c = str(cuit).strip()
    if c.endswith('.0'): c = c[:-2]
    return c.replace("-", "").replace(" ", "").replace(".", "").zfill(11)

# --- MOTOR DE LECTURA ---
def cargar_dataframe_con_busqueda(archivo_subido, palabra_clave):
    nombre = archivo_subido.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        hojas = pd.read_excel(archivo_subido, sheet_name=None, dtype=str)
        for _, df in hojas.items():
            df.columns = df.columns.astype(str).str.strip().str.upper()
            for col in df.columns:
                if palabra_clave in col.replace('.', ''):
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
            
        df.columns = df.columns.astype(str).str.strip().str.upper()
        for col in df.columns:
            if palabra_clave in col.replace('.', ''):
                return df, col
    return None, None

# ==========================================
# INTERFAZ DE USUARIO
# ==========================================
st.title("Generador de LSD con Memoria Histórica 🧠")
st.markdown("### PASO 2: Procesamiento Inteligente de Personal")

# --- BARRA LATERAL ---
st.sidebar.header("1. Parámetros Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="00")

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

# --- ÁREA PRINCIPAL ---
st.info("Subí únicamente la sábana de liquidación. El sistema cruzará los datos con su base interna.")
archivo_liq = st.file_uploader("Subir archivo 'Conceptos y Totales'", type=["csv", "xls", "xlsx"])

if archivo_liq:
    df_liq, col_cuil_liq = cargar_dataframe_con_busqueda(archivo_liq, 'CUIL')
    
    if df_liq is None:
        st.error("❌ No encontré el CUIL en la liquidación.")
        st.stop()
        
    df_liq_validos = df_liq.dropna(subset=[col_cuil_liq]).copy()
    empleados_mes = df_liq_validos[col_cuil_liq].unique()
    cantidad_empleados = len(empleados_mes)
    
    db_historica = cargar_db()
    
    # --- LA ADUANA: DETECTAR NUEVOS EMPLEADOS ---
    cuils_faltantes = [limpiar_cuit(c) for c in empleados_mes if limpiar_cuit(c) not in db_historica]
    
    if cuils_faltantes:
        st.warning(f"⚠️ ¡Atención! Se detectaron {len(cuils_faltantes)} empleados nuevos en esta liquidación que no están en la base de datos.")
        st.markdown("Por favor, **completá sus datos en la siguiente tabla** para agregarlos al repositorio histórico antes de generar el TXT:")
        
        # Preparamos una tabla interactiva para que llenes los datos faltantes
        df_faltantes = pd.DataFrame({
            "CUIL": cuils_faltantes,
            "Legajo": ["" for _ in cuils_faltantes],
            "Dependencia": ["ADMINISTRACION" for _ in cuils_faltantes], # Valor por defecto
            "CBU (O dejar vacío)": ["" for _ in cuils_faltantes],
            "Forma Pago (1,2,3,4)": ["3" for _ in cuils_faltantes]
        })
        
        # Tabla editable en pantalla
        datos_editados = st.data_editor(df_faltantes, num_rows="fixed", use_container_width=True, hide_index=True)
        
        if st.button("💾 Guardar Nuevos Empleados en Repositorio"):
            for index, row in datos_editados.iterrows():
                cuil = row["CUIL"]
                db_historica[cuil] = {
                    "legajo": str(row["Legajo"]).strip(),
                    "dependencia": str(row["Dependencia"]).strip(),
                    "cbu": str(row["CBU (O dejar vacío)"]).strip().replace("-", "").replace(" ", ""),
                    "forma_pago": str(row["Forma Pago (1,2,3,4)"]).strip()
                }
            guardar_db(db_historica)
            st.success("✅ ¡Base de datos actualizada! Ahora podés generar el TXT.")
            st.rerun() # Recarga la app automáticamente
            
    else:
        st.success(f"✅ Los {cantidad_empleados} empleados liquidados ya están registrados en el repositorio. ¡Vía libre para generar el archivo!")
        
        # --- PROCESAMIENTO Y GENERACIÓN DEL TXT (Solo aparece si no hay faltantes) ---
        if st.button("🚀 Generar TXT (Registros 01 y 02)", type="primary"):
            lineas_txt = []
            cuit_l = limpiar_cuit(cuit_empresa_input)
            t_liq = tipo_liq[0]
            
            # Registro 01
            r01 = f"01{cuit_l}SJ{periodo}{t_liq}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
            lineas_txt.append(r01)

            f_pago_txt = fecha_pago.strftime("%Y%m%d")
            f_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")

            # Registro 02
            for cuil_raw in empleados_mes:
                cuil_l = limpiar_cuit(cuil_raw)
                emp_data = db_historica[cuil_l] # Ahora estamos 100% seguros de que existe

                legajo_raw = emp_data.get('legajo', '')
                if not legajo_raw: legajo_raw = '0'
                legajo_fixed = legajo_raw if len(legajo_raw) == 10 else legajo_raw.rjust(10, ' ')

                dep_raw = emp_data.get('dependencia', 'ADMINISTRACION')
                dependencia_fixed = dep_raw.ljust(50)[:50]

                cbu_raw = emp_data.get('cbu', '')
                if len(cbu_raw) == 22:
                    cbu_txt = cbu_raw
                    fp_calc = '3'
                else:
                    cbu_txt = ' ' * 22 
                    fp_calc = '1'

                forma_pago = emp_data.get('forma_pago', '')
                if forma_pago not in ['1', '2', '3', '4']:
                    forma_pago = fp_calc

                dias_tope_fixed = str(dias_base).zfill(3)

                r02 = f"02{cuil_l}{legajo_fixed}{dependencia_fixed}{cbu_txt}{dias_tope_fixed}{f_pago_txt}{f_rubrica_txt}{forma_pago}"
                lineas_txt.append(r02)

            texto_final = "\n".join(lineas_txt)
            
            with st.expander("👀 Ver Archivo Generado"):
                st.code(texto_final)

            st.download_button(
                label="⬇️ Descargar TXT Validado",
                data=texto_final,
                file_name=f"LSD_Validado_{periodo}.txt",
                mime="text/plain",
                type="primary"
            )
