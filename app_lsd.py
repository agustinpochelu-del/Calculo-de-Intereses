import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

# =================================================================
# --- 1. FUNCIONES DE REPOSITORIO Y PROCESAMIENTO ---
# =================================================================

def cargar_json(nombre_archivo):
    if os.path.exists(nombre_archivo):
        with open(nombre_archivo, "r") as file:
            return json.load(file)
    return {}

def guardar_json(nombre_archivo, datos):
    with open(nombre_archivo, "w") as file:
        json.dump(datos, file, indent=4)

def procesar_txt_afip(archivo_txt_afip):
    df = pd.read_csv(archivo_txt_afip, sep=';', dtype=str, encoding='latin1')
    if len(df.columns) < 4:
        st.error("❌ Archivo incorrecto. Asegurate de subir el archivo 'Conceptos_Contribuyente...txt'.")
        return {}
    df_mapeo = df.iloc[:, [2, 3, 0, 1]].copy()
    df_mapeo.columns = ['codigo_sistema', 'descripcion_sistema', 'codigo_afip', 'descripcion_afip']
    df_mapeo = df_mapeo.dropna(subset=['codigo_sistema'])
    
    def clasificar_por_afip(cod_afip):
        if pd.isna(cod_afip): return 0
        if cod_afip.startswith('1') or cod_afip.startswith('2'): return 1
        if cod_afip.startswith('5'): return 2
        if cod_afip.startswith('8'): return 3
        return 0
        
    df_mapeo['tipo'] = df_mapeo['codigo_afip'].apply(clasificar_por_afip)
    return df_mapeo.set_index('codigo_sistema').to_dict('index')

# --- FUNCIONES FORMATEADORAS ---
def limpiar_cuit_cuil(cuit_cuil):
    return str(cuit_cuil).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

def form_imp(valor, enteros=13, decimales=2):
    if pd.isna(valor): valor = 0
    val_int = int(round(float(valor), decimales) * (10**decimales))
    return str(abs(val_int)).zfill(enteros + decimales)

def form_cant(valor, longitud=5):
    if pd.isna(valor): valor = 0
    return str(int(float(valor))).zfill(longitud)


# =================================================================
# --- 2. INTERFAZ DE USUARIO (UI) ---
# =================================================================

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("Herramienta para convertir liquidaciones de Excel al formato `.txt` de AFIP.")

# --- BARRA LATERAL ---
st.sidebar.header("1. Cerebro de la App")

archivo_afip = st.sidebar.file_uploader("Subir TXT de AFIP", type=["txt"])
if archivo_afip:
    if st.sidebar.button("Procesar y Guardar Conceptos"):
        nuevo_mapeo = procesar_txt_afip(archivo_afip)
        if nuevo_mapeo:
            guardar_json("mapeo_conceptos.json", nuevo_mapeo)
            st.sidebar.success("✅ ¡Conceptos actualizados con éxito!")
            st.rerun()

st.sidebar.divider()
st.sidebar.header("2. Datos de la Empresa y Período")

cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, placeholder="Ej: 30-64496559-3")

topes_db = cargar_json("topes_historicos.json")
mapeo_conceptos_db = cargar_json("mapeo_conceptos.json") 
config_app = cargar_json("config_app.json") # Para guardar la fecha de rúbrica

periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, placeholder="Ej: 202604")

tope_min, tope_max = 0.0, 0.0
fecha_pago_defecto = datetime.today()

if periodo and len(periodo) == 6:
    # Lógica de Topes
    if periodo in topes_db:
        tope_min = topes_db[periodo]["min"]
        tope_max = topes_db[periodo]["max"]
        st.sidebar.success(f"✅ Topes de {periodo} cargados.")
    else:
        st.sidebar.warning(f"⚠️ No hay topes para {periodo}.")
        nuevo_min = st.sidebar.number_input("Tope Mínimo", min_value=0.0, format="%.2f")
        nuevo_max = st.sidebar.number_input("Tope Máximo", min_value=0.0, format="%.2f")
        if st.sidebar.button("💾 Guardar Topes"):
            topes_db[periodo] = {"min": nuevo_min, "max": nuevo_max}
            guardar_json("topes_historicos.json", topes_db)
            st.sidebar.success("¡Topes guardados!")
            st.rerun()
            
    # Lógica Matemática de Fecha de Pago (Día 5 del mes anterior)
    try:
        año = int(periodo[:4])
        mes = int(periodo[4:])
        fecha_periodo = datetime(año, mes, 1)
        fecha_pago_defecto = fecha_periodo - relativedelta(months=1)
        fecha_pago_defecto = fecha_pago_defecto.replace(day=5)
    except:
        pass

st.sidebar.divider()
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, placeholder="Ej: 1", value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

st.sidebar.divider()
st.sidebar.header("Fechas (Registro 02)")

fecha_pago = st.sidebar.date_input("Fecha de Pago (Se autocalcula)", value=fecha_pago_defecto)

# Lógica de Fecha de Rúbrica Persistente
rubrica_guardada_str = config_app.get("fecha_rubrica", None)
rubrica_defecto = datetime.strptime(rubrica_guardada_str, "%Y-%m-%d") if rubrica_guardada_str else None

fecha_rubrica = st.sidebar.date_input("Fecha de Rúbrica", value=rubrica_defecto)

if st.sidebar.button("💾 Guardar Rúbrica por defecto"):
    if fecha_rubrica:
        config_app["fecha_rubrica"] = fecha_rubrica.strftime("%Y-%m-%d")
        guardar_json("config_app.json", config_app)
        st.sidebar.success("¡Rúbrica guardada!")

# --- ÁREA PRINCIPAL: CARGA DE ARCHIVOS ---
st.subheader("3. Insumos de la Liquidación Mensual")

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📊 A. Datos Transaccionales (Sábana)")
    archivo_liq = st.file_uploader("Subir Excel/CSV de 'Conceptos y Totales'", type=["xls", "xlsx", "csv"])

with col2:
    st.markdown("### 🗂️ B. Datos Maestros (Base)")
    archivo_emp = st.file_uploader("Subir Excel Maestro de Empleados", type=["xls", "xlsx", "csv"])

st.divider()

# =================================================================
# --- PROCESAMIENTO ---
# =================================================================
if st.button("Procesar y Generar TXT", type="primary"):
    if not mapeo_conceptos_db:
        st.error("❌ Detenido: Falta cargar el mapeo de conceptos de AFIP.")
    elif archivo_liq is None:
        st.warning("⚠️ Por favor, subí el archivo de liquidación ('Conceptos y totales').")
    elif archivo_emp is None:
        st.warning("⚠️ Por favor, subí la Base Maestra de Empleados para armar el Registro 02.")
    elif not cuit_empresa_input:
        st.warning("⚠️ Por favor, ingresá el CUIT de la empresa en la barra lateral.")
    elif not periodo or len(periodo) != 6:
        st.warning("⚠️ El Período debe tener 6 dígitos.")
    elif not fecha_rubrica:
        st.warning("⚠️ Falta ingresar la Fecha de Rúbrica.")
    else:
        with st.spinner("🔄 Procesando liquidación y cruzando bases..."):
            try:
                # 1. LEER SÁBANA DE LIQUIDACIÓN
                nombre_liq = archivo_liq.name.lower()
                if nombre_liq.endswith('.xlsx') or nombre_liq.endswith('.xls'):
                    df_liq = pd.read_excel(archivo_liq)
                else:
                    try:
                        df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=',')
                        if len(df_liq.columns) < 5: 
                            archivo_liq.seek(0)
                            df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                    except:
                        archivo_liq.seek(0)
                        df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                
                df_liq.columns = df_liq.columns.str.strip() 
                col_cuil_liq = next((col for col in df_liq.columns if 'CUIL' in col.upper().replace('.', '')), None)
                if not col_cuil_liq:
                    st.error("❌ No encontré la columna CUIL en la liquidación.")
                    st.stop()
                df_liq.rename(columns={col_cuil_liq: 'CUIL_INTERNO'}, inplace=True)
                df_liq = df_liq.dropna(subset=['CUIL_INTERNO'])
                
                # 2. LEER BASE MAESTRA DE EMPLEADOS
                nombre_emp = archivo_emp.name.lower()
                if nombre_emp.endswith('.xlsx') or nombre_emp.endswith('.xls'):
                    df_emp = pd.read_excel(archivo_emp)
                else:
                    df_emp = pd.read_csv(archivo_emp, sep=';', encoding='latin1')
                    
                df_emp.columns = df_emp.columns.str.strip()
                col_cuil_emp = next((col for col in df_emp.columns if 'CUIL' in col.upper().replace('.', '')), None)
                if not col_cuil_emp:
                    st.error("❌ No encontré la columna CUIL en la Base Maestra de Empleados.")
                    st.stop()
                
                # Armamos un diccionario con los datos del empleado para búsqueda ultrarrápida
                base_empleados_dict = {}
                for index, row in df_emp.iterrows():
                    c_limpio = limpiar_cuit_cuil(row[col_cuil_emp])
                    base_empleados_dict[c_limpio] = {
                        'legajo': str(row.get('Legajo', '')).strip(),
                        'dependencia': str(row.get('Dependencia', '')).strip(),
                        'cbu': str(row.get('CBU', '')).strip().replace('-', '').replace(' ', '')
                    }
                
                cuit_empresa = limpiar_cuit_cuil(cuit_empresa_input) 
                empleados_procesados = df_liq['CUIL_INTERNO'].unique()
                cantidad_empleados = len(empleados_procesados)
                lineas_txt = []
                
                # --- REGISTRO 01 ---
                tipo_liq_letra = tipo_liq[0]
                registro_01 = f"01{cuit_empresa}SJ{periodo}{tipo_liq_letra}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
                lineas_txt.append(registro_01)
                
                # Fechas formateadas para el TXT (AAAAMMDD)
                fecha_pago_txt = fecha_pago.strftime("%Y%m%d")
                fecha_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")
                
                # --- RECORRIDO DE EMPLEADOS ---
                for cuil in empleados_procesados:
                    cuil_limpio = limpiar_cuit_cuil(cuil)
                    df_empleado_liq = df_liq[df_liq['CUIL_INTERNO'] == cuil]
                    
                    # Buscamos al empleado en la Base Maestra
                    datos_maestros = base_empleados_dict.get(cuil_limpio, {})
                    
                    # --- REGISTRO 02 ---
                    # Longitudes exactas s/AFIP: Legajo(10), Dependencia(50), CBU(22), DiasTope(3)
                    legajo_txt = datos_maestros.get('legajo', '').ljust(10)[:10]
                    dependencia_txt = datos_maestros.get('dependencia', '').ljust(50)[:50]
                    cbu_raw = datos_maestros.get('cbu', '')
                    
                    # Lógica Forma de Pago y CBU
                    if len(cbu_raw) == 22:
                        forma_pago_txt = '3'
                        cbu_txt = cbu_raw
                    else:
                        forma_pago_txt = '1'
                        cbu_txt = ' ' * 22  # Si no hay CBU, se mandan 22 espacios
                        
                    dias_tope_txt = str(dias_base).zfill(3)
                    
                    registro_02 = f"02{cuil_limpio}{legajo_txt}{dependencia_txt}{cbu_txt}{dias_tope_txt}{fecha_pago_txt}{fecha_rubrica_txt}{forma_pago_txt}"
                    lineas_txt.append(registro_02)
                    
                    # --- REGISTRO 03 ---
                    for index, row in df_empleado_liq.iterrows():
                        try:
                            cod_sistema = str(row['Número de concepto'])
                            importe = row['Importe liquidado']
                            cantidad = row['Cantidad liquidada']
                        except KeyError as e:
                            st.error(f"❌ Falta la columna {e} en la Sábana.")
                            st.stop()
                        
                        if cod_sistema in mapeo_conceptos_db:
                            cod_afip = mapeo_conceptos_db[cod_sistema]['codigo_afip'].zfill(6)
                            tipo_concepto = mapeo_conceptos_db[cod_sistema]['tipo']
                        else:
                            continue 
                        
                        indicador_dc = 'D' if tipo_concepto == 3 else 'C'
                        cant_formateada = form_cant(cantidad)
                        imp_formateado = form_imp(importe)
                        
                        registro_03 = f"03{cuil_limpio}{cod_afip}{cant_formateada}  {imp_formateado}{indicador_dc}"
                        lineas_txt.append(registro_03)

                # --- FINAL ---
                texto_final = "\n".join(lineas_txt)
                
                st.success("✅ ¡Liquidación procesada con éxito!")
                st.download_button(
                    label="⬇️ Descargar archivo LSD (.txt)",
                    data=texto_final,
                    file_name=f"LSD_{cuit_empresa}_{periodo}.txt",
                    mime="text/plain",
                    type="primary"
                )
                
                with st.expander("👀 Ver previsualización del archivo"):
                    st.code(texto_final[:1500] + "\n... (mostrando los primeros caracteres)")

            except Exception as e:
                st.error(f"❌ Ocurrió un error inesperado: {e}")
