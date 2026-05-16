import streamlit as st
import pandas as pd
import json
import os

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

# =================================================================
# --- 1. FUNCIONES DE REPOSITORIO Y PROCESAMIENTO ---
# =================================================================

def cargar_json(nombre_archivo):
    """Función genérica para cargar archivos JSON (Topes o Conceptos)"""
    if os.path.exists(nombre_archivo):
        with open(nombre_archivo, "r") as file:
            return json.load(file)
    return {}

def guardar_json(nombre_archivo, datos):
    """Función genérica para guardar archivos JSON"""
    with open(nombre_archivo, "w") as file:
        json.dump(datos, file, indent=4)

def procesar_txt_afip(archivo_txt_afip):
    """Lee el TXT de AFIP, cruza los conceptos y los clasifica (1, 2, 3)"""
    df = pd.read_csv(archivo_txt_afip, sep=';', dtype=str, encoding='latin1')
    
    # Nos quedamos solo con las columnas que nos importan
    columnas_clave = ['Código contribuyente', 'Descripción.1', 'Código AFIP', 'Descripción']
    df_mapeo = df[columnas_clave].copy()
    
    # Renombramos
    df_mapeo.columns = ['codigo_sistema', 'descripcion_sistema', 'codigo_afip', 'descripcion_afip']
    df_mapeo = df_mapeo.dropna(subset=['codigo_sistema'])
    
    # NUEVA LÓGICA NUMÉRICA (1, 2, 3)
    def clasificar_por_afip(cod_afip):
        if not isinstance(cod_afip, str): return 0 # 0 = IGNORADO/DESCONOCIDO
        if cod_afip.startswith('1') or cod_afip.startswith('2'): return 1 # Remunerativo
        if cod_afip.startswith('5'): return 2 # No Remunerativo
        if cod_afip.startswith('8'): return 3 # Retención
        return 0 # Ignorado
        
    df_mapeo['tipo'] = df_mapeo['codigo_afip'].apply(clasificar_por_afip)
    
    # Convertimos a diccionario
    diccionario_mapeo = df_mapeo.set_index('codigo_sistema').to_dict('index')
    return diccionario_mapeo


# =================================================================
# --- 2. INTERFAZ DE USUARIO (UI) ---
# =================================================================

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("Herramienta para convertir liquidaciones de Excel al formato `.txt` de AFIP.")

# --- BARRA LATERAL ---
st.sidebar.header("1. Cerebro de la App")

# Carga de la parametrización de AFIP
st.sidebar.markdown("**A. Actualizar Conceptos AFIP**")
archivo_afip = st.sidebar.file_uploader("Subir TXT de AFIP", type=["txt"])
if archivo_afip:
    if st.sidebar.button("Procesar y Guardar Conceptos"):
        nuevo_mapeo = procesar_txt_afip(archivo_afip)
        guardar_json("mapeo_conceptos.json", nuevo_mapeo)
        st.sidebar.success("✅ ¡Conceptos actualizados con éxito!")
        st.rerun()

st.sidebar.divider()
st.sidebar.header("2. Parámetros Globales")

# Cargamos las bases de datos a la memoria de la app
topes_db = cargar_json("topes_historicos.json")
mapeo_conceptos_db = cargar_json("mapeo_conceptos.json") 

# El usuario ingresa el período
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, placeholder="Ej: 202604")

# Lógica inteligente de Topes
tope_min = 0.0
tope_max = 0.0

if periodo:
    if len(periodo) == 6: 
        if periodo in topes_db:
            tope_min = topes_db[periodo]["min"]
            tope_max = topes_db[periodo]["max"]
            st.sidebar.success(f"✅ Topes de {periodo} cargados automáticamente.")
        else:
            st.sidebar.warning(f"⚠️ Atención: No hay topes registrados para {periodo}.")
            nuevo_min = st.sidebar.number_input("Ingresar Tope Mínimo", min_value=0.0, format="%.2f")
            nuevo_max = st.sidebar.number_input("Ingresar Tope Máximo", min_value=0.0, format="%.2f")
            
            if st.sidebar.button("💾 Guardar Topes en Repositorio"):
                topes_db[periodo] = {"min": nuevo_min, "max": nuevo_max}
                guardar_json("topes_historicos.json", topes_db)
                st.sidebar.success("¡Topes guardados exitosamente!")
                st.rerun() 
    else:
        st.sidebar.info("Ingresá los 6 dígitos del período.")

st.sidebar.divider()

tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, placeholder="Ej: 1", value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

st.sidebar.divider()
st.sidebar.header("Fechas (Registro 02)")
fecha_pago = st.sidebar.date_input("Fecha de Pago")
fecha_rubrica = st.sidebar.date_input("Fecha de Rúbrica (Opcional)", value=None)

# --- ÁREA PRINCIPAL: CARGA DE ARCHIVOS ---
st.subheader("3. Carga de Liquidación Mensual")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**A. Liquidación del Sistema**")
    archivo_liq = st.file_uploader("Subir Excel/CSV de Liquidación", type=["xls", "xlsx", "csv"])

with col2:
    st.markdown("**B. Base de Empleados (Opcional)**")
    archivo_emp = st.file_uploader("Subir Excel con CBU/Forma de Pago", type=["xls", "xlsx", "csv"])

st.divider()

# --- ESTADO DEL SISTEMA ---
# Chequeo para avisarte si te olvidaste de subir los conceptos
if not mapeo_conceptos_db:
    st.warning("⚠️ Todavía no cargaste los conceptos de AFIP. Usá el menú lateral para subir tu archivo .txt exportado de ARCA.")

# --- PROCESAMIENTO ---
if st.button("Procesar y Generar TXT", type="primary"):
    if not mapeo_conceptos_db:
        st.error("❌ Detenido: Falta cargar el mapeo de conceptos de AFIP.")
    elif archivo_liq is None:
        st.warning("⚠️ Por favor, subí el archivo de liquidación para comenzar.")
    elif not periodo or len(periodo) != 6:
        st.warning("⚠️ El Período es obligatorio y debe tener 6 dígitos.")
    elif tope_min == 0 or tope_max == 0:
        st.error("❌ Faltan los topes para este período. Cargalos en la barra lateral y guardalos.")
    else:
        st.info("🔄 Archivos validados. Iniciando cruce de datos...")
        # Acá vendrá el motor que arma las líneas del archivo .txt
