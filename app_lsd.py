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
    
    if len(df.columns) < 4:
        st.error("❌ Archivo incorrecto. Asegurate de subir el archivo 'Conceptos_Contribuyente...txt'.")
        return {}
    
    # Col 2: Código sistema | Col 3: Desc sistema | Col 0: Código AFIP | Col 1: Desc AFIP
    df_mapeo = df.iloc[:, [2, 3, 0, 1]].copy()
    df_mapeo.columns = ['codigo_sistema', 'descripcion_sistema', 'codigo_afip', 'descripcion_afip']
    df_mapeo = df_mapeo.dropna(subset=['codigo_sistema'])
    
    # CONVENCIÓN: 1 = Remunerativo, 2 = No Remunerativo, 3 = Retención, 0 = Ignorado
    def clasificar_por_afip(cod_afip):
        if pd.isna(cod_afip): return 0
        if cod_afip.startswith('1') or cod_afip.startswith('2'): return 1
        if cod_afip.startswith('5'): return 2
        if cod_afip.startswith('8'): return 3
        return 0
        
    df_mapeo['tipo'] = df_mapeo['codigo_afip'].apply(clasificar_por_afip)
    
    diccionario_mapeo = df_mapeo.set_index('codigo_sistema').to_dict('index')
    return diccionario_mapeo

# =================================================================
# --- 1.5 FUNCIONES FORMATEADORAS PARA EL TXT DE AFIP ---
# =================================================================

def limpiar_cuit_cuil(cuit_cuil):
    """Quita guiones y espacios, devuelve 11 dígitos."""
    return str(cuit_cuil).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

def form_imp(valor, enteros=13, decimales=2):
    """Formatea importes sin coma y rellena con ceros a la izquierda (Ej: 15 dígitos)"""
    if pd.isna(valor): valor = 0
    total_len = enteros + decimales
    val_int = int(round(float(valor), decimales) * (10**decimales))
    return str(abs(val_int)).zfill(total_len)

def form_cant(valor, longitud=5):
    """Formatea cantidades (enteros) rellenando con ceros."""
    if pd.isna(valor): valor = 0
    return str(int(float(valor))).zfill(longitud)


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
    if st.sidebar.button("Pro
