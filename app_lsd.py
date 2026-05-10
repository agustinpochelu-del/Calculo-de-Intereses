import streamlit as st
import pandas as pd
import json
import os

# --- 1. FUNCIONES DE REPOSITORIO (Cerebro de la app) ---

def cargar_topes():
    if os.path.exists("topes_historicos.json"):
        with open("topes_historicos.json", "r") as file:
            return json.load(file)
    return {}

def guardar_topes(diccionario_topes):
    with open("topes_historicos.json", "w") as file:
        json.dump(diccionario_topes, file, indent=4)

# =================================================================
# --- 2. LIBRERÍA DE RANGOS DE CONCEPTOS (¡Acá va lo que me pedís!) ---
# =================================================================

RANGOS = {
    "remunerativos": (1, 199),       # Reemplazar con tus números reales
    "no_remunerativos": (400, 599),  # Reemplazar con tus números reales
    "retenciones_ss": (200, 399),    # Reemplazar con tus números reales
    "ignorados": (600, 999),     # Conceptos que el sistema escupe pero AFIP no quiere
}

def clasificar_concepto(codigo_sistema):
    """Evalúa un código de tu sistema y devuelve a qué familia pertenece."""
    try:
        codigo = int(codigo_sistema)
        if RANGOS["remunerativos"][0] <= codigo <= RANGOS["remunerativos"][1]:
            return "REMUNERATIVO"
        elif RANGOS["no_remunerativos"][0] <= codigo <= RANGOS["no_remunerativos"][1]:
            return "NO_REMUNERATIVO"
        elif RANGOS["retenciones_ss"][0] <= codigo <= RANGOS["retenciones_ss"][1]:
            return "RETENCION"
        elif RANGOS["ignorados"][0] <= codigo <= RANGOS["ignorados"][1]:
            return "IGNORADO"
        else:
            return "DESCONOCIDO"
    except ValueError:
        return "DESCONOCIDO" # Por si viene algún código con letras

# =================================================================

# --- 3. CONFIGURACIÓN DE LA PÁGINA Y UI (Lo que sigue igual...) ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

# ... (acá sigue todo el resto del código de la barra lateral y carga de archivos) ...

def cargar_topes():
    """Carga los topes históricos desde el archivo JSON."""
    if os.path.exists("topes_historicos.json"):
        with open("topes_historicos.json", "r") as file:
            return json.load(file)
    return {} # Si no existe, devuelve un diccionario vacío

def guardar_topes(diccionario_topes):
    """Guarda los topes actualizados en el archivo JSON."""
    with open("topes_historicos.json", "w") as file:
        json.dump(diccionario_topes, file, indent=4)

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("Herramienta para convertir liquidaciones de Excel al formato `.txt` de AFIP.")

# --- BARRA LATERAL: DATOS GLOBALES Y TOPES ---
st.sidebar.header("1. Parámetros Globales")

# Cargamos la base de datos de topes
topes_db = cargar_topes()

# El usuario ingresa el período
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, placeholder="Ej: 202604")

# Lógica inteligente de Topes (Graceful Fallback)
tope_min = 0.0
tope_max = 0.0

if periodo:
    if len(periodo) == 6: # Validación simple para que no busque mientras tipeás el año
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
                guardar_topes(topes_db)
                st.sidebar.success("¡Topes guardados exitosamente!")
                st.rerun() # Recarga la app para aplicar los cambios
    else:
        st.sidebar.info("Ingresá los 6 dígitos del período.")

st.sidebar.divider()

# Resto de los parámetros globales
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, placeholder="Ej: 1", value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

st.sidebar.divider()
st.sidebar.header("Fechas (Registro 02)")
fecha_pago = st.sidebar.date_input("Fecha de Pago")
fecha_rubrica = st.sidebar.date_input("Fecha de Rúbrica (Opcional)", value=None)

# --- ÁREA PRINCIPAL: CARGA DE ARCHIVOS ---
st.subheader("2. Carga de Datos")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**A. Liquidación del Sistema**")
    archivo_liq = st.file_uploader("Subir Excel/CSV de Liquidación", type=["xls", "xlsx", "csv"])

with col2:
    st.markdown("**B. Base de Empleados (Opcional)**")
    archivo_emp = st.file_uploader("Subir Excel con CBU/Forma de Pago", type=["xls", "xlsx", "csv"])

st.divider()

# --- PROCESAMIENTO (Botón de Acción) ---
if st.button("Procesar y Generar TXT", type="primary"):
    if archivo_liq is None:
        st.warning("⚠️ Por favor, subí el archivo de liquidación para comenzar.")
    elif not periodo or len(periodo) != 6:
        st.warning("⚠️ El Período es obligatorio y debe tener 6 dígitos.")
    elif tope_min == 0 or tope_max == 0:
        st.error("❌ Faltan los topes para este período. Cargalos en la barra lateral y guardalos.")
    else:
        st.info("🔄 Procesando datos... (Acá irá el motor de cálculo en el próximo paso)")
