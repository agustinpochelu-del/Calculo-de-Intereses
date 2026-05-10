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
    "no_remunerativos": (200, 399),  # Reemplazar con tus números reales
    "retenciones_ss": (400, 599),    # Reemplazar con tus números reales
    "descuentos_varios": (600, 799), # Reemplazar con tus números reales
    "ignorados": (800, 999)          # Conceptos que el sistema escupe pero AFIP no quiere
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
        elif RANGOS["descuentos_varios"][0] <= codigo <= RANGOS["descuentos_varios"][1]:
            return "DESCUENTO"
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
