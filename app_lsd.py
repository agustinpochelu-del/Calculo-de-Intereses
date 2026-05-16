import streamlit as st
import pandas as pd

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

def limpiar_cuit(cuit):
    return str(cuit).replace("-", "").replace(" ", "").replace(".", "").zfill(11)

st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("### PASO 1: Registro 01 con Conteo Inteligente")

# --- BARRA LATERAL ---
st.sidebar.header("Datos para Registro 01")
cuit_empresa_input = st.sidebar.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
periodo = st.sidebar.text_input("Período (AAAAMM)", max_chars=6, value="202604")
tipo_liq = st.sidebar.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
nro_liq = st.sidebar.text_input("Número de Liquidación", max_chars=5, value="1")
dias_base = st.sidebar.text_input("Días Base (F931)", max_chars=2, value="30")

# --- ÁREA PRINCIPAL ---
st.markdown("Subí tu planilla de liquidación. El sistema buscará automáticamente la columna **C.U.I.L.** para contar los empleados.")
archivo_liq = st.file_uploader("Subir Sábana de Liquidación", type=["xls", "xlsx"])

# --- PROCESAMIENTO ---
if st.button("Generar TXT (Solo Registro 01)", type="primary"):
    if not cuit_empresa_input or not periodo or archivo_liq is None:
        st.warning("⚠️ Faltan cargar datos o subir el archivo.")
    else:
        try:
            # 1. LECTURA CON VISIÓN LÁSER (Buscamos en qué fila están los títulos)
            # Leemos el Excel sin definir cabecera para escanearlo
            df_temp = pd.read_excel(archivo_liq, header=None)
            
            fila_titulos = None
            # Buscamos la fila exacta que tiene la palabra "C.U.I.L."
            for index, row in df_temp.iterrows():
                # Convertimos la fila a texto y buscamos
                if row.astype(str).str.contains('C.U.I.L.').any():
                    fila_titulos = index
                    break
            
            if fila_titulos is None:
                st.error("❌ No encontré la palabra 'C.U.I.L.' en ninguna fila del archivo.")
                st.stop()
                
            # 2. Volvemos a leer el archivo, pero arrancando exactamente desde los títulos
            archivo_liq.seek(0)
            df_liq = pd.read_excel(archivo_liq, header=fila_titulos)
            df_liq.columns = df_liq.columns.str.strip() # Limpiamos espacios
            
            # Contamos los CUILs únicos (dropna() ignora celdas vacías por si hay filas de totales abajo)
            cantidad_empleados = df_liq['C.U.I.L.'].dropna().nunique()
            
            # 3. ARMADO DEL REGISTRO 01
            cuit_limpio = limpiar_cuit(cuit_empresa_input)
            tipo_liq_letra = tipo_liq[0]
            nro_liq_form = str(nro_liq).zfill(5)
            dias_base_form = str(dias_base).zfill(2)
            cant_empleados_form = str(cantidad_empleados).zfill(6)
            
            # Formato AFIP: 01 + CUIT(11) + SJ + Periodo(6) + Tipo(1) + Nro(5) + Dias(2) + CantEmpleados(6)
            registro_01 = f"01{cuit_limpio}SJ{periodo}{tipo_liq_letra}{nro_liq_form}{dias_base_form}{cant_empleados_form}"
            
            st.success(f"✅ ¡Éxito! Se detectaron automáticamente **{cantidad_empleados} empleados únicos**.")
            
            # Mostramos la línea en pantalla
            st.code(registro_01)
            st.caption(f"Longitud de la línea: {len(registro_01)} caracteres (El manual exige exactamente 35).")
            
            st.download_button(
                label="⬇️ Descargar TXT de prueba",
                data=registro_01,
                file_name=f"LSD_R01_{periodo}.txt",
                mime="text/plain"
            )

        except Exception as e:
            st.error(f"Ocurrió un error al leer el archivo: {e}")
