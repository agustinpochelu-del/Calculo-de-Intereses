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
    return str(cuit_cuil).replace("-", "").replace(" ", "").zfill(11)

def form_imp(valor, enteros=13, decimales=2):
    """Formatea importes sin coma y rellena con ceros a la izquierda (Ej: 15 dígitos)"""
    if pd.isna(valor): valor = 0
    total_len = enteros + decimales
    # Redondea, multiplica por 100 para sacar la coma y convierte a entero
    val_int = int(round(float(valor), decimales) * (10**decimales))
    # Para importes negativos (descuentos), usamos valor absoluto porque AFIP usa la D o C al final
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
if not mapeo_conceptos_db:
    st.warning("⚠️ Todavía no cargaste los conceptos de AFIP. Usá el menú lateral para subir tu archivo .txt exportado de ARCA.")


# --- NUEVA SECCIÓN: VISUALIZADOR DE BASES DE DATOS PRECARGADAS ---
st.subheader("📊 Consulta de Datos Guardados en Repositorio")

tab1, tab2 = st.tabs(["📌 Topes Previsionales Guardados", "🗂️ Conceptos Mapeados (AFIP vs Sistema)"])

with tab1:
    if topes_db:
        # Convertimos el diccionario de topes a un DataFrame lindo para mostrar
        df_topes_mostrar = pd.DataFrame.from_dict(topes_db, orient='index')
        df_topes_mostrar.index.name = 'Período'
        df_topes_mostrar.columns = ['Tope Mínimo', 'Tope Máximo']
        # Mostramos la tabla formateada con dos decimales
        st.dataframe(df_topes_mostrar.style.format("{:,.2f}"), use_container_width=True)
    else:
        st.info("Aún no hay topes guardados en el repositorio.")

with tab2:
    if mapeo_conceptos_db:
        # Convertimos el diccionario de conceptos a DataFrame
        df_conceptos_mostrar = pd.DataFrame.from_dict(mapeo_conceptos_db, orient='index')
        df_conceptos_mostrar.index.name = 'Cód. Sistema'
        df_conceptos_mostrar.columns = ['Descripción Sistema', 'Cód. AFIP', 'Descripción AFIP', 'Condición (1,2,3)']
        
        # Mapeo visual para que entienda qué significa el número de condición
        def mapear_nombre_condicion(val):
            if val == 1: return "1 - Remunerativo"
            if val == 2: return "2 - No Remunerativo"
            if val == 3: return "3 - Retención"
            return "0 - Ignorado"
            
        df_conceptos_mostrar['Condición (1,2,3)'] = df_conceptos_mostrar['Condición (1,2,3)'].apply(mapear_nombre_condicion)
        
        st.dataframe(df_conceptos_mostrar, use_container_width=True)
    else:
        st.info("Aún no hay conceptos guardados. Subí el archivo .txt en la barra lateral.")

st.divider()


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
        with st.spinner("🔄 Procesando liquidación..."):
            try:
                # 1. Leemos el archivo (A prueba de balas: detecta coma o punto y coma)
                try:
                    df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=',')
                    # Si no encuentra la columna CUIL, es porque el separador era punto y coma
                    if 'C.U.I.L.' not in df_liq.columns:
                        archivo_liq.seek(0) # Volvemos a poner el archivo al principio
                        df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                except Exception:
                    archivo_liq.seek(0)
                    df_liq = pd.read_csv(archivo_liq, encoding='latin1', sep=';')
                
                # CUIT de tu empresa
                cuit_empresa = limpiar_cuit_cuil("30-64496559-3") 
                
                # Variables para armar el TXT
                lineas_txt = []
                empleados_procesados = df_liq['C.U.I.L.'].unique()
                cantidad_empleados = len(empleados_procesados)
                
                # =================================================================
                # ARMADO REGISTRO 01: Cabecera de la Liquidación (35 caracteres)
                # =================================================================
                tipo_liq_letra = tipo_liq[0] # Saca la 'M', 'Q' o 'S'
                nro_liq_formateado = str(nro_liq).zfill(5)
                dias_base_formateado = str(dias_base).zfill(2)
                cant_empleados_formateado = str(cantidad_empleados).zfill(6)
                
                registro_01 = f"01{cuit_empresa}SJ{periodo}{tipo_liq_letra}{nro_liq_formateado}{dias_base_formateado}{cant_empleados_formateado}"
                lineas_txt.append(registro_01)
                
                # =================================================================
                # RECORREMOS CADA EMPLEADO PARA ARMAR SUS REGISTROS
                # =================================================================
                for cuil in empleados_procesados:
                    df_empleado = df_liq[df_liq['C.U.I.L.'] == cuil]
                    cuil_limpio = limpiar_cuit_cuil(cuil)
                    
                    # (ACÁ IRÁ EL REGISTRO 02 Y EL 04 DE ESTE EMPLEADO EN EL PRÓXIMO PASO)
                    
                    # =================================================================
                    # ARMADO REGISTRO 03: Detalle de Conceptos
                    # =================================================================
                    for index, row in df_empleado.iterrows():
                        cod_sistema = str(row['Número de concepto'])
                        importe = row['Importe liquidado']
                        cantidad = row['Cantidad liquidada']
                        
                        # Buscamos este concepto en tu base de datos de AFIP
                        if cod_sistema in mapeo_conceptos_db:
                            cod_afip = mapeo_conceptos_db[cod_sistema]['codigo_afip'].zfill(6)
                            tipo_concepto = mapeo_conceptos_db[cod_sistema]['tipo']
                        else:
                            # Si es un concepto raro o ignorado, lo salteamos
                            continue 
                        
                        # Definimos Débito (Descuento) o Crédito (Remunerativo/No Remunerativo)
                        indicador_dc = 'D' if tipo_concepto == 3 else 'C'
                        
                        # Formateamos números
                        cant_formateada = form_cant(cantidad)
                        imp_formateado = form_imp(importe)
                        unidades = '  ' # 2 espacios en blanco
                        
                        registro_03 = f"03{cuil_limpio}{cod_afip}{cant_formateada}{unidades}{imp_formateado}{indicador_dc}"
                        lineas_txt.append(registro_03)

                # =================================================================
                # FINAL: MOSTRAMOS EL RESULTADO
                # =================================================================
                texto_final = "\n".join(lineas_txt)
                
                st.success("✅ ¡Liquidación procesada con éxito!")
                st.markdown(f"**Empleados procesados:** {cantidad_empleados}")
                st.markdown(f"**Líneas generadas:** {len(lineas_txt)}")
                
                st.download_button(
                    label="⬇️ Descargar archivo LSD (.txt)",
                    data=texto_final,
                    file_name=f"LSD_{periodo}_{nro_liq}.txt",
                    mime="text/plain",
                    type="primary"
                )
                
                # Previsualización opcional para que veas cómo va quedando
                with st.expander("👀 Ver previsualización del archivo"):
                    st.code(texto_final[:1000] + "\n... (mostrando los primeros caracteres)")

            except Exception as e:
                st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")
