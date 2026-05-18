import streamlit as st
import pandas as pd
import json
import os
import unicodedata
from datetime import datetime
from dateutil.relativedelta import relativedelta

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Generador LSD - AFIP", page_icon="📝", layout="wide")

DB_EMPLEADOS = "repositorio_empleados.json"

def cargar_db():
    if os.path.exists(DB_EMPLEADOS):
        try:
            with open(DB_EMPLEADOS, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if content:
                    return json.loads(content)
        except:
            pass
    return {}

def guardar_db(datos):
    with open(DB_EMPLEADOS, "w", encoding="utf-8") as f:
        json.dump(datos, f, indent=4, ensure_ascii=False)

def limpiar_cuit(cuit):
    c = str(cuit).strip()
    if c.endswith('.0'): c = c[:-2]
    return c.replace("-", "").replace(" ", "").replace(".", "").zfill(11)

def clean_column_name(col):
    if pd.isna(col): return ""
    text = str(col).strip().upper()
    text = ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')
    return text.replace('.', '')

def buscar_columna(df, palabras_clave):
    for col in df.columns:
        c_str = str(col).upper().replace('.', '').replace(' ', '')
        for pk in palabras_clave:
            pk_clean = pk.upper().replace('.', '').replace(' ', '')
            if pk_clean in c_str:
                return col
    return None

def leer_archivo_sueldos(archivo, palabra_clave="CUIL"):
    nombre = archivo.name.lower()
    if nombre.endswith('.xlsx') or nombre.endswith('.xls'):
        xl = pd.ExcelFile(archivo)
        for hoja in xl.sheet_names:
            df = pd.read_excel(archivo, sheet_name=hoja, dtype=str)
            for col in df.columns:
                if palabra_clave in str(col).upper().replace('.', ''):
                    return df, hoja
        return pd.read_excel(archivo, sheet_name=xl.sheet_names[0], dtype=str), xl.sheet_names[0]
    else:
        try:
            df = pd.read_csv(archivo, encoding='utf-8', sep=',', dtype=str)
        except:
            archivo.seek(0)
            df = pd.read_csv(archivo, encoding='latin1', sep=';', dtype=str)
        return df, "CSV"

# ==========================================
# INTERFAZ DE USUARIO CON PESTAÑAS
# ==========================================
st.title("Generador de Libro de Sueldos Digital (LSD)")
st.markdown("Estructura profesional para la generación masiva de declaraciones de AFIP.")

tab_liq, tab_repo = st.tabs(["📊 Procesar Liquidación Mensual", "🗂️ Repositorio de Legajos (Datos Maestros)"])

# Cargar la base histórica al inicio
db_historica = cargar_db()

# --- PESTAÑA 2: GESTIÓN DEL REPOSITORIO ---
with tab_repo:
    st.subheader("Gestión del Repositorio de Empleados")
    st.markdown("Este módulo almacena los datos estables de los trabajadores (como CBU y Forma de Pago) para evitar depender de cruces de archivos manuales todos los meses.")
    
    # Importador masivo
    with st.expander("📥 Importar o Actualizar desde un Listado Maestro (Excel/CSV)"):
        archivo_maestro = st.file_uploader("Subir listado de empleados", type=["csv", "xls", "xlsx"], key="repo_upload")
        if archivo_maestro:
            if st.button("Procesar e Importar al Repositorio"):
                try:
                    df_m, _ = leer_archivo_sueldos(archivo_maestro, "CUIL")
                    
                    col_cuil = buscar_columna(df_m, ['CUIL'])
                    col_legajo = buscar_columna(df_m, ['LEGAJO'])
                    col_dep = buscar_columna(df_m, ['DEPENDENCIA', 'REVISTA', 'LUGAR'])
                    col_cbu = buscar_columna(df_m, ['CBU'])
                    col_fp = buscar_columna(df_m, ['FORMA PAGO', 'PAGO'])
                    
                    if not col_cuil:
                        st.error("❌ No se detectó ninguna columna que contenga 'CUIL' en el archivo maestro.")
                    else:
                        contador = 0
                        for _, row in df_m.iterrows():
                            c_raw = row.get(col_cuil)
                            if pd.isna(c_raw): continue
                            c_limpio = limpiar_cuit(c_raw)
                            if not c_limpio or len(c_limpio) != 11: continue
                            
                            legajo = str(row.get(col_legajo, '')).replace('.0', '').strip() if col_legajo else ""
                            dep = str(row.get(col_dep, '')).strip() if col_dep else ""
                            cbu = str(row.get(col_cbu, '')).strip().replace('-', '').replace(' ', '') if col_cbu else ""
                            fp = str(row.get(col_fp, '')).replace('.0', '').strip() if col_fp else ""
                            
                            if legajo.lower() == 'nan': legajo = ""
                            if dep.lower() == 'nan': dep = ""
                            if cbu.lower() == 'nan': cbu = ""
                            if fp.lower() == 'nan': fp = ""
                            
                            db_historica[c_limpio] = {
                                "legajo": legajo,
                                "dependencia": dep,
                                "cbu": cbu,
                                "forma_pago": fp
                            }
                            contador += 1
                        
                        guardar_db(db_historica)
                        st.success(f"🎉 Se importaron/actualizaron {contador} empleados correctamente.")
                        st.rerun()
                except Exception as e:
                    st.error(f"Error al importar archivo: {e}")
                    
    st.divider()
    
    # Editor Manual en Pantalla
    st.markdown("### 🗂️ Base de Datos Registrada")
    if db_historica:
        data_for_df = []
        for cuil, datos in db_historica.items():
            data_for_df.append({
                "CUIL": cuil,
                "Legajo": datos.get("legajo", ""),
                "Dependencia": datos.get("dependencia", ""),
                "CBU": datos.get("cbu", ""),
                "Forma Pago (1 o 3)": datos.get("forma_pago", "3")
            })
        df_repo = pd.DataFrame(data_for_df)
        
        st.caption("Podés editar los valores directamente sobre las celdas o agregar nuevas filas abajo de todo:")
        df_editado = st.data_editor(df_repo, num_rows="dynamic", use_container_width=True, key="editor_tabla_repo")
        
        if st.button("💾 Guardar Modificaciones Manuales", type="primary"):
            nuevo_dict = {}
            for _, r in df_editado.iterrows():
                c_clean = limpiar_cuit(r["CUIL"])
                if c_clean and len(c_clean) == 11:
                    nuevo_dict[c_clean] = {
                        "legajo": str(r["Legajo"]).strip(),
                        "dependencia": str(r["Dependencia"]).strip(),
                        "cbu": str(r["CBU"]).strip().replace('-', '').replace(' ', ''),
                        "forma_pago": str(r["Forma Pago (1 o 3)"]).strip()
                    }
            guardar_db(nuevo_dict)
            st.success("✅ ¡Repositorio actualizado con éxito!")
            st.rerun()
    else:
        st.info("El repositorio está vacío. Podés completarlo subiendo tu listado maestro arriba o ingresando registros manuales.")
        df_vacio = pd.DataFrame(columns=["CUIL", "Legajo", "Dependencia", "CBU", "Forma Pago (1 o 3)"])
        df_editado = st.data_editor(df_vacio, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Guardar Primer Registro"):
            nuevo_dict = {}
            for _, r in df_editado.iterrows():
                c_clean = limpiar_cuit(r["CUIL"])
                if c_clean and len(c_clean) == 11:
                    nuevo_dict[c_clean] = {
                        "legajo": str(r["Legajo"]).strip(),
                        "dependencia": str(r["Dependencia"]).strip(),
                        "cbu": str(r["CBU"]).strip(),
                        "forma_pago": str(r["Forma Pago (1 o 3)"]).strip()
                    }
            if nuevo_dict:
                guardar_db(nuevo_dict)
                st.success("✅ Primer registro guardado.")
                st.rerun()


# --- PESTAÑA 1: PROCESAMIENTO DE LIQUIDACIÓN ---
with tab_liq:
    st.subheader("Parámetros Generales de Envío")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        cuit_empresa_input = st.text_input("C.U.I.T. de la Empresa", max_chars=13, value="30-64496559-3")
        periodo = st.text_input("Período (AAAAMM)", max_chars=6, value="202604")
    with col_p2:
        tipo_liq = st.selectbox("Tipo de Liquidación", ["M - Mensual", "Q - Quincenal", "S - Semanal"])
        nro_liq = st.text_input("Número de Liquidación", max_chars=5, value="1")
    with col_p3:
        dias_base = st.text_input("Días Base (F931)", max_chars=2, value="00")
        
    st.divider()
    
    st.markdown("#### Fechas Requeridas para el Registro 02")
    col_f1, col_f2 = st.columns(2)
    fecha_pago_defecto = datetime.today()
    try:
        if len(periodo) == 6:
            año = int(periodo[:4])
            mes = int(periodo[4:])
            fecha_pago_defecto = datetime(año, mes, 1) - relativedelta(months=1)
            fecha_pago_defecto = fecha_pago_defecto.replace(day=5)
    except:
        pass
        
    with col_f1:
        fecha_pago = st.date_input("Fecha de Pago (Calculada al 5 del mes anterior)", value=fecha_pago_defecto)
    with col_f2:
        fecha_rubrica = st.date_input("Fecha de Rúbrica", value=datetime.today())
        
    st.divider()
    
    st.markdown("#### Insumo Mensual")
    archivo_liq = st.file_uploader("Subir Sábana de Liquidación ('Conceptos y Totales')", type=["csv", "xls", "xlsx"], key="liq_main")
    
    if archivo_liq:
        try:
            df_liq, hoja_liq = leer_archivo_sueldos(archivo_liq, "CUIL")
            
            col_cuil_liq = buscar_columna(df_liq, ['CUIL'])
            col_legajo_liq = buscar_columna(df_liq, ['LEGAJO'])
            col_lugar_liq = buscar_columna(df_liq, ['LUGAR DE TRABAJO', 'DEPENDENCIA', 'SUCURSAL'])
            
            if not col_cuil_liq:
                st.error("❌ No se encontró la columna de CUIL en el archivo de liquidación cargado.")
                st.stop()
                
            df_liq_validos = df_liq.dropna(subset=[col_cuil_liq]).copy()
            empleados_mes = df_liq_validos[col_cuil_liq].unique()
            cantidad_empleados = len(empleados_mes)
            
            # --- ADUANA DE CONTROL PARA NUEVOS / TEMPORARIOS ---
            cuils_nuevos = [limpiar_cuit(c) for c in empleados_mes if limpiar_cuit(c) not in db_historica]
            
            if cuils_nuevos:
                st.warning(f"⚠️ Se detectaron {len(cuils_nuevos)} empleados de temporada o ingresantes nuevos que no figuran en el repositorio.")
                st.markdown("Completá sus datos mínimos comerciales a continuación para registrarlos e incluirlos en el proceso:")
                
                nuevos_data = []
                for c in cuils_nuevos:
                    sub_df = df_liq_validos[df_liq_validos[col_cuil_liq].apply(limpiar_cuit) == c]
                    leg_pre = ""
                    lug_pre = ""
                    if not sub_df.empty:
                        if col_legajo_liq: leg_pre = str(sub_df.iloc[0].get(col_legajo_liq, '')).replace('.0', '').strip()
                        if col_lugar_liq: lug_pre = str(sub_df.iloc[0].get(col_lugar_liq, '')).strip()
                    
                    nuevos_data.append({
                        "CUIL": c,
                        "Legajo": leg_pre if leg_pre and leg_pre.lower() != 'nan' else "",
                        "Lugar de Trabajo": lug_pre if lug_pre and lug_pre.lower() != 'nan' else "ADMINISTRACION",
                        "CBU (22 dígitos)": "",
                        "Forma Pago (1=Efectivo, 3=CBU)": "3" if leg_pre else "1"
                    })
                df_nuevos_form = pd.DataFrame(nuevos_data)
                editado_nuevos = st.data_editor(df_nuevos_form, use_container_width=True, hide_index=True, key="tabla_nuevos_mes")
                
                if st.button("💾 Registrar Nuevos e Incorporar", type="primary"):
                    for _, r in editado_nuevos.iterrows():
                        c_new = r["CUIL"]
                        db_historica[c_new] = {
                            "legajo": str(r["Legajo"]).strip(),
                            "dependencia": str(r["Lugar de Trabajo"]).strip(),
                            "cbu": str(r["CBU (22 dígitos)"]).strip().replace('-', '').replace(' ', ''),
                            "forma_pago": str(r["Forma Pago (1=Efectivo, 3=CBU)"]).strip()
                        }
                    guardar_db(db_historica)
                    st.sidebar.success("✅ Personal incorporado.")
                    st.rerun()
                    
            else:
                st.success(f"✅ Los {cantidad_empleados} empleados activos este mes están correctamente identificados en el sistema.")
                
                if st.button("🚀 Procesar y Generar Archivo de Importación TXT", type="primary"):
                    lineas_txt = []
                    
                    # ---- REGISTRO 01 ----
                    cuit_l = limpiar_cuit(cuit_empresa_input)
                    t_liq = tipo_liq[0]
                    r01 = f"01{cuit_l}SJ{periodo}{t_liq}{str(nro_liq).zfill(5)}{str(dias_base).zfill(2)}{str(cantidad_empleados).zfill(6)}"
                    lineas_txt.append(r01)
                    
                    f_pago_txt = fecha_pago.strftime("%Y%m%d")
                    f_rubrica_txt = fecha_rubrica.strftime("%Y%m%d")
                    
                    tabla_control = []
                    
                    # ---- REGISTRO 02 ----
                    for cuil_raw in empleados_mes:
                        cuil_l = limpiar_cuit(cuil_raw)
                        
                        sub_df = df_liq_validos[df_liq_validos[col_cuil_liq] == cuil_raw]
                        row_liq = sub_df.iloc[0]
                        
                        legajo_s = str(row_liq.get(col_legajo_liq, '')).replace('.0', '').strip() if col_legajo_liq else ""
                        lugar_s = str(row_liq.get(col_lugar_liq, '')).strip() if col_lugar_liq else ""
                        
                        if legajo_s.lower() == 'nan': legajo_s = ""
                        if lugar_s.lower() == 'nan': lugar_s = ""
                        
                        emp_repo = db_historica.get(cuil_l, {})
                        
                        # Legajo (Sábana manda, respalda repositorio)
                        legajo_final = legajo_s if legajo_s else emp_repo.get('legajo', '0')
                        legajo_fixed = legajo_final.rjust(10, ' ')
                        
                        # Dependencia (Sábana manda, respalda repositorio)
                        dep_final = lugar_s if lugar_s else emp_repo.get('dependencia', 'ADMINISTRACION')
                        dependencia_fixed = dep_final.ljust(50, ' ')[:50]
                        
                        # CBU (Repositorio)
                        cbu_final = emp_repo.get('cbu', '')
                        cbu_fixed = cbu_final if len(cbu_final) == 22 else (' ' * 22)
                        
                        # Forma de Pago (Repositorio)
                        fp_final = emp_repo.get('forma_pago', '')
                        if fp_final not in ['1', '2', '3', '4']:
                            fp_final = '3' if len(cbu_fixed.strip()) == 22 else '1'
                            
                        dias_tope_fixed = str(dias_base).zfill(3)
                        
                        # Armado final Registro 02
                        r02 = f"02{cuil_l}{legajo_fixed}{dependencia_fixed}{cbu_fixed}{dias_tope_fixed}{f_pago_txt}{f_rubrica_txt}{fp_final}"
                        lineas_txt.append(r02)
                        
                        tabla_control.append({
                            "CUIL": cuil_l,
                            "Legajo": f"'{legajo_fixed}'",
                            "Lugar de Trabajo": dependencia_fixed.strip(),
                            "CBU": cbu_fixed if cbu_fixed.strip() else "[Efectivo]",
                            "Forma Pago": fp_final,
                            "Largo": len(r02)
                        })
                    
                    texto_final = "\n".join(lineas_txt)
                    st.success("✅ ¡Estructura de importación completada sin alteraciones!")
                    
                    st.markdown("### 🔍 Vista Previa de Control")
                    st.dataframe(pd.DataFrame(tabla_control), use_container_width=True)
                    
                    with st.expander("👀 Ver Bloque TXT Plano"):
                        st.code(texto_final)
                        
                    st.download_button(
                        label="⬇️ Descargar Archivo LSD Validado",
                        data=texto_final,
                        file_name=f"LSD_Final_{periodo}.txt",
                        mime="text/plain",
                        type="primary"
                    )
        except Exception as e:
            st.error(f"Error durante el procesamiento mensual: {e}")
