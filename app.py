import streamlit as st
import pandas as pd
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Liquidador ARCA", page_icon="📊")

st.title("📊 Liquidador de Intereses ARCA")
st.markdown("Subí tu archivo Excel con las hojas **Deudas** y **Tasas** para calcular los intereses resarcitorios actualizados con precisión exacta.")

# --- ZONA DE SUBIDA DE ARCHIVO ---
archivo_subido = st.file_uploader("Arrastrá tu Excel aquí", type=["xlsx"])

if archivo_subido is not None:
    try:
        # Mostramos un mensaje de carga mientras piensa
        with st.spinner('Procesando liquidación...'):
            
            # 1. CARGA DE DATOS
            df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
            df_tasas = pd.read_excel(archivo_subido, sheet_name='Tasas')
            
            # Limpieza para evitar errores de espacios en blanco
            df_deudas.columns = df_deudas.columns.str.strip()
            df_tasas.columns = df_tasas.columns.str.strip()
            df_deudas = df_deudas.dropna(subset=['Vencimiento'])
            df_tasas = df_tasas[df_tasas['Desde'].notna() & (df_tasas['Desde'] != 'Total')]
            
            # Conversión de fechas
            df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
            df_deudas['fecha_pago'] = pd.to_datetime(df_deudas['fecha_pago'])
            df_tasas['Desde'] = pd.to_datetime(df_tasas['Desde'])
            df_tasas['Hasta'] = pd.to_datetime(df_tasas['Hasta'])
            
            # 2. LA FUNCIÓN DE CÁLCULO INTELIGENTE
            def calcular_interes_arca_inteligente(vencimiento, capital, fecha_pago):
                interes_acumulado = 0
                for _, tramo in df_tasas.iterrows():
                    inicio = max(vencimiento, tramo['Desde'])
                    fin = min(fecha_pago, tramo['Hasta'])
                    
                    if inicio < fin:
                        dias_interseccion = (fin - inicio).days
                        
                        # Verificamos si la deuda atraviesa el tramo completo
                        es_tramo_completo = (vencimiento <= tramo['Desde']) and (fecha_pago >= tramo['Hasta'])
                        
                        if es_tramo_completo:
                            # Aplica el ajuste manual de días
                            dias_tramo_python = (tramo['Hasta'] - tramo['Desde']).days
                            ajuste = tramo['dias'] - dias_tramo_python
                            dias_finales = dias_interseccion + ajuste
                        else:
                            # Cálculo estricto sin ajuste
                            dias_finales = dias_interseccion
                            
                        # Seguridad para evitar días negativos
                        dias_finales = max(0, dias_finales)
                        
                        # Cálculo y redondeo por tramo
                        interes_del_tramo = capital * (tramo['Tasa_Diaria'] * dias_finales)
                        interes_acumulado += round(interes_del_tramo, 2)
                        
                return round(interes_acumulado, 2)
            
            # 3. EJECUCIÓN
            df_deudas['Interes_Resarcitorio'] = df_deudas.apply(
                lambda x: calcular_interes_arca_inteligente(x['Vencimiento'], x['Capital'], x['fecha_pago']), axis=1
            )
            df_deudas['Total_Actualizado'] = df_deudas['Capital'] + df_deudas['Interes_Resarcitorio']
            
            # Emprolijamos las fechas para mostrarlas en la web y en el Excel final
            df_deudas['Vencimiento'] = df_deudas['Vencimiento'].dt.strftime('%d/%m/%Y')
            df_deudas['fecha_pago'] = df_deudas['fecha_pago'].dt.strftime('%d/%m/%Y')
            
            st.success("¡Liquidación finalizada sin errores!")
            
            # --- MOSTRAR RESULTADOS EN PANTALLA ---
            st.subheader("Vista Previa")
            st.dataframe(df_deudas[['Impuesto', 'Vencimiento', 'Capital', 'Interes_Resarcitorio', 'Total_Actualizado']])
            
            # --- PREPARAR EL EXCEL PARA DESCARGAR ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_deudas.to_excel(writer, index=False, sheet_name='Liquidacion_Final')
            procesado = output.getvalue()
            
            # --- BOTÓN DE DESCARGA ---
            st.download_button(
                label="📥 Descargar Liquidación Completa (Excel)",
                data=procesado,
                file_name="Liquidacion_ARCA_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo. Asegurate de subir el formato correcto. Detalle técnico: {e}")