import streamlit as st
import pandas as pd
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
# Usamos "wide" para que la app ocupe todo el ancho de la pantalla
st.set_page_config(page_title="Liquidador ARCA", page_icon="📊", layout="wide")

# --- DISEÑO: ESTILOS DE AZULES SOBRIOS ---
st.markdown("""
<style>
    /* Fondo de la aplicación gris/azul muy claro */
    .stApp {
        background-color: #F0F4F8;
    }
    /* Color de los títulos */
    h1, h2, h3 {
        color: #102A43 !important;
    }
    /* Estilo de los números grandes en el resumen */
    [data-testid="stMetricValue"] {
        color: #1E3A8A;
        font-weight: bold;
    }
    /* Estilo del botón de descarga */
    .stDownloadButton button {
        background-color: #2C3E50;
        color: white;
        border: none;
        border-radius: 5px;
    }
    .stDownloadButton button:hover {
        background-color: #1A252F;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Liquidador de Intereses ARCA")
st.markdown("Subí tu archivo de Excel para generar el tablero de liquidación actualizado.")

# --- ZONA DE SUBIDA DE ARCHIVO ---
archivo_subido = st.file_uploader("Arrastrá tu Excel aquí", type=["xlsx"])

# Función cortita para formato moneda argentino ($ 1.000,50)
def formato_arg(numero):
    return f"${numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

if archivo_subido is not None:
    try:
        with st.spinner('Procesando liquidación...'):
            
            # 1. CARGA Y LIMPIEZA
            df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
            df_tasas = pd.read_excel(archivo_subido, sheet_name='Tasas')
            
            df_deudas.columns = df_deudas.columns.str.strip()
            df_tasas.columns = df_tasas.columns.str.strip()
            df_deudas = df_deudas.dropna(subset=['Vencimiento'])
            df_tasas = df_tasas[df_tasas['Desde'].notna() & (df_tasas['Desde'] != 'Total')]
            
            df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
            df_deudas['fecha_pago'] = pd.to_datetime(df_deudas['fecha_pago'])
            df_tasas['Desde'] = pd.to_datetime(df_tasas['Desde'])
            df_tasas['Hasta'] = pd.to_datetime(df_tasas['Hasta'])
            
            # 2. FUNCIÓN INTELIGENTE
            def calcular_interes_arca_inteligente(vencimiento, capital, fecha_pago):
                interes_acumulado = 0
                for _, tramo in df_tasas.iterrows():
                    inicio = max(vencimiento, tramo['Desde'])
                    fin = min(fecha_pago, tramo['Hasta'])
                    
                    if inicio < fin:
                        dias_interseccion = (fin - inicio).days
                        es_tramo_completo = (vencimiento <= tramo['Desde']) and (fecha_pago >= tramo['Hasta'])
                        
                        if es_tramo_completo:
                            dias_tramo_python = (tramo['Hasta'] - tramo['Desde']).days
                            ajuste = tramo['dias'] - dias_tramo_python
                            dias_finales = max(0, dias_interseccion + ajuste)
                        else:
                            dias_finales = max(0, dias_interseccion)
                            
                        interes_del_tramo = capital * (tramo['Tasa_Diaria'] * dias_finales)
                        interes_acumulado += round(interes_del_tramo, 2)
                        
                return round(interes_acumulado, 2)
            
            # 3. CÁLCULO
            df_deudas['Interes_Resarcitorio'] = df_deudas.apply(
                lambda x: calcular_interes_arca_inteligente(x['Vencimiento'], x['Capital'], x['fecha_pago']), axis=1
            )
            df_deudas['Total_Actualizado'] = df_deudas['Capital'] + df_deudas['Interes_Resarcitorio']
            
            df_deudas['Vencimiento'] = df_deudas['Vencimiento'].dt.strftime('%d/%m/%Y')
            df_deudas['fecha_pago'] = df_deudas['fecha_pago'].dt.strftime('%d/%m/%Y')
            
            # --- ESTRUCTURA DEL DASHBOARD ---
            st.success("¡Liquidación calculada con éxito!")
            st.markdown("### 📋 Resumen de Liquidación")
            
            # Totales para los cuadros
            tot_capital = df_deudas['Capital'].sum()
            tot_interes = df_deudas['Interes_Resarcitorio'].sum()
            tot_pagar = df_deudas['Total_Actualizado'].sum()
            
            # Tres columnas para los indicadores
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Capital Original", formato_arg(tot_capital))
            with col2:
                st.metric("Intereses Resarcitorios", formato_arg(tot_interes))
            with col3:
                st.metric("TOTAL A PAGAR", formato_arg(tot_pagar))
                
            st.divider()
            
            # Tabla gigante oculta en un desplegable
            with st.expander("🔍 Ver detalle de las obligaciones procesadas", expanded=False):
                # Le decimos que ocupe todo el ancho disponible
                st.dataframe(
                    df_deudas[['Impuesto', 'Vencimiento', 'Capital', 'Interes_Resarcitorio', 'Total_Actualizado']],
                    use_container_width=True
                )
            
            # --- BOTÓN DE DESCARGA ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_deudas.to_excel(writer, index=False, sheet_name='Liquidacion_Final')
            procesado = output.getvalue()
            
            st.markdown("<br>", unsafe_allow_html=True)
            # Centramos el botón poniéndolo en la columna del medio
            c_vacia1, c_boton, c_vacia2 = st.columns([1, 2, 1])
            with c_boton:
                st.download_button(
                    label="📥 Descargar Liquidación Completa (Excel)",
                    data=procesado,
                    file_name="Liquidacion_ARCA_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Error técnico al procesar. Detalles: {e}")
