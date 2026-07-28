import streamlit as st
import pandas as pd
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Liquidador ARCA - Judicial", page_icon="⚖️", layout="wide")

# --- DISEÑO: ESTILOS DE AZULES SOBRIOS ---
st.markdown("""
<style>
    .stApp { background-color: #F0F4F8; }
    h1, h2, h3 { color: #102A43 !important; }
    [data-testid="stMetricValue"] { color: #1E3A8A; font-weight: bold; }
    .stDownloadButton button { background-color: #2C3E50; color: white; border: none; border-radius: 5px; }
    .stDownloadButton button:hover { background-color: #1A252F; color: white; }
</style>
""", unsafe_allow_html=True)

st.title("⚖️ Liquidador ARCA - Ejecución Fiscal")
st.markdown("Subí el Excel con las hojas **Deudas**, **Tasas** y **Tasas Punitorios**.")

archivo_subido = st.file_uploader("Arrastrá tu Excel aquí", type=["xlsx"])

def formato_arg(numero):
    return f"${numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

if archivo_subido is not None:
    try:
        with st.spinner('Procesando liquidación judicial...'):

            # 1. CARGA Y LIMPIEZA
            df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
            df_tasas_res = pd.read_excel(archivo_subido, sheet_name='Tasas')
            df_tasas_pun = pd.read_excel(archivo_subido, sheet_name='Tasas Punitorios')

            df_deudas.columns = df_deudas.columns.str.strip()
            df_tasas_res.columns = df_tasas_res.columns.str.strip()
            df_tasas_pun.columns = df_tasas_pun.columns.str.strip()

            df_tasas_res = df_tasas_res.rename(columns={'Dias': 'dias'})
            df_tasas_pun = df_tasas_pun.rename(columns={'Dias': 'dias'})

            df_deudas = df_deudas.dropna(subset=['Vencimiento'])
            df_tasas_res = df_tasas_res[df_tasas_res['Desde'].notna() & (df_tasas_res['Desde'] != 'Total')]
            df_tasas_pun = df_tasas_pun[df_tasas_pun['Desde'].notna() & (df_tasas_pun['Desde'] != 'Total')]

            # Compatibilidad con planillas viejas que todavía no tienen "F. Pago Capital"
            if 'F. Pago Capital' not in df_deudas.columns:
                df_deudas['F. Pago Capital'] = pd.NaT

            # Conversión de fechas
            df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
            df_deudas['F. Pago Capital'] = pd.to_datetime(df_deudas['F. Pago Capital'])
            df_deudas['fecha_Demanda'] = pd.to_datetime(df_deudas['fecha_Demanda'])
            df_deudas['Fecha_Liquidacion'] = pd.to_datetime(df_deudas['Fecha_Liquidacion'])

            df_tasas_res['Desde'] = pd.to_datetime(df_tasas_res['Desde'])
            df_tasas_res['Hasta'] = pd.to_datetime(df_tasas_res['Hasta'])
            df_tasas_pun['Desde'] = pd.to_datetime(df_tasas_pun['Desde'])
            df_tasas_pun['Hasta'] = pd.to_datetime(df_tasas_pun['Hasta'])

            # Fechas de corte para visualización
            ultima_fecha_demanda = df_deudas['fecha_Demanda'].max()
            ultima_fecha_liq = df_deudas['Fecha_Liquidacion'].max()

            # 2. MOTOR DE CÁLCULO
            def calcular_interes(fecha_inicio_calculo, fecha_fin_calculo, capital, df_tabla_tasas):
                interes_acumulado = 0
                if pd.isna(fecha_inicio_calculo) or pd.isna(fecha_fin_calculo):
                    return 0.0
                if fecha_inicio_calculo >= fecha_fin_calculo:
                    return 0.0
                for _, tramo in df_tabla_tasas.iterrows():
                    inicio = max(fecha_inicio_calculo, tramo['Desde'])
                    fin = min(fecha_fin_calculo, tramo['Hasta'])
                    if inicio < fin:
                        dias_interseccion = (fin - inicio).days
                        es_tramo_completo = (fecha_inicio_calculo <= tramo['Desde']) and (fecha_fin_calculo >= tramo['Hasta'])
                        if es_tramo_completo and 'dias' in tramo:
                            dias_tramo_python = (tramo['Hasta'] - tramo['Desde']).days
                            ajuste = tramo['dias'] - dias_tramo_python
                            dias_finales = max(0, dias_interseccion + ajuste)
                        else:
                            dias_finales = max(0, dias_interseccion)
                        interes_del_tramo = capital * (tramo['Tasa_Diaria'] * dias_finales)
                        interes_acumulado += round(interes_del_tramo, 2)
                return round(interes_acumulado, 2)

            # 3. RESARCITORIOS + PUNITORIOS + CAPITALIZABLES (por fila, dependen entre sí)
            #
            # Los tres tramos son continuos, sin "salto" de un día entre uno y el
            # siguiente (el fin de un tramo es el inicio exacto del próximo).
            # El orden de Punitorios y Capitalizables se invierte según cuál de los
            # dos eventos (Demanda o Pago del Capital) ocurre primero:
            #
            #   Caso A: Demanda ANTES del Pago del Capital (aún debe el capital al demandar)
            #       Resarcitorios:  Vencimiento -> Demanda            (sobre Capital)
            #       Punitorios:     Demanda -> Pago Capital           (sobre Capital)
            #       Capitalizables: Pago Capital -> Liquidación       (sobre monto Resarcitorios)
            #
            #   Caso B: Pago del Capital ANTES de la Demanda (pagó el capital y luego lo demandaron)
            #       Resarcitorios:  Vencimiento -> Pago Capital       (sobre Capital)
            #       Capitalizables: Pago Capital -> Demanda           (sobre monto Resarcitorios)
            #       Punitorios:     Demanda -> Liquidación            (sobre Capital)
            #
            #   Sin fecha de Pago de Capital cargada: se mantiene el comportamiento
            #   clásico de 2 tramos (Resarcitorios hasta la Demanda, Punitorios desde
            #   la Demanda hasta la Liquidación), sin Capitalizables.
            def procesar_intereses(fila):
                vencimiento = fila['Vencimiento']
                pago_capital = fila['F. Pago Capital']
                fecha_demanda = fila['fecha_Demanda']
                fecha_liq = fila['Fecha_Liquidacion']
                capital = fila['Capital']

                if pd.isna(pago_capital):
                    interes_resarcitorio = calcular_interes(vencimiento, fecha_demanda, capital, df_tasas_res)
                    interes_capitalizable = 0.0
                    interes_punitorio = calcular_interes(fecha_demanda, fecha_liq, capital, df_tasas_pun)
                elif fecha_demanda <= pago_capital:
                    # Caso A
                    interes_resarcitorio = calcular_interes(vencimiento, fecha_demanda, capital, df_tasas_res)
                    interes_punitorio = calcular_interes(fecha_demanda, pago_capital, capital, df_tasas_pun)
                    interes_capitalizable = calcular_interes(pago_capital, fecha_liq, interes_resarcitorio, df_tasas_res)
                else:
                    # Caso B
                    interes_resarcitorio = calcular_interes(vencimiento, pago_capital, capital, df_tasas_res)
                    interes_capitalizable = calcular_interes(pago_capital, fecha_demanda, interes_resarcitorio, df_tasas_res)
                    interes_punitorio = calcular_interes(fecha_demanda, fecha_liq, capital, df_tasas_pun)

                return pd.Series({
                    'Interes_Resarcitorio': interes_resarcitorio,
                    'Interes_Capitalizable': interes_capitalizable,
                    'Interes_Punitorio': interes_punitorio
                })

            resultado = df_deudas.apply(procesar_intereses, axis=1)
            df_deudas['Interes_Resarcitorio'] = resultado['Interes_Resarcitorio']
            df_deudas['Interes_Capitalizable'] = resultado['Interes_Capitalizable']
            df_deudas['Interes_Punitorio'] = resultado['Interes_Punitorio']

            fecha_inicio_punitorios_global = ultima_fecha_demanda

            # Antigüedad global y Totales
            antiguedad_juicio = (ultima_fecha_liq - fecha_inicio_punitorios_global).days
            antiguedad_juicio = max(0, antiguedad_juicio)
            df_deudas['Total_Actualizado'] = (
                df_deudas['Capital']
                + df_deudas['Interes_Resarcitorio']
                + df_deudas['Interes_Capitalizable']
                + df_deudas['Interes_Punitorio']
            )

            # --- TABLAS DE REFERENCIA PARA PANTALLA ---
            display_res = df_tasas_res.copy()
            display_pun = df_tasas_pun.copy()

            # Ajuste de cierre resarcitorios (la tabla resarcitoria se usa tanto para
            # resarcitorios como para capitalizables, ambos topeados en fecha_Demanda)
            display_res.iloc[-1, display_res.columns.get_loc('Hasta')] = ultima_fecha_demanda

            # Ajuste inteligente del inicio de punitorios
            display_pun = display_pun[display_pun['Hasta'] >= fecha_inicio_punitorios_global].copy()
            if not display_pun.empty:
                display_pun.iloc[0, display_pun.columns.get_loc('Desde')] = max(display_pun.iloc[0]['Desde'], fecha_inicio_punitorios_global)
                display_pun.iloc[-1, display_pun.columns.get_loc('Hasta')] = ultima_fecha_liq
                # Recalculamos los días expuestos para que coincidan con la liquidación real
                display_pun['dias_calculados'] = (display_pun['Hasta'] - display_pun['Desde']).dt.days
                # Compensamos si el día final es inclusivo
                display_pun['dias'] = display_pun['dias_calculados']

            display_res['Desde'] = display_res['Desde'].dt.strftime('%d/%m/%Y')
            display_res['Hasta'] = display_res['Hasta'].dt.strftime('%d/%m/%Y')
            display_pun['Desde'] = display_pun['Desde'].dt.strftime('%d/%m/%Y')
            display_pun['Hasta'] = display_pun['Hasta'].dt.strftime('%d/%m/%Y')

            # Formateo final de fechas de deudas
            df_deudas['Vencimiento'] = df_deudas['Vencimiento'].dt.strftime('%d/%m/%Y')
            df_deudas['F. Pago Capital'] = df_deudas['F. Pago Capital'].apply(
                lambda d: d.strftime('%d/%m/%Y') if pd.notna(d) else ''
            )
            df_deudas['fecha_Demanda'] = df_deudas['fecha_Demanda'].dt.strftime('%d/%m/%Y')
            df_deudas['Fecha_Liquidacion'] = df_deudas['Fecha_Liquidacion'].dt.strftime('%d/%m/%Y')

            # --- DASHBOARD ---
            st.success("¡Liquidación judicial finalizada!")
            st.markdown("### 📋 Resumen del Juicio")

            tot_capital = df_deudas['Capital'].sum()
            tot_res = df_deudas['Interes_Resarcitorio'].sum()
            tot_capi = df_deudas['Interes_Capitalizable'].sum()
            tot_pun = df_deudas['Interes_Punitorio'].sum()
            tot_pagar = df_deudas['Total_Actualizado'].sum()

            col1, col2, col3, col4 = st.columns(4)
            with col1: st.metric("Capital Original", formato_arg(tot_capital))
            with col2: st.metric("Resarcitorios", formato_arg(tot_res))
            with col3: st.metric("Capitalizables", formato_arg(tot_capi))
            with col4: st.metric("Punitorios", formato_arg(tot_pun))

            col5, col6 = st.columns(2)
            with col5: st.metric("Total Actualizado", formato_arg(tot_pagar))
            with col6: st.metric("Antigüedad Juicio", f"{antiguedad_juicio} días")

            st.divider()

            # --- DETALLE DE OBLIGACIONES ---
            with st.expander("🔍 Ver detalle completo de obligaciones", expanded=False):
                st.dataframe(
                    df_deudas[['Impuesto', 'Vencimiento', 'Capital', 'F. Pago Capital',
                               'Interes_Resarcitorio', 'Interes_Capitalizable',
                               'fecha_Demanda', 'Interes_Punitorio',
                               'Fecha_Liquidacion', 'Total_Actualizado']],
                    use_container_width=True
                )

            st.divider()

            # --- CUADRO DE TASAS DINÁMICO (AL FINAL) ---
            st.markdown("### 📈 Tasas de Interés de Referencia")
            col_t1, col_t2 = st.columns(2)

            with col_t1:
                st.write("**Resarcitorios y Capitalizables (Hasta fecha de Demanda)**")
                display_res['Tasa_Diaria'] = display_res['Tasa_Diaria'].apply(lambda x: f"{x*100:.4f}%")
                if 'dias' in display_res.columns:
                    st.dataframe(display_res[['Desde', 'Hasta', 'Tasa_Diaria', 'dias']], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(display_res[['Desde', 'Hasta', 'Tasa_Diaria']], use_container_width=True, hide_index=True)

            with col_t2:
                st.write("**Punitorios (Desde inicio de ejecución)**")
                display_pun['Tasa_Diaria'] = display_pun['Tasa_Diaria'].apply(lambda x: f"{x*100:.4f}%")
                if 'dias' in display_pun.columns:
                    st.dataframe(display_pun[['Desde', 'Hasta', 'Tasa_Diaria', 'dias']], use_container_width=True, hide_index=True)
                else:
                    st.dataframe(display_pun[['Desde', 'Hasta', 'Tasa_Diaria']], use_container_width=True, hide_index=True)

            # --- DESCARGA ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_deudas.to_excel(writer, index=False, sheet_name='Liquidacion_Apremio')
            procesado = output.getvalue()

            st.markdown("<br>", unsafe_allow_html=True)
            c1, c_boton, c2 = st.columns([1, 2, 1])
            with c_boton:
                st.download_button(
                    label="📥 Descargar Planilla (Excel)",
                    data=procesado,
                    file_name="Liquidacion_ARCA_Apremio.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    except Exception as e:
        st.error(f"Error al procesar: {e}")
