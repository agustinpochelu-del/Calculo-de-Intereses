import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

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

def formato_arg(numero):
    return f"${numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# =====================================================================================
# MOTOR DE CÁLCULO (compartido por las dos lengüetas)
# =====================================================================================
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


def cargar_tasas(archivo_subido, hoja):
    df = pd.read_excel(archivo_subido, sheet_name=hoja)
    df.columns = df.columns.str.strip()
    df = df.rename(columns={'Dias': 'dias'})
    df = df[df['Desde'].notna() & (df['Desde'] != 'Total')].copy()
    df['Desde'] = pd.to_datetime(df['Desde'])
    df['Hasta'] = pd.to_datetime(df['Hasta'])
    return df


def mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq,
                         titulo_res="**Resarcitorios y Capitalizables (Hasta fecha de Demanda)**"):
    display_res = df_tasas_res.copy()
    display_pun = df_tasas_pun.copy()

    display_res.iloc[-1, display_res.columns.get_loc('Hasta')] = ultima_fecha_demanda

    display_pun = display_pun[display_pun['Hasta'] >= fecha_inicio_punitorios_global].copy()
    if not display_pun.empty:
        display_pun.iloc[0, display_pun.columns.get_loc('Desde')] = max(display_pun.iloc[0]['Desde'], fecha_inicio_punitorios_global)
        display_pun.iloc[-1, display_pun.columns.get_loc('Hasta')] = ultima_fecha_liq
        display_pun['dias_calculados'] = (display_pun['Hasta'] - display_pun['Desde']).dt.days
        display_pun['dias'] = display_pun['dias_calculados']

    display_res['Desde'] = display_res['Desde'].dt.strftime('%d/%m/%Y')
    display_res['Hasta'] = display_res['Hasta'].dt.strftime('%d/%m/%Y')
    display_pun['Desde'] = display_pun['Desde'].dt.strftime('%d/%m/%Y')
    display_pun['Hasta'] = display_pun['Hasta'].dt.strftime('%d/%m/%Y')

    st.markdown("### 📈 Tasas de Interés de Referencia")
    col_t1, col_t2 = st.columns(2)

    with col_t1:
        st.write(titulo_res)
        display_res['Tasa_Diaria'] = display_res['Tasa_Diaria'].apply(lambda x: f"{x*100:.4f}%")
        cols = ['Desde', 'Hasta', 'Tasa_Diaria', 'dias'] if 'dias' in display_res.columns else ['Desde', 'Hasta', 'Tasa_Diaria']
        st.dataframe(display_res[cols], use_container_width=True, hide_index=True)

    with col_t2:
        st.write("**Punitorios (Desde inicio de ejecución)**")
        display_pun['Tasa_Diaria'] = display_pun['Tasa_Diaria'].apply(lambda x: f"{x*100:.4f}%")
        cols = ['Desde', 'Hasta', 'Tasa_Diaria', 'dias'] if 'dias' in display_pun.columns else ['Desde', 'Hasta', 'Tasa_Diaria']
        st.dataframe(display_pun[cols], use_container_width=True, hide_index=True)


def boton_descarga(df_deudas, nombre_archivo):
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
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=nombre_archivo
        )


# =====================================================================================
# LENGÜETA 1: JUICIO POR CAPITAL + INTERESES (capital impositivo impago, con o sin pago posterior)
# =====================================================================================
def procesar_juicio_capital(archivo_subido):
    df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
    df_tasas_res = cargar_tasas(archivo_subido, 'Tasas')
    df_tasas_pun = cargar_tasas(archivo_subido, 'Tasas Punitorios')

    df_deudas.columns = df_deudas.columns.str.strip()
    df_deudas = df_deudas.dropna(subset=['Vencimiento'])

    if 'F. Pago Capital' not in df_deudas.columns:
        df_deudas['F. Pago Capital'] = pd.NaT

    df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
    df_deudas['F. Pago Capital'] = pd.to_datetime(df_deudas['F. Pago Capital'])
    df_deudas['fecha_Demanda'] = pd.to_datetime(df_deudas['fecha_Demanda'])
    df_deudas['Fecha_Liquidacion'] = pd.to_datetime(df_deudas['Fecha_Liquidacion'])

    ultima_fecha_demanda = df_deudas['fecha_Demanda'].max()
    ultima_fecha_liq = df_deudas['Fecha_Liquidacion'].max()

    # Los tres tramos son continuos, sin salto de un día entre uno y el siguiente.
    # El orden de Punitorios y Capitalizables se invierte según cuál de los dos
    # eventos (Demanda o Pago del Capital) ocurre primero:
    #
    #   Caso A: Demanda ANTES del Pago del Capital
    #       Resarcitorios:  Vencimiento -> Demanda            (sobre Capital)
    #       Punitorios:     Demanda -> Pago Capital            (sobre Capital)
    #       Capitalizables: Pago Capital -> Liquidación         (sobre monto Resarcitorios)
    #
    #   Caso B: Pago del Capital ANTES de la Demanda
    #       Resarcitorios:  Vencimiento -> Pago Capital        (sobre Capital)
    #       Capitalizables: Pago Capital -> Demanda            (sobre monto Resarcitorios)
    #       Punitorios:     Demanda -> Liquidación              (sobre Capital)
    #
    #   Sin fecha de Pago de Capital cargada: Resarcitorios hasta la Demanda,
    #   Punitorios desde la Demanda hasta la Liquidación, sin Capitalizables.
    def procesar_fila(fila):
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
            interes_resarcitorio = calcular_interes(vencimiento, fecha_demanda, capital, df_tasas_res)
            interes_punitorio = calcular_interes(fecha_demanda, pago_capital, capital, df_tasas_pun)
            interes_capitalizable = calcular_interes(pago_capital, fecha_liq, interes_resarcitorio, df_tasas_res)
        else:
            interes_resarcitorio = calcular_interes(vencimiento, pago_capital, capital, df_tasas_res)
            interes_capitalizable = calcular_interes(pago_capital, fecha_demanda, interes_resarcitorio, df_tasas_res)
            interes_punitorio = calcular_interes(fecha_demanda, fecha_liq, capital, df_tasas_pun)

        return pd.Series({
            'Interes_Resarcitorio': interes_resarcitorio,
            'Interes_Capitalizable': interes_capitalizable,
            'Interes_Punitorio': interes_punitorio
        })

    resultado = df_deudas.apply(procesar_fila, axis=1)
    df_deudas['Interes_Resarcitorio'] = resultado['Interes_Resarcitorio']
    df_deudas['Interes_Capitalizable'] = resultado['Interes_Capitalizable']
    df_deudas['Interes_Punitorio'] = resultado['Interes_Punitorio']

    fecha_inicio_punitorios_global = ultima_fecha_demanda
    antiguedad_juicio = max(0, (ultima_fecha_liq - fecha_inicio_punitorios_global).days)

    df_deudas['Total_Actualizado'] = (
        df_deudas['Capital'] + df_deudas['Interes_Resarcitorio']
        + df_deudas['Interes_Capitalizable'] + df_deudas['Interes_Punitorio']
    )

    df_deudas_fmt = df_deudas.copy()
    df_deudas_fmt['Vencimiento'] = df_deudas_fmt['Vencimiento'].dt.strftime('%d/%m/%Y')
    df_deudas_fmt['F. Pago Capital'] = df_deudas_fmt['F. Pago Capital'].apply(lambda d: d.strftime('%d/%m/%Y') if pd.notna(d) else '')
    df_deudas_fmt['fecha_Demanda'] = df_deudas_fmt['fecha_Demanda'].dt.strftime('%d/%m/%Y')
    df_deudas_fmt['Fecha_Liquidacion'] = df_deudas_fmt['Fecha_Liquidacion'].dt.strftime('%d/%m/%Y')

    st.success("¡Liquidación judicial finalizada!")
    st.markdown("### 📋 Resumen del Juicio")

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("Capital Original", formato_arg(df_deudas['Capital'].sum()))
    with col2: st.metric("Resarcitorios", formato_arg(df_deudas['Interes_Resarcitorio'].sum()))
    with col3: st.metric("Capitalizables", formato_arg(df_deudas['Interes_Capitalizable'].sum()))
    with col4: st.metric("Punitorios", formato_arg(df_deudas['Interes_Punitorio'].sum()))

    col5, col6 = st.columns(2)
    with col5: st.metric("Total Actualizado", formato_arg(df_deudas['Total_Actualizado'].sum()))
    with col6: st.metric("Antigüedad Juicio", f"{antiguedad_juicio} días")

    st.divider()

    with st.expander("🔍 Ver detalle completo de obligaciones", expanded=False):
        st.dataframe(
            df_deudas_fmt[['Impuesto', 'Vencimiento', 'Capital', 'F. Pago Capital',
                            'Interes_Resarcitorio', 'Interes_Capitalizable',
                            'fecha_Demanda', 'Interes_Punitorio',
                            'Fecha_Liquidacion', 'Total_Actualizado']],
            use_container_width=True
        )

    st.divider()
    mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq)

    boton_descarga(df_deudas_fmt, "Liquidacion_ARCA_Apremio.xlsx")


# =====================================================================================
# LENGÜETA 2: JUICIO A LOS INTERESES (el capital impositivo ya está pago; se demanda por
# los intereses resarcitorios impagos, que pasan a ser la nueva "base" de la deuda)
# =====================================================================================
def procesar_juicio_intereses(archivo_subido):
    df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
    df_tasas_res = cargar_tasas(archivo_subido, 'Tasas')
    df_tasas_pun = cargar_tasas(archivo_subido, 'Tasas Punitorios')

    df_deudas.columns = df_deudas.columns.str.strip()
    df_deudas = df_deudas.dropna(subset=['Vencimiento'])

    df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
    df_deudas['fecha_Demanda'] = pd.to_datetime(df_deudas['fecha_Demanda'])
    df_deudas['Fecha_Liquidacion'] = pd.to_datetime(df_deudas['Fecha_Liquidacion'])

    ultima_fecha_demanda = df_deudas['fecha_Demanda'].max()
    ultima_fecha_liq = df_deudas['Fecha_Liquidacion'].max()

    # Acá "Capital" ya es el monto de intereses que se convirtió en la base del juicio
    # (el capital impositivo original está pago, no interviene). Solo hay dos tramos,
    # continuos, sin salto de un día entre uno y el siguiente:
    #   Resarcitorios: Vencimiento -> Demanda   (sobre ese monto base)
    #   Punitorios:    Demanda -> Liquidación   (sobre ese monto base)
    def procesar_fila(fila):
        interes_resarcitorio = calcular_interes(fila['Vencimiento'], fila['fecha_Demanda'], fila['Capital'], df_tasas_res)
        interes_punitorio = calcular_interes(fila['fecha_Demanda'], fila['Fecha_Liquidacion'], fila['Capital'], df_tasas_pun)
        return pd.Series({'Interes_Resarcitorio': interes_resarcitorio, 'Interes_Punitorio': interes_punitorio})

    resultado = df_deudas.apply(procesar_fila, axis=1)
    df_deudas['Interes_Resarcitorio'] = resultado['Interes_Resarcitorio']
    df_deudas['Interes_Punitorio'] = resultado['Interes_Punitorio']

    fecha_inicio_punitorios_global = ultima_fecha_demanda
    antiguedad_juicio = max(0, (ultima_fecha_liq - fecha_inicio_punitorios_global).days)

    df_deudas['Total_Actualizado'] = df_deudas['Capital'] + df_deudas['Interes_Resarcitorio'] + df_deudas['Interes_Punitorio']

    df_deudas_fmt = df_deudas.copy()
    df_deudas_fmt['Vencimiento'] = df_deudas_fmt['Vencimiento'].dt.strftime('%d/%m/%Y')
    df_deudas_fmt['fecha_Demanda'] = df_deudas_fmt['fecha_Demanda'].dt.strftime('%d/%m/%Y')
    df_deudas_fmt['Fecha_Liquidacion'] = df_deudas_fmt['Fecha_Liquidacion'].dt.strftime('%d/%m/%Y')

    st.success("¡Liquidación judicial finalizada!")
    st.markdown("### 📋 Resumen del Juicio (a los Intereses)")

    col1, col2, col3 = st.columns(3)
    with col1: st.metric("Intereses (Base del Juicio)", formato_arg(df_deudas['Capital'].sum()))
    with col2: st.metric("Resarcitorios", formato_arg(df_deudas['Interes_Resarcitorio'].sum()))
    with col3: st.metric("Punitorios", formato_arg(df_deudas['Interes_Punitorio'].sum()))

    col4, col5 = st.columns(2)
    with col4: st.metric("Total Actualizado", formato_arg(df_deudas['Total_Actualizado'].sum()))
    with col5: st.metric("Antigüedad Juicio", f"{antiguedad_juicio} días")

    st.divider()

    with st.expander("🔍 Ver detalle completo de obligaciones", expanded=False):
        st.dataframe(
            df_deudas_fmt[['Impuesto', 'concepto', 'Periodo', 'Vencimiento', 'Capital',
                            'fecha_Demanda', 'Interes_Resarcitorio', 'Interes_Punitorio',
                            'Fecha_Liquidacion', 'Total_Actualizado']],
            use_container_width=True
        )

    st.divider()
    mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq,
                         titulo_res="**Resarcitorios (Hasta fecha de Demanda)**")

    boton_descarga(df_deudas_fmt, "Liquidacion_ARCA_Intereses.xlsx")


# =====================================================================================
# GENERADOR DE PLANTILLAS EN BLANCO (para descargar y completar)
# =====================================================================================
_AZUL_HEADER = "1E3A8A"
_GRIS_EJEMPLO = "F0F4F8"
_FUENTE = "Arial"

def _estilo_header(ws, fila, columnas):
    for idx, col in enumerate(columnas, start=1):
        celda = ws.cell(row=fila, column=idx, value=col)
        celda.font = Font(name=_FUENTE, bold=True, color="FFFFFF", size=11)
        celda.fill = PatternFill(start_color=_AZUL_HEADER, end_color=_AZUL_HEADER, fill_type="solid")
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = 30

def _fila_ejemplo(ws, fila, valores, formatos=None):
    for idx, val in enumerate(valores, start=1):
        celda = ws.cell(row=fila, column=idx, value=val)
        celda.font = Font(name=_FUENTE, italic=True, color="6B7280", size=10)
        celda.fill = PatternFill(start_color=_GRIS_EJEMPLO, end_color=_GRIS_EJEMPLO, fill_type="solid")
        if formatos and formatos[idx - 1]:
            celda.number_format = formatos[idx - 1]

def _ajustar_anchos(ws, anchos):
    for idx, ancho in enumerate(anchos, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = ancho

def _nota(ws, fila, rango, texto):
    ws.merge_cells(rango)
    ws.cell(row=fila, column=1, value=texto).font = Font(name=_FUENTE, italic=True, size=9, color="9CA3AF")

def _hoja_tasas(wb, nombre, tasa_diaria_ejemplo, dias_ejemplo):
    ws = wb.create_sheet(nombre)
    _estilo_header(ws, 1, ["Desde", "Hasta", "Tasa_Diaria", "Dias"])
    _fila_ejemplo(ws, 2,
                  [datetime.date(2025, 7, 1), datetime.date(2029, 12, 31), tasa_diaria_ejemplo, dias_ejemplo],
                  formatos=["DD/MM/YYYY", "DD/MM/YYYY", "0.000000", "0"])
    _ajustar_anchos(ws, [14, 14, 14, 10])
    ws["A1"].comment = Comment("Fecha de inicio del tramo de tasa vigente.", "Liquidador ARCA")
    ws["B1"].comment = Comment("Fecha de fin del tramo (podés poner una fecha lejana como 31/12/2029 para el tramo vigente actual).", "Liquidador ARCA")
    ws["C1"].comment = Comment("Tasa diaria en formato decimal. Ej: 0,0917% diario = 0.000917", "Liquidador ARCA")
    ws["D1"].comment = Comment("Cantidad real de días calendario que abarca el tramo cuando se aplica completo (ajuste fino del cálculo). Podés dejarlo igual a Hasta-Desde si no tenés el dato exacto.", "Liquidador ARCA")
    _nota(ws, 3, "A3:D3", "↑ Fila de ejemplo (gris/cursiva): borrala y cargá tus propios tramos de tasa. Podés agregar tantas filas como cambios de tasa haya.")
    return ws

def _hoja_instrucciones(wb, titulo, lineas):
    ws = wb.create_sheet("Instrucciones", 0)
    ws["A1"] = titulo
    ws["A1"].font = Font(name=_FUENTE, bold=True, size=14, color=_AZUL_HEADER)
    for i, linea in enumerate(lineas, start=3):
        ws.cell(row=i, column=1, value=linea).font = Font(name=_FUENTE, size=11)
        ws.row_dimensions[i].height = 18
    ws.column_dimensions["A"].width = 110
    ws.sheet_view.showGridLines = False


def generar_plantilla_capital():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _hoja_instrucciones(wb, "Plantilla - Juicio por Capital + Intereses", [
        "Completá la hoja 'Deudas' con una fila por cada obligación (impuesto/período) del juicio.",
        "",
        "Columnas de 'Deudas':",
        "  • Impuesto: nombre del impuesto (ej: IMPUESTO A LAS GANANCIAS, IVA, etc.)",
        "  • concepto: detalle del concepto (ej: Saldo DDJJ, Anticipo, etc.)",
        "  • Periodo: período fiscal (ej: 2024-1)",
        "  • Vencimiento: fecha de vencimiento original de la obligación",
        "  • Capital: monto de capital adeudado (impuesto original, sin intereses)",
        "  • F. Pago Capital: fecha en que se pagó el capital. DEJAR VACÍO si el capital todavía no fue pagado.",
        "  • fecha_Demanda: fecha de inicio de la demanda de ejecución fiscal",
        "  • Fecha_Liquidacion: fecha a la que querés calcular la liquidación (fecha de pago de intereses)",
        "",
        "Completá también las hojas 'Tasas' (resarcitorios/capitalizables) y 'Tasas Punitorios' con los",
        "tramos de tasa vigentes en el período del juicio. Cada fila es un tramo con su tasa diaria.",
        "",
        "Borrá las filas de ejemplo (en gris/cursiva) antes de cargar tus propios datos.",
    ])

    ws = wb.create_sheet("Deudas")
    _estilo_header(ws, 1, ["Impuesto", "concepto", "Periodo", "Vencimiento", "Capital", "F. Pago Capital", "fecha_Demanda", "Fecha_Liquidacion"])
    _fila_ejemplo(ws, 2,
                  ["IMPUESTO A LAS GANANCIAS", "Saldo DDJJ", "2024-1",
                   datetime.date(2024, 5, 20), 150000.00, None,
                   datetime.date(2025, 1, 15), datetime.date(2025, 6, 30)],
                  formatos=[None, None, None, "DD/MM/YYYY", "#,##0.00", "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY"])
    _ajustar_anchos(ws, [26, 20, 12, 14, 16, 16, 14, 16])
    ws["F1"].comment = Comment("Dejar vacío si el capital todavía no fue pagado.", "Liquidador ARCA")
    _nota(ws, 3, "A3:H3", "↑ Fila de ejemplo: borrala y cargá tus propias obligaciones (podés agregar tantas filas como necesites).")

    _hoja_tasas(wb, "Tasas", 0.000917, 253)
    _hoja_tasas(wb, "Tasas Punitorios", 0.001167, 207)

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def generar_plantilla_intereses():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _hoja_instrucciones(wb, "Plantilla - Juicio a los Intereses", [
        "Usá esta planilla cuando el capital impositivo original YA fue pagado, y el juicio se inicia",
        "por los intereses resarcitorios que quedaron impagos (esos intereses pasan a ser la nueva",
        "'base' de la deuda). No incluye columna de pago de capital porque no aplica en este caso.",
        "",
        "Columnas de 'Deudas':",
        "  • Impuesto: nombre del impuesto (ej: IMPUESTO A LAS GANANCIAS, IVA, etc.)",
        "  • concepto: detalle del concepto (ej: INTERESES Capitalizables)",
        "  • Periodo: período fiscal de origen (ej: 2025-1)",
        "  • Vencimiento: fecha desde la que corren los intereses sobre este monto (vencimiento de origen)",
        "  • Capital: monto de intereses adeudados que se convierte en la base de este juicio",
        "  • fecha_Demanda: fecha de inicio de la demanda de ejecución fiscal",
        "  • Fecha_Liquidacion: fecha a la que querés calcular la liquidación (fecha de pago de intereses)",
        "",
        "Completá también las hojas 'Tasas' (resarcitorios) y 'Tasas Punitorios' con los tramos de",
        "tasa vigentes en el período del juicio. Cada fila es un tramo con su tasa diaria.",
        "",
        "Borrá las filas de ejemplo (en gris/cursiva) antes de cargar tus propios datos.",
    ])

    ws = wb.create_sheet("Deudas")
    _estilo_header(ws, 1, ["Impuesto", "concepto", "Periodo", "Vencimiento", "Capital", "fecha_Demanda", "Fecha_Liquidacion"])
    _fila_ejemplo(ws, 2,
                  ["IMPUESTO A LAS GANANCIAS", "INTERESES Capitalizables", "2025-1",
                   datetime.date(2025, 6, 17), 3837980.63,
                   datetime.date(2026, 3, 13), datetime.date(2026, 7, 30)],
                  formatos=[None, None, None, "DD/MM/YYYY", "#,##0.00", "DD/MM/YYYY", "DD/MM/YYYY"])
    _ajustar_anchos(ws, [26, 24, 12, 14, 16, 14, 16])
    _nota(ws, 3, "A3:G3", "↑ Fila de ejemplo: borrala y cargá tus propias obligaciones (podés agregar tantas filas como necesites).")

    _hoja_tasas(wb, "Tasas", 0.000917, 253)
    _hoja_tasas(wb, "Tasas Punitorios", 0.001167, 207)

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# =====================================================================================
# INTERFAZ: DOS LENGÜETAS
# =====================================================================================
tab1, tab2 = st.tabs(["💰 Juicio por Capital + Intereses", "📈 Juicio a los Intereses"])

with tab1:
    st.markdown("Subí el Excel con las hojas **Deudas**, **Tasas** y **Tasas Punitorios** (formato con Capital impositivo).")
    st.download_button(
        label="📄 Descargar plantilla en blanco",
        data=generar_plantilla_capital(),
        file_name="Plantilla_Juicio_Capital_Intereses.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="plantilla_capital"
    )
    archivo_1 = st.file_uploader("Arrastrá tu Excel aquí", type=["xlsx"], key="uploader_capital")
    if archivo_1 is not None:
        try:
            with st.spinner('Procesando liquidación judicial...'):
                procesar_juicio_capital(archivo_1)
        except Exception as e:
            st.error(f"Error al procesar: {e}")

with tab2:
    st.markdown("Subí el Excel con las hojas **Deudas**, **Tasas** y **Tasas Punitorios** (formato donde 'Capital' es el monto de intereses adeudados; el capital impositivo original ya está pago).")
    st.download_button(
        label="📄 Descargar plantilla en blanco",
        data=generar_plantilla_intereses(),
        file_name="Plantilla_Juicio_Intereses.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="plantilla_intereses"
    )
    archivo_2 = st.file_uploader("Arrastrá tu Excel aquí", type=["xlsx"], key="uploader_intereses")
    if archivo_2 is not None:
        try:
            with st.spinner('Procesando liquidación judicial...'):
                procesar_juicio_intereses(archivo_2)
        except Exception as e:
            st.error(f"Error al procesar: {e}")
