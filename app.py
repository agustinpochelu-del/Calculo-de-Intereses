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
# TASAS POR DEFECTO ("memoria" de la app)
#
# Las tasas se guardan como TASA MENSUAL en porcentaje, que es como las publica ARCA.
# La tasa diaria se deriva dividiendo por 30 en el momento del cálculo, sin truncar
# decimales (0,0275/30 y no 0,00091667: truncar movía centavos en cada línea).
#
# "Hasta" es INCLUSIVO: el tramo rige hasta el final de ese día, y el día siguiente
# arranca el tramo que sigue. Por eso el "Hasta" de un tramo + 1 día tiene que dar
# exactamente el "Desde" del tramo siguiente; si queda un hueco, se pierden días de
# interés sin que nadie lo note.
#
# "Dias" es la cantidad de días que ARCA computa para el tramo COMPLETO, tal como
# figura en el detalle de cálculo que emite. No siempre coincide con lo que daría la
# cuenta por meses de 30 (ver dias_arca): en el bimestre 12/2024-01/2025 ARCA computa
# 61 y no 60, en febrero 2025 computa 28 y no 30, y en 03/2025-06/2025 computa 122 y
# no 120. Son los días oficiales del tramo y mandan sobre cualquier fórmula. Solo se
# usan cuando el devengamiento cubre el tramo entero; para los tramos de las puntas,
# que se cubren a medias, se cuenta con dias_arca.
#
# Los tramos abiertos ("vigente hasta nuevo aviso") van sin "Dias": nunca se cubren
# enteros, así que el dato no aplica.
#
# 🔧 Para actualizar cuando salga una tasa nueva: agregá una fila al final de la lista
# correspondiente (Desde, Hasta, Tasa_Mensual) y cerrá el "Hasta" del tramo anterior
# con el día en que dejó de regir, anotándole los "Dias" que informe ARCA.
# =====================================================================================
TASAS_RESARCITORIAS_DEFAULT = [
    {"Desde": "2024-04-29", "Hasta": "2024-05-31", "Tasa_Mensual": 12.07, "Dias": 32},
    {"Desde": "2024-06-01", "Hasta": "2024-07-31", "Tasa_Mensual": 6.41, "Dias": 60},
    {"Desde": "2024-08-01", "Hasta": "2024-09-30", "Tasa_Mensual": 6.41, "Dias": 60},
    {"Desde": "2024-10-01", "Hasta": "2024-11-30", "Tasa_Mensual": 6.41, "Dias": 60},
    {"Desde": "2024-12-01", "Hasta": "2025-01-31", "Tasa_Mensual": 7.47, "Dias": 61},
    {"Desde": "2025-02-01", "Hasta": "2025-02-28", "Tasa_Mensual": 7.26, "Dias": 28},
    {"Desde": "2025-03-01", "Hasta": "2025-06-30", "Tasa_Mensual": 4.00, "Dias": 122},
    {"Desde": "2025-07-01", "Hasta": "2029-12-31", "Tasa_Mensual": 2.75, "Dias": None},  # vigente hasta nuevo aviso
]

TASAS_PUNITORIAS_DEFAULT = [
    {"Desde": "2026-01-01", "Hasta": "2050-12-31", "Tasa_Mensual": 3.50, "Dias": None},  # vigente hasta nuevo aviso
]


def _normalizar_tabla_tasas(df):
    """Deja la tabla lista para el motor: fechas como datetime, tasa diaria derivada
    de la mensual y tramos ordenados. Acepta tablas viejas que traigan Tasa_Diaria."""
    df = df.copy()
    df['Desde'] = pd.to_datetime(df['Desde'])
    df['Hasta'] = pd.to_datetime(df['Hasta'])

    if 'Tasa_Mensual' in df.columns:
        df['Tasa_Mensual'] = pd.to_numeric(df['Tasa_Mensual'])
        df['Tasa_Diaria'] = df['Tasa_Mensual'] / 100 / 30
    elif 'Tasa_Diaria' in df.columns:
        # Formato viejo, con la tasa diaria ya truncada (0,000917 en vez de 0,00091666...).
        # ARCA publica la mensual con dos decimales, así que redondear ahí la recupera
        # exacta (0,000917 x 30 = 2,751 -> 2,75) y se vuelve a derivar sin truncar.
        df['Tasa_Mensual'] = (pd.to_numeric(df['Tasa_Diaria']) * 30 * 100).round(2)
        df['Tasa_Diaria'] = df['Tasa_Mensual'] / 100 / 30
    else:
        raise ValueError("la tabla de tasas necesita una columna Tasa_Mensual o Tasa_Diaria")

    if 'dias' in df.columns and 'Dias' not in df.columns:
        df = df.rename(columns={'dias': 'Dias'})
    if 'Dias' not in df.columns:
        df['Dias'] = None
    df['Dias'] = pd.to_numeric(df['Dias'], errors='coerce')

    return df.sort_values('Desde').reset_index(drop=True)


def _tabla_default_a_df(tramos):
    return _normalizar_tabla_tasas(pd.DataFrame(tramos))


# =====================================================================================
# MOTOR DE CÁLCULO (compartido por las dos lengüetas)
# =====================================================================================
def dias_arca(desde, hasta):
    """Días entre dos fechas con la convención de ARCA: cada mes completo cuenta 30
    días, y el resto se cuenta por días corridos.

    Es la diferencia clave contra un conteo de días calendario. Ejemplo real, un
    anticipo con vencimiento 13/02/2026 liquidado al 10/06/2026:
        días calendario = 117   (15 de febrero + 31 + 30 + 31 + 10)
        días ARCA       = 118   (3 meses x 30 + 28 días de 13/05 a 10/06)
    Febrero corto suma 30 igual, y por eso el cálculo por días corridos da de menos
    en unos meses y de más en otros.
    """
    if pd.isna(desde) or pd.isna(hasta) or hasta <= desde:
        return 0
    meses = (hasta.year - desde.year) * 12 + (hasta.month - desde.month)
    ancla = desde + pd.DateOffset(months=meses)
    if ancla > hasta:
        meses -= 1
        ancla = desde + pd.DateOffset(months=meses)
    return meses * 30 + (hasta - ancla).days


def calcular_interes(fecha_inicio_calculo, fecha_fin_calculo, capital, df_tabla_tasas):
    """Devuelve (interés, días) del período, aplicando cada tramo de tasa vigente.

    El interés empieza a devengar el día SIGUIENTE a la fecha de origen (el
    vencimiento, la demanda o el pago, según el tramo) y corre hasta la fecha de
    corte inclusive: es lo que muestra el detalle de ARCA, donde un vencimiento del
    13/05/2024 abre el primer tramo el 14/05/2024.

    Para contar los días de cada tramo hay dos casos:

    - El devengamiento cubre el tramo ENTERO: se usan los días oficiales de la
      tabla ("Dias"), que es lo que computa ARCA aunque no siempre coincida con la
      cuenta por meses de 30.
    - El devengamiento cubre el tramo a medias (los tramos de las puntas): se
      cuentan con dias_arca sobre la parte efectivamente cubierta.

    El interés de cada tramo se redondea antes de sumarlo, igual que en el detalle
    de ARCA, donde los importes por tramo suman exacto el total informado.
    """
    if pd.isna(fecha_inicio_calculo) or pd.isna(fecha_fin_calculo):
        return 0.0, 0
    if fecha_inicio_calculo >= fecha_fin_calculo or capital == 0:
        return 0.0, 0

    un_dia = pd.Timedelta(days=1)
    # Ventana de devengamiento, con el extremo superior abierto.
    devenga_desde = fecha_inicio_calculo + un_dia
    devenga_hasta = fecha_fin_calculo + un_dia

    interes_acumulado = 0.0
    dias_acumulados = 0

    for _, tramo in df_tabla_tasas.iterrows():
        tramo_desde = tramo['Desde']
        tramo_hasta = tramo['Hasta'] + un_dia
        inicio = max(devenga_desde, tramo_desde)
        fin = min(devenga_hasta, tramo_hasta)
        if inicio >= fin:
            continue

        cubre_tramo_entero = (inicio == tramo_desde) and (fin == tramo_hasta)
        if cubre_tramo_entero and pd.notna(tramo['Dias']):
            dias = int(tramo['Dias'])
        else:
            dias = dias_arca(inicio, fin)
        if dias <= 0:
            continue

        interes_acumulado += round(capital * tramo['Tasa_Diaria'] * dias, 2)
        dias_acumulados += dias

    return round(interes_acumulado, 2), dias_acumulados


def cargar_tasas(archivo_subido, hoja, tramos_default):
    """Lee la hoja de tasas del Excel subido. Si la hoja no existe o está vacía,
    devuelve la tabla por defecto guardada en la app (ver TASAS_*_DEFAULT arriba)."""
    try:
        df = pd.read_excel(archivo_subido, sheet_name=hoja)
        df.columns = df.columns.str.strip()
        df = df[df['Desde'].notna() & (df['Desde'] != 'Total')].copy()
        if df.empty:
            raise ValueError("hoja de tasas vacía")
        return _normalizar_tabla_tasas(df), False
    except Exception:
        return _tabla_default_a_df(tramos_default), True


def avisar_tasas(uso_default_res, uso_default_pun):
    if uso_default_res:
        st.info("ℹ️ No se encontraron tasas resarcitorias en el archivo: se usaron las tasas por defecto guardadas en la app.")
    if uso_default_pun:
        st.info("ℹ️ No se encontraron tasas punitorias en el archivo: se usaron las tasas por defecto guardadas en la app.")


# =====================================================================================
# VALIDACIONES DE LA HOJA "Deudas"
# =====================================================================================
def validar_deudas(df, columnas_requeridas, df_tasas_res, df_tasas_pun):
    """Devuelve una lista de problemas encontrados. Lista vacía = todo en orden."""
    problemas = []

    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        problemas.append(f"Faltan columnas en la hoja 'Deudas': {', '.join(faltantes)}.")
        return problemas

    if df.empty:
        problemas.append("La hoja 'Deudas' no tiene ninguna obligación cargada.")
        return problemas

    sin_capital = df['Capital'].isna() | (df['Capital'] <= 0)
    if sin_capital.any():
        filas = ', '.join(str(i + 2) for i in df.index[sin_capital])
        problemas.append(f"Hay capital vacío o menor o igual a cero (filas del Excel: {filas}).")

    mal_ordenadas = df['Vencimiento'] > df['fecha_Demanda']
    if mal_ordenadas.any():
        filas = ', '.join(str(i + 2) for i in df.index[mal_ordenadas])
        problemas.append(f"El vencimiento es posterior a la fecha de demanda (filas del Excel: {filas}).")

    liq_previa = df['Fecha_Liquidacion'] < df['fecha_Demanda']
    if liq_previa.any():
        filas = ', '.join(str(i + 2) for i in df.index[liq_previa])
        problemas.append(f"La fecha de liquidación es anterior a la de demanda (filas del Excel: {filas}).")

    # El último tramo de tasa tiene que llegar hasta la fecha de liquidación.
    fin_res = df['fecha_Demanda'].max()
    fin_pun = df['Fecha_Liquidacion'].max()
    if fin_res > df_tasas_res['Hasta'].max():
        problemas.append(
            f"La tabla de tasas resarcitorias termina el {df_tasas_res['Hasta'].max():%d/%m/%Y} "
            f"y el cálculo llega hasta el {fin_res:%d/%m/%Y}: faltan tramos.")
    if fin_pun > df_tasas_pun['Hasta'].max():
        problemas.append(
            f"La tabla de tasas punitorias termina el {df_tasas_pun['Hasta'].max():%d/%m/%Y} "
            f"y el cálculo llega hasta el {fin_pun:%d/%m/%Y}: faltan tramos.")

    return problemas


def verificar_tabla_tasas(df_tasas, nombre):
    """Avisa si entre dos tramos consecutivos queda un hueco o una superposición, y si
    algún tramo cerrado no tiene cargados los días oficiales de ARCA."""
    for i in range(len(df_tasas) - 1):
        esperado = df_tasas.iloc[i]['Hasta'] + pd.Timedelta(days=1)
        siguiente = df_tasas.iloc[i + 1]['Desde']
        if siguiente != esperado:
            st.warning(
                f"⚠️ En la tabla de {nombre} el tramo que termina el "
                f"{df_tasas.iloc[i]['Hasta']:%d/%m/%Y} no empalma con el que arranca el "
                f"{siguiente:%d/%m/%Y}. Revisá las fechas: así se pierden (o se duplican) días de interés.")

    # El último tramo queda abierto, así que nunca se cubre entero: ahí "Dias" no aplica.
    sin_dias = df_tasas.iloc[:-1][df_tasas.iloc[:-1]['Dias'].isna()]
    for _, tramo in sin_dias.iterrows():
        st.warning(
            f"⚠️ El tramo de {nombre} del {tramo['Desde']:%d/%m/%Y} al {tramo['Hasta']:%d/%m/%Y} "
            f"no tiene cargados los días oficiales de ARCA. La app los va a contar sola, "
            f"pero puede quedar uno o dos días de diferencia: copiá el número del detalle de cálculo de ARCA.")


# =====================================================================================
# UI: TABLA DE TASAS DE REFERENCIA
# =====================================================================================
def mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq,
                         titulo_res="**Resarcitorios y Capitalizables (Hasta fecha de Demanda)**"):
    display_res = df_tasas_res.copy()
    display_pun = df_tasas_pun.copy()

    # Las puntas se recortan a las fechas del juicio; ahí los días oficiales del tramo
    # completo ya no aplican y hay que recalcularlos sobre el pedazo que quedó.
    display_res = display_res[display_res['Desde'] <= ultima_fecha_demanda].copy()
    if not display_res.empty:
        display_res.iloc[-1, display_res.columns.get_loc('Hasta')] = ultima_fecha_demanda
        display_res.iloc[-1, display_res.columns.get_loc('Dias')] = None

    display_pun = display_pun[display_pun['Hasta'] >= fecha_inicio_punitorios_global].copy()
    if not display_pun.empty:
        display_pun.iloc[0, display_pun.columns.get_loc('Desde')] = max(display_pun.iloc[0]['Desde'], fecha_inicio_punitorios_global)
        display_pun.iloc[-1, display_pun.columns.get_loc('Hasta')] = ultima_fecha_liq
        display_pun.iloc[0, display_pun.columns.get_loc('Dias')] = None
        display_pun.iloc[-1, display_pun.columns.get_loc('Dias')] = None

    def formatear(df):
        if df.empty:
            return df
        df = df.copy()
        df['Días'] = [int(d) if pd.notna(d) else dias_arca(des, has + pd.Timedelta(days=1))
                      for d, des, has in zip(df['Dias'], df['Desde'], df['Hasta'])]
        df['Tasa mensual'] = df['Tasa_Mensual'].apply(lambda x: f"{x:.4f}%".replace('.', ','))
        df['Tasa diaria'] = df['Tasa_Diaria'].apply(lambda x: f"{x*100:.6f}%".replace('.', ','))
        df['Desde'] = df['Desde'].dt.strftime('%d/%m/%Y')
        df['Hasta'] = df['Hasta'].dt.strftime('%d/%m/%Y')
        return df[['Desde', 'Hasta', 'Tasa mensual', 'Tasa diaria', 'Días']]

    st.markdown("### 📈 Tasas de Interés de Referencia")
    st.caption("Los tramos completos usan los días oficiales de ARCA. En los tramos de las puntas, "
               "que se cubren a medias, los días se cuentan como los cuenta ARCA: cada mes completo "
               "vale 30 días y el resto por días corridos.")
    col_t1, col_t2 = st.columns(2)

    with col_t1:
        st.write(titulo_res)
        st.dataframe(formatear(display_res), use_container_width=True, hide_index=True)

    with col_t2:
        st.write("**Punitorios (Desde inicio de ejecución)**")
        st.dataframe(formatear(display_pun), use_container_width=True, hide_index=True)


# =====================================================================================
# DESCARGA DEL EXCEL LIQUIDADO
# =====================================================================================
def boton_descarga(df_deudas, nombre_archivo, columnas_moneda, columnas_totalizar):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Liquidacion_Apremio'

    columnas = list(df_deudas.columns)
    _estilo_header(ws, 1, columnas)

    for fila_idx, (_, fila) in enumerate(df_deudas.iterrows(), start=2):
        for col_idx, col in enumerate(columnas, start=1):
            celda = ws.cell(row=fila_idx, column=col_idx, value=fila[col])
            celda.font = Font(name=_FUENTE, size=10)
            if col in columnas_moneda:
                celda.number_format = '#,##0.00'

    # Fila de totales, para no tener que sumar a mano al presentar la liquidación.
    # Se escriben importes, no fórmulas: la planilla se presenta en el expediente y
    # tiene que mostrar el total aunque se abra con un lector que no evalúe fórmulas.
    fila_total = len(df_deudas) + 2
    ws.cell(row=fila_total, column=1, value="TOTAL").font = Font(name=_FUENTE, bold=True, size=10)
    for col_idx, col in enumerate(columnas, start=1):
        if col in columnas_totalizar:
            celda = ws.cell(row=fila_total, column=col_idx, value=round(df_deudas[col].sum(), 2))
            celda.font = Font(name=_FUENTE, bold=True, size=10)
            celda.number_format = '#,##0.00'
    for col_idx in range(1, len(columnas) + 1):
        ws.cell(row=fila_total, column=col_idx).fill = PatternFill(
            start_color=_GRIS_EJEMPLO, end_color=_GRIS_EJEMPLO, fill_type="solid")

    _ajustar_anchos(ws, [max(12, min(26, len(str(c)) + 4)) for c in columnas])
    ws.freeze_panes = "A2"

    output = BytesIO()
    wb.save(output)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c_boton, c2 = st.columns([1, 2, 1])
    with c_boton:
        st.download_button(
            label="📥 Descargar Planilla (Excel)",
            data=output.getvalue(),
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
    df_tasas_res, uso_default_res = cargar_tasas(archivo_subido, 'Tasas', TASAS_RESARCITORIAS_DEFAULT)
    df_tasas_pun, uso_default_pun = cargar_tasas(archivo_subido, 'Tasas Punitorios', TASAS_PUNITORIAS_DEFAULT)
    avisar_tasas(uso_default_res, uso_default_pun)
    verificar_tabla_tasas(df_tasas_res, "tasas resarcitorias")
    verificar_tabla_tasas(df_tasas_pun, "tasas punitorias")

    df_deudas.columns = df_deudas.columns.str.strip()
    df_deudas = df_deudas.dropna(subset=['Vencimiento'])

    if 'F. Pago Capital' not in df_deudas.columns:
        df_deudas['F. Pago Capital'] = pd.NaT

    df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
    df_deudas['F. Pago Capital'] = pd.to_datetime(df_deudas['F. Pago Capital'])
    df_deudas['fecha_Demanda'] = pd.to_datetime(df_deudas['fecha_Demanda'])
    df_deudas['Fecha_Liquidacion'] = pd.to_datetime(df_deudas['Fecha_Liquidacion'])

    problemas = validar_deudas(
        df_deudas, ['Impuesto', 'Vencimiento', 'Capital', 'fecha_Demanda', 'Fecha_Liquidacion'],
        df_tasas_res, df_tasas_pun)
    if problemas:
        for p in problemas:
            st.error(f"❌ {p}")
        return

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
            resarcitorio, dias_res = calcular_interes(vencimiento, fecha_demanda, capital, df_tasas_res)
            capitalizable, dias_cap = 0.0, 0
            punitorio, dias_pun = calcular_interes(fecha_demanda, fecha_liq, capital, df_tasas_pun)
        elif fecha_demanda <= pago_capital:
            resarcitorio, dias_res = calcular_interes(vencimiento, fecha_demanda, capital, df_tasas_res)
            punitorio, dias_pun = calcular_interes(fecha_demanda, pago_capital, capital, df_tasas_pun)
            capitalizable, dias_cap = calcular_interes(pago_capital, fecha_liq, resarcitorio, df_tasas_res)
        else:
            resarcitorio, dias_res = calcular_interes(vencimiento, pago_capital, capital, df_tasas_res)
            capitalizable, dias_cap = calcular_interes(pago_capital, fecha_demanda, resarcitorio, df_tasas_res)
            punitorio, dias_pun = calcular_interes(fecha_demanda, fecha_liq, capital, df_tasas_pun)

        return pd.Series({
            'Interes_Resarcitorio': resarcitorio, 'Dias_Resarcitorios': dias_res,
            'Interes_Capitalizable': capitalizable, 'Dias_Capitalizables': dias_cap,
            'Interes_Punitorio': punitorio, 'Dias_Punitorios': dias_pun,
        })

    df_deudas = df_deudas.join(df_deudas.apply(procesar_fila, axis=1))
    for col in ['Dias_Resarcitorios', 'Dias_Capitalizables', 'Dias_Punitorios']:
        df_deudas[col] = df_deudas[col].astype(int)

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

    columnas_detalle = ['Impuesto', 'Vencimiento', 'Capital', 'F. Pago Capital',
                        'Dias_Resarcitorios', 'Interes_Resarcitorio',
                        'Dias_Capitalizables', 'Interes_Capitalizable',
                        'fecha_Demanda', 'Dias_Punitorios', 'Interes_Punitorio',
                        'Fecha_Liquidacion', 'Total_Actualizado']

    with st.expander("🔍 Ver detalle completo de obligaciones", expanded=False):
        st.dataframe(df_deudas_fmt[columnas_detalle], use_container_width=True)

    st.divider()
    mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq)

    boton_descarga(
        df_deudas_fmt[columnas_detalle], "Liquidacion_ARCA_Apremio.xlsx",
        columnas_moneda=['Capital', 'Interes_Resarcitorio', 'Interes_Capitalizable',
                         'Interes_Punitorio', 'Total_Actualizado'],
        columnas_totalizar=['Capital', 'Interes_Resarcitorio', 'Interes_Capitalizable',
                            'Interes_Punitorio', 'Total_Actualizado'])


# =====================================================================================
# LENGÜETA 2: JUICIO A LOS INTERESES (el capital impositivo ya está pago; se demanda por
# los intereses resarcitorios impagos, que pasan a ser la nueva "base" de la deuda)
# =====================================================================================
def procesar_juicio_intereses(archivo_subido):
    df_deudas = pd.read_excel(archivo_subido, sheet_name='Deudas')
    df_tasas_res, uso_default_res = cargar_tasas(archivo_subido, 'Tasas', TASAS_RESARCITORIAS_DEFAULT)
    df_tasas_pun, uso_default_pun = cargar_tasas(archivo_subido, 'Tasas Punitorios', TASAS_PUNITORIAS_DEFAULT)
    avisar_tasas(uso_default_res, uso_default_pun)
    verificar_tabla_tasas(df_tasas_res, "tasas resarcitorias")
    verificar_tabla_tasas(df_tasas_pun, "tasas punitorias")

    df_deudas.columns = df_deudas.columns.str.strip()
    df_deudas = df_deudas.dropna(subset=['Vencimiento'])

    df_deudas['Vencimiento'] = pd.to_datetime(df_deudas['Vencimiento'])
    df_deudas['fecha_Demanda'] = pd.to_datetime(df_deudas['fecha_Demanda'])
    df_deudas['Fecha_Liquidacion'] = pd.to_datetime(df_deudas['Fecha_Liquidacion'])

    problemas = validar_deudas(
        df_deudas, ['Impuesto', 'Vencimiento', 'Capital', 'fecha_Demanda', 'Fecha_Liquidacion'],
        df_tasas_res, df_tasas_pun)
    if problemas:
        for p in problemas:
            st.error(f"❌ {p}")
        return

    ultima_fecha_demanda = df_deudas['fecha_Demanda'].max()
    ultima_fecha_liq = df_deudas['Fecha_Liquidacion'].max()

    # Acá "Capital" ya es el monto de intereses que se convirtió en la base del juicio
    # (el capital impositivo original está pago, no interviene). Solo hay dos tramos,
    # continuos, sin salto de un día entre uno y el siguiente:
    #   Resarcitorios: Vencimiento -> Demanda   (sobre ese monto base)
    #   Punitorios:    Demanda -> Liquidación   (sobre ese monto base)
    def procesar_fila(fila):
        resarcitorio, dias_res = calcular_interes(fila['Vencimiento'], fila['fecha_Demanda'], fila['Capital'], df_tasas_res)
        punitorio, dias_pun = calcular_interes(fila['fecha_Demanda'], fila['Fecha_Liquidacion'], fila['Capital'], df_tasas_pun)
        return pd.Series({
            'Interes_Resarcitorio': resarcitorio, 'Dias_Resarcitorios': dias_res,
            'Interes_Punitorio': punitorio, 'Dias_Punitorios': dias_pun,
        })

    df_deudas = df_deudas.join(df_deudas.apply(procesar_fila, axis=1))
    for col in ['Dias_Resarcitorios', 'Dias_Punitorios']:
        df_deudas[col] = df_deudas[col].astype(int)

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

    columnas_detalle = ['Impuesto', 'concepto', 'Periodo', 'Vencimiento', 'Capital',
                        'fecha_Demanda', 'Dias_Resarcitorios', 'Interes_Resarcitorio',
                        'Dias_Punitorios', 'Interes_Punitorio',
                        'Fecha_Liquidacion', 'Total_Actualizado']
    columnas_detalle = [c for c in columnas_detalle if c in df_deudas_fmt.columns]

    with st.expander("🔍 Ver detalle completo de obligaciones", expanded=False):
        st.dataframe(df_deudas_fmt[columnas_detalle], use_container_width=True)

    st.divider()
    mostrar_tabla_tasas(df_tasas_res, df_tasas_pun, ultima_fecha_demanda, fecha_inicio_punitorios_global, ultima_fecha_liq,
                         titulo_res="**Resarcitorios (Hasta fecha de Demanda)**")

    boton_descarga(
        df_deudas_fmt[columnas_detalle], "Liquidacion_ARCA_Intereses.xlsx",
        columnas_moneda=['Capital', 'Interes_Resarcitorio', 'Interes_Punitorio', 'Total_Actualizado'],
        columnas_totalizar=['Capital', 'Interes_Resarcitorio', 'Interes_Punitorio', 'Total_Actualizado'])


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

def _hoja_tasas(wb, nombre, tramos_default):
    ws = wb.create_sheet(nombre)
    _estilo_header(ws, 1, ["Desde", "Hasta", "Tasa_Mensual", "Dias"])

    fila = 2
    for tramo in tramos_default:
        desde = datetime.datetime.strptime(tramo["Desde"], "%Y-%m-%d").date()
        hasta = datetime.datetime.strptime(tramo["Hasta"], "%Y-%m-%d").date()
        valores = [desde, hasta, tramo["Tasa_Mensual"], tramo["Dias"]]
        formatos = ["DD/MM/YYYY", "DD/MM/YYYY", "0.0000", "0"]
        for idx, (val, fmt) in enumerate(zip(valores, formatos), start=1):
            celda = ws.cell(row=fila, column=idx, value=val)
            celda.font = Font(name=_FUENTE, size=10)
            celda.fill = PatternFill(start_color="FFF9DB", end_color="FFF9DB", fill_type="solid")
            celda.number_format = fmt
        fila += 1

    _ajustar_anchos(ws, [14, 14, 16, 10])
    ws["A1"].comment = Comment("Fecha de inicio del tramo de tasa vigente.", "Liquidador ARCA")
    ws["B1"].comment = Comment("Último día en que rige el tramo (inclusive). El día siguiente tiene que ser el 'Desde' del tramo que sigue; el último tramo queda abierto (vigente hasta nuevo aviso).", "Liquidador ARCA")
    ws["C1"].comment = Comment("Tasa MENSUAL en porcentaje, tal como la publica ARCA. Ej: 2,75 (la app divide por 30 para obtener la diaria).", "Liquidador ARCA")
    ws["D1"].comment = Comment("Días que computa ARCA para el tramo COMPLETO, según su detalle de cálculo. Se usa solo cuando el interés cubre el tramo entero; en los tramos de las puntas la app cuenta los días sola. Dejar vacío en el tramo abierto.", "Liquidador ARCA")
    _nota(ws, fila + 1, f"A{fila + 1}:D{fila + 1}",
          "↑ Estas son las tasas vigentes conocidas (ya precargadas por defecto en la app: si dejás esta hoja vacía "
          "o la borrás, la app las usa igual). Si sale una tasa nueva, agregá una fila debajo con el nuevo tramo "
          "y cambiá el 'Hasta' del tramo anterior al último día en que rigió.")
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
        "Completá también las hojas 'Tasas' (resarcitorios/capitalizables) y 'Tasas Punitorios'. Ya vienen",
        "precargadas con las tasas vigentes conocidas (las mismas que la app usa por defecto): si cambia la tasa,",
        "agregá una fila nueva al final del tramo correspondiente. Si preferís, podés dejar estas hojas tal cual",
        "vienen, borrarlas, o directamente no completarlas: la app va a usar las tasas guardadas por defecto igual.",
        "",
        "Los días se cuentan como los cuenta ARCA: cada mes completo vale 30 días y el resto se cuenta por días",
        "corridos. Por eso la planilla ya no pide una columna de días: la calcula sola.",
        "",
        "Borrá la fila de ejemplo de 'Deudas' (en gris/cursiva) antes de cargar tus propios datos.",
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

    _hoja_tasas(wb, "Tasas", TASAS_RESARCITORIAS_DEFAULT)
    _hoja_tasas(wb, "Tasas Punitorios", TASAS_PUNITORIAS_DEFAULT)

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
        "Completá también las hojas 'Tasas' (resarcitorios) y 'Tasas Punitorios'. Ya vienen precargadas con",
        "las tasas vigentes conocidas (las mismas que la app usa por defecto): si cambia la tasa, agregá una",
        "fila nueva al final del tramo correspondiente. Si preferís, podés dejar estas hojas tal cual vienen,",
        "borrarlas, o directamente no completarlas: la app va a usar las tasas guardadas por defecto igual.",
        "",
        "Los días se cuentan como los cuenta ARCA: cada mes completo vale 30 días y el resto se cuenta por días",
        "corridos. Por eso la planilla ya no pide una columna de días: la calcula sola.",
        "",
        "Borrá la fila de ejemplo de 'Deudas' (en gris/cursiva) antes de cargar tus propios datos.",
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

    _hoja_tasas(wb, "Tasas", TASAS_RESARCITORIAS_DEFAULT)
    _hoja_tasas(wb, "Tasas Punitorios", TASAS_PUNITORIAS_DEFAULT)

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
