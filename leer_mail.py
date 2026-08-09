"""Lector de los mails de boleta de deuda que mandan los agentes fiscales.

Los agentes fiscales pegan en el mail la pantalla "Datos de la Boleta de Deuda"
del sistema de ARCA. Ese pegado viaja como una tabla HTML, y de ahí sale casi
todo lo que necesitan las planillas del liquidador:

    del mail                              a la planilla
    ─────────────────────────────────     ─────────────────────
    Impuestos - Conceptos - Subconc.  →   Impuesto / concepto
    Período                           →   Periodo
    Vencimiento                       →   Vencimiento
    Monto de la Deuda                 →   Capital
    Pagos de Capital Registrados      →   F. Pago Capital
    Fecha Sorteo                      →   fecha_Demanda

La fecha de liquidación no sale del mail: la elegís vos según cuándo presentás.

Lo que NO hace este módulo es decidir qué se calcula sobre cada fila. Eso lo
dice una nota escrita a mano por el agente fiscal al lado del importe ("DEBE",
"PRESENTO DDJJ - DEBE INT. RESARCITORIOS + PUNITORIOS + CAPITALIZADOS", "pto.
DJ debe intereses"...). El módulo la lee, propone una clasificación y marca la
fila para que la revises. Nunca decide solo: esto termina en un expediente.

Solo biblioteca estándar: se puede correr en cualquier lado sin instalar nada.
"""
import email
import email.policy
import re
import unicodedata
from html.parser import HTMLParser


# ── Los mails llegan con acentos rotos ──────────────────────────────────────
# Algunos reenvíos pierden la codificación por el camino y "Período" llega como
# "Perï¿½odo": el carácter original ya no está y no hay forma de recuperarlo.
# Por eso todas las comparaciones se hacen sobre el texto pelado a ASCII, donde
# las dos formas terminan iguales ("Perodo").

def _pelar(texto):
    """Deja el texto en ASCII, sin acentos ni basura de codificación."""
    if not texto:
        return ''
    limpio = unicodedata.normalize('NFKD', texto)
    limpio = ''.join(c for c in limpio if not unicodedata.combining(c))
    return ''.join(c if c.isascii() else '' for c in limpio)


def _comprimir(texto):
    """Pelado, en minúsculas y con los espacios normalizados, para comparar."""
    return re.sub(r'\s+', ' ', _pelar(texto)).strip().lower()


# ── Lectura del HTML ────────────────────────────────────────────────────────

class _Tablas(HTMLParser):
    """Saca las tablas del HTML como listas de celdas, respetando los saltos.

    Los saltos importan: dentro de una misma celda vienen apilados el impuesto,
    el concepto y el subconcepto.
    """

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tablas = []
        self.pila = []
        self.fila = None
        self.celda = []
        self.en_celda = False

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self.pila.append([])
        elif tag == 'tr' and self.pila:
            self.fila = []
        elif tag in ('td', 'th') and self.fila is not None:
            self.en_celda = True
            self.celda = []
        elif tag in ('br', 'p', 'div', 'tr') and self.en_celda:
            self.celda.append('\n')

    def handle_endtag(self, tag):
        if tag == 'table' and self.pila:
            self.tablas.append(self.pila.pop())
        elif tag == 'tr' and self.fila is not None:
            if self.pila and any(c.strip() for c in self.fila):
                self.pila[-1].append(self.fila)
            self.fila = None
        elif tag in ('td', 'th') and self.en_celda:
            texto = ''.join(self.celda)
            texto = re.sub(r'[^\S\n]+', ' ', texto)
            lineas = [l.strip() for l in texto.split('\n')]
            self.fila.append('\n'.join(l for l in lineas if l))
            self.en_celda = False

    def handle_data(self, dato):
        if self.en_celda:
            self.celda.append(dato)


def _tablas_del_mail(datos):
    """Devuelve (tablas, remitente, asunto) a partir de los bytes de un .eml."""
    if isinstance(datos, str):
        datos = datos.encode('utf-8', 'replace')
    msg = email.message_from_bytes(datos, policy=email.policy.default)

    parte = msg.get_body(preferencelist=('html',))
    if parte is None:
        raise ValueError(
            "El mail no trae la versión con formato (HTML), que es de donde se lee "
            "la tabla. Reenvialo sin convertirlo a texto plano."
        )

    parser = _Tablas()
    parser.feed(parte.get_content())
    parser.close()

    remitente = ''
    for cabecera in ('From', 'Reply-To'):
        if msg.get(cabecera):
            direcciones = re.findall(r'[\w.+-]+@[\w.-]+', str(msg.get(cabecera)))
            if direcciones:
                remitente = direcciones[0].lower()
                break

    return parser.tablas, remitente, str(msg.get('Subject') or '').strip()


# ── Conversión de los valores ───────────────────────────────────────────────

_RE_FECHA = re.compile(r'\b(\d{1,2})/(\d{1,2})/(\d{4})\b')
# Pelado, "Período" queda en "Periodo". Si el reenvío rompió la codificación, la
# "í" puede haber quedado en cualquier cosa: "Perodo", "Peri12odo". Lo único
# confiable es que empieza con "Per" y termina en "odo".
_PERIODO = r'Per\w{0,4}odo'
_RE_PERIODO = re.compile(_PERIODO + r'\s*:?\s*(\d{4})\s*/\s*(\d{1,2})')
_RE_CUOTA = re.compile(r'Cuota\s*:?\s*(\d+)')
# El importe abre la celda. Puede venir con separador de miles o sin él.
_RE_IMPORTE = re.compile(r'^\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+)')


def _fecha(texto, cual=0):
    """Saca la fecha número `cual` del texto, en formato ISO. '' si no hay."""
    encontradas = _RE_FECHA.findall(texto or '')
    if len(encontradas) <= cual:
        return ''
    dia, mes, anio = encontradas[cual]
    return f"{anio}-{int(mes):02d}-{int(dia):02d}"


def _importe(texto):
    """Convierte '2.452.500,32' o '366033,27' en float. None si no hay número."""
    m = _RE_IMPORTE.match(texto or '')
    if not m:
        return None
    return float(m.group(1).replace('.', '').replace(',', '.'))


# ── Clasificación de la fila ────────────────────────────────────────────────
# La nota que escribe el agente fiscal al lado del importe es lo que dice qué
# se calcula. No está normalizada: cada uno la escribe a su manera. Estas reglas
# son una PROPUESTA; todo lo que no encaja limpio se marca 'revisar'.

CAPITAL = 'capital'        # se deben capital e intereses
INTERESES = 'intereses'    # el capital está cancelado, se deben los intereses
REVISAR = 'revisar'        # no me arriesgo: decidilo vos


def clasificar(subconcepto, nota):
    """Propone a qué liquidación va la fila.

    Lo que manda es el subconcepto de ARCA, porque dice qué es el importe:

      • SALDO DE DECLARACIÓN JURADA, ANTICIPOS → el importe es capital impositivo.
        Va a "Capital + Intereses", que calcula resarcitorios, punitorios y, si el
        capital se pagó, los capitalizables desde ese pago.

      • INTERESES RESARCITORIOS → el importe YA es interés: el capital de origen
        está cancelado. Va a "Juicio a los Intereses".

    La nota del agente fiscal no decide, pero sirve de contraste: si dice que el
    capital está cancelado y en la boleta no figura el pago, algo no cierra.
    """
    n = _comprimir(nota)
    s = _comprimir(subconcepto)

    # Un plan de pagos vigente cambia el panorama: no se liquida como el resto.
    if 'plan vigente' in n or 'facilidades' in n or re.search(r'\brg\s*\d{4}\b', n):
        return REVISAR, 'Menciona un plan de pagos: revisá si corresponde liquidarla.'

    if 'interes' in s:
        return INTERESES, ''

    if 'saldo de declaracion jurada' in s or 'anticipo' in s:
        return CAPITAL, ''

    if not s:
        return REVISAR, 'No pude leer el subconcepto: clasificala a mano.'
    return REVISAR, f'No reconozco el subconcepto "{subconcepto.strip()}": clasificala a mano.'


def _contrastar_con_la_nota(fila):
    """Avisa cuando la nota del agente fiscal no coincide con lo que se leyó.

    El caso que importa: la nota dice que el capital ya se pagó ("debe intereses",
    "presentó DDJJ") pero la boleta no trae el pago registrado. Sin esa fecha, los
    punitorios se calculan hasta la liquidación en vez de hasta el pago, y el
    importe sale de más.
    """
    if fila['Destino'] != CAPITAL or fila['F. Pago Capital']:
        return
    n = _comprimir(fila['Nota'])
    dice_que_pago = (re.search(r'debe\s+int', n) or 'presento ddjj' in n
                     or 'pto. dj' in n or 'pto dj' in n)
    if dice_que_pago:
        fila['Aviso'] = (
            f'La nota dice "{fila["Nota"].strip()}", o sea que el capital estaría '
            'cancelado, pero la boleta no registra el pago. Si se pagó, cargá la '
            'fecha en "F. Pago Capital": sin ella los punitorios corren de más.')


# ── Armado de las filas de deuda ────────────────────────────────────────────

def _es_encabezado(fila, texto):
    return any(texto in _comprimir(c) for c in fila)


def _partir_descripcion(celda):
    """De la primera celda saca (impuesto, concepto, subconcepto, cuota).

    Viene apilada así, a veces con el nombre del impuesto cortado en dos líneas:

        IMPUESTO AL VALOR
        AGREGADO LEY 23349 Y SUS MODIFICACIONES
        DECLARACIÓN JURADA              <- concepto
        INTERESES RESARCITORIOS         <- subconcepto
        Rectificativa/Denuncia Z Nº: 50
        Fecha de Resolución/Intimación: 19/12/2024

    Las últimas dos líneas antes de los datos sueltos son concepto y
    subconcepto; todo lo de arriba es el nombre del impuesto.
    """
    lineas = [l.strip('  \t') for l in (celda or '').split('\n')]
    lineas = [l for l in lineas if l]

    cuota = ''
    utiles = []
    for linea in lineas:
        pelada = _pelar(linea)
        m = _RE_CUOTA.search(pelada)
        if m:
            cuota = m.group(1)
            continue
        # Datos sueltos que no forman parte del nombre.
        if re.search(r'Rectificativa|Denuncia|Fecha de|Establec|' + _PERIODO, pelada, re.I):
            continue
        utiles.append(linea)

    if len(utiles) >= 3:
        impuesto = ' '.join(utiles[:-2])
        concepto, subconcepto = utiles[-2], utiles[-1]
    elif len(utiles) == 2:
        impuesto, concepto, subconcepto = utiles[0], utiles[1], ''
    else:
        impuesto = utiles[0] if utiles else ''
        concepto = subconcepto = ''

    return impuesto.strip(), concepto.strip(), subconcepto.strip(), cuota


def _leer_detalle(tabla):
    """Convierte la tabla 'Detalle de Deuda' en filas."""
    filas = []
    for cruda in tabla:
        if len(cruda) < 3:
            continue
        primera = _comprimir(cruda[0])
        if not primera or 'impuestos - conceptos' in primera:
            continue
        if primera.startswith('suma total'):
            continue

        capital = _importe(cruda[1])
        if capital is None:
            continue

        impuesto, concepto, subconcepto, cuota = _partir_descripcion(cruda[0])

        # Lo que sigue al importe en la misma celda es la nota manuscrita.
        nota = _RE_IMPORTE.sub('', cruda[1], count=1).strip()
        nota = re.sub(r'\s*\n\s*', ' ', nota).strip()

        m = _RE_PERIODO.search(_pelar(cruda[2]))
        if not m:
            periodo = ''
        elif cuota:
            # Los anticipos vienen todos con período "/0": lo que los distingue es
            # la cuota. En las planillas del estudio se escriben "2025-1", "2025-2".
            periodo = f"{m.group(1)}-{cuota}"
        else:
            periodo = f"{m.group(1)}/{m.group(2)}"
        vencimiento = _fecha(cruda[2])

        destino, aviso = clasificar(subconcepto or concepto, nota)

        filas.append({
            'Impuesto': impuesto,
            'concepto': subconcepto or concepto,
            'Periodo': periodo,
            'Cuota': cuota,
            'Vencimiento': vencimiento,
            'Capital': capital,
            'F. Pago Capital': '',
            'Nota': nota,
            'Destino': destino,
            'Aviso': aviso,
        })
    return filas


def _clave_pago(cuota, vencimiento, importe):
    """Lo que identifica una fila entre las dos tablas del mail.

    El nombre del impuesto no sirve de clave: en la tabla de deuda y en la de
    pagos viene cortado en distinta cantidad de líneas. Vencimiento, cuota e
    importe alcanzan de sobra dentro de una misma boleta.
    """
    return (cuota, vencimiento, round(importe, 2))


def _leer_pagos(tabla):
    """De 'Pagos de Capital Registrados' saca {clave: {fechas de pago}}."""
    pagos = {}
    for cruda in tabla:
        if len(cruda) < 3:
            continue
        if 'impuestos - conceptos' in _comprimir(cruda[0]):
            continue
        importe = _importe(cruda[1])
        if importe is None:
            continue
        # En esta tabla el período y el vencimiento van dentro de la primera celda,
        # en la línea "Período: 2026/0 -  13/10/2025".
        lineas = [l for l in cruda[0].split('\n') if re.search(_PERIODO, _pelar(l))]
        vencimiento = _fecha(lineas[-1]) if lineas else _fecha(cruda[0])
        _, _, _, cuota = _partir_descripcion(cruda[0])
        fecha_pago = _fecha(cruda[2])
        if vencimiento and fecha_pago:
            pagos.setdefault(_clave_pago(cuota, vencimiento, importe), set()).add(fecha_pago)
    return pagos


def _buscar_dato(tablas, etiqueta):
    """Busca un valor de las fichas 'Etiqueta : valor' que trae la boleta."""
    objetivo = _comprimir(etiqueta)
    for tabla in tablas:
        for fila in tabla:
            for i, celda in enumerate(fila[:-1]):
                if _comprimir(celda).rstrip(' :') == objetivo:
                    # El valor puede venir cortado en varias líneas por el ancho
                    # de la tabla original: se vuelve a juntar en una sola.
                    valor = re.sub(r'\s+', ' ', fila[i + 1]).strip()
                    if valor:
                        return valor
    return ''


# ── Entrada principal ───────────────────────────────────────────────────────

def leer_boleta(datos, remitentes_habilitados=()):
    """Lee un mail de boleta de deuda y devuelve todo lo que se pudo extraer.

    `remitentes_habilitados` es una lista de direcciones. Si se pasa y el mail
    viene de otra, se avisa — pero igual se lee: el mail puede venir reenviado
    y no tiene sentido negarse a leerlo por eso.
    """
    tablas, remitente, asunto = _tablas_del_mail(datos)
    avisos = []

    if remitentes_habilitados:
        habilitados = {r.strip().lower() for r in remitentes_habilitados if r.strip()}
        if habilitados and remitente not in habilitados:
            avisos.append(
                f"El mail viene de {remitente or 'un remitente que no pude leer'}, "
                "que no está en la lista de agentes fiscales conocidos. "
                "Revisá el contenido con más atención de lo habitual."
            )

    tabla_detalle = None
    tabla_pagos = None
    for tabla in tablas:
        for fila in tabla[:2]:
            if tabla_detalle is None and _es_encabezado(fila, 'monto de la deuda'):
                tabla_detalle = tabla
            if tabla_pagos is None and _es_encabezado(fila, 'importe del pago'):
                tabla_pagos = tabla

    if tabla_detalle is None:
        raise ValueError(
            "No encontré la tabla 'Detalle de Deuda' en el mail. ¿Es una boleta de "
            "deuda de ARCA? Si el agente fiscal mandó una captura de pantalla en vez "
            "de pegar la tabla, no hay nada que leer."
        )

    filas = _leer_detalle(tabla_detalle)
    if not filas:
        raise ValueError("Encontré la tabla de deuda pero no pude leer ninguna fila.")

    # El vencimiento y el importe son lo único que entra al cálculo: sin eso la
    # fila no sirve, y prefiero decirlo antes que liquidar mal.
    for fila in filas:
        if not fila['Vencimiento']:
            fila['Destino'] = REVISAR
            fila['Aviso'] = 'No pude leer el vencimiento. Cargalo a mano.'

    pagos = _leer_pagos(tabla_pagos) if tabla_pagos else {}
    for fila in filas:
        fechas = pagos.get(_clave_pago(fila['Cuota'], fila['Vencimiento'], fila['Capital']))
        if not fechas:
            continue
        if len(fechas) == 1:
            fila['F. Pago Capital'] = next(iter(fechas))
        else:
            # Dos pagos con el mismo vencimiento e importe: no adivino cuál va.
            fila['Destino'] = REVISAR
            fila['Aviso'] = ('Hay más de un pago con este vencimiento e importe '
                             f"({', '.join(sorted(fechas))}). Elegí la fecha a mano.")

    for fila in filas:
        if not fila['Aviso']:
            _contrastar_con_la_nota(fila)

    # La fecha de sorteo es la que viene usando el estudio como fecha de demanda.
    fecha_demanda = _fecha(_buscar_dato(tablas, 'Fecha Sorteo'))

    return {
        'contribuyente': _buscar_dato(tablas, 'Contribuyente'),
        'cuit': re.sub(r'\D', '', _buscar_dato(tablas, 'C.U.I.T.').split('-')[0])[:11],
        'juicio': _buscar_dato(tablas, 'Nro. Juicio'),
        'expediente': _buscar_dato(tablas, 'Expediente'),
        'juzgado': _buscar_dato(tablas, 'Juzgado'),
        'monto_demanda': _importe(_buscar_dato(tablas, 'Monto Demanda')),
        'fecha_demanda': fecha_demanda,
        'remitente': remitente,
        'asunto': asunto,
        'filas': filas,
        'avisos': avisos,
    }
