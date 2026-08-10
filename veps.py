"""Armado del archivo TXT para generar VEPs múltiples en ARCA.

Pagar una liquidación de apremio a mano significa cargar un VEP por cada importe:
ocho vencimientos con capital, resarcitorios, capitalizables y punitorios son
treinta y dos. Con un archivo se cargan todos juntos.

El formato tiene dos tipos de registro:

    Encabezado, una sola línea de 33 caracteres:

        01 · CUIT del generador (11) · 20001 · 00100 · 003 · 003 · cantidad+1 (4)

        Los últimos cuatro dígitos son la cantidad de VEPs más uno, o sea el
        total de líneas del archivo contando el encabezado.

    Detalle, una línea por VEP:

        02 <VEP fechaExpiracion="..." nroFormulario="..." codTipoPago="..."
                contribuyenteCUIT="..." concepto="..." subConcepto="..."
                periodoFiscal="AAAAMM" importe="..."
                [anticipoCuota="..."]><Obligacion impuesto="..." importe="..."/></VEP>

**El formato sale de un archivo que ARCA aceptó de verdad**, no de la guía. La
guía publicada usa `precio` en vez de `importe` y pone espacios alrededor de los
`=`; es de 2010 y quedó vieja. Ante la duda manda el archivo que funcionó, que
está fijado en `test_veps.py`.

Los códigos de impuesto, concepto y subconcepto van **sin ceros a la izquierda**,
y cuando no hay cuota el atributo `anticipoCuota` directamente no aparece.
"""
import datetime
import json
import os
import re

ARCHIVO_CONCEPTOS = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "conceptos_veps.json")

# Días que dura un VEP antes de vencer, contados desde que se genera.
DIAS_EXPIRACION = 10

# --- Qué subconcepto le corresponde a cada importe de la liquidación ---
# El capital lleva el mismo número que el concepto (declaración jurada o
# anticipo); los intereses tienen el suyo propio.
SUBCONCEPTO_CAPITAL = None          # se resuelve con el concepto de la fila
SUBCONCEPTO_RESARCITORIO = 51
SUBCONCEPTO_CAPITALIZABLE = 52
SUBCONCEPTO_PUNITORIO = 94

COLUMNAS_IMPORTE = {
    'Capital': SUBCONCEPTO_CAPITAL,
    'Interes_Resarcitorio': SUBCONCEPTO_RESARCITORIO,
    'Interes_Capitalizable': SUBCONCEPTO_CAPITALIZABLE,
    'Interes_Punitorio': SUBCONCEPTO_PUNITORIO,
}

# --- Conceptos ---
CONCEPTO_DDJJ = 19
CONCEPTO_ANTICIPO = 191


def cargar_conceptos():
    """Lee la tabla de impuestos, conceptos y subconceptos del estudio."""
    with open(ARCHIVO_CONCEPTOS, encoding='utf-8') as f:
        datos = json.load(f)
    tabla = {(c['impuesto'], c['concepto'], c['subconcepto']): c
             for c in datos['combinaciones']}
    return tabla, datos['impuestos'], datos.get('actualizado', '')


# ── Del CUIT se deduce si es persona física o jurídica ──────────────────────
# Importa porque Ganancias tiene un código de impuesto para cada una.

FISICAS = ('20', '23', '24', '27')
JURIDICAS = ('30', '33', '34')


def tipo_de_persona(cuit):
    """'fisica', 'juridica' o '' si el prefijo no es ninguno de los conocidos."""
    solo = re.sub(r'\D', '', cuit or '')
    if len(solo) != 11:
        return ''
    if solo[:2] in FISICAS:
        return 'fisica'
    if solo[:2] in JURIDICAS:
        return 'juridica'
    return ''


# ── De los nombres de la liquidación a los códigos de ARCA ─────────────────
# La liquidación trae los nombres como los escribe ARCA en la boleta de deuda,
# que no son los mismos de la tabla de conceptos. Este puente es explícito a
# propósito: lo que no reconoce, lo dice, en vez de arriesgar un código.
#
# 'ganancias' es el único que depende del CUIT: sociedades e individuos tienen
# códigos distintos.

IMPUESTO_POR_NOMBRE = (
    (r'sicore',                                   217),
    (r'valor agregado|^iva\b',                     30),
    (r'bienes personales',                        180),
    (r'acciones o participaciones',               211),
    (r'aportes.*seguridad social|seguridad social ley 24241', 301),
    (r'contribuciones',                           351),
    (r'honorarios',                                49),
    (r'ret.*art.*79|art\.? 79',                   767),
    (r'ganancias',                          'ganancias'),
)


def impuesto_de(nombre, cuit):
    """Devuelve (codigo, aviso). El código es None si no se pudo determinar."""
    plano = re.sub(r'\s+', ' ', (nombre or '')).strip().lower()
    if not plano:
        return None, "La fila no dice de qué impuesto es."

    for patron, codigo in IMPUESTO_POR_NOMBRE:
        if re.search(patron, plano):
            if codigo != 'ganancias':
                return codigo, ''
            persona = tipo_de_persona(cuit)
            if persona == 'juridica':
                return 10, ''
            if persona == 'fisica':
                return 11, ''
            return None, (f"Es Ganancias, pero el CUIT {cuit} no empieza en un prefijo "
                          "conocido, así que no sé si va como sociedad (10) o como "
                          "persona física (11). Elegilo a mano.")

    return None, f'No reconozco el impuesto "{nombre}". Cargá el código a mano.'


def concepto_de(nombre):
    """19 para declaración jurada, 191 para anticipos."""
    plano = (nombre or '').lower()
    if 'anticipo' in plano:
        return CONCEPTO_ANTICIPO, ''
    if 'declaracion' in plano or 'declaración' in plano or 'saldo' in plano:
        return CONCEPTO_DDJJ, ''
    if 'interes' in plano or 'interés' in plano:
        # Una fila cuyo concepto ya es "intereses resarcitorios" viene de una
        # declaración jurada: el interés es el subconcepto, no el concepto.
        return CONCEPTO_DDJJ, ''
    return None, f'No reconozco el concepto "{nombre}". Cargalo a mano.'


# ── Período fiscal ─────────────────────────────────────────────────────────

def periodo_fiscal(periodo, vencimiento, concepto):
    """Devuelve (AAAAMM, cuota, aviso).

    Los anticipos se declaran con el mes en 00 y el número de cuota aparte. El
    resto lleva el mes que corresponde.

    `periodo` viene de la planilla como '2026-1' (año y cuota, para anticipos)
    o '2024/5' (año y mes). Si no se puede leer, se cae al vencimiento.
    """
    texto = str(periodo or '').strip()

    m = re.match(r'^(\d{4})\s*[-/]\s*(\d{1,2})$', texto)
    if m:
        anio, segundo = int(m.group(1)), int(m.group(2))
        if concepto == CONCEPTO_ANTICIPO:
            return f"{anio}00", segundo, ''
        return f"{anio}{segundo:02d}", 0, ''

    if vencimiento is not None:
        anio, mes = vencimiento.year, vencimiento.month
        if concepto == CONCEPTO_ANTICIPO:
            return f"{anio}00", 0, (
                f'No pude leer la cuota del período "{periodo}". Puse el año del '
                'vencimiento y la cuota en cero: revisalo.')
        return f"{anio}{mes:02d}", 0, (
            f'No pude leer el período "{periodo}": usé el mes del vencimiento.')

    return '', 0, f'No pude leer el período "{periodo}" y no hay vencimiento.'


# ── Armado del archivo ─────────────────────────────────────────────────────

def _importe(valor):
    """Con punto decimal y dos decimales, como espera ARCA."""
    return f"{round(float(valor), 2):.2f}"


def linea_vep(pago, fecha_expiracion):
    """Arma la línea de detalle de un VEP."""
    cuota = ''
    if pago.get('cuota'):
        cuota = f' anticipoCuota="{int(pago["cuota"])}"'
    monto = _importe(pago['importe'])
    return (
        f'02 <VEP fechaExpiracion="{fecha_expiracion:%Y-%m-%d}"'
        f' nroFormulario="{pago["formulario"]}"'
        f' codTipoPago="{pago["codigo_pago"]}"'
        f' contribuyenteCUIT="{pago["cuit"]}"'
        f' concepto="{pago["concepto"]}"'
        f' subConcepto="{pago["subconcepto"]}"'
        f' periodoFiscal="{pago["periodo"]}"'
        f' importe="{monto}"{cuota}>'
        f'<Obligacion impuesto="{pago["impuesto"]}" importe="{monto}"/></VEP>'
    )


def encabezado(cuit_generador, cantidad):
    """La primera línea: identifica al que sube el archivo y cuántos VEPs trae.

    Los últimos cuatro dígitos son la cantidad más uno, o sea el total de líneas
    del archivo contando esta.
    """
    solo = re.sub(r'\D', '', cuit_generador or '')
    if len(solo) != 11:
        raise ValueError(f"El CUIT del generador tiene que tener 11 dígitos: '{cuit_generador}'")
    return f"01{solo}2000100100003003{cantidad + 1:04d}"


def armar_txt(pagos, cuit_generador, fecha_expiracion=None):
    """Devuelve el contenido completo del archivo."""
    if not pagos:
        raise ValueError("No hay ningún importe seleccionado para pagar.")
    if fecha_expiracion is None:
        fecha_expiracion = datetime.date.today() + datetime.timedelta(days=DIAS_EXPIRACION)

    lineas = [encabezado(cuit_generador, len(pagos))]
    lineas += [linea_vep(p, fecha_expiracion) for p in pagos]
    return "\n".join(lineas) + "\n"


def nombre_archivo(cuit_generador, fecha=None):
    """El nombre que espera ARCA: F20001.cuit.<cuit>.fecha.<AAAAMMDD>.txt"""
    solo = re.sub(r'\D', '', cuit_generador or '')
    fecha = fecha or datetime.date.today()
    return f"F20001.cuit.{solo}.fecha.{fecha:%Y%m%d}.txt"


# ── De la liquidación a la lista de pagos ──────────────────────────────────

def preparar(filas, cuit_contribuyente):
    """Convierte las filas de una liquidación en candidatos a VEP.

    Una fila de la liquidación puede dar hasta cuatro VEPs: capital,
    resarcitorios, capitalizables y punitorios. Devuelve uno por importe con
    algo para pagar, cada uno con su aviso si algo no se pudo resolver.

    No decide qué se paga: eso lo elige Agustín tildando en pantalla.
    """
    tabla, _, _ = cargar_conceptos()
    candidatos = []

    for i, fila in enumerate(filas):
        impuesto, aviso_imp = impuesto_de(fila.get('Impuesto'), cuit_contribuyente)
        concepto, aviso_con = concepto_de(fila.get('concepto'))
        periodo, cuota, aviso_per = periodo_fiscal(
            fila.get('Periodo'), fila.get('Vencimiento'), concepto)

        for columna, subconcepto in COLUMNAS_IMPORTE.items():
            importe = fila.get(columna)
            if importe is None or round(float(importe or 0), 2) <= 0:
                continue

            # El capital lleva el mismo número que el concepto.
            sub = subconcepto if subconcepto is not None else concepto

            avisos = [a for a in (aviso_imp, aviso_con, aviso_per) if a]
            combo = tabla.get((impuesto, concepto, sub)) if impuesto and concepto else None
            if combo is None and not avisos:
                avisos.append(
                    f"La combinación impuesto {impuesto} / concepto {concepto} / "
                    f"subconcepto {sub} no está en la tabla del estudio. "
                    "Agregala a 'conceptos ARCA VEPS.xlsx' o cargá el formulario a mano.")

            candidatos.append({
                'fila': i,
                'columna': columna,
                'Impuesto': fila.get('Impuesto'),
                'concepto_texto': fila.get('concepto'),
                'Vencimiento': fila.get('Vencimiento'),
                'cuit': re.sub(r'\D', '', cuit_contribuyente or ''),
                'impuesto': impuesto,
                'concepto': concepto,
                'subconcepto': sub,
                'periodo': periodo,
                'cuota': cuota,
                'importe': round(float(importe), 2),
                'formulario': combo['formulario'] if combo else None,
                'codigo_pago': combo['codigo_pago'] if combo else None,
                'aviso': ' '.join(avisos),
            })

    return candidatos
