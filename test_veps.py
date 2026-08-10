"""Tests del generador de VEPs.

Se corren con:  python test_veps.py

El caso principal reproduce, carácter por carácter, un archivo que ARCA aceptó
de verdad. Lo único cambiado es el CUIT, por uno inventado: el repositorio es
público. El formato —nombres de atributos, orden, comillas, encabezado— es el
que se cargó. Es la única fuente confiable del formato: la guía publicada está
vieja (usa `precio` en vez de `importe` y pone espacios alrededor de los `=`).

Si alguna vez ARCA cambia el formato y hay que tocar `veps.py`, este test va a
fallar. Eso es lo que tiene que pasar: el archivo de referencia se cambia recién
cuando haya uno nuevo que ARCA haya aceptado, no antes.
"""
import datetime
import sys

import veps

fallos = []


def check(nombre, obtenido, esperado):
    ok = obtenido == esperado
    if not ok:
        fallos.append(f"{nombre}:\n    esperado {esperado!r}\n    obtenido {obtenido!r}")
    print(f"  {'ok   ' if ok else 'FALLA'} {nombre}")


# =====================================================================================
# El archivo real. Cuatro VEPs de aportes de seguridad social (impuesto 301),
# declaración jurada (concepto 19), con los cuatro subconceptos.
# La fecha de expiración vino vacía de la planilla que lo generó y quedó en la
# fecha cero de Excel; se reproduce igual porque así fue aceptado.
REAL = (
    "012099999999720001001000030030005\n"
    '02 <VEP fechaExpiracion="1899-12-30" nroFormulario="800" codTipoPago="30"'
    ' contribuyenteCUIT="20999999997" concepto="19" subConcepto="19"'
    ' periodoFiscal="202606" importe="150000.50">'
    '<Obligacion impuesto="301" importe="150000.50"/></VEP>\n'
    '02 <VEP fechaExpiracion="1899-12-30" nroFormulario="800" codTipoPago="30"'
    ' contribuyenteCUIT="20999999997" concepto="19" subConcepto="51"'
    ' periodoFiscal="202606" importe="845200.00">'
    '<Obligacion impuesto="301" importe="845200.00"/></VEP>\n'
    '02 <VEP fechaExpiracion="1899-12-30" nroFormulario="800" codTipoPago="30"'
    ' contribuyenteCUIT="20999999997" concepto="19" subConcepto="52"'
    ' periodoFiscal="202606" importe="45000.00">'
    '<Obligacion impuesto="301" importe="45000.00"/></VEP>\n'
    '02 <VEP fechaExpiracion="1899-12-30" nroFormulario="800" codTipoPago="30"'
    ' contribuyenteCUIT="20999999997" concepto="19" subConcepto="94"'
    ' periodoFiscal="202606" importe="12450.75">'
    '<Obligacion impuesto="301" importe="12450.75"/></VEP>\n'
)

CUIT = "20999999997"
PAGOS = [
    {'cuit': CUIT, 'impuesto': 301, 'concepto': 19, 'subconcepto': sub,
     'periodo': '202606', 'cuota': 0, 'importe': importe,
     'formulario': 800, 'codigo_pago': 30}
    for sub, importe in ((19, 150000.50), (51, 845200.00), (52, 45000.00), (94, 12450.75))
]

print("\n1) Reproduce el archivo que ARCA aceptó, carácter por carácter")
generado = veps.armar_txt(PAGOS, CUIT, datetime.date(1899, 12, 30))
check("el archivo entero", generado, REAL)
check("el encabezado", generado.split("\n")[0], REAL.split("\n")[0])

# =====================================================================================
print("\n2) El encabezado")
# Los últimos cuatro dígitos son la cantidad de VEPs más uno: el total de líneas.
check("33 caracteres", len(veps.encabezado(CUIT, 4)), 33)
check("4 VEPs -> 0005", veps.encabezado(CUIT, 4)[-4:], "0005")
check("1 VEP -> 0002", veps.encabezado(CUIT, 1)[-4:], "0002")
check("32 VEPs -> 0033", veps.encabezado(CUIT, 32)[-4:], "0033")
check("lleva el CUIT del generador", veps.encabezado(CUIT, 4)[2:13], CUIT)
check("acepta el CUIT con guiones", veps.encabezado("20-99999999-7", 4), veps.encabezado(CUIT, 4))
try:
    veps.encabezado("123", 1)
    fallos.append("un CUIT corto tendría que fallar")
    print("  FALLA protesta con un CUIT corto")
except ValueError:
    print("  ok    protesta con un CUIT corto")

# =====================================================================================
print("\n3) Anticipos: el mes va en 00 y la cuota aparte")
anticipo = dict(PAGOS[0], concepto=191, subconcepto=51, periodo='202600', cuota=3)
linea = veps.linea_vep(anticipo, datetime.date(2026, 8, 19))
check('el período lleva el mes en 00', 'periodoFiscal="202600"' in linea, True)
check('aparece anticipoCuota', 'anticipoCuota="3"' in linea, True)
check('va después del importe', linea.index('importe=') < linea.index('anticipoCuota='), True)

sin_cuota = veps.linea_vep(dict(PAGOS[0]), datetime.date(2026, 8, 19))
check('sin cuota, el atributo no aparece', 'anticipoCuota' not in sin_cuota, True)

# =====================================================================================
print("\n4) De qué tipo es cada CUIT")
for cuit, esperado in (("20999999997", 'fisica'), ("27123456781", 'fisica'),
                       ("23123456781", 'fisica'), ("24123456781", 'fisica'),
                       ("30999999995", 'juridica'), ("33123456781", 'juridica'),
                       ("34123456781", 'juridica'), ("99123456781", ''), ("123", '')):
    check(f"{cuit} -> {esperado or 'desconocido'}", veps.tipo_de_persona(cuit), esperado)

# =====================================================================================
print("\n5) Ganancias cambia de código según el CUIT")
check("sociedad -> 10", veps.impuesto_de("IMPUESTO A LAS GANANCIAS", "30999999995")[0], 10)
check("persona física -> 11", veps.impuesto_de("IMPUESTO A LAS GANANCIAS", "20999999997")[0], 11)
codigo, aviso = veps.impuesto_de("IMPUESTO A LAS GANANCIAS", "99123456781")
check("CUIT raro -> no adivina", codigo, None)
check("  y explica por qué", 'sociedad' in aviso, True)

print("\n   los demás impuestos no dependen del CUIT")
for nombre, esperado in (
    ("IMPUESTO AL VALOR AGREGADO LEY 23349 Y SUS MODIFICACIONES", 30),
    ("APORTES DE LA SEGURIDAD SOCIAL LEY 24241", 301),
    ("CONTRIBUCIONES SEGURIDAD SOCIAL", 351),
    ("SICORE-IMPTO.A LAS GANANCIAS", 217),
    ("IMPTO.S/BIENES PERSONALES", 180),
):
    check(f"  {nombre[:38]} -> {esperado}", veps.impuesto_de(nombre, CUIT)[0], esperado)

codigo, aviso = veps.impuesto_de("IMPUESTO A LOS DEBITOS Y CREDITOS", CUIT)
check("un impuesto que no está -> no adivina", codigo, None)
check("  y lo dice", 'No reconozco' in aviso, True)

# =====================================================================================
print("\n6) El período fiscal")
F = datetime.date
check("anticipo 2026-3", veps.periodo_fiscal('2026-3', F(2026, 5, 13), 191, 10)[:2], ('202600', 3))
check("IVA 2024/5 (mensual)", veps.periodo_fiscal('2024/5', F(2024, 6, 25), 19, 30)[:2], ('202405', 0))
check("seg. social 2025/12 (mensual)",
      veps.periodo_fiscal('2025/12', F(2026, 1, 13), 19, 301)[:2], ('202512', 0))
periodo, cuota, aviso = veps.periodo_fiscal('', F(2024, 6, 25), 19, 30)
check("sin período, cae al vencimiento", periodo, '202406')
check("  y avisa", bool(aviso), True)

print("\n   los anuales llevan el mes en 00, con cuota o sin ella")
# Ganancias y Bienes Personales son anuales: la declaración jurada va con el mes
# en 00 y sin cuota; los anticipos, con el mes en 00 y la cuota que corresponda.
check("Ganancias DDJJ 2024/0", veps.periodo_fiscal('2024/0', F(2025, 6, 26), 19, 10)[:2],
      ('202400', 0))
check("Bienes Personales DDJJ 2024/0",
      veps.periodo_fiscal('2024/0', F(2025, 6, 26), 19, 180)[:2], ('202400', 0))
check("Ganancias anticipo cuota 1", veps.periodo_fiscal('2026-1', F(2025, 10, 13), 191, 11)[:2],
      ('202600', 1))
check("Bienes Personales anticipo cuota 2",
      veps.periodo_fiscal('2026-2', F(2025, 11, 13), 191, 180)[:2], ('202600', 2))

# La red: si una planilla cargada a mano pone el mes de un anual, se corrige y se avisa.
periodo, cuota, aviso = veps.periodo_fiscal('2024/6', F(2025, 6, 26), 19, 10)
check("un anual con mes real se corrige a 00", periodo, '202400')
check("  y lo avisa", 'anual' in aviso, True)

# Sin período fiscal legible, un anual usa el año del vencimiento, no el mes.
periodo, _, aviso = veps.periodo_fiscal('', F(2025, 6, 26), 19, 10)
check("un anual sin período usa solo el año", periodo, '202500')

# =====================================================================================
print("\n7) La liquidación completa: cada fila da hasta cuatro VEPs")
# El caso que motivó todo esto: ocho anticipos, cada uno con capital y los tres
# intereses. Treinta y dos VEPs cargados de a uno.
FILAS = [{
    'Impuesto': 'IMPUESTO A LAS GANANCIAS', 'concepto': 'ANTICIPOS',
    'Periodo': f'2026-{i}', 'Vencimiento': F(2025, 10, 13),
    'Capital': 366033.27, 'Interes_Resarcitorio': 80527.32,
    'Interes_Capitalizable': 664.35, 'Interes_Punitorio': 23060.16,
} for i in range(1, 9)]

candidatos = veps.preparar(FILAS, "30999999995")
check("8 filas x 4 importes = 32 VEPs", len(candidatos), 32)
check("ninguno quedó con avisos", [c for c in candidatos if c['aviso']], [])
check("todos con formulario", all(c['formulario'] for c in candidatos), True)
check("el capital lleva el subconcepto del concepto (191)",
      next(c['subconcepto'] for c in candidatos if c['columna'] == 'Capital'), 191)
check("los resarcitorios llevan 51",
      next(c['subconcepto'] for c in candidatos if c['columna'] == 'Interes_Resarcitorio'), 51)
check("los capitalizables llevan 52",
      next(c['subconcepto'] for c in candidatos if c['columna'] == 'Interes_Capitalizable'), 52)
check("los punitorios llevan 94",
      next(c['subconcepto'] for c in candidatos if c['columna'] == 'Interes_Punitorio'), 94)
check("es Ganancias Sociedades, por el CUIT", candidatos[0]['impuesto'], 10)
check("la cuota sale del período", candidatos[0]['cuota'], 1)

# El archivo de esos 32 tiene 33 líneas y el encabezado lo declara.
txt = veps.armar_txt(candidatos, "30999999995", F(2026, 8, 19))
check("33 líneas en total", len(txt.strip().split("\n")), 33)
check("el encabezado dice 0033", txt[:33][-4:], "0033")

# Los importes en cero no generan VEP: no se paga nada.
sin_punitorios = [dict(f, Interes_Punitorio=0) for f in FILAS]
check("un importe en cero no genera VEP", len(veps.preparar(sin_punitorios, "30999999995")), 24)

# =====================================================================================
print("\n8) Lo que no se puede resolver, se marca (no se adivina)")
raras = [{'Impuesto': 'IMPUESTO A LOS SELLOS', 'concepto': 'DECLARACION JURADA',
          'Periodo': '2025/3', 'Vencimiento': F(2025, 4, 13), 'Capital': 1000.0}]
c = veps.preparar(raras, "30999999995")
check("queda el candidato", len(c), 1)
check("sin formulario", c[0]['formulario'], None)
check("con el motivo explicado", 'No reconozco' in c[0]['aviso'], True)

try:
    veps.armar_txt([], CUIT)
    fallos.append("un archivo vacío tendría que fallar")
    print("  FALLA protesta con la lista vacía")
except ValueError:
    print("  ok    protesta con la lista vacía")

# Un pago sin formulario no puede llegar al archivo. La pantalla ya los aparta,
# pero si alguna vez se llama a armar_txt() desde otro lado, tiene que cortar:
# un archivo con nroFormulario="None" tiene forma de archivo y no sirve.
try:
    veps.armar_txt([dict(PAGOS[0], formulario=None)], CUIT)
    fallos.append("un pago sin formulario tendría que fallar")
    print("  FALLA no deja pasar un pago sin formulario")
except ValueError as e:
    check("no deja pasar un pago sin formulario", 'faltan en la tabla' in str(e), True)

# =====================================================================================
print("\n9) El nombre del archivo")
check("como lo espera ARCA", veps.nombre_archivo(CUIT, F(2026, 8, 9)),
      "F20001.cuit.20999999997.fecha.20260809.txt")

# =====================================================================================
print("\n" + "=" * 60)
if fallos:
    print(f"{len(fallos)} FALLA(S):")
    for f_ in fallos:
        print("  -", f_)
    sys.exit(1)
print("Todos los tests pasaron.")
