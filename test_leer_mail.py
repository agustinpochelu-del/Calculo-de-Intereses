"""Tests del lector de mails de boleta de deuda.

Se corren con:  python test_leer_mail.py

El mail de prueba es inventado, pero reproduce la estructura exacta de los que
mandan los agentes fiscales: la tabla "Detalle de Deuda", la de "Pagos de Capital
Registrados", la ficha con la fecha de sorteo, el nombre del impuesto cortado en
varias líneas y los acentos rotos que dejan algunos reenvíos.

Los mails reales NO van al repositorio: es público y llevan CUIT, domicilio y
deuda de contribuyentes.
"""
import sys
from email.message import EmailMessage

import leer_mail

fallos = []


def check(nombre, obtenido, esperado):
    ok = obtenido == esperado
    if not ok:
        fallos.append(f"{nombre}: esperado {esperado!r}, obtenido {obtenido!r}")
    print(f"  {'ok   ' if ok else 'FALLA'} {nombre}: {obtenido!r}")


# ── El mail de prueba ───────────────────────────────────────────────────────
# "Perï¿½odo" no es un error de tipeo: así llega cuando el reenvío rompe la
# codificación, y el lector tiene que entenderlo igual.

FILA = """
<tr>
  <td>{impuesto}<br>{concepto}<br>{subconcepto}<br>{cuota}
      Rectificativa/Denuncia Z N&ordm;: 50<br>
      Fecha de Resoluci&oacute;n/Intimaci&oacute;n: 20/3/2026</td>
  <td>{importe}<br><br>{nota}</td>
  <td>Perï¿½odo: {periodo}<br>{vencimiento}</td>
  <td>NO</td>
</tr>"""

HTML = """<html><body>
<table>
  <tr><td>C.U.I.T. :</td><td>30999999995 - Establec. 00</td></tr>
  <tr><td>Contribuyente :</td><td>EMPRESA DE<br>PRUEBA S.A.</td></tr>
  <tr><td>Nro. Juicio :</td><td>434 / 111111 /<br>2026</td></tr>
  <tr><td>Monto Demanda :</td><td>1.482.000,00</td></tr>
</table>

<table>
  <tr><th>Impuestos - Conceptos - Subconceptos</th><th>Monto de la Deuda</th>
      <th>Perï¿½odo<br>Vencimiento</th><th>Hon.</th></tr>
  """ + FILA.format(
    impuesto="IMPUESTO A LAS GANANCIAS", concepto="DECLARACI&Oacute;N JURADA",
    subconcepto="SALDO DE DECLARACI&Oacute;N JURADA", cuota="",
    importe="1000000,00", nota="DEBE", periodo="2024/0", vencimiento="26/6/2025",
) + FILA.format(
    # El nombre del impuesto viene cortado en dos líneas, como en los mails reales.
    impuesto="IMPUESTO A LAS<br>GANANCIAS", concepto="ANTICIPOS",
    subconcepto="ANTICIPOS", cuota="Cuota: 1<br>",
    importe="400000,00", nota="pto. DJ debe intereses", periodo="2026/0",
    vencimiento="13/10/2025",
) + FILA.format(
    impuesto="IMPUESTO AL VALOR<br>AGREGADO LEY 23349 Y SUS MODIFICACIONES",
    concepto="DECLARACI&Oacute;N JURADA", subconcepto="INTERESES RESARCITORIOS",
    cuota="", importe="80000,00", nota="Compensaci&oacute;n de int. resarc<br>DEBE INT. CAPITALIZABLES + PUNITORIOS",
    periodo="2024/5", vencimiento="25/6/2024",
) + FILA.format(
    impuesto="APORTES DE LA SEGURIDAD SOCIAL LEY 24241",
    concepto="DECLARACI&Oacute;N JURADA", subconcepto="SALDO DE DECLARACI&Oacute;N JURADA",
    cuota="", importe="2000,00", nota="PLAN VIGENTE RG5321-W064792<br>CAPITAL + RESARC + PUNITORIOS",
    periodo="2025/12", vencimiento="13/1/2026",
) + """
  <tr><td>Suma Total del Detalle de Deuda:</td><td>1482000,00</td><td></td><td></td></tr>
</table>

<table>
  <tr><td>Expediente :</td><td>9015 / 2026</td></tr>
  <tr><td>Fecha Sorteo :</td><td>10/06/2026 18:43:48</td></tr>
</table>

<table>
  <tr><th>Impuestos - Conceptos - Subconceptos<br>Perï¿½odo / Vencimiento</th>
      <th>Importe del Pago</th><th>Detalle del Pago</th></tr>
  <tr><td>IMPUESTO A LAS GANANCIAS<br>ANTICIPOS<br>ANTICIPOS<br>Cuota: 1<br>
          Perï¿½odo: 2026/0 -  13/10/2025</td>
      <td>400000,00</td>
      <td>3/8/2026<br>Pagado en : () - Pago No Bancario</td></tr>
</table>
</body></html>"""


def armar_mail(html=HTML, de="abogado@arca.gob.ar"):
    msg = EmailMessage()
    msg['From'] = f"Agente Fiscal <{de}>"
    msg['To'] = "estudio@ejemplo.com"
    msg['Subject'] = "EMPRESA DE PRUEBA S.A."
    msg.set_content("Te paso la boleta.")
    msg.add_alternative(html, subtype='html')
    return msg.as_bytes()


# =====================================================================================
print("\n1) La carátula de la boleta")
b = leer_mail.leer_boleta(armar_mail())
check("contribuyente (venía cortado en dos líneas)", b['contribuyente'], "EMPRESA DE PRUEBA S.A.")
check("CUIT", b['cuit'], "30999999995")
check("juicio", b['juicio'], "434 / 111111 / 2026")
check("monto de demanda", b['monto_demanda'], 1482000.00)
check("fecha de demanda (sale de la fecha de sorteo)", b['fecha_demanda'], "2026-06-10")
check("remitente", b['remitente'], "abogado@arca.gob.ar")

# =====================================================================================
print("\n2) Las filas de deuda")
filas = b['filas']
check("cantidad de filas", len(filas), 4)
check("la suma da el monto de demanda", round(sum(f['Capital'] for f in filas), 2), 1482000.00)

saldo, anticipo, iva, plan = filas
check("impuesto del saldo", saldo['Impuesto'], "IMPUESTO A LAS GANANCIAS")
check("concepto del saldo", saldo['concepto'], "SALDO DE DECLARACIÓN JURADA")
check("período del saldo", saldo['Periodo'], "2024/0")
check("vencimiento del saldo", saldo['Vencimiento'], "2025-06-26")
check("capital del saldo", saldo['Capital'], 1000000.00)

# El nombre del impuesto llega cortado en dos líneas y hay que volver a pegarlo.
check("impuesto del anticipo (venía en dos líneas)", anticipo['Impuesto'], "IMPUESTO A LAS GANANCIAS")
check("período del anticipo (año y cuota)", anticipo['Periodo'], "2026-1")
check("vencimiento del anticipo", anticipo['Vencimiento'], "2025-10-13")
check("fecha de pago del capital", anticipo['F. Pago Capital'], "2026-08-03")

check("impuesto del IVA (dos líneas, nombre largo)", iva['Impuesto'],
      "IMPUESTO AL VALOR AGREGADO LEY 23349 Y SUS MODIFICACIONES")
check("vencimiento del IVA", iva['Vencimiento'], "2024-06-25")

# =====================================================================================
print("\n3) La clasificación que se propone")
# Manda el subconcepto de ARCA, no la nota que escribe el agente fiscal.
check("saldo de DDJJ -> capital", saldo['Destino'], leer_mail.CAPITAL)
check("anticipos -> capital", anticipo['Destino'], leer_mail.CAPITAL)
check("intereses resarcitorios -> intereses", iva['Destino'], leer_mail.INTERESES)
check("plan de pagos -> revisar", plan['Destino'], leer_mail.REVISAR)
check("y se explica por qué", bool(plan['Aviso']), True)

# El anticipo tiene el pago registrado, así que no hay nada que advertir.
check("anticipo pagado, sin advertencia", anticipo['Aviso'], '')

# =====================================================================================
print("\n4) Aviso cuando la nota dice que se pagó pero no figura el pago")
# Sin la fecha de pago, los punitorios corren hasta la liquidación y salen de más.
sin_pago = HTML.replace("""<tr><td>IMPUESTO A LAS GANANCIAS<br>ANTICIPOS<br>ANTICIPOS<br>Cuota: 1<br>
          Perï¿½odo: 2026/0 -  13/10/2025</td>
      <td>400000,00</td>
      <td>3/8/2026<br>Pagado en : () - Pago No Bancario</td></tr>""", "")
anticipo_solo = leer_mail.leer_boleta(armar_mail(sin_pago))['filas'][1]
check("quedó sin fecha de pago", anticipo_solo['F. Pago Capital'], '')
check("avisa que la nota no cierra con la boleta", 'punitorios' in anticipo_solo['Aviso'], True)

# =====================================================================================
print("\n5) Remitentes")
b2 = leer_mail.leer_boleta(armar_mail(de="otro@gmail.com"),
                           remitentes_habilitados=["abogado@arca.gob.ar"])
check("avisa si viene de una dirección desconocida", len(b2['avisos']), 1)
check("pero igual lee las filas", len(b2['filas']), 4)
b3 = leer_mail.leer_boleta(armar_mail(), remitentes_habilitados=["abogado@arca.gob.ar"])
check("sin aviso si el remitente está en la lista", len(b3['avisos']), 0)

# =====================================================================================
print("\n6) Mails que no se pueden leer")
for descripcion, datos in (
    ("un mail sin la tabla de deuda", armar_mail("<html><body><p>Hola, te llamo luego.</p></body></html>")),
):
    try:
        leer_mail.leer_boleta(datos)
        fallos.append(f"{descripcion}: tendría que haber fallado")
        print(f"  FALLA {descripcion}: no protestó")
    except ValueError as e:
        print(f"  ok    {descripcion}: protesta bien ({str(e)[:55]}...)")

# =====================================================================================
print("\n" + "=" * 60)
if fallos:
    print(f"{len(fallos)} FALLA(S):")
    for f in fallos:
        print("  -", f)
    sys.exit(1)
print("Todos los tests pasaron.")
