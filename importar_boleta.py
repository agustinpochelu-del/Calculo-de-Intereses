"""Convierte una boleta de deuda en planillas, desde la línea de comandos.

Lo usa la rutina automática que revisa el correo: el asistente busca el mail y
guarda el cuerpo en un archivo, y este programa hace el resto. La división es a
propósito — **ningún importe pasa por la interpretación del asistente**. Los
números salen del mismo código probado contra ARCA que usa la app.

Uso:

    python importar_boleta.py CUERPO.html --salida CARPETA [--remitente DIRECCION]
    python importar_boleta.py MAIL.eml    --salida CARPETA

Escribe en CARPETA una o dos planillas .xlsx y muestra un resumen. Con `--json`
devuelve el resumen en JSON, para que la rutina lo pueda relatar sin tener que
leer los importes uno por uno.

Códigos de salida:
    0  se leyó bien y la suma coincide con el monto de demanda
    1  se leyó, pero hay algo para mirar (la suma no cierra, o hay filas sin clasificar)
    2  no es una boleta de deuda, o no se pudo leer
"""
import argparse
import datetime
import json
import os
import re
import sys

import pandas as pd

import leer_mail
from planillas import COLUMNAS_CAPITAL, COLUMNAS_INTERESES, armar_planilla


def _nombre_archivo(texto):
    """Un nombre de archivo tranquilo, sin acentos ni caracteres raros."""
    limpio = leer_mail._pelar(texto or 'boleta').strip()
    limpio = re.sub(r'[^\w\s-]', '', limpio)
    limpio = re.sub(r'\s+', '_', limpio).strip('_')
    return (limpio or 'boleta')[:50]


def procesar(ruta, salida, remitente='', asunto='', habilitados=(), liquidacion=None):
    with open(ruta, 'rb') as f:
        crudo = f.read()

    if ruta.lower().endswith('.eml'):
        boleta = leer_mail.leer_boleta(crudo, habilitados)
    else:
        html = crudo.decode('utf-8', 'replace')
        boleta = leer_mail.leer_boleta_html(html, remitente, asunto, habilitados)

    df = pd.DataFrame(boleta['filas'])
    df['Vencimiento'] = pd.to_datetime(df['Vencimiento'], errors='coerce')
    df['F. Pago Capital'] = pd.to_datetime(df['F. Pago Capital'], errors='coerce')
    df['fecha_Demanda'] = pd.to_datetime(boleta['fecha_demanda'], errors='coerce')
    df['Fecha_Liquidacion'] = pd.Timestamp(liquidacion or datetime.date.today())

    suma = round(float(df['Capital'].sum()), 2)
    declarado = boleta['monto_demanda']
    cuadra = declarado is not None and abs(suma - declarado) < 0.01

    os.makedirs(salida, exist_ok=True)
    base = _nombre_archivo(boleta['contribuyente'])
    generadas = []
    for destino, columnas, etiqueta in (
        (leer_mail.CAPITAL, COLUMNAS_CAPITAL, 'Capital'),
        (leer_mail.INTERESES, COLUMNAS_INTERESES, 'Intereses'),
    ):
        parte = df[df['Destino'] == destino]
        if parte.empty:
            continue
        ruta_salida = os.path.join(salida, f"{base}_{etiqueta}.xlsx")
        with open(ruta_salida, 'wb') as f:
            f.write(armar_planilla(parte, columnas))
        generadas.append({
            'archivo': ruta_salida,
            'planilla': etiqueta,
            'filas': len(parte),
            'total': round(float(parte['Capital'].sum()), 2),
        })

    revisar = df[df['Destino'] == leer_mail.REVISAR]
    return {
        'contribuyente': boleta['contribuyente'],
        'cuit': boleta['cuit'],
        'juicio': boleta['juicio'],
        'fecha_demanda': boleta['fecha_demanda'],
        'fecha_liquidacion': str(liquidacion or datetime.date.today()),
        'remitente': boleta['remitente'] or remitente,
        'filas': len(df),
        'suma': suma,
        'monto_demanda': declarado,
        'cuadra': cuadra,
        'planillas': generadas,
        'a_revisar': [
            {'vencimiento': str(f['Vencimiento'].date()) if pd.notna(f['Vencimiento']) else '',
             'capital': round(float(f['Capital']), 2),
             'concepto': f['concepto'],
             'motivo': f['Aviso'] or f'Nota: {f["Nota"]}'}
            for _, f in revisar.iterrows()
        ],
        'advertencias': (
            boleta['avisos']
            + [f"{f['concepto']} venc {f['Vencimiento']:%d/%m/%Y}: {f['Aviso']}"
               for _, f in df.iterrows()
               if f['Aviso'] and f['Destino'] != leer_mail.REVISAR and pd.notna(f['Vencimiento'])]
        ),
    }


def main():
    p = argparse.ArgumentParser(description="Convierte una boleta de deuda de ARCA en planillas.")
    p.add_argument('archivo', help="cuerpo del mail en .html, o el mail entero en .eml")
    p.add_argument('--salida', required=True, help="carpeta donde dejar las planillas")
    p.add_argument('--remitente', default='', help="de quién vino (para los .html)")
    p.add_argument('--asunto', default='', help="asunto del mail (para los .html)")
    p.add_argument('--habilitados', default='',
                   help="direcciones de agentes fiscales conocidos, separadas por coma")
    p.add_argument('--liquidacion', default='',
                   help="fecha de liquidación AAAA-MM-DD (por defecto, hoy)")
    p.add_argument('--json', action='store_true', help="mostrar el resumen en JSON")
    args = p.parse_args()

    habilitados = [d.strip().lower() for d in args.habilitados.split(',') if d.strip()]
    liquidacion = datetime.date.fromisoformat(args.liquidacion) if args.liquidacion else None

    try:
        r = procesar(args.archivo, args.salida, args.remitente, args.asunto,
                     habilitados, liquidacion)
    except ValueError as e:
        # No es una boleta de deuda: es el caso normal cuando el agente fiscal
        # manda otra cosa (honorarios, una consulta suelta). No es un error.
        salida = {'boleta': False, 'motivo': str(e)}
        print(json.dumps(salida, ensure_ascii=False) if args.json else f"No es una boleta: {e}")
        return 2

    if args.json:
        print(json.dumps({'boleta': True, **r}, ensure_ascii=False, indent=2))
    else:
        print(f"{r['contribuyente']} — CUIT {r['cuit']} — juicio {r['juicio']}")
        print(f"  {r['filas']} filas por {r['suma']:,.2f}", end='')
        if r['cuadra']:
            print("  ✓ coincide con el monto de demanda")
        else:
            print(f"  ✗ la boleta declara {r['monto_demanda']}")
        print(f"  fecha de demanda: {r['fecha_demanda']}  |  liquidación: {r['fecha_liquidacion']}")
        for g in r['planillas']:
            print(f"  → {g['archivo']}  ({g['filas']} filas, {g['total']:,.2f})")
        for f in r['a_revisar']:
            print(f"  ! revisar {f['concepto']} venc {f['vencimiento']}: {f['motivo']}")
        for a in r['advertencias']:
            print(f"  ! {a}")

    return 0 if (r['cuadra'] and not r['a_revisar'] and not r['advertencias']) else 1


if __name__ == '__main__':
    sys.exit(main())
