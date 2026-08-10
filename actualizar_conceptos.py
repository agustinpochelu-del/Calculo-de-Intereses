"""Pasa la tabla de conceptos del estudio a `conceptos_veps.json`, validándola.

La fuente es la planilla **"conceptos ARCA VEPS.xlsx"**, que se edita a mano y no
va al repositorio. Este programa la lee, revisa que esté sana y escribe el JSON
que usa la app.

    python actualizar_conceptos.py

Se corre cada vez que se agrega una combinación de impuesto / concepto /
subconcepto a la planilla. Si algo no cierra, lo dice y no escribe nada: es
preferible seguir con la tabla vieja que con una tabla contradictoria.

Códigos de salida:
    0  se escribió el JSON
    1  la planilla tiene problemas y no se escribió nada
    2  no se encontró la planilla
"""
import datetime
import json
import os
import sys

import pandas as pd

CARPETA = os.path.dirname(os.path.abspath(__file__))
PLANILLA = os.path.join(CARPETA, "conceptos ARCA VEPS.xlsx")
SALIDA = os.path.join(CARPETA, "conceptos_veps.json")

COLUMNAS = ['imp_n', 'impuesto', 'con_n', 'concepto', 'sub_n', 'subconcepto',
            'form', 'cod']

# Los números que ARCA le da a cada concepto, y cómo se llaman en la planilla.
# Sirve para detectar el error fácil: escribir "Anticipo" en el texto pero dejar
# el número de la declaración jurada, o al revés.
NUMERO_DE_CONCEPTO = {'declaracion jurada': 19, 'anticipo': 191}


def revisar(df):
    """Devuelve la lista de problemas encontrados. Vacía si está sana."""
    problemas = []

    for columna, que in (('form', 'formulario'), ('cod', 'código de pago')):
        faltan = df[df[columna].isna()]
        for _, r in faltan.iterrows():
            problemas.append(
                f"fila {r.name + 2}: {r['impuesto']} / {r['concepto']} / "
                f"{r['subconcepto']} no tiene {que}.")

    # El número del concepto tiene que coincidir con su nombre.
    for _, r in df.iterrows():
        nombre = str(r['concepto']).strip().lower()
        esperado = NUMERO_DE_CONCEPTO.get(nombre)
        if esperado is not None and not pd.isna(r['con_n']) and int(r['con_n']) != esperado:
            problemas.append(
                f"fila {r.name + 2}: dice \"{r['concepto']}\" pero el número de concepto "
                f"es {int(r['con_n'])}; para \"{r['concepto']}\" tiene que ser {esperado}.")

    # Dos filas con la misma combinación que manden a formularios distintos.
    clave = ['imp_n', 'con_n', 'sub_n']
    for valores, g in df.groupby(clave):
        if len(g) > 1 and (g['form'].nunique() > 1 or g['cod'].nunique() > 1):
            filas = ", ".join(str(i + 2) for i in g.index)
            problemas.append(
                f"impuesto {valores[0]} / concepto {valores[1]} / subconcepto {valores[2]} "
                f"aparece más de una vez con formularios distintos (filas {filas}).")

    return problemas


def main():
    if not os.path.exists(PLANILLA):
        print(f"No encontré la planilla:\n  {PLANILLA}")
        return 2

    df = pd.read_excel(PLANILLA)
    if len(df.columns) < len(COLUMNAS):
        print(f"La planilla tiene {len(df.columns)} columnas y esperaba {len(COLUMNAS)}: "
              f"{COLUMNAS}")
        return 1
    df.columns = COLUMNAS[:len(df.columns)]

    problemas = revisar(df)
    if problemas:
        print(f"La planilla tiene {len(problemas)} problema(s). No escribí nada:\n")
        for p in problemas:
            print("  •", p)
        print(f"\nCorregilos en:\n  {PLANILLA}\ny volvé a correr esto.")
        return 1

    impuestos = {str(int(n)): nom for (n, nom), _ in df.groupby(['imp_n', 'impuesto'])}
    combinaciones = [{
        'impuesto': int(r['imp_n']),
        'concepto': int(r['con_n']),
        'subconcepto': int(r['sub_n']),
        'formulario': int(r['form']),
        'codigo_pago': int(r['cod']),
        'texto': f"{r['impuesto']} - {r['concepto']} - {r['subconcepto']}",
    } for _, r in df.iterrows()]

    with open(SALIDA, 'w', encoding='utf-8') as f:
        json.dump({
            'origen': 'Planilla "conceptos ARCA VEPS.xlsx" del estudio',
            'actualizado': str(datetime.date.today()),
            'impuestos': impuestos,
            'combinaciones': combinaciones,
        }, f, ensure_ascii=False, indent=2)

    print(f"Escrito {os.path.basename(SALIDA)}: {len(combinaciones)} combinaciones, "
          f"{len(impuestos)} impuestos.")
    for n, nom in sorted(impuestos.items(), key=lambda x: int(x[0])):
        cuantas = sum(1 for c in combinaciones if c['impuesto'] == int(n))
        print(f"  {n:>4}  {nom:<34} {cuantas} combinaciones")
    return 0


if __name__ == '__main__':
    sys.exit(main())
