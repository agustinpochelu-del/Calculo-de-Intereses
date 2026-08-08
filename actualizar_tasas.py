"""Control de tasas contra la página oficial de ARCA.

Baja la tabla "Evolución de Tasas de Intereses" del sitio de ARCA, la compara con
`tasas.json` y avisa si hay diferencias.

    python actualizar_tasas.py              # solo informa
    python actualizar_tasas.py --escribir   # además actualiza tasas.json

Códigos de salida: 0 = sin novedades, 1 = hay diferencias, 2 = no se pudo verificar.
Los usa el control automático de .github/workflows/control-tasas.yml.

Solo usa la librería estándar, así que corre en cualquier Python 3 sin instalar nada.

Ojo con una limitación: la página publica las fechas y las tasas, pero NO los días
que ARCA computa por tramo (el campo "dias" de tasas.json, que sale del detalle de
cálculo). Un tramo nuevo va a entrar con "dias": null, y la app lo va a contar sola
avisando en pantalla. Para dejarlo exacto hay que pedirle a ARCA un detalle de
cálculo que atraviese ese tramo y copiar el número.
"""
import argparse
import datetime
import json
import os
import ssl
import sys
import urllib.error
import urllib.request
from html.parser import HTMLParser

URL = "https://serviciosweb.afip.gob.ar/genericos/calculoInteres/punitorios.aspx"
ARCHIVO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tasas.json")
ABIERTO = "2999-12-31"  # cierre del tramo vigente, tal como lo publica ARCA


# =====================================================================================
# BAJADA Y PARSEO DE LA TABLA OFICIAL
# =====================================================================================
class _Tablas(HTMLParser):
    """Extrae las tablas de la página como listas de listas de texto."""

    def __init__(self):
        super().__init__()
        self.tablas = []
        self.fila = []
        self.celda = []
        self.en_celda = False

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.tablas.append([])
        elif tag == "tr":
            self.fila = []
        elif tag in ("td", "th"):
            self.en_celda, self.celda = True, []

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self.en_celda = False
            self.fila.append(" ".join("".join(self.celda).split()))
        elif tag == "tr" and self.tablas and self.fila:
            self.tablas[-1].append(self.fila)
            self.fila = []

    def handle_data(self, data):
        if self.en_celda:
            self.celda.append(data)


def _a_iso(fecha_dmy):
    return datetime.datetime.strptime(fecha_dmy, "%d/%m/%Y").date().isoformat()


def bajar_tramos(url=URL, timeout=60):
    """Devuelve los tramos de la tabla oficial, del más nuevo al más viejo."""
    contexto = ssl.create_default_context()
    # El certificado del sitio de ARCA no siempre valida con el store local; la
    # respuesta se controla igual por estructura más abajo.
    contexto.check_hostname = False
    contexto.verify_mode = ssl.CERT_NONE
    pedido = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    html = urllib.request.urlopen(pedido, context=contexto, timeout=timeout).read()
    html = html.decode("utf-8", "replace")

    parser = _Tablas()
    parser.feed(html)
    if not parser.tablas:
        raise ValueError("la página no trajo ninguna tabla")

    filas = max(parser.tablas, key=len)
    tramos = []
    for fila in filas:
        # Desde | Hasta | Norma | Resarc. % mensual | % diario | Punit. % mensual | % diario
        if len(fila) != 7:
            continue
        try:
            desde, hasta = _a_iso(fila[0]), _a_iso(fila[1])
            resarcitoria, punitoria = float(fila[3]), float(fila[5])
        except ValueError:
            continue  # fila de encabezado o basura
        tramos.append({
            "desde": desde,
            "hasta": hasta,
            "norma": fila[2],
            "resarcitoria_mensual": resarcitoria,
            "punitoria_mensual": punitoria,
        })

    if len(tramos) < 20:
        raise ValueError(
            f"solo se reconocieron {len(tramos)} tramos: la página debe haber cambiado de formato")
    return tramos


# =====================================================================================
# COMPARACIÓN CONTRA tasas.json
# =====================================================================================
CAMPOS = ("norma", "resarcitoria_mensual", "punitoria_mensual")


def comparar(guardados, oficiales):
    """Devuelve (nuevos, cambiados, faltantes) comparando por el par de fechas."""
    por_fecha_g = {(t["desde"], t["hasta"]): t for t in guardados}
    por_fecha_o = {(t["desde"], t["hasta"]): t for t in oficiales}

    nuevos = [o for k, o in por_fecha_o.items() if k not in por_fecha_g]
    faltantes = [g for k, g in por_fecha_g.items() if k not in por_fecha_o]
    cambiados = []
    for k, o in por_fecha_o.items():
        g = por_fecha_g.get(k)
        if g and any(g.get(c) != o.get(c) for c in CAMPOS):
            cambiados.append((g, o))
    return nuevos, cambiados, faltantes


def verificar_continuidad(tramos):
    """Devuelve los pares de tramos consecutivos que no empalman."""
    ordenados = sorted(tramos, key=lambda t: t["desde"])
    huecos = []
    for a, b in zip(ordenados, ordenados[1:]):
        esperado = datetime.date.fromisoformat(a["hasta"]) + datetime.timedelta(days=1)
        if esperado.isoformat() != b["desde"]:
            huecos.append((a["hasta"], b["desde"]))
    return huecos


def fusionar(guardados, oficiales):
    """Arma la lista nueva conservando los días oficiales que ya estaban cargados.

    Se buscan por fecha de inicio, no por el par de fechas: cuando sale una tasa nueva,
    ARCA le cierra el "hasta" al tramo que estaba vigente, y con la clave doble se
    perdían los días de todos los tramos. Si además cambió el "hasta" de un tramo ya
    cerrado, los días se descartan: eran de otro período.
    """
    por_desde = {t["desde"]: t for t in guardados}
    resultado = []
    for o in oficiales:
        tramo = dict(o)
        previo = por_desde.get(o["desde"])
        tramo["dias"] = previo.get("dias") if previo and previo["hasta"] == o["hasta"] else None
        resultado.append(tramo)
    return resultado


# =====================================================================================
def main():
    ap = argparse.ArgumentParser(description="Controla las tasas contra la página de ARCA.")
    ap.add_argument("--escribir", action="store_true", help="actualiza tasas.json si hay diferencias")
    ap.add_argument("--archivo", default=ARCHIVO)
    args = ap.parse_args()

    with open(args.archivo, encoding="utf-8") as fh:
        doc = json.load(fh)
    guardados = doc["tramos"]

    try:
        oficiales = bajar_tramos()
    except (urllib.error.URLError, ValueError, TimeoutError) as e:
        print(f"NO SE PUDO VERIFICAR: {type(e).__name__}: {e}")
        print("La página de ARCA no respondió o cambió de formato. Las tasas guardadas "
              "siguen siendo las últimas verificadas.")
        return 2

    nuevos, cambiados, faltantes = comparar(guardados, oficiales)
    print(f"tramos en tasas.json: {len(guardados)}   en la página de ARCA: {len(oficiales)}")

    if not (nuevos or cambiados or faltantes):
        print(f"\nSIN NOVEDADES: las tasas guardadas coinciden con las de ARCA.")
        if args.escribir:
            doc["verificado"] = datetime.date.today().isoformat()
            _guardar(args.archivo, doc)
            print("Se actualizó la fecha de verificación.")
        return 0

    print()
    for t in nuevos:
        vigente = " (tramo vigente)" if t["hasta"] == ABIERTO else ""
        print(f"TRAMO NUEVO{vigente}: {_fmt(t['desde'])} al {_fmt(t['hasta'])} — "
              f"resarcitorios {t['resarcitoria_mensual']}% / punitorios {t['punitoria_mensual']}% "
              f"— {t['norma']}")
    for g, o in cambiados:
        print(f"TRAMO CAMBIADO: {_fmt(o['desde'])} al {_fmt(o['hasta'])}")
        for c in CAMPOS:
            if g.get(c) != o.get(c):
                print(f"    {c}: {g.get(c)} -> {o.get(c)}")
    for t in faltantes:
        print(f"TRAMO QUE YA NO ESTA EN ARCA: {_fmt(t['desde'])} al {_fmt(t['hasta'])} "
              f"(revisar a mano, no se borra solo)")

    if nuevos:
        print("\nATENCION: los tramos nuevos entran sin los dias oficiales de ARCA "
              "(campo \"dias\"). La app los va a contar sola por meses de 30 dias y va a "
              "avisar en pantalla cuando le toque usar uno completo. Para dejarlo exacto, "
              "pedile a ARCA un detalle de calculo que atraviese el tramo y copia el numero.")

    if args.escribir:
        fusionados = fusionar(guardados, oficiales)
        huecos = verificar_continuidad(fusionados)
        if huecos:
            print(f"\nNO SE ESCRIBIO: la tabla bajada tiene huecos de continuidad: {huecos}")
            return 2
        doc["tramos"] = fusionados
        doc["verificado"] = datetime.date.today().isoformat()
        _guardar(args.archivo, doc)
        print(f"\ntasas.json actualizado ({len(fusionados)} tramos).")

    return 1


def _fmt(iso):
    return datetime.date.fromisoformat(iso).strftime("%d/%m/%Y")


def _guardar(ruta, doc):
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=2)
        fh.write("\n")


if __name__ == "__main__":
    sys.exit(main())
