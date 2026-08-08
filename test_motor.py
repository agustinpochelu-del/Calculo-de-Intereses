"""Tests de regresión del motor de cálculo.

Se corren con:  python test_motor.py

Los casos son liquidaciones reales emitidas por el sistema de ARCA. El motor tiene
que dar los mismos importes al centavo, tramo por tramo.
"""
import sys
import pandas as pd


# --- streamlit no hace falta para probar el motor: se reemplaza por un stub ---
class _Any:
    def __call__(self, *a, **k): return _Any()
    def __getattr__(self, name): return _Any()
    def __enter__(self): return self
    def __exit__(self, *a): return False


class _St(_Any):
    def columns(self, spec, **k):
        return [_Any() for _ in range(spec if isinstance(spec, int) else len(spec))]
    def tabs(self, labels): return [_Any() for _ in labels]
    def file_uploader(self, *a, **k): return None
    def cache_data(self, fn=None, **k):
        # @st.cache_data tiene que dejar la funcion tal cual para poder llamarla aca.
        return fn if fn is not None else (lambda f: f)


sys.modules['streamlit'] = _St()
import app  # noqa: E402

F = pd.Timestamp
RES, PUN, META = app.cargar_tasas()

fallos = []


def check(nombre, obtenido, esperado, tol=0.0):
    if isinstance(obtenido, (int, float)) and isinstance(esperado, (int, float)):
        ok = abs(obtenido - esperado) <= tol
    else:
        ok = obtenido == esperado
    if not ok:
        fallos.append(f"{nombre}: esperado {esperado}, obtenido {obtenido}")
    print(f"  {'ok   ' if ok else 'FALLA'} {nombre}: {obtenido} (esperado {esperado})")


# =====================================================================================
print("\n1) Conteo de días con meses de 30 (convención ARCA)")
# 3 meses completos (14/02 -> 14/05) + 28 días corridos (14/05 -> 11/06).
# Por días calendario darían 117: febrero corto suma 30 igual.
check("14/02/2026 -> 11/06/2026", app.dias_arca(F('2026-02-14'), F('2026-06-11')), 118)
check("14/10/2025 -> 11/06/2026", app.dias_arca(F('2025-10-14'), F('2026-06-11')), 238)
check("01/07/2025 -> 14/03/2026", app.dias_arca(F('2025-07-01'), F('2026-03-14')), 253)
check("14/03/2026 -> 13/08/2026", app.dias_arca(F('2026-03-14'), F('2026-08-13')), 150)
check("mismo día", app.dias_arca(F('2026-06-10'), F('2026-06-10')), 0)
check("fin anterior al inicio", app.dias_arca(F('2026-06-10'), F('2026-01-01')), 0)

# =====================================================================================
# Detalle de cálculo de ARCA del 07/08/2026, capital 156.366,00.
# Es el caso importante: el devengamiento cruza los ocho tramos de tasa, así que
# ejercita tanto los días oficiales de los tramos completos como el conteo por
# meses de 30 en los tramos de las puntas.
print("\n2) Detalle ARCA - venc 13/05/2024, demanda 13/03/2026, pago capital 12/08/2026")
CAPITAL = 156366.00
res, dias_res = app.calcular_interes(F('2024-05-13'), F('2026-03-13'), CAPITAL, RES)
pun, dias_pun = app.calcular_interes(F('2026-03-13'), F('2026-08-12'), CAPITAL, PUN)
check("total resarcitorios", res, 167507.60, tol=0.01)
check("total punitorios", pun, 27364.05, tol=0.01)
check("total de días", dias_res + dias_pun, 812)
check("total general de intereses", round(res + pun, 2), 194871.65, tol=0.01)

# =====================================================================================
print("\n3) Detalle ARCA - venc 26/06/2025 (arranca sobre el final de un tramo)")
CAPITAL = 2452500.32
res, dias_res = app.calcular_interes(F('2025-06-26'), F('2026-03-13'), CAPITAL, RES)
pun, dias_pun = app.calcular_interes(F('2026-03-13'), F('2026-08-12'), CAPITAL, PUN)
check("total resarcitorios", res, 581855.70, tol=0.01)
check("total punitorios", pun, 429187.56, tol=0.01)
check("total de días", dias_res + dias_pun, 407)
check("total general de intereses", round(res + pun, 2), 1011043.26, tol=0.01)

# =====================================================================================
# Este caso confirma además el orden de los tramos cuando la demanda es anterior al
# pago del capital: resarcitorios hasta la demanda, punitorios hasta el pago del
# capital, y capitalizables desde ahí sobre el monto de los resarcitorios.
print("\n4) Detalle ARCA - con intereses capitalizables")
CAPITAL = 973922.83
res, d1 = app.calcular_interes(F('2026-01-13'), F('2026-03-13'), CAPITAL, RES)
pun, d2 = app.calcular_interes(F('2026-03-13'), F('2026-05-13'), CAPITAL, PUN)
cap, d3 = app.calcular_interes(F('2026-05-13'), F('2026-08-12'), res, RES)
check("resarcitorios", res, 53565.76, tol=0.01)
check("punitorios", pun, 68174.60, tol=0.01)
check("capitalizables", cap, 4419.18, tol=0.01)
check("total de días", d1 + d2 + d3, 210)
check("total general de intereses", round(res + pun + cap, 2), 126159.54, tol=0.01)

# =====================================================================================
# Los 8 anticipos de Ganancias liquidados por ARCA (mismo capital, un vencimiento
# por mes). Fue el caso que destapó el problema del conteo de días.
print("\n5) Liquidación ARCA de los 8 anticipos de Ganancias")
CAPITAL = 366033.27
DEMANDA, PAGO, LIQUIDACION = F('2026-06-10'), F('2026-08-03'), F('2026-08-12')
ARCA = {  # vencimiento -> (resarcitorios, capitalizables)
    '2025-10-13': (79856.26, 658.81), '2025-11-13': (69790.34, 575.77),
    '2025-12-15': (59053.37, 487.19), '2026-01-13': (49658.51, 409.68),
    '2026-02-13': (39592.60, 326.64), '2026-03-13': (29526.68, 243.60),
    '2026-04-13': (19460.77, 160.55), '2026-05-13': (9394.85, 77.51),
}
for venc, (res_arca, cap_arca) in ARCA.items():
    res, _ = app.calcular_interes(F(venc), DEMANDA, CAPITAL, RES)
    cap, _ = app.calcular_interes(PAGO, LIQUIDACION, res, RES)
    check(f"resarcitorios venc {venc}", res, res_arca, tol=0.01)
    check(f"capitalizables venc {venc}", cap, cap_arca, tol=0.01)
pun, dias_pun = app.calcular_interes(DEMANDA, PAGO, CAPITAL, PUN)
check("punitorios", pun, 23060.10, tol=0.01)
check("días punitorios", dias_pun, 54)

# =====================================================================================
print("\n6) La tabla oficial de tasas.json está completa y sin huecos")
# cargar_tasas() revienta si los tramos no empalman, así que llegar hasta acá ya es
# parte de la prueba.
check("tramos cargados", len(RES), META['tramos'])
check("arranca en 1901", RES['Desde'].min(), F('1901-01-01'))
# Las dos tablas salen de las mismas filas de la tabla oficial: mismas fechas.
check("resarcitorias y punitorias comparten los tramos",
      int(bool((RES['Desde'].values == PUN['Desde'].values).all())), 1)
print(f"  ok    verificada contra ARCA el {META['verificado']}")

# =====================================================================================
print("\n7) La tasa se guarda mensual y se deriva sin truncar")
tasa = RES[RES['Desde'] == F('2025-07-01')].iloc[0]
check("resarcitoria mensual vigente", tasa['Tasa_Mensual'], 2.75)
check("resarcitoria diaria vigente", tasa['Tasa_Diaria'], 0.0275 / 30, tol=1e-15)
check("punitoria mensual vigente", PUN[PUN['Desde'] == F('2025-07-01')].iloc[0]['Tasa_Mensual'], 3.5)

# Antes la tabla de punitorios arrancaba en 2026: una demanda anterior daba punitorios
# en cero, sin avisar. Ahora la historia está completa.
jul24 = F('2024-07-01')
pun_2024 = PUN[(PUN['Desde'] <= jul24) & (PUN['Hasta'] >= jul24)].iloc[0]
check("punitoria de julio 2024 (antes daba cero)", pun_2024['Tasa_Mensual'], 7.39)

# =====================================================================================
print("\n" + "=" * 60)
if fallos:
    print(f"{len(fallos)} FALLA(S):")
    for f_ in fallos:
        print("  -", f_)
    sys.exit(1)
print("Todos los tests pasaron.")
