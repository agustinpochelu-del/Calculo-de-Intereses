# Liquidador ARCA — Resumen Ejecutivo del Desarrollo

## 1. Qué es

App en **Streamlit** (`app.py`, un solo archivo, deployada vía GitHub) que calcula liquidaciones
judiciales de deuda impositiva ante ARCA (ex-AFIP) en juicios de ejecución fiscal, incluyendo
intereses resarcitorios, punitorios y capitalizables.

El objetivo del cálculo no es aproximarse: la liquidación se presenta en un expediente, así que
tiene que dar **el mismo importe al centavo** que el detalle de cálculo que emite ARCA.

## 2. Estado actual

La app tiene **dos lengüetas independientes** (`st.tabs`), cada una con su propio uploader de Excel
y su propio motor de procesamiento, porque parten de datos de origen distinto:

| Lengüeta | Función principal | Cuándo se usa |
|---|---|---|
| 💰 Juicio por Capital + Intereses | `procesar_juicio_capital()` | El capital impositivo original está impago (o se pagó en algún momento del juicio) |
| 📈 Juicio a los Intereses | `procesar_juicio_intereses()` | El capital YA fue pagado; se demanda por los intereses resarcitorios impagos, que pasan a ser la nueva "base" de la deuda |

Cada lengüeta tiene, además, un botón para **descargar una plantilla Excel en blanco** con el
formato exacto esperado (`generar_plantilla_capital()` / `generar_plantilla_intereses()`), con
instrucciones, fila de ejemplo y las tasas vigentes precargadas.

## 3. Mapa de funciones en `app.py`

```
formato_arg()                     # formatea números en pesos arg ($1.234,56)

dias_arca()                       # convención de días de ARCA (ver punto 4)
calcular_interes()                # MOTOR NUMÉRICO. Recibe (inicio, fin, capital, tabla_tasas)
                                  # y devuelve (interés, días) por tramos de tasa.
_normalizar_tabla_tasas()         # deriva la tasa diaria de la mensual, ordena tramos,
                                  # y acepta el formato viejo con Tasa_Diaria truncada.
cargar_tasas()                    # lee hoja Tasas/Tasas Punitorios del Excel subido;
                                  # si no existe o está vacía, cae a TASAS_*_DEFAULT.
validar_deudas()                  # chequea la hoja Deudas y devuelve problemas concretos
verificar_tabla_tasas()           # avisa huecos entre tramos y tramos sin días oficiales

mostrar_tabla_tasas()             # UI: cuadro de tasas de referencia al pie de cada resultado
boton_descarga()                  # UI: arma el Excel liquidado (openpyxl) con fila de totales

procesar_juicio_capital()         # Lengüeta 1: lee Excel, corre el motor Caso A/B (ver punto 5),
                                  # arma el dashboard, tabla detalle y descarga.
procesar_juicio_intereses()       # Lengüeta 2: motor simple de 2 tramos (Resarcitorios + Punitorios)
                                  # aplicado sobre la columna Capital (que ya es un monto de intereses).

_estilo_header(), _fila_ejemplo(),
_ajustar_anchos(), _nota(),
_hoja_tasas(), _hoja_instrucciones()  # helpers de openpyxl para armar las plantillas descargables

generar_plantilla_capital()       # arma el .xlsx en blanco para la Lengüeta 1
generar_plantilla_intereses()     # arma el .xlsx en blanco para la Lengüeta 2

TASAS_RESARCITORIAS_DEFAULT       # constantes con los tramos de tasa "conocidos" (ver punto 7)
TASAS_PUNITORIAS_DEFAULT
```

## 4. Cómo cuenta los días ARCA (la parte más delicada)

Reconstruido a partir de cuatro liquidaciones reales emitidas por el sistema de ARCA, incluidos
los detalles de cálculo tramo por tramo. Son cuatro reglas y todas importan:

**a) El interés devenga desde el día SIGUIENTE a la fecha de origen**, hasta la fecha de corte
inclusive. Un vencimiento del 13/05/2024 abre el primer tramo el 14/05/2024; los punitorios de una
demanda del 13/03/2026 arrancan el 14/03/2026.

**b) Cada mes completo cuenta 30 días, y solo el resto se cuenta por días corridos**
(función `dias_arca`). Es la diferencia que arrastraba la versión anterior, que contaba días
calendario:

| venc. 13/02/2026 → demanda 10/06/2026 | días |
|---|---|
| días calendario (15 de febrero + 31 + 30 + 31 + 10) | 117 |
| días ARCA (3 meses × 30 + 28 días de 13/05 a 10/06) | **118** |

Febrero corto suma 30 igual, así que el error del conteo calendario no es parejo: en unos meses da
de menos y en otros de más.

**c) Los tramos que el devengamiento cubre ENTEROS usan los días oficiales de ARCA**, que no
siempre salen de la regla (b). Están en la columna `Dias` de la hoja de tasas y mandan sobre
cualquier fórmula:

| tramo | días ARCA | lo que daría la regla (b) |
|---|---|---|
| 01/06/2024 – 31/07/2024 | 60 | 60 ✓ |
| 01/12/2024 – 31/01/2025 | **61** | 60 ✗ |
| 01/02/2025 – 28/02/2025 | **28** | 30 ✗ |
| 01/03/2025 – 30/06/2025 | **122** | 120 ✗ |

En los tramos de las puntas, que se cubren a medias, no aplican: ahí se cuenta con `dias_arca`.

**d) El interés se redondea por tramo antes de sumarse.** En el detalle de ARCA los importes de
cada tramo suman exacto el total informado.

Además, el `Hasta` de cada tramo es **inclusivo**: el `Hasta` + 1 día tiene que dar exactamente el
`Desde` del tramo siguiente. Si queda un hueco se pierde un día de interés en cada cambio de tasa,
sin que nadie lo note. `verificar_tabla_tasas()` avisa si eso pasa.

### Validación

`test_motor.py` (se corre con `python test_motor.py`) fija como regresión las cuatro liquidaciones
reales de ARCA. **Las cuatro cierran al centavo**, tramo por tramo, incluidos los totales de días
que informa ARCA (812, 407, 210).

## 5. Reglas de negocio del motor (Lengüeta 1)

Cada obligación tiene 4 fechas: `Vencimiento` (V), `F. Pago Capital` (P, puede estar vacía),
`fecha_Demanda` (D), `Fecha_Liquidacion` (L). Los tramos son continuos: el fin de uno es el inicio
del próximo.

El orden de Punitorios y Capitalizables **se invierte** según cuál evento ocurre primero:

- **Caso A — Demanda antes del Pago del Capital** (`D <= P`):
  `Resarcitorios: V→D (sobre Capital)` → `Punitorios: D→P (sobre Capital)` → `Capitalizables: P→L (sobre monto de Resarcitorios)`
- **Caso B — Pago del Capital antes de la Demanda** (`P < D`):
  `Resarcitorios: V→P (sobre Capital)` → `Capitalizables: P→D (sobre monto de Resarcitorios)` → `Punitorios: D→L (sobre Capital)`
- **Sin `F. Pago Capital` cargada**: `Resarcitorios: V→D`, `Punitorios: D→L`, sin Capitalizables.

El Caso A quedó confirmado contra un detalle de ARCA con capitalizables (venc 13/01/2026, demanda
13/03/2026, pago capital 13/05/2026, pago intereses 12/08/2026).

La Lengüeta 2 usa siempre el caso "sin pago": `Resarcitorios: V→D`, `Punitorios: D→L`, sobre la
columna `Capital` (que en ese formato ya es un monto de intereses).

## 6. Formato de datos esperado

**Hoja `Deudas`** — Lengüeta 1: `Impuesto | concepto | Periodo | Vencimiento | Capital | F. Pago Capital | fecha_Demanda | Fecha_Liquidacion`
(F. Pago Capital puede dejarse vacía).

**Hoja `Deudas`** — Lengüeta 2: igual pero **sin** `F. Pago Capital`.

**Hojas `Tasas` y `Tasas Punitorios`** (mismo formato en ambas lengüetas):
`Desde | Hasta | Tasa_Mensual | Dias`.

- `Tasa_Mensual` es el porcentaje mensual tal como lo publica ARCA (ej. `2.75`). La tasa diaria se
  deriva dividiendo por 30 en el momento del cálculo, **sin truncar**.
- `Dias` son los días oficiales del tramo completo (punto 4c). Se deja vacío en el tramo abierto.

Los archivos viejos con `Tasa_Diaria` se siguen leyendo: la app reconstruye la mensual redondeando
a dos decimales, que es como ARCA la publica (`0,000917 × 30 = 2,751 → 2,75`), y vuelve a derivar
la diaria sin truncar.

**Salida**: el Excel descargado incluye columnas `Dias_Resarcitorios`, `Dias_Capitalizables` y
`Dias_Punitorios` para poder auditar de dónde sale cada importe, más una fila de totales con
importes (no fórmulas, para que se lea igual en cualquier visor).

## 7. Tasas por defecto ("memoria" de la app)

Como los tramos de tasa cambian poco, están **hardcodeadas** en `app.py` (constantes
`TASAS_RESARCITORIAS_DEFAULT` / `TASAS_PUNITORIAS_DEFAULT`, arriba del todo). Si el Excel subido no
trae hoja de tasas (o la trae vacía), `cargar_tasas()` cae automáticamente a estos valores y la UI
muestra un aviso. El último tramo de cada tabla queda con `Hasta` abierto (2029-12-31 /
2050-12-31) representando "vigente hasta nuevo aviso".

**Para actualizar cuando salga una tasa nueva**: agregar una fila a la lista correspondiente y
cerrar el `Hasta` del tramo anterior con el día en que dejó de regir, anotándole los `Dias` que
informe el detalle de cálculo de ARCA. Mismo criterio para las plantillas descargables, que
precargan estas mismas tablas.

## 8. Deployment

- Repo en GitHub, deployado en Streamlit (Community Cloud u otro — confirmar con Agustín).
- Dependencias: `streamlit`, `pandas`, `openpyxl`.
- Sin base de datos ni almacenamiento persistente: todo el estado vive en el Excel que sube el
  usuario en cada corrida, más las constantes de tasas embebidas en el código.
- Sin autenticación conocida en esta app (confirmar si hace falta agregar).

## 9. Pendientes / backlog

1. **Confirmar el capital de la liquidación de anticipos de Ganancias.** ARCA liquidó sobre
   $366.033,27 y el Excel de entrada dice $336.033,27 — diferencia exacta de $30.000, con los
   dígitos 3 y 6 dados vuelta. Es un dato de origen, no un problema de cálculo.
2. Evaluar si conviene mover `TASAS_*_DEFAULT` a un archivo de config aparte (JSON/YAML) en vez de
   constantes en `app.py`, para que actualizar una tasa no requiera tocar el código Python.
3. Sumar al `test_motor.py` los casos nuevos que vayan apareciendo, en especial los que crucen
   cambios de tasa (son los que más ejercitan el motor).

### Resuelto

- ~~Diferencia residual en los Resarcitorios contra la calculadora oficial de ARCA~~ → era el
  conteo de días (punto 4), no un tramo de tasa faltante.
- ~~Inconsistencia del campo `Dias` (207 vs 208)~~ → el campo son los días oficiales del tramo
  completo; en los tramos abiertos no aplica y va vacío.
- ~~Validación de datos al subir el Excel~~ → `validar_deudas()` y `verificar_tabla_tasas()`.
- ~~Tests automatizados de regresión~~ → `test_motor.py`.
- ~~Precisión de las tasas~~ → se guardan mensuales y la diaria se deriva sin truncar.

## 10. Archivos de referencia (casos reales usados para validar)

- Cuatro detalles de cálculo emitidos por ARCA el 07/08/2026, fijados en `test_motor.py`:
  - venc 13/05/2024 → demanda 13/03/2026, capital $156.366,00 (cruza los 8 tramos de tasa)
  - venc 26/06/2025 → demanda 13/03/2026, capital $2.452.500,32 (arranca sobre el final de un tramo)
  - venc 13/01/2026, con intereses capitalizables, capital $973.922,83
  - 8 anticipos de Ganancias, demanda 10/06/2026, capital $366.033,27 c/u
- `Luis_Rojas.xlsx` — Lengüeta 1, caso con `F. Pago Capital` posterior a la demanda (Caso A).
- `Int_Lima_Sur.xlsx` — Lengüeta 2, formato "Juicio a los Intereses".
