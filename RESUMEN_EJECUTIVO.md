# Liquidador ARCA — Resumen Ejecutivo del Desarrollo

## 1. Qué es

App en **Streamlit** (`app.py`, un solo archivo, deployada vía GitHub) que calcula liquidaciones
judiciales de deuda impositiva ante ARCA (ex-AFIP) en juicios de ejecución fiscal, incluyendo
intereses resarcitorios, punitorios y capitalizables.

El objetivo del cálculo no es aproximarse: la liquidación se presenta en un expediente, así que
tiene que dar **el mismo importe al centavo** que el detalle de cálculo que emite ARCA.

## 2. Estado actual

La app tiene **tres lengüetas** (`st.tabs`). Las dos de liquidación son independientes entre sí,
cada una con su propio uploader de Excel y su propio motor, porque parten de datos de origen
distinto. La primera no liquida: prepara.

| Lengüeta | Función principal | Cuándo se usa |
|---|---|---|
| 📧 Importar desde el mail | `procesar_mail()` | Llega el mail del agente fiscal con la boleta de deuda y hay que pasarlo a planilla |
| 💰 Juicio por Capital + Intereses | `procesar_juicio_capital()` | El capital impositivo original está impago (o se pagó en algún momento del juicio) |
| 📈 Juicio a los Intereses | `procesar_juicio_intereses()` | El capital YA fue pagado; se demanda por los intereses resarcitorios impagos, que pasan a ser la nueva "base" de la deuda |

Las dos de liquidación tienen, además, un botón para **descargar una plantilla Excel en blanco**
con el formato exacto esperado (`generar_plantilla_capital()` / `generar_plantilla_intereses()`),
con instrucciones, fila de ejemplo y las tasas vigentes precargadas.

### El circuito completo

```
mail del agente fiscal (.eml)
        │
        ▼  lengüeta 📧 — leer_mail.py
   tabla revisable  ──► control: la suma de las filas contra el "Monto Demanda" de la boleta
        │
        ▼  se bajan una o dos planillas .xlsx
        │
        ▼  lengüetas 💰 y 📈 (el camino de siempre, ya validado contra ARCA)
   liquidación
```

El mail **no se liquida directo, a propósito**. Convertirlo en planilla deja un archivo revisable
en el medio y hace que la importación termine usando el mismo motor de cálculo ya probado, sin
abrir un segundo camino que habría que validar aparte.

## 3. Mapa de funciones en `app.py`

```
formato_arg()                     # formatea números en pesos arg ($1.234,56)

dias_arca()                       # convención de días de ARCA (ver punto 4)
calcular_interes()                # MOTOR NUMÉRICO. Recibe (inicio, fin, capital, tabla_tasas)
                                  # y devuelve (interés, días) por tramos de tasa.
cargar_tasas()                    # lee tasas.json, arma las dos tablas y verifica que
                                  # los tramos empalmen (ver punto 7)
validar_deudas()                  # chequea la hoja Deudas y devuelve problemas concretos
avisar_tramos_sin_dias()          # avisa si hubo que usar un tramo completo sin los
                                  # días oficiales de ARCA

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

procesar_mail()                   # Lengüeta 0: lee el .eml, muestra la carátula de la boleta,
                                  # contrasta la suma contra el Monto Demanda, deja revisar
                                  # fila por fila y baja las planillas ya cargadas.
armar_planilla()                  # arma el .xlsx con la hoja Deudas a partir de lo revisado
```

Archivos que acompañan a `app.py`:

```
tasas.json                        # las tasas. UNICO lugar donde se actualizan (punto 7)
actualizar_tasas.py               # compara tasas.json contra la página de ARCA
leer_mail.py                      # lector de los mails de boleta de deuda (punto 6 bis)
test_motor.py                     # regresión contra las liquidaciones reales de ARCA
test_leer_mail.py                 # regresión del lector, con un mail inventado
.github/workflows/control-tasas.yml   # corre el control una vez por semana
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

**No hay que cargar tasas en el Excel.** La app las tiene (punto 7). Si el archivo trae
hojas `Tasas` de una plantilla vieja, se ignoran y la app lo avisa en pantalla: una hoja
desactualizada metida en el medio cambiaría un importe que va a un expediente sin que
nadie lo note.

**Salida**: el Excel descargado incluye columnas `Dias_Resarcitorios`, `Dias_Capitalizables` y
`Dias_Punitorios` para poder auditar de dónde sale cada importe, más una fila de totales con
importes (no fórmulas, para que se lea igual en cualquier visor).

## 6 bis. El lector de mails: `leer_mail.py`

Los agentes fiscales pegan en el mail la pantalla **"Datos de la Boleta de Deuda"** del sistema
de ARCA. Ese pegado viaja como tabla HTML, y de ahí sale casi toda la planilla:

| Del mail | A la planilla |
|---|---|
| Impuestos - Conceptos - Subconceptos | `Impuesto`, `concepto` |
| Período (+ Cuota, si es un anticipo) | `Periodo` (`2026-1`, `2026-2`…) |
| Vencimiento | `Vencimiento` |
| Monto de la Deuda | `Capital` |
| Pagos de Capital Registrados | `F. Pago Capital` |
| **Fecha Sorteo** | `fecha_Demanda` |

`Fecha_Liquidacion` no sale del mail: la elige Agustín según cuándo presenta.

**La fecha de demanda es la Fecha Sorteo, no la "Fecha de Inicio".** Están a un día de
distancia y es fácil confundirlas. Se verificó contra dos boletas reales cuyas liquidaciones
ya estaban hechas: en las dos, la fecha usada coincide con la de sorteo.

### Qué se calcula sobre cada fila: manda el subconcepto

Al lado del importe, el agente fiscal escribe una nota a mano ("DEBE", "pto. DJ debe
intereses", "PRESENTO DDJJ - DEBE INT. RESARCITORIOS + PUNITORIOS + CAPITALIZADOS"…). **Esa
nota no está normalizada y no se usa para decidir.** Lo que decide es el subconcepto de ARCA,
porque dice qué *es* el importe:

| Subconcepto | El importe es | Va a |
|---|---|---|
| SALDO DE DECLARACIÓN JURADA, ANTICIPOS | capital impositivo | 💰 Capital + Intereses |
| INTERESES RESARCITORIOS | interés ya devengado | 📈 Juicio a los Intereses |
| cualquier otro, o nota con plan de pagos | — | queda en **revisar** |

La nota igual se lee y se muestra, y sirve de contraste: si dice que el capital está cancelado
pero la boleta no registra el pago, la fila se marca. Sin esa fecha los punitorios corren hasta
la liquidación en vez de hasta el pago, y el importe sale de más.

### Controles

- **La suma de las filas contra el `Monto Demanda` de la boleta.** Es el mejor control que hay,
  porque los dos números vienen del mismo mail. Si no coinciden, algo se leyó mal o quedó afuera,
  y se avisa en rojo.
- Las filas en `revisar` **no salen en ninguna planilla** hasta que se clasifiquen a mano.
- El remitente se contrasta contra una lista de agentes fiscales conocidos, si se pasa una. No
  bloquea: el mail puede venir reenviado. Solo avisa.

### Limitaciones conocidas

- Lee **`.eml`** (Gmail: *⋮ → Descargar mensaje*). No lee `.msg` de Outlook, que es un formato
  binario distinto y necesitaría una dependencia nueva.
- Necesita la versión con formato del mail. Si el agente fiscal manda una **captura de pantalla**
  en vez de pegar la tabla, no hay nada que leer y la app lo dice.
- Los acentos rotos de algunos reenvíos ("Perï¿½odo") están contemplados: las comparaciones se
  hacen sobre el texto pelado a ASCII.

⚠️ **Los mails reales no van al repositorio**, que es público: llevan CUIT, domicilio y deuda de
contribuyentes. `.gitignore` bloquea `*.eml` y `*.msg`. `test_leer_mail.py` usa un mail inventado.

## 6 ter. Generación de VEPs: `veps.py`

Pagar una liquidación a mano significa cargar un VEP por cada importe: ocho vencimientos
con capital, resarcitorios, capitalizables y punitorios son **treinta y dos**. Con un
archivo se cargan todos juntos.

### El formato

Dos tipos de registro. El encabezado, una sola línea de 33 caracteres:

```
01 · CUIT del generador (11) · 20001 · 00100 · 003 · 003 · cantidad+1 (4)
```

Los últimos cuatro dígitos son la cantidad de VEPs **más uno**: el total de líneas del
archivo contando el encabezado.

Y una línea por VEP:

```
02 <VEP fechaExpiracion="AAAA-MM-DD" nroFormulario="…" codTipoPago="…"
        contribuyenteCUIT="…" concepto="…" subConcepto="…" periodoFiscal="AAAAMM"
        importe="…" [anticipoCuota="…"]><Obligacion impuesto="…" importe="…"/></VEP>
```

✅ **Probado contra ARCA el 09/08/2026**: se subió un archivo generado por `veps.py` con
importes inventados y los quince VEPs salieron bien. Cubría Ganancias declaración jurada
con los cuatro subconceptos, anticipos con cuotas 1 y 2, Bienes Personales, IVA y aportes
de seguridad social —incluido el caso de formularios distintos en la misma boleta (1931
para el capital de seguridad social, 800 para sus intereses).

⚠️ **El formato sale de ese archivo, no de la guía publicada.** La guía
(`VerGuia.aspx?id=365`) usa `precio` en vez de `importe` y pone espacios alrededor de
los `=`; es de 2010 y quedó vieja. `test_veps.py` fija el archivo real y lo reproduce
carácter por carácter: si hay que tocar `veps.py`, ese test falla, y el archivo de
referencia se cambia recién cuando haya uno nuevo que ARCA haya aceptado. No hay otra
forma de validarlo que subirlo.

Detalles que importan: los códigos van **sin ceros a la izquierda** (`concepto="19"`, no
`"019"`), y cuando no hay cuota el atributo `anticipoCuota` **no aparece**.

### De la liquidación a los VEPs

Cada fila da hasta cuatro VEPs, y lo que los distingue es el subconcepto:

| Columna de la liquidación | subConcepto |
|---|---|
| `Capital` | el mismo número que el concepto (19 DDJJ / 191 Anticipo) |
| `Interes_Resarcitorio` | 51 |
| `Interes_Capitalizable` | 52 |
| `Interes_Punitorio` | 94 |

Los anticipos llevan el período con el **mes en 00** (`202600`) y el número de cuota en
`anticipoCuota`. El resto lleva su mes.

**Ganancias depende del CUIT**: 10 para sociedades, 11 para personas físicas. Se deduce
del prefijo — 20, 23, 24 y 27 son personas físicas; 30, 33 y 34, jurídicas. Cualquier
otro prefijo no se adivina: se marca para que lo elija Agustín.

`conceptos_veps.json` tiene las 45 combinaciones de impuesto / concepto / subconcepto con
su formulario y código de pago, salidas de la planilla del estudio. Lo que no está en esa
tabla **no sale en el archivo**: se marca con el motivo, igual que las filas en «revisar».

### En pantalla

Lengüeta 🧾: se sube una liquidación ya calculada y aparece un renglón por importe, con
una casilla **destildada**. Casi nunca se paga todo lo liquidado, así que nada se
selecciona solo. Abajo, la cantidad de VEPs y el total, y el archivo con el nombre que
espera ARCA (`F20001.cuit.<cuit>.fecha.<AAAAMMDD>.txt`).

Se sube en ARCA en **Presentación de DDJJ y Pagos → VEP → Generación masiva**.

## 7. Tasas: `tasas.json`

Las tasas viven en **`tasas.json`**, al lado de `app.py`. Es el único lugar donde se
actualizan, y replica la tabla "Evolución de Tasas de Intereses" que publica ARCA: una
fila por tramo, con la resarcitoria y la punitoria juntas, igual que en el original.
Están cargados los **42 tramos**, desde 1901 hasta el vigente.

```json
{ "desde": "2025-07-01", "hasta": "2999-12-31", "norma": "R (MEC) 823/2025",
  "resarcitoria_mensual": 2.75, "punitoria_mensual": 3.5, "dias": null }
```

- `hasta` es inclusivo y los tramos tienen que empalmar (`hasta` + 1 día = `desde` del
  siguiente). `cargar_tasas()` lo verifica al arrancar y corta con un error si no cierra.
- Las tasas son **mensuales en porcentaje**, como las publica ARCA; la diaria se deriva
  dividiendo por 30 sin truncar.
- `dias` son los días oficiales del tramo completo (punto 4c). ARCA **no los publica** en
  la tabla de tasas: salen del detalle de cálculo. Los tramos sin el dato quedan en `null`
  y la app cuenta sola, avisando en pantalla si le tocó usar uno completo sin él.
- El tramo vigente cierra en `2999-12-31`, tal como lo publica ARCA. Internamente se
  recorta a 2262 porque pandas no puede representar fechas más lejanas.

### Actualización automática

`actualizar_tasas.py` baja la tabla oficial y la compara con `tasas.json`. Solo usa la
librería estándar, así que corre en cualquier Python 3:

```
python actualizar_tasas.py              # solo informa
python actualizar_tasas.py --escribir   # además actualiza tasas.json
```

`.github/workflows/control-tasas.yml` lo corre **todos los lunes a las 9 de la mañana**
(también se puede disparar a mano desde la pestaña Actions). Si encuentra una tasa nueva:

1. deja el cambio hecho en una rama `tasas/AAAA-MM-DD`,
2. abre un issue con el link para crear el pull request y el detalle de lo que cambió.

**No mergea solo, a propósito**: la tasa entra en liquidaciones que van a un expediente,
así que la última palabra la tiene una persona. Si la página no responde o cambia de
formato, abre un issue avisando y la app sigue liquidando con las últimas tasas
verificadas.

Cuando entra un tramo nuevo llega sin el campo `dias`. Para dejarlo exacto hay que pedirle
a ARCA un detalle de cálculo que atraviese ese tramo y copiar el número.

## 8. Acceso

La app es de **uso interno del estudio** y no la controla el código: está marcada como
**privada en Streamlit Community Cloud**, que pide login con cuenta de Google antes de
servir nada. Quien no esté invitado no llega ni a ejecutar `app.py`.

- **Para dar de alta o de baja a alguien**: en Streamlit Cloud, en la app →
  Settings → Sharing.
- La app se entra por `https://liquidador-pochelu.streamlit.app`, enlazada desde el
  Área del Estudio de la web (`estudiopochelu.com/estudio`).

La puerta está en la app y **no** en la web del estudio, a propósito: la app corre en
otro servidor, así que quien tenga la URL entra haya pasado o no por la web. Un link
escondido detrás de una pantalla de acceso no es control de acceso.

⚠️ **Si alguna vez se pasa la app a pública, o se muda a un hosting sin login propio,
queda abierta a cualquiera con la URL.** Hubo una puerta propia dentro de la app
(`acceso.py`: usuario y contraseña, con hash scrypt y sal por usuario, más
`generar_clave.py` y `test_acceso.py`). Se sacó para no encadenar dos logins, ya que
Streamlit ya pedía uno. Está en el historial de git: si se muda la app, se recupera.

## 9. Deployment

- Repo en GitHub, deployado en **Streamlit Community Cloud**.
- Community Cloud no permite dominio propio: si se quiere entrar por
  `liquidador.estudiopochelu.com`, ese subdominio tiene que **redirigir** a la URL
  `.streamlit.app`. Para que el dominio propio quede en la barra de direcciones habría
  que mudar la app a otro hosting.
- Dependencias: `streamlit`, `pandas`, `openpyxl`. El control de tasas no suma ninguna
  (solo librería estándar).
- Sin base de datos ni almacenamiento persistente: todo el estado vive en el Excel que sube el
  usuario en cada corrida, más las constantes de tasas embebidas en el código.

## 10. Pendientes / backlog

1. **Confirmar el capital de la liquidación de anticipos de Ganancias.** ARCA liquidó sobre
   $366.033,27 y el Excel de entrada dice $336.033,27 — diferencia exacta de $30.000, con los
   dígitos 3 y 6 dados vuelta. Es un dato de origen, no un problema de cálculo.
2. Completar el campo `dias` de los tramos que todavía están en `null` (36 de 42). No hace falta
   hacerlo de golpe: la app avisa en pantalla cuando le toca usar uno, y ahí se pide el detalle
   de cálculo a ARCA. Los seis que ya están cargados cubren 2024-2025, que es lo que se usa hoy.
3. Sumar al `test_motor.py` los casos nuevos que vayan apareciendo, en especial los que crucen
   cambios de tasa (son los que más ejercitan el motor).
4. **Ampliar el vocabulario de subconceptos del lector de mails.** Hoy reconoce SALDO DE
   DECLARACIÓN JURADA, ANTICIPOS e INTERESES RESARCITORIOS, que es lo que apareció en los mails
   vistos hasta ahora. Cualquier otro cae en `revisar` — que es el comportamiento correcto, pero
   si empieza a repetirse alguno conviene agregarlo a `clasificar()`.
5. **Traer los mails solos.** Hoy hay que bajar el `.eml` y subirlo. Conectar la casilla para que
   los busque por remitente es una capa aparte, y recién tiene sentido ahora que el lector anda.
   Pide permisos sobre el correo, así que es una decisión, no solo trabajo.

### Resuelto

- ~~Diferencia residual en los Resarcitorios contra la calculadora oficial de ARCA~~ → era el
  conteo de días (punto 4), no un tramo de tasa faltante.
- ~~Inconsistencia del campo `Dias` (207 vs 208)~~ → el campo son los días oficiales del tramo
  completo; en los tramos abiertos no aplica y va vacío.
- ~~Validación de datos al subir el Excel~~ → `validar_deudas()` y `verificar_tabla_tasas()`.
- ~~Tests automatizados de regresión~~ → `test_motor.py`.
- ~~Precisión de las tasas~~ → se guardan mensuales y la diaria se deriva sin truncar.

## 11. Archivos de referencia (casos reales usados para validar)

- Cuatro detalles de cálculo emitidos por ARCA el 07/08/2026, fijados en `test_motor.py`:
  - venc 13/05/2024 → demanda 13/03/2026, capital $156.366,00 (cruza los 8 tramos de tasa)
  - venc 26/06/2025 → demanda 13/03/2026, capital $2.452.500,32 (arranca sobre el final de un tramo)
  - venc 13/01/2026, con intereses capitalizables, capital $973.922,83
  - 8 anticipos de Ganancias, demanda 10/06/2026, capital $366.033,27 c/u
- Dos planillas de clientes que quedan **solo en la carpeta local**, nunca en el
  repositorio: una de Lengüeta 1 con `F. Pago Capital` posterior a la demanda (Caso A) y
  otra de Lengüeta 2, en formato "Juicio a los Intereses".
- Un archivo de VEPs que ARCA aceptó, reproducido en `test_veps.py` con el CUIT cambiado.

⚠️ **Este repositorio es público.** No van nombres de contribuyentes, CUITs reales,
mails ni planillas de clientes. `.gitignore` bloquea `*.eml`, `*.msg` y `*.xlsx`; el
resto es cuestión de no escribirlo.
