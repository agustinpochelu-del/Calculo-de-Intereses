# CLAUDE.md

Liquidador de intereses de ARCA para juicios de ejecución fiscal, del estudio
contable de Agustín Pochelú. Streamlit + pandas + openpyxl, desplegado en
Streamlit Community Cloud.

## Antes de tocar nada

**Leé `Memoria.md`** (privado, fuera de git). Son las correcciones y el contexto
que se fue acumulando: cómo se lee una boleta de deuda, qué campo es la fecha de
demanda, por qué el formato del VEP sale de un archivo y no de la guía. Cada
entrada dice *por qué* y *cómo se aplica*.

`RESUMEN_EJECUTIVO.md` tiene el detalle técnico: mapa de funciones, la convención
de días de ARCA, el formato del TXT de VEPs y el backlog.

## Memoria

**Cuando algo salga mal —un error propio, un mail que no se pudo leer, una
clasificación equivocada que Agustín corrigió, un formato que cambió— escribilo
en `Memoria.md` en el momento**, sin esperar a que te lo pida. Entrada arriba de
todo, con fecha, etiqueta (`corrección` / `decisión` / `negocio` / `técnico`) y
las líneas **Por qué** y **Cómo lo aplico**. El formato está explicado adentro.

Si no falló nada, no escribas: es un registro de aprendizajes, no un parte de
actividad. Si ya hay una entrada del mismo tema, actualizala en vez de duplicar.
Cuando una nota se vuelve una regla permanente, mudala a `RESUMEN_EJECUTIVO.md`.

La rutina programada `boletas-arca` también escribe ahí.

## Lo mínimo indispensable

- **Los números salen de código probado, nunca de una lectura tuya.** Ni un
  importe, ni una fecha, ni una clasificación se deducen "a ojo". Esto termina en
  un expediente judicial: un peso de más es un problema real.
- **Corré los tres tests antes de dar algo por terminado**, y no toques los
  archivos de referencia que fijan:

  ```bash
  python test_motor.py && python test_leer_mail.py && python test_veps.py
  ```

  `test_motor.py` compara contra liquidaciones reales de ARCA al centavo.
  `test_veps.py` reproduce un archivo TXT que ARCA aceptó, carácter por carácter.
  Si alguno falla, algo se rompió: no se ajusta la referencia para que pase.

- **Este repositorio es público.** No van nombres de contribuyentes, CUITs
  reales, mails ni planillas de clientes — ni siquiera como ejemplo en un
  marcador de campo. `.gitignore` bloquea `*.eml`, `*.msg` y `*.xlsx`; el resto
  es cuestión de no escribirlo.
- **No pushees sin que Agustín lo pida.** Editá y commiteá; el push lo decide él.
- **Escribí en rioplatense con voseo**, sobrio, sin marketing.

## Cuando algo no se puede resolver, se marca

El patrón del proyecto, en las tres pantallas: lo que no se puede determinar con
certeza queda en «revisar», con el motivo escrito, y **no entra en la salida**.
Nunca se completa con una suposición. Vale para las filas de una boleta, para las
combinaciones que faltan en la tabla de conceptos y para los remitentes
desconocidos.

## Datos que se editan a mano

- `tasas.json` — las tasas. Se controlan solas una vez por semana contra la
  página de ARCA (`.github/workflows/control-tasas.yml`).
- `conceptos ARCA VEPS.xlsx` — impuestos, conceptos y subconceptos con su
  formulario y código de pago. **Está fuera de git.** Después de editarla:

  ```bash
  python actualizar_conceptos.py
  ```

  Valida y escribe `conceptos_veps.json`. Si algo no cierra, no escribe nada.

## Correr la app

```bash
streamlit run app.py
```
