"""Tests de la puerta de acceso.

Se corren con:  python test_acceso.py

Lo que se verifica acá no es estético: si esto falla, la app queda abierta.
"""
import sys
import time


class _Any:
    def __call__(self, *a, **k): return _Any()
    def __getattr__(self, name): return _Any()
    def __enter__(self): return self
    def __exit__(self, *a): return False


class _St(_Any):
    """Stub de streamlit con secrets y session_state de verdad, que es lo que usa acceso.py."""
    def __init__(self, usuarios=None):
        self.secrets = {'usuarios': usuarios or {}}
        self.session_state = {}

    def columns(self, spec, **k):
        return [_Any() for _ in range(spec if isinstance(spec, int) else len(spec))]


USUARIOS = {}
st = _St()
sys.modules['streamlit'] = st
import acceso  # noqa: E402

fallos = []


def check(nombre, ok):
    if not ok:
        fallos.append(nombre)
    print(f"  {'ok   ' if ok else 'FALLA'} {nombre}")


# =====================================================================================
print("\n1) El hash es reproducible y depende de la sal")
SAL_A, SAL_B = 'a' * 32, 'b' * 32
h1 = acceso.hashear('contraseña larga de prueba', SAL_A)
h2 = acceso.hashear('contraseña larga de prueba', SAL_A)
h3 = acceso.hashear('contraseña larga de prueba', SAL_B)
check("la misma clave y la misma sal dan el mismo hash", h1 == h2)
check("la misma clave con otra sal da otro hash", h1 != h3)
check("el hash no contiene la contraseña", 'contraseña' not in h1)
check("el hash mide 64 caracteres hexadecimales", len(h1) == 64 and all(c in '0123456789abcdef' for c in h1))

# =====================================================================================
print("\n2) La verificación acepta la clave correcta y rechaza el resto")
st.secrets = {'usuarios': {
    'tester': {'nombre': 'Tester', 'sal': SAL_A, 'hash': acceso.hashear('clave-de-prueba-123', SAL_A)},
}}
check("entra con la clave correcta", acceso._verificar('tester', 'clave-de-prueba-123'))
check("no entra con la clave equivocada", not acceso._verificar('tester', 'clave-de-prueba-124'))
check("no entra con la clave vacía", not acceso._verificar('tester', ''))
check("no entra un usuario que no existe", not acceso._verificar('otro', 'clave-de-prueba-123'))
check("el usuario no distingue mayúsculas", acceso._verificar('  TESTER ', 'clave-de-prueba-123'))
check("la clave sí distingue mayúsculas", not acceso._verificar('tester', 'CLAVE-DE-PRUEBA-123'))

# =====================================================================================
print("\n3) Sin usuarios configurados no se puede entrar")
st.secrets = {'usuarios': {}}
check("no entra nadie si no hay usuarios cargados", not acceso._verificar('tester', 'clave-de-prueba-123'))
st.secrets = {}
check("no revienta si no hay secrets", acceso._usuarios() == {})

# =====================================================================================
print("\n4) Un usuario inexistente tarda lo mismo que uno real")
# Si respondiera al toque, se podría averiguar quién tiene cuenta probando usuarios.
st.secrets = {'usuarios': {
    'tester': {'nombre': 'Tester', 'sal': SAL_A, 'hash': acceso.hashear('clave-de-prueba-123', SAL_A)},
}}


def medir(usuario, clave, vueltas=3):
    t = time.perf_counter()
    for _ in range(vueltas):
        acceso._verificar(usuario, clave)
    return (time.perf_counter() - t) / vueltas


real = medir('tester', 'clave-mala')
inexistente = medir('fantasma', 'clave-mala')
proporcion = inexistente / real if real else 0
check(f"tiempos comparables (proporción {proporcion:.2f}, esperado entre 0,5 y 2)",
      0.5 <= proporcion <= 2.0)

# =====================================================================================
print("\n5) El bloqueo por intentos fallidos se activa y se libera")
st.session_state = {'acceso_intentos': acceso._MAX_INTENTOS, 'acceso_ultimo_intento': time.time()}
check("bloquea al llegar al máximo de intentos", acceso._bloqueado_hasta() > 0)
st.session_state = {'acceso_intentos': acceso._MAX_INTENTOS,
                    'acceso_ultimo_intento': time.time() - acceso._ESPERA_SEGUNDOS - 1}
check("libera cuando pasó la espera", acceso._bloqueado_hasta() == 0)
st.session_state = {'acceso_intentos': acceso._MAX_INTENTOS - 1, 'acceso_ultimo_intento': time.time()}
check("no bloquea antes de llegar al máximo", acceso._bloqueado_hasta() == 0)

# =====================================================================================
print("\n" + "=" * 60)
if fallos:
    print(f"{len(fallos)} FALLA(S):")
    for f in fallos:
        print("  -", f)
    sys.exit(1)
print("Todos los tests pasaron.")
