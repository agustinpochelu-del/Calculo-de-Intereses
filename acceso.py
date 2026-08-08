"""Puerta de acceso del Liquidador: usuario y contraseña, solo personal del estudio.

La puerta está acá y no en la página del estudio a propósito. La app corre en otro
servidor: quien tenga la URL entra, haya pasado o no por la web. Un link escondido
detrás de una pantalla de acceso no es control de acceso.

Las credenciales NO están en el repositorio. Viven en los secrets de Streamlit, y de
la contraseña solo se guarda un hash con scrypt (no la contraseña). El formato es:

    [usuarios.agustin]
    nombre = "Agustín"
    sal = "a1b2c3..."
    hash = "d4e5f6..."

Para generar ese bloque: `python generar_clave.py`. La contraseña no se escribe en
ningún archivo, ni en el repositorio ni en el disco: se pega directo en el panel de
secrets de Streamlit.
"""
import hashlib
import hmac
import time

import streamlit as st

# Parámetros de scrypt. n alto encarece cada intento a propósito: es lo que hace lento
# probar contraseñas al voleo. Si se cambian, hay que regenerar todos los hashes.
_N, _R, _P, _LARGO = 2 ** 14, 8, 1, 32

_MAX_INTENTOS = 5
_ESPERA_SEGUNDOS = 300  # 5 minutos de bloqueo al pasarse de intentos


def hashear(clave, sal_hex):
    """Devuelve el hash hexadecimal de la contraseña para la sal dada."""
    return hashlib.scrypt(
        clave.encode('utf-8'), salt=bytes.fromhex(sal_hex),
        n=_N, r=_R, p=_P, dklen=_LARGO,
    ).hex()


def _usuarios():
    """Lee los usuarios de los secrets. Devuelve {} si no hay nada configurado."""
    try:
        return dict(st.secrets.get('usuarios', {}))
    except Exception:
        # Sin archivo de secrets (por ejemplo, corriendo local sin configurar nada).
        return {}


def _verificar(usuario, clave):
    """True si el usuario existe y la contraseña coincide.

    Se compara con compare_digest para que el tiempo de respuesta no delate cuántos
    caracteres del hash coincidían.
    """
    registro = _usuarios().get((usuario or '').strip().lower())
    if not registro:
        # Se hashea igual contra una sal descartable para que un usuario inexistente
        # tarde lo mismo que uno real y no se pueda averiguar quién tiene cuenta.
        hashear(clave or '', '00' * 16)
        return False
    try:
        esperado = registro['hash']
        obtenido = hashear(clave or '', registro['sal'])
    except (KeyError, ValueError):
        return False
    return hmac.compare_digest(esperado, obtenido)


def _bloqueado_hasta():
    intentos = st.session_state.get('acceso_intentos', 0)
    ultimo = st.session_state.get('acceso_ultimo_intento', 0)
    if intentos >= _MAX_INTENTOS:
        restante = _ESPERA_SEGUNDOS - (time.time() - ultimo)
        if restante > 0:
            return restante
        st.session_state['acceso_intentos'] = 0
    return 0


def _pantalla_de_acceso():
    _, centro, _ = st.columns([1, 1.6, 1])
    with centro:
        with st.container(key='lq-acceso'):
            st.markdown('### Acceso del estudio')
            st.caption('Esta herramienta es de uso interno. Si no tenés usuario, pedíselo a Agustín.')

            espera = _bloqueado_hasta()
            if espera:
                st.error(f"❌ Demasiados intentos fallidos. Probá de nuevo en "
                         f"{int(espera // 60) + 1} minuto(s).")
                return

            with st.form('form_acceso'):
                usuario = st.text_input('Usuario', autocomplete='username')
                clave = st.text_input('Contraseña', type='password', autocomplete='current-password')
                entrar = st.form_submit_button('Ingresar', use_container_width=True)

            if entrar:
                if _verificar(usuario, clave):
                    st.session_state['acceso_ok'] = True
                    st.session_state['acceso_usuario'] = usuario.strip().lower()
                    st.session_state['acceso_intentos'] = 0
                    st.rerun()
                else:
                    st.session_state['acceso_intentos'] = st.session_state.get('acceso_intentos', 0) + 1
                    st.session_state['acceso_ultimo_intento'] = time.time()
                    restantes = _MAX_INTENTOS - st.session_state['acceso_intentos']
                    if restantes > 0:
                        st.error(f"❌ Usuario o contraseña incorrectos. "
                                 f"Te quedan {restantes} intento(s).")
                    else:
                        st.error("❌ Demasiados intentos fallidos. Esperá unos minutos.")


def requiere_acceso():
    """Corta la ejecución si no hay sesión iniciada. Se llama antes de mostrar la app."""
    if st.session_state.get('acceso_ok'):
        return

    if not _usuarios():
        st.error(
            "❌ No hay usuarios configurados. Cargá al menos uno en los secrets de "
            "Streamlit antes de publicar la app: si esta pantalla queda así, cualquiera "
            "con la URL podría entrar."
        )
        st.stop()

    _pantalla_de_acceso()
    st.stop()


def nombre_de_quien_entro():
    usuario = st.session_state.get('acceso_usuario', '')
    registro = _usuarios().get(usuario, {})
    return registro.get('nombre', usuario)


def boton_salir():
    if st.button('Salir', key='lq_salir'):
        for k in ('acceso_ok', 'acceso_usuario', 'acceso_intentos', 'acceso_ultimo_intento'):
            st.session_state.pop(k, None)
        st.rerun()
