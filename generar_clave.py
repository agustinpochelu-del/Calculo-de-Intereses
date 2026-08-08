"""Genera el bloque de credenciales para pegar en los secrets de Streamlit.

    python generar_clave.py

Pide el usuario y la contraseña (la contraseña no se ve mientras la escribís) y
imprime en pantalla el bloque a copiar. La contraseña NO se guarda en ningún archivo:
del bloque que sale solo hay un hash, del que no se puede volver a la contraseña.

Dónde pegarlo: en Streamlit Cloud, en la app → Settings → Secrets. Nunca en el
repositorio.
"""
import getpass
import secrets
import sys

from acceso import hashear


def main():
    print("Generador de credenciales del Liquidador ARCA\n")

    usuario = input("Usuario (sin espacios, ej: agustin): ").strip().lower()
    if not usuario or ' ' in usuario:
        print("\nEl usuario no puede estar vacío ni tener espacios.")
        return 1

    nombre = input("Nombre para mostrar (ej: Agustín): ").strip() or usuario

    clave = getpass.getpass("Contraseña: ")
    if len(clave) < 12:
        print("\nUsá una contraseña de 12 caracteres o más. Es la única barrera "
              "entre la app y cualquiera que tenga la URL.")
        return 1
    if clave != getpass.getpass("Repetir contraseña: "):
        print("\nLas contraseñas no coinciden.")
        return 1

    sal = secrets.token_hex(16)
    print("\n" + "=" * 68)
    print("Pegá esto en Streamlit Cloud → Settings → Secrets")
    print("(si ya hay otros usuarios, agregá solo el bloque, no borres los demás)")
    print("=" * 68 + "\n")
    print(f'[usuarios.{usuario}]')
    print(f'nombre = "{nombre}"')
    print(f'sal = "{sal}"')
    print(f'hash = "{hashear(clave, sal)}"')
    print("\n" + "=" * 68)
    print("La contraseña no quedó guardada en ningún lado: anotala donde la guardes")
    print("habitualmente, porque de acá no se puede recuperar.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
