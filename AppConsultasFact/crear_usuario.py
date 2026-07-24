"""
Script auxiliar para dar de alta el PRIMER admin (bootstrap) o crear/actualizar
usuarios desde consola, sin editar usuarios.json a mano.

Como sin admin no se puede entrar a la vista de gestión web, este script
permite crear ese primer admin. Puede crearse con debe_cambiar=False para
poder ingresar directamente.

Usa el almacenamiento centralizado (usuarios.py): lock de archivo + escritura
atómica, igual que la app.

Uso:
    python crear_usuario.py
"""
import getpass

from werkzeug.security import generate_password_hash

from usuarios import cargar_usuarios, modificar_usuarios

ROLES_VALIDOS = ("admin", "consulta")
LONGITUD_MIN = 6


def main():
    usuarios = cargar_usuarios()

    usuario = input("Usuario: ").strip()
    if not usuario:
        print("El nombre de usuario no puede estar vacío.")
        return

    if usuario in usuarios:
        resp = input(f"'{usuario}' ya existe. ¿Actualizar? (s/N): ").strip().lower()
        if resp != "s":
            print("Cancelado.")
            return

    password = getpass.getpass("Contraseña: ")
    if len(password) < LONGITUD_MIN:
        print(f"La contraseña debe tener al menos {LONGITUD_MIN} caracteres.")
        return
    if password != getpass.getpass("Repetir contraseña: "):
        print("Las contraseñas no coinciden.")
        return

    rol = input(f"Rol {ROLES_VALIDOS}: ").strip().lower()
    if rol not in ROLES_VALIDOS:
        print(f"Rol inválido. Debe ser uno de: {', '.join(ROLES_VALIDOS)}")
        return

    # Para el primer admin conviene NO forzar cambio (poder entrar directo).
    resp = input("¿Forzar cambio de contraseña en el primer ingreso? (s/N): ").strip().lower()
    debe_cambiar = resp == "s"

    nuevo_hash = generate_password_hash(password)

    def _mutar(usuarios_dict):
        usuarios_dict[usuario] = {
            "password_hash": nuevo_hash,
            "rol": rol,
            "debe_cambiar": debe_cambiar,
        }

    modificar_usuarios(_mutar)
    print(
        f"Usuario '{usuario}' guardado con rol '{rol}' "
        f"(debe_cambiar={debe_cambiar})."
    )


if __name__ == "__main__":
    main()
