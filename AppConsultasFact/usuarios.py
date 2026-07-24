"""
Almacenamiento centralizado de usuarios en usuarios.json.

Lectura/escritura seguras ante múltiples workers de Gunicorn:
- Un lock de archivo (portalocker) serializa el acceso.
- La escritura es atómica: se escribe a un temporal y se hace os.replace,
  de modo que nunca queda un JSON a medio escribir.

Estructura de cada usuario:
    { "nombre": {"password_hash": "...", "rol": "admin|consulta",
                 "debe_cambiar": true|false} }
"""
import json
import os
import tempfile

import portalocker

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USUARIOS_JSON = os.path.join(BASE_DIR, "usuarios.json")
# Lock en archivo aparte: el de datos se reemplaza con os.replace y no sirve
# para sostener un lock estable entre escrituras.
LOCK_FILE = USUARIOS_JSON + ".lock"
LOCK_TIMEOUT = 10  # segundos


def _leer_sin_lock():
    """Lee el JSON. Devuelve {} si no existe o está corrupto."""
    try:
        with open(USUARIOS_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _escribir_atomico(usuarios):
    """Escribe el JSON de forma atómica (temporal + os.replace)."""
    directorio = os.path.dirname(USUARIOS_JSON)
    fd, tmp = tempfile.mkstemp(dir=directorio, prefix=".usuarios_", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(usuarios, f, indent=2, ensure_ascii=False)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, USUARIOS_JSON)  # atómico en el mismo sistema de archivos
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def cargar_usuarios():
    """Devuelve el dict de usuarios (lectura bajo lock)."""
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        return _leer_sin_lock()


def guardar_usuarios(usuarios):
    """Reemplaza por completo el archivo de usuarios (escritura bajo lock)."""
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        _escribir_atomico(usuarios)


def modificar_usuarios(mutador):
    """
    Realiza un read-modify-write atómico: bajo un único lock lee el dict,
    aplica mutador(dict) y lo guarda. Así dos workers no pisan sus cambios.

    'mutador' recibe el dict y lo modifica in place; puede lanzar ValueError
    para abortar sin guardar. Devuelve el valor que retorne el mutador.
    """
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        usuarios = _leer_sin_lock()
        resultado = mutador(usuarios)
        _escribir_atomico(usuarios)
        return resultado
