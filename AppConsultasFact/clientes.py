"""
Almacenamiento CIFRADO de la relación producto -> cliente(s) en clientes.dat.

La relación producto->cliente(s) es información sensible: se guarda CIFRADA en
disco (Fernet, cifrado simétrico autenticado) y solo se descifra en memoria
cuando la app la necesita. Al abrir clientes.dat con un editor no se lee nada
en claro.

Seguridad de acceso (igual criterio que usuarios.json):
- La clave de cifrado se lee de la variable de entorno CLIENTES_KEY (.env);
  NUNCA se hardcodea. Genérala con:
      python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
- Un lock de archivo (portalocker) serializa el acceso entre workers.
- La escritura es atómica: temporal + os.replace, para no dejar nunca un
  archivo a medio escribir.
- clientes.dat vive junto a este módulo (fuera de /static), no es público.

Estructura (en memoria, ya descifrada), código de producto como clave:
    { "2A03008128": {"nombre": "MANGA/LAY FLAT 120 X 0.32 mm",
                     "clientes": ["BIOGALENIC"], "tamano_bolsa": null}, ... }

La primera vez (clientes.dat no existe) se inicializa con el contenido de
clientes_precarga.json y se escribe ya cifrado.
"""
import json
import os
import re
import tempfile

import portalocker
from cryptography.fernet import Fernet, InvalidToken

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENTES_DAT = os.path.join(BASE_DIR, "clientes.dat")
PRECARGA_JSON = os.path.join(BASE_DIR, "clientes_precarga.json")
# Lock en archivo aparte: el de datos se reemplaza con os.replace y no sirve
# para sostener un lock estable entre escrituras.
LOCK_FILE = CLIENTES_DAT + ".lock"
LOCK_TIMEOUT = 10  # segundos


class ClientesError(Exception):
    """
    Error controlado de la relación de clientes (clave ausente, clave inválida
    o archivo corrupto/no descifrable). Los llamadores lo capturan para degradar
    con un mensaje claro sin romper el resto del aplicativo.
    """


def _fernet():
    """Construye el cifrador Fernet con la clave de CLIENTES_KEY (.env)."""
    clave = os.getenv("CLIENTES_KEY")
    if not clave:
        raise ClientesError(
            "Falta la variable CLIENTES_KEY en el entorno (.env). Genérala con: "
            'python -c "from cryptography.fernet import Fernet; '
            'print(Fernet.generate_key().decode())"'
        )
    try:
        return Fernet(clave.encode("utf-8") if isinstance(clave, str) else clave)
    except (ValueError, TypeError) as exc:
        raise ClientesError(
            "CLIENTES_KEY no es una clave Fernet válida (debe ser base64 de 32 "
            "bytes, generada con Fernet.generate_key())."
        ) from exc


def _precarga():
    """Contenido inicial desde clientes_precarga.json (o {} si no existe)."""
    try:
        with open(PRECARGA_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _escribir_sin_lock(datos):
    """Cifra y escribe clientes.dat de forma atómica (temporal + os.replace)."""
    fernet = _fernet()
    cifrado = fernet.encrypt(json.dumps(datos, ensure_ascii=False).encode("utf-8"))
    directorio = os.path.dirname(CLIENTES_DAT)
    fd, tmp = tempfile.mkstemp(dir=directorio, prefix=".clientes_", suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(cifrado)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, CLIENTES_DAT)  # atómico en el mismo sistema de archivos
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _leer_sin_lock():
    """
    Descifra clientes.dat y devuelve el dict. Si el archivo no existe, lo
    inicializa (cifrado) con la precarga. Lanza ClientesError si la clave falta
    o el archivo está corrupto/no descifrable.
    """
    fernet = _fernet()  # valida la clave antes de tocar el archivo
    if not os.path.exists(CLIENTES_DAT):
        datos = _precarga()
        _escribir_sin_lock(datos)  # crea el archivo cifrado inicial
        return datos
    try:
        with open(CLIENTES_DAT, "rb") as f:
            cifrado = f.read()
        plano = fernet.decrypt(cifrado)
        return json.loads(plano.decode("utf-8"))
    except InvalidToken as exc:
        raise ClientesError(
            "No se pudo descifrar clientes.dat: la clave CLIENTES_KEY no "
            "corresponde a este archivo, o el archivo está dañado."
        ) from exc
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise ClientesError(
            "clientes.dat descifrado no contiene un JSON válido."
        ) from exc


def cargar_clientes():
    """Devuelve el dict de la relación (lectura/descifrado bajo lock)."""
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        return _leer_sin_lock()


def guardar_clientes(datos):
    """Reemplaza por completo la relación (escritura cifrada bajo lock)."""
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        _escribir_sin_lock(datos)


def modificar_clientes(mutador):
    """
    Read-modify-write atómico: bajo un único lock descifra el dict, aplica
    mutador(dict) in place y lo vuelve a cifrar/guardar. El mutador puede lanzar
    ValueError para abortar sin guardar. Devuelve lo que retorne el mutador.
    """
    with portalocker.Lock(LOCK_FILE, "a", timeout=LOCK_TIMEOUT):
        datos = _leer_sin_lock()
        resultado = mutador(datos)
        _escribir_sin_lock(datos)
        return resultado


def normalizar_clientes(entrada):
    """
    Normaliza la entrada de clientes a una lista limpia. Acepta una lista o un
    texto separado por '+' o comas. Quita espacios y vacíos, pasa a MAYÚSCULAS
    (coherente con el dataset y estable para agrupar), conserva el orden y
    elimina duplicados.
    """
    if isinstance(entrada, (list, tuple)):
        crudos = [str(x) for x in entrada]
    else:
        crudos = re.split(r"[+,]", str(entrada or ""))
    vistos = set()
    limpio = []
    for c in crudos:
        c = c.strip().upper()
        if not c or c in vistos:
            continue
        vistos.add(c)
        limpio.append(c)
    return limpio
