# CLAUDE.md — Consultas Fact (Módulo de Cálculo de Desperdicio)

Instrucciones para trabajar en este repositorio. Léelas antes de hacer cambios.

---

## Qué es este proyecto

App **Flask** (Python) que automatiza el cálculo del **% de desperdicio** de las órdenes de producción del ERP **Factory**. Es una **interfaz gráfica de consulta y cálculo de solo lectura** sobre Factory — no es un sistema con datos propios. Empresa: **PLASTITEC**, área de Tecnología de Información. Usuario principal: Administración de Producción.

Todo (UI, mensajes, código de cara al usuario, commits) va en **español**.

---

## Reglas críticas (NO romper)

1. **Solo lectura sobre Factory.** Nunca INSERT/UPDATE/DELETE contra la base del ERP. Todas las consultas son de lectura.
2. **Consultas SIEMPRE parametrizadas** (marcadores `?`, nunca concatenar valores en el SQL). Aplica también a filtros `LIKE` y a las tablas Kardex.
3. **Sin dependencias nuevas** salvo autorización explícita. La única dependencia especial ya aprobada es `cryptography` (para el cifrado de clientes). No añadir frameworks ni build steps.
4. **Sin build de frontend.** CSS plano con variables CSS (design tokens) + JS inline + Jinja2. Nada de Tailwind/Node/compilación. Gráficos con Chart.js por CDN.
5. **No exponer secretos.** Credenciales y claves viven en `.env` (no versionado). Nunca hardcodear claves. `.gitignore` cubre `.env`, `clientes.dat`, `usuarios.json`, entorno virtual.
6. **No romper lo que funciona.** Cálculo, tabs (Resultado/Análisis/Resumen), Excel, PDF, login/roles y gestión de usuarios/clientes deben seguir operativos tras cualquier cambio.
7. **Modo cuidadoso en cambios de cálculo:** implementar, luego VALIDAR con un caso conocido y reportar antes/después, y confirmar que lo demás no cambió. No dar por cerrado sin validación.

---

## Convenciones de código

- **Idioma:** español en UI, mensajes de error, etiquetas y comentarios de cara al negocio.
- **Nombres internos:** no renombrar rutas/funciones existentes por estética (p. ej. la ruta interna sigue siendo `/ordenes-produccion` aunque el módulo se muestre como "Cálculo de desperdicio").
- **Persistencia:** archivos, no BD propia. `usuarios.json` (usuarios) y `clientes.dat` (clientes, cifrado). Usar el patrón de portalocker + escritura atómica que ya existe.
- **Diseño / design tokens** (definidos en `:root` de la hoja de estilos; referenciar SIEMPRE, no hardcodear colores):
  - Primario `#4F46E5`, primario oscuro `#4338CA`/`#3730A3`, primario suave `#EEF0FE`.
  - Fondo `#FBFBFC`, superficie `#FFFFFF`, borde `#EAEBEF`.
  - Texto `#141519`, texto-2 `#5C6270`, texto-3 `#9296A1`.
  - Tipografía **Inter**; números de KPI en peso 800 con `tabular-nums`. Sin serif.
  - Botones: **Excel = verde `#1D6F42`**, **PDF = rojo `#C8322B`**, resto índigo primario.
- **Escala de color del % de desperdicio** (verde/ámbar/naranja):
  - **Verde `#0E9F6E`** → `< 10 %` (aceptable)
  - **Ámbar `#E8A317`** (texto `#A67908`) → `10 %–15 %` (atención)
  - **Naranja `#F07A1E`** (texto `#D9631A`) → `> 15 %` (alto)

---

## Lógica de negocio del cálculo (núcleo — tratar con máximo cuidado)

### Principio general
- **% siempre PONDERADO**, en todos los niveles (orden, producto, mes, cliente, KPI): `% = Σ(numerador) / Σ(denominador) × 100`. **Nunca** promedio simple de porcentajes.
- Cada regla de cálculo expone `num_ponderado` y `den_ponderado`, de modo que un grupo (p. ej. por mes) pueda mezclar métodos y siga siendo `Σnum / Σden × 100`.
- El ponderado se calcula **solo sobre órdenes con desperdicio calculado**; el consumo de órdenes "sin cálculo" NO entra al denominador (decisión documentada en UI).
- **Conteo por construcción:** `Todas = Con desperdicio + Sin cálculo` (la categoría "Sin cálculo" = Todas − Con desperdicio; no contar por separado).

### Arquitectura de reglas por producto
`REGLAS_DESPERDICIO_POR_PRODUCTO` + `_regla_desperdicio(cod)` / `_calcular_desperdicio_producto`. Prefijo sin regla → regla general. **Para agregar un producto con cálculo distinto, registrar una entrada por prefijo; no reescribir las demás reglas.**

### Prefijos de producto consultables
`2A02, 2A03, 2A04, 2A05, 2A06, 2A10` (y `2B`, `2C` al implementarse el farmacéutico). Prefijo inválido → mensaje de validación y no ejecutar.

### Regla GENERAL — por NOTA (manga 2A03, ducto 2A04, película 2A10, etc.)
- Desperdicio: se parsea de la NOTA de la orden con `_parse_desperdicio_nota` (patrón tipo `"DES X 325 KG X"`).
- Consumo (denominador): suma del detalle op1 con **`COD LIKE '2A0%' OR COD LIKE '2A62%'`** (filtro ÚNICO para todos; el detalle de cada orden ya trae solo lo realmente consumido, por eso el filtro amplio es correcto — NO crear consumo por producto).
- `% = valor_nota / consumo × 100`.

### Regla ESPECIAL — compuesto 2A02 (mezcla/peletizado), por DIFERENCIA
- Producido = **CANTE** (cantidad entregada, cabecera OP).
- Consumido = suma del detalle **`2A01%`** (mezcla).
- Desperdicio = Consumido − Entregado.
- **% = Desperdicio / Entregado × 100** (denominador = ENTREGADO).
- Validación de referencia: **OP 436615 → 2,87 %** (3.672,40 − 3.570 = 102,40; 102,40/3.570).
- Casos límite: CANTE=0 → sin cálculo; consumido<entregado → 0 %; sin filas 2A01 → sin cálculo.

### Regla FARMACÉUTICO — prefijos 2B / 2C (en implementación)
- Fuente: tablas **Kardex / KardexA / Kardexhi** (mismos campos `COD, CANT, OP_OC, DES`), consultadas con **UNION ALL** por `OP_OC` (= número de OP). El dato migra entre tablas; por eso se consultan las tres, no se predice una.
- Filtro de filas: solo las cuyo campo **`DES` empieza con `"DES X"`** (patrón de desperdicio). No filtrar por COD aquí.
- Parseo: reutilizar `_parse_desperdicio_nota` **generalizado a cualquier unidad** (UN, KG, …), no solo KG. Verificar que no rompe los formatos KG de la extrusión.
- Cálculo: `% = Σ(DES parseado) / Σ(CANT) × 100` sobre las filas DES de la orden.
- Unidades: DES y CANT de una misma fila van en la misma unidad (válido). NO sumar magnitudes absolutas entre productos de distinta unidad; usar el % ponderado.
- Solo lectura sobre Kardex; verificar permisos de lectura del usuario de BD sobre las tres tablas.

---

## Clientes (análisis por cliente)

- Relación **producto → cliente(s)** en `clientes.dat`, **cifrado** con `cryptography` (Fernet). Clave en `.env` como `CLIENTES_KEY` (nunca hardcodeada). Módulo `clientes.py` (espejo de `usuarios.py`, cifrado).
- Modelo: `{ "COD": { "nombre":..., "clientes": ["A","B"], "tamano_bolsa": null } }`. Clientes en **lista** (una referencia puede tener varios). `tamano_bolsa` reservado a futuro.
- **Agrupación Opción B:** referencias compartidas = **fila única** "A + B" (no sumar a cada cliente por separado). Cada orden se cuenta una vez; **Σ grupos = total** (sin doble conteo). Sin mapeo → "Sin cliente asignado".
- Gestión **solo admin**, protegida en **backend** (no solo ocultar enlaces).
- Si se pierde `CLIENTES_KEY`, `clientes.dat` es irrecuperable.

---

## Roles y seguridad

- Roles: **admin** (gestiona usuarios y clientes, ve `/health`) y **consulta** (usa el módulo).
- Contraseñas con **hash scrypt** (werkzeug), nunca en texto plano.
- Rutas sensibles protegidas por rol **en el backend** (`@admin_required`), no solo ocultando enlaces. Aplica a `/admin/usuarios`, `/admin/clientes`, `/health`.

---

## Despliegue (producción)

- Servidor **Rocky Linux 9**. App servida con **Gunicorn** (WSGI) tras **Nginx** (proxy inverso), gestionada por **systemd**. `debug` SIEMPRE desactivado en producción.
- **Dato crítico:** el servicio define `Environment=OPENSSL_CONF=.../openssl_sql2012.cnf`. Ese `.cnf` es necesario para que **OpenSSL 3 acepte conectar a SQL Server 2012**. Sin él, la app arranca pero **falla la conexión a BD**. No olvidarlo en ningún servicio nuevo.
- El entorno virtual en el servidor se llama **`env`** (no `venv`) → ejecutable en `env/bin/gunicorn`.
- Conviven otros servicios en la VM que **NO se tocan**: SST (Django, `gunicorn.service` + `django_app.conf`) y `tomcat.service`.
- Claves de producción distintas a las de desarrollo; usuario de BD con permisos mínimos (no `sa`); `CLIENTES_KEY`/`SECRET_KEY` guardadas también fuera del servidor.

---

## Al hacer cambios de cálculo o de datos, SIEMPRE

1. Implementar de forma aislada (registrar regla por prefijo si aplica; no tocar las demás).
2. **Validar con un caso conocido** y reportar **antes/después** (p. ej. OP 436615 → 2,87 %; farma OP_OC 463021).
3. Confirmar que **los demás productos NO cambiaron** su resultado.
4. Confirmar que no se rompió: tabs, Excel, PDF, ponderado, conteo de categorías.
5. Documentar decisiones no obvias con un comentario en el código.

---

## Pendientes / roadmap

- **Farmacéutico (2B/2C):** en implementación (regla Kardex descrita arriba). Enfoque incremental: arrancar con el formato de `DES` de ejemplo e ir cubriendo variantes según pruebas.
- **Refactor de escalabilidad (B5):** modularizar (blueprints), aislar capa de datos a Factory, centralizar config/auth/utilidades. Estructural: hacer aparte, sin cambiar funcionalidad, validando equivalencia.
- **Documentación:** mantener actualizados Stack Tecnológico y Manual Técnico (dependencia `cryptography`, `clientes.dat`, estructura modular tras B5).
- **Mejoras futuras:** servicio con usuario dedicado (no `root`); tamaño de bolsa por referencia; ajustes screen/bocas vs Colpitts/Plumatts.

---

## Glosario rápido

- **OP:** orden de producción. **CANTE:** cantidad entregada (producido). **CANT1 / detalle op1:** consumo.
- **Nota:** campo de la OP con el desperdicio (extrusión). **Concepto/DES:** texto de desperdicio en Kardex (farmacéutico).
- **Ponderado:** `Σnum/Σden×100`. **Sin cálculo:** orden sin % (nota no parseable, CANT/CANTE=0, sin datos).
