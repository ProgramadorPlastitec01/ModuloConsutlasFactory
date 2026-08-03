# Consultas Fact — Documento de Contexto (Triple V)
 
> **Proyecto:** Consultas Fact — Módulo de Cálculo de Desperdicio
> **Empresa:** PLASTITEC · Área de Tecnología de Información
> **Dirigido a:** Administración de Producción (Rodrigo Peña Muñoz)
> **Naturaleza:** Interfaz gráfica que automatiza el cálculo del % de desperdicio de las órdenes de producción del ERP Factory. **No** es un aplicativo en sentido estricto; es una capa de consulta/cálculo de solo lectura sobre Factory.
> **Fecha del documento:** julio de 2026
 
Este documento resume **de dónde Venimos, dónde Vamos (estamos) y hacia dónde Vamos**, con todo el contexto técnico, las decisiones tomadas y su justificación, para poder retomar el proyecto sin perder nada.
 
---
 
## 0. Resumen ejecutivo en una línea
 
Un proceso manual de horas (consultar orden por orden en Factory, leer el desperdicio de la nota, sumar consumos y calcular en Excel) se convirtió en una consulta web de segundos, con análisis, indicadores y exportación. Ya está en despliegue a producción y se proyecta como base de una plataforma de más módulos de consulta.
 
---
 
# 1. DE DÓNDE VENIMOS
 
## 1.1 El problema original
- El cálculo del % de desperdicio de las órdenes de producción se hacía **a mano**.
- El ERP **Factory** no lista el contenido de todas las órdenes juntas: había que entrar **orden por orden**, leer el valor de desperdicio consignado en la **nota**, reunir y sumar el **consumo**, pasar todo a **Excel** y calcular allí.
- Costaba **un par de horas** para reportes de varios meses, más para un anual, y era **propenso a error humano**.
## 1.2 La solicitud formal
- **Solicitante:** Rodrigo Peña Muñoz — Administrador de Producción.
- **Fecha:** 2 de julio de 2026 (correo formal "Solicitud Consulta Órdenes de Producción Factory").
- **Receptor:** Jefatura de TI (Ivan Javier Rairán); desarrollo por el Programador TI.
- **Enfoque inicial:** área de extrusión, código **2A03 (manga)** y su consumo de compuesto.
- **Ejemplo aportado (OP 459634):** 14.850 metros de manga, consumo 698,16 + 1.260,34 = 1.958,5 kg, desperdicio 325 kg (nota "DES X 325 KG X").
- **Salida pedida:** tabla con columnas `OC | OP | FECHA | METROS | KG COMPUESTO | KG DESP | ANCHO | %Desp`.
## 1.3 Punto de partida técnico (primera versión / MVP)
- App **Flask** (Python), plantillas **Jinja2**, CSS propio + JS inline.
- Acceso **solo lectura** a **SQL Server** (base `EMP001_INV`, Factory) vía **pyodbc**, consultas **parametrizadas**.
- **Sin base de datos propia**: la única persistencia era `usuarios.json` (login por usuario/rol; contraseñas con **hash scrypt** vía werkzeug; bloqueo con portalocker).
- Módulo único: Órdenes de Producción (filtro, consulta, cálculo de desperdicio por orden, export a Excel con openpyxl).
- Endpoint `/health` para diagnóstico de conexión a BD.
- Cálculo por orden: `% = valor_nota / suma_consumo × 100`, con el consumo sumado del detalle.
- Estaba pendiente: servidor WSGI de producción, credenciales fuera de texto plano, versiones fijadas, etc.
## 1.4 Lógica de cálculo original (funciones clave en app.py)
- `_parse_desperdicio_nota`: extrae el valor de desperdicio de la NOTA con 3 patrones regex (formatos tipo "DES X 136.12 KG X").
- `_sumar_cant1_2a`: suma el consumo (denominador).
- `_calcular_desperdicio`: aplica `valor_nota / consumo × 100`.
---
 
# 2. DÓNDE VAMOS (ESTADO ACTUAL)
 
## 2.1 Identidad y diseño visual (rediseño completado)
El front se rediseñó de un estilo "genérico/IA" a un lenguaje **SaaS limpio y profesional** (referencia Linear/Vercel/Stripe), **sin Tailwind ni build step** — CSS plano con **variables CSS como design tokens** y macros/estructura de Jinja.
 
**Design tokens (paleta índigo):**
- Primario `#4F46E5`, primario oscuro `#4338CA`/`#3730A3`, primario suave `#EEF0FE`.
- Fondo `#FBFBFC` (gris muy claro, no blanco puro), superficie `#FFFFFF`, borde `#EAEBEF`.
- Texto `#141519`, texto-2 `#5C6270`, texto-3 `#9296A1`.
- Tipografía **Inter** (Google Fonts). Números de KPI en Inter 800 con `tabular-nums`. Sin serif.
**Regla de color del % de desperdicio (escala de 3 niveles):**
- **Verde** `#0E9F6E` → **< 10 %** (aceptable)
- **Ámbar** `#E8A317` (texto `#A67908`) → **10 %–15 %** (atención)
- **Naranja** `#F07A1E` (texto `#D9631A`) → **> 15 %** (alto)
- (Nota: el rojo original `#E0523A` se descartó por "alarmante".)
**Layout:** sidebar fijo con navegación, topbar con breadcrumb + acciones, contenido con ancho máx. ~1080px. Botón **Exportar Excel = verde** (`#1D6F42`), **Descargar PDF = rojo** (`#C8322B`), primario índigo para el resto.
 
## 2.2 Estructura de la vista de resultados (3 pestañas)
- **Resultado:** lista de órdenes (tarjetas), con filtros rápidos (Todas / Con desperdicio / Sin cálculo) y **búsqueda en tiempo real** (JS sobre el DOM, sin recargar).
- **Análisis:** detalle navegable. Tres sub-vistas (Por producto, Por mes, Por cliente) como **tablas con filas expandibles** (clic despliega las órdenes del grupo).
- **Resumen:** vista ejecutiva. **KPIs** (órdenes, desperdicio más alto/más bajo con su OP, promedio ponderado), **gráfico de barras horizontales Top 8** (visualización, con enlace "Ver los N en Análisis →") y **tabla de tendencia por producto** con sparklines (única de Resumen).
**Diferenciación Análisis vs Resumen (A1):** la tabla-ranking vive **solo** en Análisis; en Resumen el ranking existe **solo como barras**. Sin duplicación.
 
## 2.3 Cálculo del desperdicio — estado actual
**Promedio ponderado en todos los niveles** (orden, producto, mes, cliente, KPI):
`% = Σ(numerador) / Σ(denominador) × 100`. **Nunca** promedio simple de porcentajes.
- Se calcula **solo sobre las órdenes con desperdicio calculado**; el consumo de órdenes sin dato **no** entra al denominador (decisión documentada en la UI).
**Auditoría realizada (hallazgos y correcciones):**
- **M1 (bug real, corregido):** órdenes con nota pero consumo 0 inflaban los ponderados de grupo (sumaban al numerador sin aportar al denominador). Corregido: una orden solo acumula si `consumo > 0`.
- **M2 (decisión):** el consumo de órdenes sin nota NO entra al denominador; documentado en UI.
- **M3/M4 (verificado):** valor de nota (KG) y CANT1 están en la **misma unidad (kg)** — confirmado con negocio; parseo tolera coma decimal.
**Prefijos de producto consultables:** `2A02, 2A03, 2A04, 2A05, 2A06, 2A10` (validados; prefijo inválido → mensaje y no ejecuta).
 
**Denominador del consumo (único para todos):** `COD LIKE '2A0%' OR COD LIKE '2A62%'`.
- **Decisión clave:** el detalle (op1) de cada orden **ya trae solo lo que esa orden realmente consumió** (Escenario A confirmado). Por eso el filtro amplio es correcto y **no** se necesita consumo por producto. (Se descartó explícitamente un mapeo `CONSUMO_POR_PREFIJO` que se había implementado por error y se revirtió.)
**Métodos de cálculo por producto (arquitectura de reglas):**
- Estructura `REGLAS_DESPERDICIO_POR_PRODUCTO` + `_regla_desperdicio(cod)` / `_calcular_desperdicio_producto`. Prefijo sin regla → regla general.
- **Regla general (nota):** manga (2A03), ducto (2A04), película (2A10) y demás → desperdicio leído de la **NOTA**; consumo `2A0%+2A62%`.
- **Regla especial — compuesto 2A02 (mezcla/peletizado):** desperdicio por **diferencia**, no por nota.
  - Producido = **CANTE** (cantidad entregada, cabecera OP).
  - Consumido = suma del detalle **`2A01%`** (la mezcla que consume el compuesto).
  - Desperdicio (kg) = Consumido − Entregado.
  - **% = Desperdicio / Entregado × 100** (denominador = **ENTREGADO**, decisión confirmada).
  - Validación de referencia: **OP 436615** → 3.672,40 − 3.570 = 102,40 kg → 102,40/3.570 = **2,87 %** ✅ (no 2,79 %).
  - Casos límite: CANTE=0 → sin cálculo; consumido<entregado → 0 %; sin filas 2A01 → sin cálculo.
- **Ponderado mixto:** cada regla expone `num_ponderado` / `den_ponderado`; un grupo (p. ej. por mes) puede mezclar métodos y sigue siendo Σnum/Σden×100.
## 2.4 Conteo de categorías (corregido)
La categoría "Sin nota" se renombró/ampliáá a **"Sin cálculo"** (cubre todo % None sea por nota no interpretable, CANTE=0, sin 2A01, etc.). Regla: **Todas = Con desperdicio + Sin cálculo**, por construcción (sin cálculo = Todas − Con desperdicio). (Surgió porque un conteo daba 181 = 178 + 0, faltaban 3.)
 
## 2.5 Análisis por cliente (A4 — implementado)
- Relación **producto → cliente(s)** almacenada **cifrada** en disco (`clientes.dat`), ilegible al abrirlo, descifrada en memoria.
- Cifrado con **`cryptography` (Fernet)**; clave en `.env` como **`CLIENTES_KEY`** (nunca hardcodeada). Nueva dependencia: `cryptography==48.0.0`.
- Módulo `clientes.py` (espejo de `usuarios.py`: portalocker + escritura atómica, pero cifrado). Errores controlados con `ClientesError`.
- Precarga inicial de **38 registros** desde `clientes_precarga.json`.
- Modelo: `{ "COD": { "nombre":..., "clientes": ["A","B"], "tamano_bolsa": null } }`. Clientes en **lista** (una referencia puede tener varios). `tamano_bolsa` preparado para futuro (hoy null, sin UI).
- **Agrupación Opción B (grupo conjunto):** referencias compartidas → una **fila única** "FMCA + BEKER" (no se suma a cada cliente por separado). Cada orden se cuenta una vez; Σ grupos = total (sin doble conteo). Productos sin mapeo → "Sin cliente asignado".
- **Gestión solo admin** desde el app (`/admin/clientes` + guardar/eliminar), protegida en backend con `@admin_required`.
- Validado: cifrado ilegible, round-trip, grupo compartido como fila única, sin doble conteo, permisos por rol.
## 2.6 Otros ajustes funcionales implementados
- **Reporte PDF** para gerencia (servidor, WeasyPrint/ReportLab): portada con criterios, resumen ejecutivo (KPIs), ranking por producto, tendencia y detalle, paginado y con texto real. Botón "Descargar PDF".
- **Loader + timeout** en el frontend al consultar (overlay que bloquea interacción; si tarda mucho, mensaje de error).
- `/health` restringido a admin (protegido en backend, no solo ocultando el enlace).
- Etiquetas de método: órdenes 2A02 muestran "por diferencia" / "consumido − entregado"; PDF renombró NOTA(KG) → DESPERD.(KG).
## 2.7 Documentación generada (Word, estilo unificado)
1. **Levantamiento de Requerimientos** (solicitud original de Rodrigo, RF-01..RF-13, columnas pedidas).
2. **Acta de Presentación N.° 001** (viernes 17 jul 2026; asistentes Rodrigo + TI; compromiso: publicar en servidor y enviar acceso por correo).
3. **Acta de Acuerdos N.° 002** (nuevos requerimientos del correo de adiciones; ejecutados / en curso / futuros).
4. **Presentación en diapositivas** (.pptx, 8 slides, enfoque valor para directivos).
5. **Stack Tecnológico** (.docx).
6. **Manual Técnico** (.docx).
7. **Manual de Usuario** (.docx, con capturas paso a paso).
> Pendiente de actualizar en Stack/Manual Técnico tras cierre: dependencia `cryptography`, archivo cifrado de clientes, y (a futuro) la estructura modular del refactor.
 
## 2.8 DESPLIEGUE A PRODUCCIÓN — estado en curso (lo más reciente)
**Servidor:** VM Linux **Rocky Linux 9** (`localhost.localdomain`), IP interna 172.16.2.240. Conviven varios servicios:
- `appconsultas.service` → **Consultas Fact VIEJO**, corre desde `/opt/AppConsultasFact`, gunicorn 3 workers en **127.0.0.1:8001**. Sigue activo (da servicio).
- `gunicorn.service` → app **Django** (SST u otra). **No tocar.**
- `tomcat.service` → app Java. **No tocar.**
- SST también referenciado en Nginx (`/opt/sst/Sst_betaProyect/testOne/static/`). **No tocar.**
**Detalle crítico del viejo:** su `.service` incluye `Environment=OPENSSL_CONF=/opt/AppConsultasFact/openssl_sql2012.cnf` — necesario para que **OpenSSL 3 (Rocky 9) acepte conectar a SQL Server 2012**. Sin este `.cnf`, la app arranca pero **falla la conexión a BD**.
 
**Estrategia de despliegue:** en **paralelo**, sin borrar el viejo. Repo subido a **GitHub por primera vez** (sin `.env`, sin `clientes.dat`, sin `usuarios.json`, sin entorno virtual; sí `.env.example`, `requirements.txt`, `clientes_precarga.json`).
 
**Nuevo montado en:** `/opt/ModuloConsultas/ModuloConsutlasFactory/AppConsultasFact`
- Entorno virtual se llama **`env`** (¡no `venv`!): ejecutable en `env/bin/gunicorn`.
- Se repuso `usuarios.json` y `.env` desde respaldo; se **añadió `CLIENTES_KEY`** al `.env` (generada con `Fernet.generate_key()`); se copió `openssl_sql2012.cnf` a la carpeta nueva.
- Prueba manual: `gunicorn --bind 127.0.0.1:8055 app:app` → **arrancó OK**.
- Servicio nuevo creado: **`appconsultas-nuevo.service`** (calcado del viejo, ruta nueva, `OPENSSL_CONF` nuevo, puerto **8055**, `env/bin/gunicorn`).
**Nginx actual:** `location / { proxy_pass http://127.0.0.1:8001; ... }` (apunta al viejo). SST en `/etc/nginx/conf.d/django_app.conf`.
 
**ÚLTIMO PASO EN CURSO / BLOQUEO ACTUAL:**
- Al arrancar `appconsultas-nuevo.service` da `Address already in use` en **8055**, porque la **prueba manual de gunicorn** que se lanzó antes en 8055 **quedó corriendo** (nunca se cerró con Ctrl+C).
- **Acción pendiente inmediata:** parar el servicio nuevo, matar el proceso gunicorn manual que ocupa el 8055 (identificar con `ss -tlnp | grep 8055`, `kill <pid>` — SIN tocar 8001/SST/Tomcat), liberar el puerto, y reiniciar `appconsultas-nuevo.service` limpio.
- **Luego:** `curl http://127.0.0.1:8055/health` para confirmar que conecta a BD (valida el `.cnf`).
---
 
# 3. HACIA DÓNDE VAMOS
 
## 3.1 Terminar el despliegue (inmediato)
1. Liberar 8055 (matar gunicorn manual huérfano) y dejar `appconsultas-nuevo.service` en `active (running)` limpio.
2. Validar `/health` en 8055 (conexión a BD OK).
3. Prueba funcional del nuevo por su puerto (consulta conocida, tabs, clientes) conviviendo con el viejo.
4. **Cambio de apuntamiento:** o bien cambiar el `proxy_pass` de Nginx de 8001 → 8055, o intercambiar puertos (que el nuevo pase a 8001). Recargar Nginx (`nginx -t && systemctl reload nginx`).
5. Habilitar el servicio nuevo en el arranque (`systemctl enable appconsultas-nuevo.service`).
6. **Detener** el viejo (`systemctl stop appconsultas.service`) y dejarlo **deshabilitado pero sin borrar** unos días como respaldo. SST y Tomcat intactos.
7. Verificar en navegador desde la red interna.
8. Comunicar el enlace de acceso por correo a los usuarios (compromiso del acta).
**Recordatorios de despliegue:**
- Guardar `CLIENTES_KEY` y `SECRET_KEY` de producción en lugar seguro **fuera del servidor**. Si se pierde `CLIENTES_KEY`, `clientes.dat` es irrecuperable.
- Claves de producción **distintas** a las de desarrollo. Usuario de BD con permisos mínimos (no `sa`).
- `debug` desactivado en producción.
- Mejora futura: correr el servicio con un usuario de servicio dedicado, no `root`.
## 3.2 Refactor de escalabilidad (B5 — planificado, al final)
- Reorganizar el `app.py` único en estructura modular: **blueprints** por módulo, **capa de acceso a datos** a Factory reutilizable, config/auth/utilidades centralizadas.
- Objetivo: que nuevos módulos de consulta "se enchufen" sin reescribir la base.
- **Se hace aparte y con validación** de que todo funciona igual que antes (es estructural, delicado). Por eso se dejó para después de lo funcional y del despliegue; el commit limpio en GitHub es la red de seguridad.
## 3.3 Farmacéutico (fase futura — ya viable, ya NO depende de que Factory cambie)
**Hallazgo (analizado en sitio con Sandra Núñez):**
- La orden se crea **vacía**; tras generarse, la **nota de la OP queda bloqueada**. El producto se ingresa por **movimientos** relacionados a la orden, sin tocar la nota. Por eso la nota nunca sirvió para farma.
- **Vía encontrada:** en el registro del movimiento hay un campo **"concepto"** que **sí es editable y sí guarda** (probado). Ahí se consignaría el desperdicio.
- El desperdicio vendría en un patrón tipo **`DES`** dentro del concepto → probablemente **reutilizable** con la lógica de parseo existente.
- Relación movimiento↔orden por el campo **`OP_OC`** (pendiente confirmar con script desde la BD).
**Complejidad principal — 3 tablas Kardex por antigüedad (mismos campos):**
- `Kardex` → mes actual · `KardexA` → mes anterior · `Kardeshi` → histórico (+2 meses).
- **El dato SE MUEVE entre tablas al envejecer** (Kardex → KardexA → Kardeshi). Implica: consultas que crucen meses deben unir las tres (UNION) y filtrar por fecha; no asumir tabla fija, porque un movimiento migra con el tiempo.
**Estado:** documentar el hallazgo con el script real de la BD (nombres de campos de las Kardex + OP_OC + patrón DES) y luego planificar la implementación. Decisión: **se deja para después del despliegue actual**; no se metió con calzador.
 
## 3.4 Otras adiciones futuras (del correo de adiciones de Rodrigo)
- **Extrusión PP:** ducto (2A04) y película (2A10) **ya incluidos** (consultables, cálculo por nota, consumo por el filtro amplio). ✅ Hecho.
- **Ajustes de fabricación:** revisar screen/bocas vs Colpitts/Plumatts (unas consumen metros de manga directo, otras bolsas parcialmente fabricadas). Refinamiento futuro.
- **Tamaño de bolsa** por referencia y afinado de referencias compartidas por varios clientes (estructura ya preparada, campo `tamano_bolsa`).
- Alimentar y mantener la relación producto–cliente (responsabilidad de Producción).
## 3.5 Visión de largo plazo
El activo más valioso es la **conexión establecida con Factory**: es una capa reutilizable. Sobre ella (y sobre el refactor modular) se planea construir **más módulos de consulta** de otros temas, con Consultas Fact como plataforma base. Por eso la prioridad de una base sólida y escalable.
 
---
 
# 4. Referencias rápidas (cheat-sheet)
 
| Tema | Valor |
|---|---|
| Prefijos consultables | 2A02, 2A03, 2A04, 2A05, 2A06, 2A10 |
| Denominador consumo (general) | `COD LIKE '2A0%' OR COD LIKE '2A62%'` |
| Regla 2A02 (compuesto) | desp = Σ2A01 − CANTE ; % = desp / CANTE × 100 |
| Validación 2A02 | OP 436615 → 2,87 % |
| Escala color | verde <10 % · ámbar 10–15 % · naranja >15 % |
| Ponderado | Σnum / Σden × 100 (nunca promedio simple) |
| Clientes | `clientes.dat` cifrado (Fernet), clave `CLIENTES_KEY` en .env, 38 precargados, Opción B |
| Servidor | Rocky Linux 9, /opt/... , OpenSSL cnf para SQL Server 2012 |
| Viejo | `appconsultas.service`, /opt/AppConsultasFact, 8001 |
| Nuevo | `appconsultas-nuevo.service`, /opt/ModuloConsultas/ModuloConsutlasFactory/AppConsultasFact, env/, 8055 |
| NO tocar | SST (gunicorn.service + django_app.conf), Tomcat |
| Persistencia | usuarios.json (hash scrypt), clientes.dat (cifrado). Sin BD propia. |
 
---
 
# 5. Personas
 
| Nombre | Rol |
|---|---|
| Rodrigo Peña Muñoz | Administrador de Producción (solicitante / usuario) |
| Ivan Javier Rairán | Jefatura de TI |
| Sandra Núñez | Producción (apoyó el análisis del proceso farmacéutico) |
| Programador TI | Desarrollo (Área de Tecnología de Información) |
 
---
 
*Fin del documento de contexto — Consultas Fact · PLASTITEC · Tecnología de Información.*