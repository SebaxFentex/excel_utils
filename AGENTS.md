# AGENTS.md — excel_utils

Guia para agentes (y humanos) que trabajan en este repo. Cada regla con su
"porque" viene de un defecto real; no relajarla sin evidencia nueva.

## 1. Proyecto en 30 segundos

Libreria interna PAQUETE (`uv_build`, layout `src/`) para automatizar libros
de Excel con `xlwings`: apertura segura en instancia propia, actualizacion de
tablas y tablas dinamicas por nombre, y formulas M de Power Query. Se publica
en `https://github.com/SebaxFentex/excel_utils` y los proyectos de
`D:\Tigo\Proyectos` la declaran por Git con tag fijo en su
`[tool.uv.sources]`; no se opera con el Programador de tareas y no pide
credenciales a ningun humano.

Requiere Windows con Excel instalado (`xlwings` maneja Excel por COM): en
Linux/macOS el import carga pero `abrir()` no puede funcionar.

Stack: Python 3.14, `uv`, `xlwings==0.37.8`.

```text
src/excel_utils/
  __init__.py        NullHandler + API publica (sin re-exports de modulos internos)
  _excepciones.py    ExcepcionExcel (1), ExcepcionTablaNoEncontrada (2),
                     ExcepcionTablaDinamicaNoEncontrada (3),
                     ExcepcionConsultaPowerQuery (4)
  cliente.py         ClienteExcel
Ejecutables/
  build.bat          uv build y lista dist/
  abrir_vscode.bat   abre la raiz de la libreria
  tests/             un .bat por test (espejo de tests/)
tests/               los tres tests; el unico lugar del ecosistema con tests
logs/                logs de desarrollo (ignorado por git)
```

## 2. Arquitectura y API

- **La libreria no lee entorno.** `import excel_utils` no exige ninguna variable y no
  abre Excel. Todo entra por parametro: la ruta del libro, `visible` y los
  nombres de tabla/consulta. Los tests leen `EXCEL_NIVEL_LOG` y lo pasan a
  `log_utils.configurar_log`.
- `ClienteExcel(ruta_archivo=None, *, visible=False)`: context manager. Al
  salir del `with` cierra el libro **sin guardar** (guardar es siempre
  explicito, tambien con excepcion: nunca se persiste a medias); si el bloque
  termino limpio con cambios sin guardar, se avisa con un warning en el log.
  Orden de metodos: dunder → publicos → privados.
- `abrir()` exige ruta (argumento o constructor, `ValueError` si falta) y no
  permite cambiar de libro sin `cerrar()` previo (`RuntimeError`, salvo misma
  ruta que es idempotente). `cerrar()`/`salir()` son idempotentes: cerrar dos
  veces no es error.
- `definir_formula_sql()` escapa con `str.replace` (dobla comillas, convierte
  saltos a `#(lf)`): a proposito sin `re.sub`, cuyo reemplazo interpretaria
  los backslash del valor (una ruta `C:\nuevo` se corromperia).
- `configurar_libro()` es el punto explicito del flujo: hoy `abrir()` ya deja
  el libro configurado; si la configuracion crece, vive ahi, no en `abrir()`.

## 3. Reglas inviolables

1. **Cero lecturas de entorno en la libreria.** Ningun modulo de `src/`
   importa `os`: los tests leen `EXCEL_NIVEL_LOG` y lo pasan como parametro a
   `log_utils.configurar_log`. Un valor que cambia por maquina se pasa como
   parametro.
2. **Codigos por donde ocurrio, no por el driver.** Tabla inexistente (2) y
   tabla dinamica inexistente (3) son estados distintos del libro: no se
   fusionan en un solo "no encontrado". Formula rechazada (4) es error del
   libro, no del parametro.
3. **La libreria emite, el entrypoint configura.** `logging.getLogger(__name__)`
   en cada modulo y `NullHandler` en el paquete; `configurar_log()` solo se
   llama desde desarrollo. Prohibido `basicConfig()` y prohibido que un modulo
   que emite agregue manejadores: el unico que agrega es
   `log_utils.configurar_log()`, y lo hace sobre el logger "excel_utils", nunca sobre el
   raiz del consumidor.
4. **Sin `print()`.** El unico destino de las trazas es `log`.
5. **Sin efectos al importar.** Prohibido abrir Excel a nivel de modulo y
   tragarte la excepcion con `except Exception: app = None`: el consumidor
   descubriria el problema mucho despues.
6. **Los datos de entrada malformados se dejan nativos.** `TypeError` si el
   valor no es texto y `ValueError` si es texto pero llega vacio o solo con
   espacios. Nunca se envuelven en `ExcepcionExcel`.
7. **`.bat` de proceso en `Ejecutables/` invocando `uv run`.**
   `abrir_vscode.bat` es utilitario: no lleva `uv run`.
8. **Nada especulativo.** Sin abstracciones de un solo uso ni parametros de
   "por si acaso". Lo muerto se borra: queda en git si se necesita.
9. **La descripcion del docstring dice que hace la funcion, nada mas.** Los
   primeros parrafos (antes de `Args`/`Returns`/`Raises`/`Notes`) describen la
   funcion misma; lo que hace la dependencia (`xlwings` por COM, Excel en
   segundo plano) va en `Notes` o en el README. La descripcion no compara con
   otras librerias ni con versiones anteriores: el contrato se enuncia en
   positivo ("devuelve `None`", no "no devuelve filas").
10. **Espanol en la API.** Modulos, funciones, metodos, parametros y mensajes
    en espanol; `snake_case`, constantes `MAYUSCULAS_CON_GUIONES_BAJOS`.
    Conceptos tecnicos consagrados conservan el ingles (`log`, `ListObject`,
    `Sql.Database`, formula M).

## 4. Logging, codigos de salida y `.env`

- Los tests llaman `log_utils.configurar_log(nombre="excel_utils",
  carpeta_logs=<repo>/logs, nivel_archivo/nivel_consola=os.environ.get("EXCEL_NIVEL_LOG"),
  prefijo="excel_utils-")`: archivo en DEBUG y consola en INFO, cada uno con
  su formato, purga los antiguos e idempotente. Devuelve el `Logger`
  configurado.
- Codigos: `0` correcto · `1` fallo generico · `2` tabla · `3` tabla dinamica ·
  `4` Power Query. El consumidor hace `sys.exit(exc.codigo_salida)`.
- Sin `.env` (D2: tests locales sin configuracion externa). Si un test futuro
  necesita config, su `.bat` lleva `--env-file` y se versiona `.env.example`.
- Nunca loguear rutas sensibles si el consumidor asi lo considera: los `INFO`
  registran nombres de tabla/consulta y la ruta abierta.

## 5. Verificacion antes de dar por hecho

1. `uv run python -m compileall -q src/excel_utils`.
2. `import excel_utils` con el entorno vacio: sin error, `handlers` con el `NullHandler`
   y sin instancia de Excel creada.
3. `ty check src/excel_utils tests` y `ruff check src tests` + `ruff format --check`
   en verde (los `noqa: BLE001` llevan su razon: COM de Excel lanza
   `com_error` donde un hasattr/None no alcanza; el aviso
   `reportOptionalMemberAccess` del editor no lo ve el CLI: da
   `All checks passed!`).
4. `uv build` sin advertencias de `build_system.requires`; el wheel debe traer
   `excel_utils/__init__.py`.
5. Greps de guardia: cero `getenv`/`os.environ` en `src/`, cero
   `print(`, cero `basicConfig`, cero `load_dotenv`, cero `cwd(`.
6. Los `.bat` se prueban ejecutandolos: `cmd.exe /c Ejecutables\build.bat` debe
   llegar al `pause` final listando `dist\`.
7. Tres tests (`test_abrir`, `test_actualizar_tablas`,
   `test_consulta_power_query`): locales, sin credenciales, cada uno con su
   `.bat` espejo. La tabla dinamica positiva (RefreshTable real) queda
   pendiente de un libro fixture con dinamica: el COM para crearla es fragil
   entre versiones de Excel.

## 6. Versionado

`X.Y.Z`: `Z` correcciones/refactor/docs sin cambio observable; `Y`
funcionalidad nueva o ruptura migrable; `X` primera estable o rediseno.
Tras un cambio verificado se pregunta si se sube version con su tag o se
espera; commit y push si van siempre. Cada bump: `pyproject` → `uv lock` →
`uv build` → tag `vX.Y.Z` → push `--follow-tags` → release.

## 7. Historial

- **0.4.0 (2026-11-02)** — Logging via `log_utils` (v0.1.0, dependencia por
  Git con tag): se elimina `src/excel_utils/_log.py`; los tests llaman
  `log_utils.configurar_log` con archivo en DEBUG y consola en INFO.
- **0.3.0 (2026-11-02)** — Renombre del paquete `excel`→`excel_utils` (Fase 2):
  `src/excel/`→`src/excel_utils/`, imports absolutos, `configurar_log`/
  `purgar_logs` fuera del `__all__` (siguen en `excel_utils._log`), logger
  `excel`→`excel_utils`, tests en `tests/archivos/`, `LICENSE` MIT, LF +
  `.gitattributes`. Mapeo: `import excel`→`import excel_utils`,
  `from excel._log import ...`→`from excel_utils._log import ...`.
- **0.2.1 (2026-11-02)** — Suite de tests locales (`test_abrir`,
  `test_actualizar_tablas`, `test_consulta_power_query`) con sus `.bat`
  espejo; dos correcciones COM halladas por los tests (`AutoSaveOn` en libro
  local, `QueryTable` en tabla plana); ruff en verde (`Self`, tipos de
  `__exit__`, `noqa: BLE001` con razon, format).
- **0.2.0 (2026-11-02)** — D5 (opcion C): `abrir()` falla rapido (`ValueError`
  sin ruta, `RuntimeError` con otro libro abierto; misma ruta idempotente),
  `__exit__` avisa con warning si hay cambios sin guardar y nunca guarda con
  excepcion; flag `_modificado` (lo ponen `actualizar_*`/`definir_*`, lo limpia
  `guardar()`/`cerrar()`).
- **0.1.3 (2026-11-02)** — Primera version publicada. Reestructura a norma de
  librerias: layout `src/`, backend `uv_build`, `requires-python>=3.14`,
  `_log.py` + `NullHandler`, `_excepciones.py` con codigos, `README.md` en
  7 bloques y este `AGENTS.md`. Renombre total a espanol (D1); el consumidor
  sera redisenado y no se migra aqui. Mapeo viejo→nuevo:
  `ExcelClient`→`ClienteExcel`, `client.py`→`cliente.py`,
  `open`→`abrir`, `save`→`guardar`, `close`→`cerrar`, `quit`→`salir`,
  `configure_workbook`→`configurar_libro`,
  `refresh_table_by_name`→`actualizar_tabla_por_nombre`,
  `refresh_pivot_table_by_name`→`actualizar_tabla_dinamica_por_nombre`,
  `update_power_query_formula`→`actualizar_formula_power_query`,
  `set_power_query_sql_formula`→`definir_formula_sql`
  (`server`→`servidor`, `database`→`base_datos`, `sql_query`→`consulta_sql`),
  `ValueError` de tabla→`ExcepcionTablaNoEncontrada` (2),
  `ValueError` de dinamica→`ExcepcionTablaDinamicaNoEncontrada` (3),
  `ValueError` de consulta→`ExcepcionConsultaPowerQuery` (4).
