# excel

Libreria interna para automatizar libros de Excel con `xlwings`: apertura
segura, actualizacion de tablas y tablas dinamicas por nombre, y actualizacion
de formulas M de consultas Power Query (incluido el caso `Sql.Database`).

```python
from excel import ClienteExcel

with ClienteExcel("ruta/al/archivo.xlsx", visible=False) as cliente:
    cliente.configurar_libro()
    cliente.definir_formula_sql(
        "Tab_clientes",
        servidor="MI-SERVIDOR",
        base_datos="MI_BD",
        consulta_sql="SELECT * FROM dbo.clientes",
    )
    cliente.actualizar_tabla_por_nombre("Tab_clientes")
    cliente.actualizar_tabla_dinamica_por_nombre("Din_clientes")
    cliente.guardar()
```

## 1. Que hace y que no

**Hace**

- Abre el libro en una instancia propia de Excel (no la del usuario), sin
  avisos ni autoguardado, con consultas OLEDB en modo sincronico.
- Actualiza tablas (`ListObject`) y tablas dinamicas por nombre, esperando a
  que terminen las consultas asincronas.
- Actualiza la formula M de una consulta de Power Query por nombre, o la
  construye con `Sql.Database(servidor, base_datos, consulta)`.
- Traduce los fallos del libro a excepciones propias con codigo de salida.

**No hace**

- No lee variables de entorno ni archivos `.env`: todo entra por parametro.
  El `.env` es del proyecto consumidor.
- No extrae a DataFrame ni reemplaza a `pandas.read_excel`. Si se necesita
  volumen, el consumidor lee el libro ya actualizado con pandas.
- No reintenta, no programa corridas y no arma el correo de salida: eso es del
  proyecto que la usa.
- No configura logging: solo emite. Si el consumidor no configura nada, no se
  imprime ni se escribe ningun archivo.
- No pide credenciales ni MFA: opera sobre archivos locales.

## 2. Como declararla en el proyecto consumidor

En el `pyproject.toml` de `python_<nombre>`, por Git con el tag fijo (nunca
`sys.path` ni copiar carpetas):

```toml
[project]
dependencies = ["excel"]

[tool.uv.sources]
excel = { git = "https://github.com/SebaxFentex/excel_utils", tag = "v0.2.1" }
```

El repositorio es publico: no hacen falta credenciales. El `tag` fija la version
del codigo que se instala; para tomar otra hay que cambiarlo y resolver de nuevo.

```bash
uv lock && uv sync
```

## 3. Que necesita el proyecto consumidor

| Requisito | Detalle |
|---|---|
| Sistema | Windows con Microsoft Excel instalado (`xlwings` maneja Excel por COM) |
| Python | `>=3.14` |
| Gestor | `uv` (cada proyecto con su propio `.venv` y `uv.lock`) |
| Variables de entorno | Ninguna: rutas y parametros entran por argumento |
| Login manual | No aplica: opera sobre archivos locales, sin MFA |
| Nivel de log | Opcional: `EXCEL_NIVEL_LOG` (`DEBUG`, `INFO`…; default `INFO`), solo si el consumidor llama a `configurar_log()` |

No hay `.bat` de la libreria que el consumidor deba ejecutar.

## 4. Ejemplo minimo

```python
from excel import ClienteExcel

with ClienteExcel("C:/datos/reporte.xlsx") as cliente:
    cliente.actualizar_tabla_por_nombre("Tab_ventas")
    cliente.guardar()
```

Con consulta SQL sobre Power Query y manejo de errores:

```python
import sys
from excel import ClienteExcel, ExcepcionExcel

try:
    with ClienteExcel("C:/datos/reporte.xlsx", visible=False) as cliente:
        cliente.definir_formula_sql(
            "Tab_ventas",
            servidor="SERVIDOR",
            base_datos="BASE",
            consulta_sql="SELECT * FROM dbo.ventas WHERE periodo = '202609'",
        )
        cliente.actualizar_tabla_por_nombre("Tab_ventas")
        cliente.actualizar_tabla_dinamica_por_nombre("Din_ventas")
        cliente.guardar()
except ExcepcionExcel as exc:
    sys.exit(exc.codigo_salida)
```

## 5. API publica

### `ClienteExcel(ruta_archivo=None, *, visible=False)`

Abre un libro de Excel y ejecuta operaciones comunes de automatizacion.
`visible=False` trabaja en segundo plano. Es context manager: al salir del
`with` cierra el libro **sin guardar** (guardar es explicito con `guardar()`);
si quedan cambios sin guardar se avisa con un warning en el log, y con
excepcion dentro del bloque nunca se guarda.

### `abrir(ruta_archivo=None) -> None`

Abre el libro en una instancia dedicada de Excel con avisos y actualizacion
de pantalla apagados, autoguardado desactivado y consultas OLEDB sincronicas.
Sin argumento usa la ruta del constructor. **Lanza** `ValueError` si no hay
ruta en ningun lado y `RuntimeError` si ya hay otro libro abierto (misma
ruta: no hace nada, es idempotente).

### `guardar() -> None`

Guarda el libro abierto. **Lanza** `RuntimeError` si no hay libro abierto.

### `cerrar(*, guardar=False) -> None`

Cierra el libro, guardando primero solo si `guardar` es `True`.

### `salir(*, guardar=False) -> None`

Cierra el libro y termina la instancia de Excel.

### `configurar_libro() -> None`

Punto explicito del flujo del consumidor: la apertura ya deja el libro
configurado. **Lanza** `RuntimeError` si no hay libro abierto.

### `actualizar_tabla_por_nombre(nombre_tabla) -> None`

Actualiza la tabla (`ListObject`) con ese nombre y espera a las consultas
asincronas. **Lanza** `ExcepcionTablaNoEncontrada` (codigo 2) si ninguna hoja
tiene una tabla con ese nombre.

### `actualizar_tabla_dinamica_por_nombre(nombre_tabla_dinamica) -> None`

Actualiza la tabla dinamica con ese nombre y espera a las consultas
asincronas. **Lanza** `ExcepcionTablaDinamicaNoEncontrada` (codigo 3) si
ninguna hoja tiene una tabla dinamica con ese nombre.

### `actualizar_formula_power_query(nombre_consulta, formula) -> None`

Actualiza la formula M de la consulta. **Lanza**
`ExcepcionConsultaPowerQuery` (codigo 4) si el libro rechaza la formula.

### `definir_formula_sql(nombre_consulta, *, servidor, base_datos, consulta_sql) -> None`

Construye `Origen = Sql.Database(servidor, base_datos, [Query=consulta])`
(escapando comillas y saltos de linea al formato M) y la asigna a la
consulta. **Lanza** `ExcepcionConsultaPowerQuery` (codigo 4) si el libro
rechaza la formula.

### `configurar_log(nivel=None, prefijo="excel_utils-", conservar=30) -> Logger`

Configura archivo + consola para el logger `excel` (solo desarrollo).
Nivel por parametro, `EXCEL_NIVEL_LOG` o `INFO`; idempotente.

## 6. Codigos de salida

| Codigo | Significado | Cuando se usa |
|---|---|---|
| `0` | correcto | la ejecucion termino sin errores |
| `1` | fallo generico | convencion Python (`ExcepcionExcel` base, errores no clasificados) |
| `2` | tabla no encontrada | ninguna hoja tiene una tabla (`ListObject`) con ese nombre |
| `3` | tabla dinamica no encontrada | ninguna hoja tiene una tabla dinamica con ese nombre |
| `4` | consulta Power Query | el libro rechazo la formula M de la consulta |

Los codigos propios (2 en adelante) son de esta libreria y no se comparten con
otras. El consumidor propaga sin repetir el numero (`sys.exit(exc.codigo_salida)`).
Los datos de entrada malformados llegan nativos: `TypeError` si el tipo es
incorrecto y `ValueError` si el contenido esta vacio.

## 7. Desarrollo

```bash
uv sync                  # crea .venv con las dependencias del lock
uv run python tests/test_abrir.py
uv run python tests/test_actualizar_tablas.py
uv run python tests/test_consulta_power_query.py
uv build                 # sin advertencias de build_system.requires
```

Los tests son locales (sin credenciales): crean sus libros en
`tests/archivos/descarga/`, que se limpia en cada corrida.

```text
src/excel/
  __init__.py        NullHandler + API publica
  _log.py            configurar_log() y purgar_logs() para desarrollo
  _excepciones.py    ExcepcionExcel (1), ExcepcionTablaNoEncontrada (2),
                     ExcepcionTablaDinamicaNoEncontrada (3),
                     ExcepcionConsultaPowerQuery (4)
  cliente.py         ClienteExcel
ejecutables/         build.bat, abrir_vscode.bat y tests/ (un .bat por test)
tests/               los tres tests; el unico lugar del ecosistema con tests
logs/                logs de desarrollo (ignorado por git)
```

El detalle de arquitectura, reglas inviolables e historial de rupturas esta en
`AGENTS.md`.
