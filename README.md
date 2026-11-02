# excel

Libreria para automatizar archivos de Excel con `xlwings`, enfocada en:

- Configuracion segura de libro (sin autoguardado ni popups molestos)
- Refresco de tablas y tablas dinamicas por nombre
- Actualizacion de formulas de consultas Power Query

## Instalacion local con uv

Desde otro proyecto:

```bash
uv add ../librerias/excel
```

## Uso rapido

```python
from excel import ExcelClient

with ExcelClient("ruta/al/archivo.xlsx", visible=False) as client:
    client.configure_workbook()

    client.set_power_query_sql_formula(
        "Tab_clientes",
        server="MI-SERVIDOR",
        database="MI_BD",
        sql_query="SELECT * FROM dbo.clientes",
    )

    client.refresh_table_by_name("Tab_clientes")
    client.refresh_pivot_table_by_name("Din_clientes")

    client.save()
```

## API publica

- `ExcelClient(file_path, *, visible=False)`
- `ExcelClient.open()`
- `ExcelClient.save()`
- `ExcelClient.close(save=False)`
- `ExcelClient.configure_workbook()`
- `ExcelClient.refresh_table_by_name(table_name)`
- `ExcelClient.refresh_pivot_table_by_name(pivot_table_name)`
- `ExcelClient.update_power_query_formula(query_name, formula)`
- `ExcelClient.set_power_query_sql_formula(query_name, *, server, database, sql_query)`

