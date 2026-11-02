"""Cliente reutilizable para manipular libros de Excel con xlwings."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import xlwings as xw


class ExcelClient:
    """Cliente para abrir un archivo de Excel y ejecutar operaciones comunes de automatizacion."""

    def __init__(self, file_path: str | Path | None = None, *, visible: bool = False) -> None:
        self.file_path = Path(file_path) if file_path is not None else None
        self.visible = visible
        self.app: xw.App | None = None
        self.workbook: Any | None = None

    def __enter__(self) -> ExcelClient:
        self.open()
        return self

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        self.close(save=False)

    def open(self, file_path: str | Path | None = None) -> None:
        if self.app is None:
            self.app = xw.App(visible=self.visible, add_book=False)
            self.app.display_alerts = False
            self.app.screen_updating = False

        path = Path(file_path) if file_path is not None else self.file_path

        if path is None:
            return

        if self.workbook is not None:
            return

        self.workbook = self.app.books.open(path)

        if hasattr(self.workbook.api, "AutoSaveOn"):
            self.workbook.api.AutoSaveOn = False

        if hasattr(self.workbook.api, "ConflictResolution"):
            self.workbook.api.ConflictResolution = 2

        for i in range(1, self.workbook.api.Connections.Count + 1):
            conn = self.workbook.api.Connections.Item(i)
            if hasattr(conn, "OLEDBConnection") and conn.OLEDBConnection is not None:
                conn.OLEDBConnection.BackgroundQuery = False

    def save(self) -> None:
        wb = self._require_workbook()
        wb.save()

    def close(self, *, save: bool = False) -> None:
        if self.workbook is not None:
            if save:
                self.workbook.save()
            self.workbook.close()
            self.workbook = None
    
    def quit(self, *, save: bool = False) -> None:
        self.close(save=save)

        if self.app is not None:
            self.app.quit()

        self.app = None


    def refresh_table_by_name(self, table_name: str) -> None:
        """Refresca una tabla de Excel por nombre y espera a que terminen consultas asincronas."""
        wb = self._require_workbook()

        for worksheet in wb.api.Worksheets:
            for i in range(1, worksheet.ListObjects.Count + 1):
                table = worksheet.ListObjects(i)
                if table.Name != table_name:
                    continue

                if hasattr(table, "QueryTable") and table.QueryTable is not None:
                    table.QueryTable.Refresh(False)

                wb.app.api.CalculateUntilAsyncQueriesDone()
                return

        raise ValueError(f"No se encontro la tabla '{table_name}'.")

    def refresh_pivot_table_by_name(self, pivot_table_name: str) -> None:
        """Refresca una tabla dinamica por nombre y espera a que terminen consultas asincronas."""
        wb = self._require_workbook()

        for worksheet in wb.api.Worksheets:
            if worksheet.PivotTables().Count == 0:
                continue

            for i in range(1, worksheet.PivotTables().Count + 1):
                pivot_table = worksheet.PivotTables(i)
                if pivot_table.Name != pivot_table_name:
                    continue

                pivot_table.RefreshTable()
                wb.app.api.CalculateUntilAsyncQueriesDone()
                return

        raise ValueError(f"No se encontro la tabla dinamica '{pivot_table_name}'.")

    def update_power_query_formula(self, query_name: str, formula: str) -> None:
        """Actualiza la formula M de una consulta de Power Query por nombre."""
        wb = self._require_workbook()

        try:
            wb.api.Queries(query_name).Formula = formula
        except Exception as exc:
            raise ValueError(
                f"No se pudo actualizar la consulta de Power Query '{query_name}'."
            ) from exc

    def set_power_query_sql_formula(
        self,
        query_name: str,
        *,
        server: str,
        database: str,
        sql_query: str,
    ) -> None:
        """Construye y asigna una formula M usando Sql.Database para una consulta de Power Query."""
        server_escaped = self._escape_m_string(server)
        database_escaped = self._escape_m_string(database)
        sql_query_escaped = self._escape_m_string(sql_query)

        formula = (
            "let\r\n"
            f'    Origen = Sql.Database("{server_escaped}", "{database_escaped}", '
            f'[Query="{sql_query_escaped}"])\r\n'
            "in\r\n"
            "    Origen"
        )
        self.update_power_query_formula(query_name, formula)

    def _require_workbook(self) -> Any:
        if self.workbook is None:
            raise RuntimeError("Debes abrir el libro primero con open().")
        return self.workbook

    @staticmethod
    def _escape_m_string(value: str) -> str:
        """Escapa texto para incrustarlo como literal string dentro de una formula M."""
        escaped = value.replace('"', '""')
        escaped = escaped.replace("\r\n", "#(lf)").replace("\n", "#(lf)").replace("\r", "#(lf)")
        return escaped

