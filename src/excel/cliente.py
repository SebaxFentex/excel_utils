"""Cliente reutilizable para manipular libros de Excel con xlwings."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import xlwings as xw

from ._excepciones import (
    ExcepcionConsultaPowerQuery,
    ExcepcionTablaDinamicaNoEncontrada,
    ExcepcionTablaNoEncontrada,
)

log = logging.getLogger(__name__)


class ClienteExcel:
    """Abre un libro de Excel y ejecuta operaciones comunes de automatizacion.

    Es context manager: al salir del `with` cierra el libro sin guardar
    (guardar es explicito con `guardar()`); si quedan cambios sin guardar,
    se avisa con un warning en el log. Con excepcion dentro del bloque,
    nunca se guarda.
    """

    def __init__(self, ruta_archivo: str | Path | None = None, *, visible: bool = False) -> None:
        self.ruta_archivo = Path(ruta_archivo) if ruta_archivo is not None else None
        self.visible = visible
        self.aplicacion: xw.App | None = None
        self.libro: Any | None = None
        self._ruta_abierta: Path | None = None
        self._modificado = False

    def __enter__(self) -> ClienteExcel:
        self.abrir()
        return self

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        if exc_type is None and self._modificado:
            log.warning("El libro se cierra con cambios sin guardar: se descartan.")
        self.cerrar(guardar=False)

    def abrir(self, ruta_archivo: str | Path | None = None) -> None:
        """Abre el libro en una instancia propia de Excel.

        Notes:
            Usa una instancia dedicada de Excel (no la del usuario) con
            avisos y actualizacion de pantalla apagados, autoguardado
            desactivado y consultas OLEDB en modo sincronico.

        Raises:
            ValueError: si no hay ruta ni en el argumento ni en el constructor.
            RuntimeError: si ya hay otro libro abierto (cerralo primero).
        """
        ruta = Path(ruta_archivo) if ruta_archivo is not None else self.ruta_archivo

        if ruta is None:
            raise ValueError(
                "abrir() necesita una ruta: pasala por argumento o en el constructor (ruta_archivo)."
            )

        if self.aplicacion is None:
            self.aplicacion = xw.App(visible=self.visible, add_book=False)
            self.aplicacion.display_alerts = False
            self.aplicacion.screen_updating = False

        if self.libro is not None:
            if self._ruta_abierta is not None and ruta.resolve() == self._ruta_abierta:
                return
            raise RuntimeError(
                f"Ya hay un libro abierto ({self._ruta_abierta}); "
                "cerralo con cerrar() antes de abrir otro."
            )

        log.info("Abriendo libro: %s", ruta)
        self.libro = self.aplicacion.books.open(ruta)
        self.ruta_archivo = ruta
        self._ruta_abierta = ruta.resolve()

        # AutoSaveOn solo existe en libros en la nube (OneDrive/SharePoint):
        # en un libro local el hasattr pasa pero asignar lanza com_error.
        if hasattr(self.libro.api, "AutoSaveOn"):
            try:
                self.libro.api.AutoSaveOn = False
            except Exception:
                log.debug("No se pudo desactivar AutoSaveOn (libro local).")

        if hasattr(self.libro.api, "ConflictResolution"):
            self.libro.api.ConflictResolution = 2

        for i in range(1, self.libro.api.Connections.Count + 1):
            conexion = self.libro.api.Connections.Item(i)
            if hasattr(conexion, "OLEDBConnection") and conexion.OLEDBConnection is not None:
                conexion.OLEDBConnection.BackgroundQuery = False

    def guardar(self) -> None:
        """Guarda el libro abierto."""
        libro = self._requerir_libro()
        libro.save()
        self._modificado = False
        log.info("Libro guardado.")

    def cerrar(self, *, guardar: bool = False) -> None:
        """Cierra el libro, guardando primero solo si `guardar` es True."""
        if self.libro is not None:
            if guardar:
                self.libro.save()
            self.libro.close()
            self.libro = None
            self._ruta_abierta = None
            self._modificado = False

    def salir(self, *, guardar: bool = False) -> None:
        """Cierra el libro y termina la instancia de Excel."""
        self.cerrar(guardar=guardar)

        if self.aplicacion is not None:
            self.aplicacion.quit()

        self.aplicacion = None

    def configurar_libro(self) -> None:
        """Aplica la configuracion segura al libro abierto.

        Notes:
            Hoy la apertura en `abrir()` ya deja el libro configurado
            (sin autoguardado ni consultas en segundo plano); este metodo
            existe como punto explicito del flujo del consumidor.
        """
        self._requerir_libro()
        log.debug("Libro configurado.")

    def actualizar_tabla_por_nombre(self, nombre_tabla: str) -> None:
        """Actualiza una tabla de Excel por nombre y espera a las consultas asincronas.

        Raises:
            ExcepcionTablaNoEncontrada: si ninguna hoja tiene una tabla con ese nombre.
        """
        libro = self._requerir_libro()

        for hoja in libro.api.Worksheets:
            for i in range(1, hoja.ListObjects.Count + 1):
                tabla = hoja.ListObjects(i)
                if tabla.Name != nombre_tabla:
                    continue

                # QueryTable solo existe en tablas vinculadas a datos externos:
                # en una tabla plana el acceso lanza com_error, no retorna None.
                try:
                    query_table = tabla.QueryTable
                except Exception:
                    query_table = None
                if query_table is not None:
                    query_table.Refresh(False)

                libro.app.api.CalculateUntilAsyncQueriesDone()
                self._modificado = True
                log.info("Tabla actualizada: %s", nombre_tabla)
                return

        raise ExcepcionTablaNoEncontrada(f"No se encontro la tabla '{nombre_tabla}'.")

    def actualizar_tabla_dinamica_por_nombre(self, nombre_tabla_dinamica: str) -> None:
        """Actualiza una tabla dinamica por nombre y espera a las consultas asincronas.

        Raises:
            ExcepcionTablaDinamicaNoEncontrada: si ninguna hoja tiene una
                tabla dinamica con ese nombre.
        """
        libro = self._requerir_libro()

        for hoja in libro.api.Worksheets:
            if hoja.PivotTables().Count == 0:
                continue

            for i in range(1, hoja.PivotTables().Count + 1):
                tabla_dinamica = hoja.PivotTables(i)
                if tabla_dinamica.Name != nombre_tabla_dinamica:
                    continue

                tabla_dinamica.RefreshTable()
                libro.app.api.CalculateUntilAsyncQueriesDone()
                self._modificado = True
                log.info("Tabla dinamica actualizada: %s", nombre_tabla_dinamica)
                return

        raise ExcepcionTablaDinamicaNoEncontrada(
            f"No se encontro la tabla dinamica '{nombre_tabla_dinamica}'."
        )

    def actualizar_formula_power_query(self, nombre_consulta: str, formula: str) -> None:
        """Actualiza la formula M de una consulta de Power Query por nombre.

        Raises:
            ExcepcionConsultaPowerQuery: si el libro rechaza la formula.
        """
        libro = self._requerir_libro()

        try:
            libro.api.Queries(nombre_consulta).Formula = formula
        except Exception as exc:
            raise ExcepcionConsultaPowerQuery(
                f"No se pudo actualizar la consulta de Power Query '{nombre_consulta}'."
            ) from exc
        self._modificado = True
        log.info("Consulta Power Query actualizada: %s", nombre_consulta)

    def definir_formula_sql(
        self,
        nombre_consulta: str,
        *,
        servidor: str,
        base_datos: str,
        consulta_sql: str,
    ) -> None:
        """Construye y asigna una formula M usando Sql.Database para una consulta.

        Raises:
            ExcepcionConsultaPowerQuery: si el libro rechaza la formula.
        """
        servidor_escapado = self._escapar_texto_m(servidor)
        base_datos_escapada = self._escapar_texto_m(base_datos)
        consulta_sql_escapada = self._escapar_texto_m(consulta_sql)

        formula = (
            "let\r\n"
            f'    Origen = Sql.Database("{servidor_escapado}", "{base_datos_escapada}", '
            f'[Query="{consulta_sql_escapada}"])\r\n'
            "in\r\n"
            "    Origen"
        )
        self.actualizar_formula_power_query(nombre_consulta, formula)

    def _requerir_libro(self) -> Any:
        if self.libro is None:
            raise RuntimeError("Debes abrir el libro primero con abrir().")
        return self.libro

    @staticmethod
    def _escapar_texto_m(valor: str) -> str:
        """Escapa texto para incrustarlo como literal string dentro de una formula M."""
        escapado = valor.replace('"', '""')
        escapado = escapado.replace("\r\n", "#(lf)").replace("\n", "#(lf)").replace("\r", "#(lf)")
        return escapado
