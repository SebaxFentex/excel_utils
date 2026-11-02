"""Test local de actualizacion de tablas y tablas dinamicas.

Crea el libro en tests/archivos/descarga/ (se limpia antes de cada corrida),
lo actualiza por nombre y verifica los codigos 2 y 3 con nombres inexistentes.
La tabla dinamica positiva (RefreshTable real) queda pendiente de un libro
fixture con dinamica: el COM para crearla es fragil entre versiones de Excel.

Uso: uv run python tests/test_actualizar_tablas.py
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

import xlwings as xw

from excel import (
    ClienteExcel,
    ExcepcionExcel,
    ExcepcionTablaDinamicaNoEncontrada,
    ExcepcionTablaNoEncontrada,
)
from excel._log import configurar_log

log = logging.getLogger("excel.tests")

CARPETA_DESCARGA = Path(__file__).resolve().parent / "archivos" / "descarga"
RUTA_LIBRO = CARPETA_DESCARGA / "tablas.xlsx"


def _limpiar_descarga() -> None:
    if CARPETA_DESCARGA.exists():
        shutil.rmtree(CARPETA_DESCARGA)
    CARPETA_DESCARGA.mkdir(parents=True, exist_ok=True)


def _crear_libro_vacio(ruta: Path) -> None:
    app = xw.App(visible=False)
    try:
        libro = app.books.add()
        libro.save(str(ruta))
        libro.close()
    finally:
        app.quit()


def _crear_libro_con_tabla(ruta: Path) -> None:
    _crear_libro_vacio(ruta)
    cliente = ClienteExcel(ruta)
    try:
        cliente.abrir()
        libro = cliente.libro
        assert libro is not None, "el libro no quedo abierto"
        hoja = libro.sheets[0]
        hoja["A1"].value = [["Nombre", "Monto"], ["Ana", 10], ["Luis", 20]]
        tabla = hoja.api.ListObjects.Add(1, hoja["A1:B3"].api, None, 1)
        tabla.Name = "Tab_prueba"
        cliente.guardar()
    finally:
        cliente.salir()
    log.info("OK fixture creado: libro con Tab_prueba")


def _probar_actualizar_tabla() -> None:
    cliente = ClienteExcel(RUTA_LIBRO)
    try:
        cliente.abrir()
        cliente.actualizar_tabla_por_nombre("Tab_prueba")
        log.info("OK actualizar_tabla_por_nombre con tabla existente")

        try:
            cliente.actualizar_tabla_por_nombre("Tab_inexistente")
        except ExcepcionTablaNoEncontrada as exc:
            assert exc.codigo_salida == 2, f"codigo de salida {exc.codigo_salida}"
            log.info("OK tabla inexistente traducida a codigo 2")
        else:
            raise AssertionError("una tabla inexistente deberia lanzar ExcepcionTablaNoEncontrada")

        try:
            cliente.actualizar_tabla_dinamica_por_nombre("Din_inexistente")
        except ExcepcionTablaDinamicaNoEncontrada as exc:
            assert exc.codigo_salida == 3, f"codigo de salida {exc.codigo_salida}"
            log.info("OK dinamica inexistente traducida a codigo 3")
        else:
            raise AssertionError(
                "una dinamica inexistente deberia lanzar ExcepcionTablaDinamicaNoEncontrada"
            )
    finally:
        cliente.salir()


def main() -> None:
    configurar_log()
    log.info("Test de actualizacion de tablas")

    _limpiar_descarga()
    _crear_libro_con_tabla(RUTA_LIBRO)
    _probar_actualizar_tabla()

    log.info("OK test_actualizar_tablas completo")


if __name__ == "__main__":
    try:
        main()
    except ExcepcionExcel as exc:
        raise SystemExit(exc.codigo_salida) from exc
