"""Test local del ciclo de vida del libro, sin configuracion externa.

Verifica abrir/guardar/cerrar/salir, el fallo rapido de D5 (sin ruta,
doble libro) y el warning al salir del `with` con cambios sin guardar.

Uso: uv run python tests/test_abrir.py
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path

import xlwings as xw

from excel_utils import ClienteExcel, ExcepcionExcel
from excel_utils._log import configurar_log

log = logging.getLogger("excel_utils.tests")

CARPETA_ARCHIVOS = Path(__file__).resolve().parent / "archivos"
RUTA_LIBRO = CARPETA_ARCHIVOS / "ciclo.xlsx"


def _limpiar_archivos() -> None:
    if CARPETA_ARCHIVOS.exists():
        shutil.rmtree(CARPETA_ARCHIVOS)
    CARPETA_ARCHIVOS.mkdir(parents=True, exist_ok=True)


def _crear_libro_vacio(ruta: Path) -> None:
    app = xw.App(visible=False)
    try:
        libro = app.books.add()
        libro.save(str(ruta))
        libro.close()
    finally:
        app.quit()


def _probar_sin_ruta() -> None:
    cliente = ClienteExcel()
    try:
        cliente.abrir()
    except ValueError:
        log.info("OK abrir() sin ruta lanza ValueError")
    else:
        raise AssertionError("abrir() sin ruta deberia lanzar ValueError")
    assert cliente.aplicacion is None, "abrir() sin ruta creo la instancia de Excel"
    log.info("OK abrir() sin ruta no crea la instancia de Excel")


def _probar_ciclo() -> None:
    _crear_libro_vacio(RUTA_LIBRO)
    cliente = ClienteExcel(RUTA_LIBRO)
    try:
        cliente.abrir()
        libro = cliente.libro
        assert libro is not None, "el libro no quedo abierto"
        libro.sheets[0]["A1"].value = "hola"
        cliente.guardar()
        assert RUTA_LIBRO.exists(), "guardar() no dejo el archivo en disco"
        log.info("OK abrir + escribir + guardar")

        cliente.cerrar()
        assert cliente.libro is None, "cerrar() no solto el libro"

        cliente.abrir()
        libro = cliente.libro
        assert libro is not None, "el libro no quedo abierto"
        assert libro.sheets[0]["A1"].value == "hola", "el valor no persistio"
        log.info("OK reabrir: el valor persistio")

        cliente.abrir(RUTA_LIBRO)
        log.info("OK abrir() la misma ruta es idempotente")

        try:
            cliente.abrir(CARPETA_ARCHIVOS / "otro.xlsx")
        except RuntimeError:
            log.info("OK abrir() otro libro con uno abierto lanza RuntimeError")
        else:
            raise AssertionError("abrir() un segundo libro deberia lanzar RuntimeError")

        cliente.cerrar()
        cliente.cerrar()
        log.info("OK cerrar() idempotente")
    finally:
        cliente.salir()


def _probar_salir_idempotente() -> None:
    cliente = ClienteExcel(RUTA_LIBRO)
    try:
        cliente.abrir()
    finally:
        cliente.salir()
    cliente.salir()
    assert cliente.aplicacion is None, "salir() no solto la aplicacion"
    log.info("OK salir() idempotente")


def _probar_warning_sin_guardar() -> None:
    registros: list[tuple[str, str]] = []

    class Manejador(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            registros.append((record.levelname, record.getMessage()))

    manejador = Manejador()
    logging.getLogger("excel_utils").addHandler(manejador)
    try:
        cliente = ClienteExcel(RUTA_LIBRO)
        try:
            cliente.abrir()
            cliente._modificado = True
            cliente.__exit__(None, None, None)
        finally:
            cliente.salir()
    finally:
        logging.getLogger("excel_utils").removeHandler(manejador)
    assert any(n == "WARNING" and "sin guardar" in m for n, m in registros), (
        f"falto el warning: {registros}"
    )
    log.info("OK __exit__ avisa con warning si hay cambios sin guardar")

    registros.clear()
    manejador2 = Manejador()
    logging.getLogger("excel_utils").addHandler(manejador2)
    try:
        cliente2 = ClienteExcel(RUTA_LIBRO)
        try:
            cliente2.abrir()
            cliente2._modificado = True
            cliente2.__exit__(RuntimeError, RuntimeError("x"), None)
        finally:
            cliente2.salir()
    finally:
        logging.getLogger("excel_utils").removeHandler(manejador2)
    assert not any(n == "WARNING" for n, _ in registros), (
        f"con excepcion no debe avisar: {registros}"
    )
    log.info("OK __exit__ con excepcion no avisa ni guarda")


def main() -> None:
    configurar_log()
    log.info("Test del ciclo de vida del libro")

    _limpiar_archivos()
    _probar_sin_ruta()
    _probar_ciclo()
    _probar_salir_idempotente()
    _probar_warning_sin_guardar()

    log.info("OK test_abrir completo")


if __name__ == "__main__":
    try:
        main()
    except ExcepcionExcel as exc:
        raise SystemExit(exc.codigo_salida) from exc
