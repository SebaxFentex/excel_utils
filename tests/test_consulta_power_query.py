"""Test local de formulas M de Power Query, sin servidor real.

Verifica el escapado de texto M como funcion pura, que definir_formula_sql()
arma la formula Sql.Database esperada (capturando la llamada, sin Excel) y que
una consulta inexistente en un libro real se traduce a codigo 4.

Uso: uv run python tests/test_consulta_power_query.py
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from unittest import mock

import xlwings as xw

from excel import ClienteExcel, ExcepcionConsultaPowerQuery, ExcepcionExcel
from excel._log import configurar_log

log = logging.getLogger("excel.tests")

CARPETA_DESCARGA = Path(__file__).resolve().parent / "archivos" / "descarga"
RUTA_LIBRO = CARPETA_DESCARGA / "consultas.xlsx"


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


def _probar_escapado() -> None:
    casos = (
        ("texto plano", "texto plano"),
        ('con "comillas"', 'con ""comillas""'),
        ("linea1\nlinea2", "linea1#(lf)linea2"),
        ("retorno\r\nwindows", "retorno#(lf)windows"),
        ("retorno\rmac", "retorno#(lf)mac"),
        ("C:\\nuevo\\dato", "C:\\nuevo\\dato"),
    )
    for entrada, esperado in casos:
        obtenido = ClienteExcel._escapar_texto_m(entrada)
        assert obtenido == esperado, f"escapado de {entrada!r}: {obtenido!r}"
    log.info("OK escapado de texto M (comillas y saltos)")


def _probar_formula_sql() -> None:
    cliente = ClienteExcel("nada.xlsx")
    capturada: dict[str, str] = {}

    def _capturar(nombre_consulta: str, formula: str) -> None:
        capturada["nombre"] = nombre_consulta
        capturada["formula"] = formula

    with mock.patch.object(
        cliente, "actualizar_formula_power_query", side_effect=_capturar
    ):
        cliente.definir_formula_sql(
            "Tab_ventas",
            servidor='SRV"01',
            base_datos="BASE",
            consulta_sql="SELECT *\nFROM dbo.ventas",
        )

    assert capturada["nombre"] == "Tab_ventas", f"nombre: {capturada}"
    esperada = (
        "let\r\n"
        '    Origen = Sql.Database("SRV""01", "BASE", '
        '[Query="SELECT *#(lf)FROM dbo.ventas"])\r\n'
        "in\r\n"
        "    Origen"
    )
    assert capturada["formula"] == esperada, f"formula: {capturada['formula']!r}"
    log.info("OK definir_formula_sql arma la formula Sql.Database escapada")


def _probar_consulta_inexistente() -> None:
    _crear_libro_vacio(RUTA_LIBRO)
    cliente = ClienteExcel(RUTA_LIBRO)
    try:
        cliente.abrir()
        libro = cliente.libro
        assert libro is not None, "el libro no quedo abierto"
        libro.sheets[0]["A1"].value = "x"
        cliente.guardar()
        try:
            cliente.actualizar_formula_power_query("Consulta_inexistente", "let Origen = 1 in Origen")
        except ExcepcionConsultaPowerQuery as exc:
            assert exc.codigo_salida == 4, f"codigo de salida {exc.codigo_salida}"
            log.info("OK consulta inexistente traducida a codigo 4")
        else:
            raise AssertionError(
                "una consulta inexistente deberia lanzar ExcepcionConsultaPowerQuery"
            )
    finally:
        cliente.salir()


def main() -> None:
    configurar_log()
    log.info("Test de formulas M de Power Query")

    _limpiar_descarga()
    _probar_escapado()
    _probar_formula_sql()
    _probar_consulta_inexistente()

    log.info("OK test_consulta_power_query completo")


if __name__ == "__main__":
    try:
        main()
    except ExcepcionExcel as exc:
        raise SystemExit(exc.codigo_salida) from exc
