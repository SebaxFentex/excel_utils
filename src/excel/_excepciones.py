"""Excepciones propias de la libreria excel.

El codigo de salida viaja como atributo de clase para que el proyecto
consumidor haga ``sys.exit(exc.codigo_salida)`` sin repetir el numero.

Tabla de codigos de la libreria:

    0  ejecucion correcta (universal)
    1  fallo generico (convencion Python)
    2  tabla de Excel no encontrada en el libro
    3  tabla dinamica no encontrada en el libro
    4  no se pudo actualizar la consulta de Power Query

Los datos de entrada malformados llegan como ``TypeError`` (tipo incorrecto) o
``ValueError`` (texto vacio), sin envolverse.
"""


class ExcepcionExcel(Exception):
    """Base de los errores propios de la libreria (fallo generico)."""

    codigo_salida = 1


class ExcepcionTablaNoEncontrada(ExcepcionExcel):
    """Ninguna hoja del libro tiene una tabla (ListObject) con ese nombre."""

    codigo_salida = 2


class ExcepcionTablaDinamicaNoEncontrada(ExcepcionExcel):
    """Ninguna hoja del libro tiene una tabla dinamica con ese nombre."""

    codigo_salida = 3


class ExcepcionConsultaPowerQuery(ExcepcionExcel):
    """El libro rechazo la formula M de la consulta de Power Query."""

    codigo_salida = 4
