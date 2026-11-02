"""Automatizacion de libros de Excel y consultas Power Query con xlwings.

La libreria emite logs y el entrypoint los configura: aqui solo se agrega un
``NullHandler`` para no ensuciar la salida de quien importa.

La instancia de Excel y las rutas entran por parametro.
"""

import logging

from ._excepciones import (
    ExcepcionConsultaPowerQuery,
    ExcepcionExcel,
    ExcepcionTablaDinamicaNoEncontrada,
    ExcepcionTablaNoEncontrada,
)
from ._log import configurar_log, purgar_logs
from .cliente import ClienteExcel

logging.getLogger(__name__).addHandler(logging.NullHandler())

__all__ = [
    "ClienteExcel",
    "ExcepcionConsultaPowerQuery",
    "ExcepcionExcel",
    "ExcepcionTablaDinamicaNoEncontrada",
    "ExcepcionTablaNoEncontrada",
    "configurar_log",
    "purgar_logs",
]
