"""Modelos de datos compartidos por las capas de Expedientes ETL."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Record:
    """Registro normalizado que llega a la salida principal."""

    codigo: str
    parte: str
    clientes: str = ""
