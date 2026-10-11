"""Lectores de archivos de entrada de Expedientes ETL."""

from expedientes_etl import (
    convert_legacy_doc,
    process_excel,
    process_word,
)

__all__ = ["convert_legacy_doc", "process_excel", "process_word"]
