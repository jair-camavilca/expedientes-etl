"""Punto de acceso a las reglas de normalización de datos.

Las funciones se mantienen detrás de este módulo para que lectores,
servicios y futuras interfaces no dependan de la estructura interna del
motor legado.
"""

from expedientes_etl import (
    _format_person_name,
    _person_without_comma,
    _first_surname_words,
    choose_party,
    extract_surnames,
    is_legal_entity,
    normalize_code,
    normalize_text,
)

__all__ = [
    "choose_party",
    "extract_surnames",
    "is_legal_entity",
    "normalize_code",
    "normalize_text",
]
