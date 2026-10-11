"""Servicio de aplicación compartido por CLI, BAT e interfaz web."""

from pathlib import Path

from expedientes_etl import (
    process_file_to_outputs as _process_file_to_outputs,
    process_single_file,
)


def process_file_to_outputs(
    input_path: Path,
    output_dir: Path,
    quality_dir: Path,
    code_column: str | None = None,
    party_column: str | None = None,
) -> dict[str, object]:
    """Procesa una fuente y devuelve resultados y controles generados."""
    return _process_file_to_outputs(
        input_path,
        output_dir,
        quality_dir,
        code_column,
        party_column,
    )


__all__ = ["process_file_to_outputs", "process_single_file"]
