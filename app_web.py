"""Interfaz web local para utilizar Expedientes ETL desde el navegador."""

from __future__ import annotations

import io
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st

PROJECT_DIR = Path(__file__).resolve().parent
TMP_ROOT = PROJECT_DIR / "tmp" / "web"

if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import expedientes_etl as etl  # noqa: E402


def _save_upload(uploaded_file, work_dir: Path) -> Path:
    source_dir = work_dir / "entrada"
    source_dir.mkdir(parents=True, exist_ok=True)
    destination = source_dir / Path(uploaded_file.name).name
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def _process_file(source: Path, work_dir: Path) -> dict[str, object]:
    output_dir = work_dir / "salida"
    quality_dir = output_dir / "control_calidad"
    result = etl.process_file_to_outputs(source, output_dir, quality_dir)
    return {
        "nombre": source.name,
        "salida": result["destination"],
        "control": result["quality_path"],
        "registros": result["records"],
        "revision": result["review"],
        "conflictos": result["conflicts"],
        "alertas": result["alerts"],
    }


def _zip_results(results: list[dict[str, object]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            for key in ("salida", "control"):
                path = Path(result[key])
                archive.write(path, Path("resultados") / path.name)
    return buffer.getvalue()


st.set_page_config(page_title="Expedientes ETL", page_icon="📄", layout="wide")
st.title("Expedientes ETL")
st.write("Procesa archivos Excel y Word localmente y genera resultados simplificados.")
st.info("Los archivos se procesan en esta computadora; no se envían a un servicio externo.")

uploaded_files = st.file_uploader(
    "Seleccione uno o varios archivos",
    type=["xlsx", "docx", "doc"],
    accept_multiple_files=True,
    help="Se aceptan archivos Excel, Word moderno y Word antiguo.",
)

if st.button("Procesar archivos", type="primary", disabled=not uploaded_files):
    session_name = datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = TMP_ROOT / session_name
    if work_dir.exists():
        shutil.rmtree(work_dir)
    results: list[dict[str, object]] = []
    progress = st.progress(0, text="Iniciando procesamiento...")
    for index, uploaded_file in enumerate(uploaded_files, start=1):
        try:
            source = _save_upload(uploaded_file, work_dir)
            result = _process_file(source, work_dir)
            results.append(result)
            st.success(f"Procesado: {uploaded_file.name}")
        except (OSError, RuntimeError, ValueError) as exc:
            st.error(f"No se pudo procesar {uploaded_file.name}: {exc}")
        progress.progress(index / len(uploaded_files), text=f"Procesando {index} de {len(uploaded_files)}")

    if results:
        st.subheader("Resumen del procesamiento")
        st.dataframe(
            [
                {
                    "archivo": result["nombre"],
                    "registros": result["registros"],
                    "revisión manual": result["revision"],
                    "conflictos": result["conflictos"],
                    "alertas": result["alertas"],
                }
                for result in results
            ],
            use_container_width=True,
            hide_index=True,
        )
        st.subheader("Descargas")
        for result in results:
            output_path = Path(result["salida"])
            st.download_button(
                f"Descargar {output_path.name}",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"output-{output_path.name}",
            )
        st.download_button(
            "Descargar resultados y control de calidad en ZIP",
            data=_zip_results(results),
            file_name=f"expedientes_etl_resultados_{session_name}.zip",
            mime="application/zip",
            key=f"zip-{session_name}",
        )
        st.caption("Los archivos de esta sesión se conservan temporalmente en tmp/web.")
