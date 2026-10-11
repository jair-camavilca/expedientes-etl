"""Interfaz web local para utilizar Expedientes ETL desde el navegador."""

from __future__ import annotations

import io
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import streamlit as st

PROJECT_DIR = Path(__file__).resolve().parent
TMP_ROOT = PROJECT_DIR / "tmp" / "web"
SESSION_MAX_AGE_SECONDS = 2 * 60 * 60
MAX_UPLOAD_SIZE_BYTES = 50 * 1024 * 1024
ALLOWED_UPLOAD_SUFFIXES = {".xlsx", ".docx"}

if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from core.servicio import process_file_to_outputs  # noqa: E402


def _save_upload(uploaded_file, work_dir: Path) -> Path:
    original_name = Path(uploaded_file.name).name
    suffix = Path(original_name).suffix.lower()
    if suffix not in ALLOWED_UPLOAD_SUFFIXES:
        raise ValueError("solo se admiten archivos .xlsx y .docx en la versión web")
    if uploaded_file.size > MAX_UPLOAD_SIZE_BYTES:
        raise ValueError("el archivo supera el límite permitido de 50 MB")

    source_dir = work_dir / "entrada"
    source_dir.mkdir(parents=True, exist_ok=True)
    destination = source_dir / original_name
    if destination.exists():
        destination = source_dir / f"{Path(original_name).stem}_{uuid4().hex[:8]}{Path(original_name).suffix}"
    destination.write_bytes(uploaded_file.getbuffer())
    return destination


def _remove_expired_sessions() -> None:
    """Elimina sesiones temporales antiguas de la aplicación web."""
    if not TMP_ROOT.exists():
        return
    now = datetime.now().timestamp()
    for session_dir in TMP_ROOT.iterdir():
        if not session_dir.is_dir():
            continue
        try:
            if now - session_dir.stat().st_mtime > SESSION_MAX_AGE_SECONDS:
                shutil.rmtree(session_dir)
        except OSError:
            continue


def _process_file(source: Path, work_dir: Path) -> dict[str, object]:
    output_dir = work_dir / "salida"
    quality_dir = output_dir / "control_calidad"
    result = process_file_to_outputs(source, output_dir, quality_dir)
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


st.set_page_config(page_title="Expedientes ETL", page_icon=":material/description:", layout="wide")
st.markdown(
    """
    <style>
    .block-container {{
        max-width: 1100px;
        padding-top: 1.25rem;
        padding-bottom: 0.75rem;
    }}
    h1 {{
        letter-spacing: -0.04em;
        margin-bottom: 0.15rem;
        font-size: 2.15rem;
    }}
    h2, h3 {{
        margin-top: 0.65rem;
        margin-bottom: 0.35rem;
    }}
    .stMarkdown p {{
        margin-bottom: 0.35rem;
    }}
    div[data-testid="stAlert"] {{
        padding: 0.55rem 0.85rem;
        margin-bottom: 0.55rem;
    }}
    div[data-testid="stFileUploader"] {{
        padding-top: 0.1rem;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        border-radius: 0.7rem;
    }}
    div[data-testid="stAlert"] {{
        border-radius: 0.6rem;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("Expedientes ETL")
st.write("Carga archivos Excel o Word y obtén un resultado ordenado con los datos principales del expediente.")
st.caption("Carga únicamente archivos autorizados. Se procesan temporalmente y las sesiones antiguas se eliminan automáticamente.")
_remove_expired_sessions()

st.subheader("Seleccionar archivos")
st.caption("Excel (.xlsx) o Word (.docx); puedes seleccionar uno o varios archivos.")
uploaded_files = st.file_uploader(
    "Archivos de entrada",
    type=["xlsx", "docx"],
    accept_multiple_files=True,
    help="La versión web acepta Excel (.xlsx) y Word moderno (.docx). Para .doc antiguo use el flujo local de Windows.",
)
process_clicked = st.button("Procesar archivos", type="primary", disabled=not uploaded_files)

if process_clicked:
    session_name = f"{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:8]}"
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
        st.subheader("Resultados y descargas")
        st.success(f"Se procesaron {len(results)} archivo(s) correctamente.")
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
            hide_index=True,
        )
        st.subheader("Descargas")
        download_column, zip_column = st.columns(2)
        with download_column:
            for result in results:
                output_path = Path(result["salida"])
                st.download_button(
                    f"Descargar {output_path.name}",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"output-{output_path.name}",
                )
        with zip_column:
            st.download_button(
                "Descargar paquete ZIP",
                data=_zip_results(results),
                file_name=f"expedientes_etl_resultados_{session_name}.zip",
                mime="application/zip",
                key=f"zip-{session_name}",
                help="Incluye los resultados y los reportes de control de calidad.",
            )
        st.caption("Los archivos de esta sesión se conservan temporalmente en tmp/web.")
