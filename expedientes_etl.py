"""ETL de registros judiciales para Expedientes ETL.

Admite fuentes Excel (.xlsx) y Word (.docx), y genera exactamente las
columnas ``codigo`` y ``parte`` en un libro Excel formateado.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator, Sequence

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet


COURT_CODE_RE = re.compile(
    r"(?<!\d)(?P<code>\d{1,6}\s*-\s*\d{2,4}\s*-\s*\d+(?:/\d+)*\s*-\s*[A-Z0-9]{1,5}\s*-\s*[A-Z]{1,5}\s*-\s*[A-Z]{1,5}\s*-\s*\d{1,4})(?!\d)",
    re.IGNORECASE,
)
LOOSE_COURT_CODE_RE = re.compile(
    r"(?<!\d)\d{1,6}\s*-\s*\d{2,4}\s*-\s*[\d/]+\s*-\s*[A-Z0-9]{1,5}\s*-\s*[A-Z]{1,5}\s*-\s*[A-Z]{1,5}\s*-\s*\d{1,4}(?!\d)",
    re.IGNORECASE,
)
PARTIAL_CASE_RE = re.compile(
    r"\bexp(?:ediente)?\.?\s*(?:n(?:ro|úm|°|º)?\.?)?\s*"
    r"(?P<code>\d{1,6}\s*-\s*\d{2,4}(?:\s*-\s*\d+)?)",
    re.IGNORECASE,
)
EMPTY_EXP_LABEL_RE = re.compile(
    r"(?im)^\s*exp(?:ediente)?\.?\s*n(?:ro|úm|°|º|�)?\.?\s*$"
)
CASE_LABEL_RE = re.compile(r"\b(?:exp(?:ediente)?\.?\s*(?:n(?:ro|úm|°|º)?\.?\s*)?)?", re.IGNORECASE)
NUMBERED_BLOCK_RE = re.compile(r"(?m)^\s*(?P<number>\d+)\s*[.)]\s*(?P<body>.*?)(?=^\s*\d+\s*[.)]\s*|\Z)", re.DOTALL)
PAREN_RE = re.compile(r"\([^()]*\)")
VERSUS_RE = re.compile(r"\s+(?:vs?\.?|contra|c\/|en\s+contra\s+de)\s+", re.IGNORECASE)
PARTY_CONNECTOR_RE = re.compile(
    r"\s+(?:contra|c\/|con|por|ante|frente\s+a|y\s+otros?|y\s+dem[aá]s)\b",
    re.IGNORECASE,
)
SUCCESSION_RE = re.compile(r"^\s*(?:sucesi[oó]n|herencia|masa\s+hereditaria)\b", re.IGNORECASE)
LEGAL_ENTITY_RE = re.compile(
    r"\b(?:S\.?\s*A\.?\s*C?\.?|S\.?\s*R\.?\s*L\.?|E\.?\s*I\.?\s*R\.?\s*L\.?|S\.?\s*A\.?\s*A?\.?|SBN|SUNARP|INDECOPI|IPAE|COPAMO|INTERSEGURO|MINISTERIO|MUNICIPALIDAD|PODER\s+JUDICIAL|GOBIERNO\s+REGIONAL|SUPERINTENDENCIA|ASOCIACI(?:O|Ó)N(?:ES)?|ASOC\.?|CLUB|COOPERATIVA|COOPERATIVAS|EMPRESA|SOCIEDAD|COMUNIDAD|COMIT[EÉ]|JUNTA|UNIVERSIDAD|INSTITUTO|BANCO|CAJA|FUNDACI(?:O|Ó)N|SINDICATO|COLEGIO\s+PROFESIONAL|IGLESIA|CONGREGACI(?:O|Ó)N|INMOBILIARIA|INMOBILIARIO|CONDOMINIO|CONSTRUCTORA|PROMOTORA|ARRENDADORA|INVERSIONES|AGR[IÍ]COLA|SUCESORES?)\b",
    re.IGNORECASE,
)
SURNAME_PARTICLES = {"de", "del", "la", "las", "los", "y", "da", "do", "van", "von"}
CIVIL_STATUS_SUFFIX_RE = re.compile(
    r"\s+(?:vda?\.?|viud[ao]|q\.?e\.?(?:p\.?(?:d\.?)?)?)"
    r"(?:\s+de\s+[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’-]+)?$",
    re.IGNORECASE,
)
SCRIPT_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = SCRIPT_DIR
DEFAULT_OUTPUT = WORKSPACE_DIR / "salida" / "expedientes_etl_resultado.xlsx"
QUALITY_DIR = WORKSPACE_DIR / "notas"
INPUT_DIRS = (WORKSPACE_DIR / "entrada",)
ENTITY_CATALOG_PATH = SCRIPT_DIR / "entidades_juridicas.txt"


@dataclass(frozen=True)
class Record:
    codigo: str
    parte: str


def normalize_text(value: object) -> str:
    """Convierte un valor a texto limpio, sin alterar letras acentuadas."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def load_entity_terms(path: Path) -> set[str]:
    """Carga términos jurídicos editables, uno por línea, desde UTF-8."""
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return set()
    return {
        normalize_text(line).casefold()
        for line in lines
        if normalize_text(line) and not normalize_text(line).startswith("#")
    }


ENTITY_CATALOG = load_entity_terms(ENTITY_CATALOG_PATH)
ENTITY_CATALOG_RE = re.compile(
    r"(?<!\w)(?:"
    + "|".join(re.escape(term) for term in sorted(ENTITY_CATALOG, key=len, reverse=True))
    + r")(?!\w)",
    re.IGNORECASE,
) if ENTITY_CATALOG else None


def is_legal_entity(text: str) -> bool:
    """Detecta entidades con reglas internas y catálogo editable."""
    return bool(LEGAL_ENTITY_RE.search(text) or (ENTITY_CATALOG_RE and ENTITY_CATALOG_RE.search(text)))


def normalize_code(value: object) -> str | None:
    """Normaliza un expediente y rellena el correlativo inicial a cinco dígitos."""
    text = normalize_text(value).upper().replace("–", "-").replace("—", "-")
    match = COURT_CODE_RE.search(text)
    if not match:
        return None
    chunks = [re.sub(r"\s+", "", chunk) for chunk in match.group("code").split("-")]
    chunks[0] = chunks[0].zfill(5)
    return "-".join(chunks)


def iter_codes(value: object) -> Iterator[str]:
    """Devuelve todos los códigos judiciales detectados en un valor."""
    text = normalize_text(value).upper().replace("–", "-").replace("—", "-")
    full_matches = list(COURT_CODE_RE.finditer(text))
    for match in full_matches:
        raw_chunks = [re.sub(r"\s+", "", chunk) for chunk in match.group("code").split("-")]
        for subcase in raw_chunks[2].split("/"):
            chunks = [*raw_chunks]
            chunks[2] = subcase
            normalized = normalize_code("-".join(chunks))
            if normalized:
                yield normalized
    if not full_matches:
        for match in PARTIAL_CASE_RE.finditer(text):
            yield normalize_text(match.group("code")).replace(" ", "")


def _strip_accents(value: str) -> str:
    return "".join(
        char for char in unicodedata.normalize("NFD", value)
        if unicodedata.category(char) != "Mn"
    )


def _looks_like_natural_person(text: str) -> bool:
    """Heurística conservadora para distinguir nombres de entidades."""
    if not text or is_legal_entity(text):
        return False
    words = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’-]+", text)
    return len(words) >= 2 and not any(char.isdigit() for char in text)


def _clean_candidate(text: str) -> str:
    text = PAREN_RE.sub(" ", text)
    text = re.sub(r"\b(?:demandante|demandado|actor|actora|emplazado|emplazada)\b\s*:?[ ]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip(" ,;:-")
    return text


def _first_litigant_segment(text: str) -> str:
    """Conserva el primer litigante y elimina el texto de la contraparte."""
    text = VERSUS_RE.split(text, maxsplit=1)[0]
    text = PARTY_CONNECTOR_RE.split(text, maxsplit=1)[0]
    # Una sigla institucional seguida de coma suele preceder una dirección o
    # descripción, no nombres de persona (por ejemplo: ``COPAMO, Playa...``).
    comma_head = text.split(",", 1)[0].strip()
    if re.fullmatch(r"[A-ZÁÉÍÓÚÜÑ0-9 .&'-]+", comma_head) and len(comma_head.split()) <= 2:
        text = comma_head
    return _clean_candidate(text)


def _surname_words(surname_text: str) -> list[str]:
    """Obtiene apellidos sin cortar partículas de apellidos compuestos."""
    surname_text = CIVIL_STATUS_SUFFIX_RE.sub("", surname_text).strip()
    words = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’-]+", surname_text)
    if not words:
        return []
    lowered = [word.lower() for word in words]
    if lowered[0] in SURNAME_PARTICLES:
        # ``De Los Heros Ballen De Van Walleghem``: primer apellido
        # compuesto ``De Los Heros`` + segundo apellido ``Ballen``.
        first_end = 0
        while first_end < len(lowered) and lowered[first_end] in SURNAME_PARTICLES:
            first_end += 1
        if first_end < len(lowered):
            first_end += 1
        second_end = first_end
        if second_end < len(lowered) and lowered[second_end] in SURNAME_PARTICLES:
            while second_end < len(lowered) and lowered[second_end] in SURNAME_PARTICLES:
                second_end += 1
            if second_end < len(lowered):
                second_end += 1
        elif second_end < len(lowered):
            second_end += 1
        words = words[:second_end]
    # En el patrón habitual ``Apellido1 Apellido2 de ApellidoCasada`` se
    # descarta la parte de casada, pero se conserva ``De La Torre`` inicial.
    if lowered[0] not in SURNAME_PARTICLES:
        # ``Rivera de Gonzales`` y ``Gagliuffi de Castagnino`` se conservan
        # completos; un ``de`` posterior a dos apellidos suele introducir
        # apellido de casada y se excluye.
        if len(lowered) > 1 and lowered[1] != "de":
            for index, word in enumerate(lowered[2:], start=2):
                if word == "de":
                    words = words[:index]
                    break
            else:
                words = words[:2]
    while len(words) > 1 and words[-1].lower() in SURNAME_PARTICLES:
        words.pop()
    return words


def extract_surnames(value: object) -> str:
    """Extrae los dos primeros apellidos de una persona natural.

    Se interpreta el formato judicial habitual ``APELLIDOS, Nombres``. En
    formatos sin coma se toma el primer segmento plausible del litigio.
    """
    original = normalize_text(value)
    if not original:
        return ""

    # Se evalúa cada participante y se conserva la primera persona natural.
    candidates = [part.strip() for part in re.split(r";|\n", original) if part.strip()]
    expanded: list[str] = []
    for candidate in candidates:
        expanded.extend(piece.strip() for piece in VERSUS_RE.split(candidate) if piece.strip())
    candidates = expanded or [original]

    for candidate in candidates:
        candidate = _clean_candidate(candidate)
        if not _looks_like_natural_person(candidate):
            continue
        if "," in candidate:
            surname_text = candidate.split(",", 1)[0]
        else:
            words = re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ'’-]+", candidate)
            # Sin coma solo se aceptan dos tokens que no sean partículas.
            if len(words) >= 2 and words[1].lower() in SURNAME_PARTICLES:
                return words[0].title()
            surname_text = " ".join(words[:2])
        words = _surname_words(surname_text)
        if len(words) >= 2:
            return " ".join(words).title()
        if len(words) == 1 and "," in candidate:
            return words[0].title()

    # Si no hay persona natural, se conserva la primera razón social completa,
    # sin arrastrar el texto de la contraparte.
    fallback = _first_litigant_segment(original)
    cleaned_fallback = CIVIL_STATUS_SUFFIX_RE.sub("", fallback).strip()
    return cleaned_fallback.title() if cleaned_fallback != fallback else fallback


def choose_party(value: object) -> str:
    """Selecciona una persona natural; si no existe, devuelve la entidad."""
    text = normalize_text(value)
    if not text:
        return ""
    if SUCCESSION_RE.match(text):
        succession = _first_litigant_segment(text)
        succession = re.sub(r"^\s*sucesion\b", "Sucesión", succession, flags=re.IGNORECASE)
        return succession.title()
    # Los paréntesis suelen contener la razón social secundaria en casos mixtos.
    without_parentheses = PAREN_RE.sub(" ", text)
    natural = extract_surnames(without_parentheses)
    if natural and _looks_like_natural_person(without_parentheses):
        return natural
    return extract_surnames(text)


def _find_column(columns: Sequence[object], aliases: Iterable[str]) -> str | None:
    normalized = {_strip_accents(normalize_text(column)).lower(): str(column) for column in columns}
    for alias in aliases:
        alias_key = _strip_accents(alias).lower()
        for column_key, original in normalized.items():
            if alias_key == column_key or alias_key in column_key:
                return original
    return None


def process_excel(path: Path, code_column: str | None = None, party_column: str | None = None, quality_issues: list[dict[str, str]] | None = None) -> list[Record]:
    """Lee todas las hojas de un Excel y transforma sus filas."""
    sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    records: list[Record] = []
    code_aliases = ("codigo", "código", "expediente", "id expediente", "identificador del expediente", "número de expediente", "nro expediente")
    party_aliases = ("parte", "partes", "litigante", "litigantes", "parte procesal", "demandante", "demandado", "titular")
    for frame in sheets.values():
        if frame.empty:
            continue
        selected_code = code_column or _find_column(frame.columns, code_aliases)
        selected_party = party_column or _find_column(frame.columns, party_aliases)
        if not selected_code:
            raise ValueError("No se encontró una columna de código/expediente en el Excel.")
        if not selected_party:
            raise ValueError("No se encontró una columna de parte/litigante en el Excel.")
        for _, row in frame.iterrows():
            party = choose_party(row.get(selected_party, ""))
            codes = list(iter_codes(row.get(selected_code, "")))
            if not codes and quality_issues is not None:
                quality_issues.append({"codigo": "", "parte": party, "nivel": "ALTA", "motivo": "Fila sin código judicial detectable"})
            for code in codes:
                records.append(Record(code, party))
    return records


def _block_party(body: str) -> str:
    """Obtiene exclusivamente el titular de la primera línea del bloque Word."""
    first_line = next((line.strip() for line in body.splitlines() if line.strip()), "")
    # Los bloques pueden comenzar como ``24.CARLINI...`` o ``24. CARLINI...``.
    first_line = re.sub(r"^\s*\d+\s*[.)]\s*", "", first_line)
    first_line = re.sub(r"^\s*(?:por|con|ddte\.?|demandante)\s*:\s*", "", first_line, flags=re.IGNORECASE)
    # Si el titular es una entidad, una aclaración entre paréntesis puede
    # formar parte del nombre registrado (por ejemplo, ``INVERSIONES
    # PLATINUM (DIEGO FARAH)``). En ese caso no se debe truncar el titular.
    if "(" in first_line and is_legal_entity(first_line) and "," not in first_line:
        return first_line
    return choose_party(first_line)


def _party_before_code(body: str, code_start: int, default: str) -> str:
    """Detecta un nuevo titular de entidad antes de un expediente del bloque."""
    exp_matches = list(re.finditer(r"\bexp(?:ediente)?", body[:code_start], re.IGNORECASE))
    if exp_matches:
        current_label = exp_matches[-1]
        region_start = exp_matches[-2].end() if len(exp_matches) > 1 else 0
        candidate_text = body[region_start : current_label.start()]
    else:
        candidate_text = body[:code_start]
    candidate_lines = candidate_text.splitlines()
    for line in reversed(candidate_lines):
        candidate = line.strip()
        if not candidate or re.match(r"^(?:con|por|dte\.?|ddte\.?|demandante|seguido|seguida|sobre|cas\.?|exp)\b", candidate, re.IGNORECASE):
            continue
        if "(" in candidate and is_legal_entity(candidate) and not re.match(r"^\d", candidate):
            return _block_party(candidate)
    return default


def process_word(path: Path, quality_issues: list[dict[str, str]] | None = None) -> list[Record]:
    """Parsea bloques numerados de un DOCX sin dependencias adicionales.

    Un `.docx` es un contenedor ZIP con XML. Se extraen los párrafos ``w:p``
    en orden, incluyendo los que se encuentren dentro de tablas.
    """
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    try:
        with zipfile.ZipFile(path) as archive:
            document_xml = archive.read("word/document.xml")
    except (KeyError, zipfile.BadZipFile) as exc:
        raise ValueError(f"El archivo no es un DOCX válido: {path.name}") from exc
    root = ET.fromstring(document_xml)
    paragraphs: list[str] = []
    for paragraph in root.iter(f"{namespace}p"):
        paragraph_text = "".join(node.text or "" for node in paragraph.iter(f"{namespace}t"))
        paragraph_text = normalize_text(paragraph_text)
        if paragraph_text:
            paragraphs.append(paragraph_text)
    text = "\n".join(paragraphs)
    records: list[Record] = []
    for match in NUMBERED_BLOCK_RE.finditer(text):
        body = match.group("body")
        party = _block_party(body)
        full_matches = list(COURT_CODE_RE.finditer(body))
        codes = list(iter_codes(body))
        has_empty_exp = bool(EMPTY_EXP_LABEL_RE.search(body))
        if has_empty_exp:
            records.append(Record("NO DEFINIDO", party))
            if quality_issues is not None:
                quality_issues.append(
                    {
                        "codigo": "NO DEFINIDO",
                        "parte": party,
                        "nivel": "ALTA",
                        "motivo": "Etiqueta de expediente sin número; completar desde la fuente",
                    }
                )
        if not codes and not has_empty_exp and quality_issues is not None:
            loose_code = LOOSE_COURT_CODE_RE.search(body)
            partial_code = PARTIAL_CASE_RE.search(body)
            if loose_code:
                issue_code = normalize_text(loose_code.group(0))
                reason = "Código judicial con formato anómalo; revisar segmento o año"
            elif partial_code:
                issue_code = normalize_text(partial_code.group("code"))
                reason = "Expediente incompleto; falta la estructura judicial completa"
            else:
                issue_code = ""
                reason = "Bloque Word sin expediente judicial detectable"
            quality_issues.append({"codigo": issue_code, "parte": party, "nivel": "ALTA", "motivo": reason})
        if full_matches:
            for code_match in full_matches:
                raw_chunks = [re.sub(r"\s+", "", chunk) for chunk in code_match.group("code").split("-")]
                local_party = _party_before_code(body, code_match.start(), party)
                for subcase in raw_chunks[2].split("/"):
                    chunks = [*raw_chunks]
                    chunks[2] = subcase
                    code = normalize_code("-".join(chunks))
                    if code:
                        records.append(Record(code, local_party))
        else:
            for code in codes:
                records.append(Record(code, party))
    return records


def convert_legacy_doc(path: Path) -> Path:
    """Convierte un Word `.doc` a `.docx` usando Microsoft Word instalado."""
    temporary_dir = WORKSPACE_DIR / "salida" / ".conversion_doc"
    temporary_dir.mkdir(parents=True, exist_ok=True)
    target = temporary_dir / f"{path.stem}_{path.stat().st_mtime_ns}.docx"
    source_ps = str(path.resolve()).replace("'", "''")
    target_ps = str(target.resolve()).replace("'", "''")
    powershell_script = f"""
$source = '{source_ps}'
$target = '{target_ps}'
$word = $null
$document = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open($source, $false, $true)
    $document.SaveAs2($target, 16)
}} finally {{
    if ($document -ne $null) {{ $document.Close($false) }}
    if ($word -ne $null) {{ $word.Quit() }}
}}
"""
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", powershell_script],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    if completed.returncode != 0 or not target.exists():
        detail = normalize_text(completed.stderr or completed.stdout)
        raise RuntimeError(
            f"No se pudo convertir {path.name} con Microsoft Word. "
            "Verifique que Word esté instalado y pueda abrir el archivo."
            + (f" Detalle: {detail}" if detail else "")
        )
    return target


def deduplicate(records: Iterable[Record]) -> pd.DataFrame:
    """Elimina solo duplicados exactos de la pareja código-parte."""
    unique_pairs: dict[tuple[str, str], None] = {}
    for record in records:
        unique_pairs[(record.codigo, record.parte)] = None
    return pd.DataFrame(
        [{"codigo": code, "parte": party} for code, party in unique_pairs],
        columns=["codigo", "parte"],
    )


def select_preferred_party(parties: Iterable[str]) -> str:
    """Prefiere una parte no vacía, no truncada y con mayor información."""
    unique = list(dict.fromkeys(party for party in parties if party))
    if not unique:
        return ""
    return max(
        unique,
        key=lambda party: (
            not party.lower().endswith(tuple(f" {particle}" for particle in SURNAME_PARTICLES)),
            "," not in party,
            len(party.split()),
            len(party),
        ),
    )


def build_quality_reports(records: Sequence[Record], source_issues: Sequence[dict[str, str]] | None = None, verified_cases: set[tuple[str, str]] | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Construye reportes auxiliares sin alterar la salida principal."""
    review_rows: list[dict[str, str]] = []
    alert_rows: list[dict[str, str]] = list(source_issues or [])
    by_code: dict[str, set[str]] = {}
    for record in records:
        by_code.setdefault(record.codigo, set()).add(record.parte)
        words = record.parte.split()
        if not record.parte:
            review_rows.append({"codigo": record.codigo, "parte": "", "motivo": "Parte vacía"})
            alert_rows.append({"codigo": record.codigo, "parte": "", "nivel": "ALTA", "motivo": "Parte vacía; revisar titular en la fuente"})
        elif SUCCESSION_RE.match(record.parte):
            continue
        elif is_legal_entity(record.parte):
            continue
        elif record.parte.lower().endswith(tuple(f" {particle}" for particle in SURNAME_PARTICLES)):
            review_rows.append({"codigo": record.codigo, "parte": record.parte, "motivo": "Termina en partícula; revisar apellido compuesto"})
        if not is_legal_entity(record.parte) and not SUCCESSION_RE.match(record.parte):
            if len(words) == 1:
                alert_rows.append({"codigo": record.codigo, "parte": record.parte, "nivel": "MEDIA", "motivo": "Solo un apellido disponible; confirmar en la fuente"})
            elif len(words) >= 5:
                alert_rows.append({"codigo": record.codigo, "parte": record.parte, "nivel": "MEDIA", "motivo": "Titular con muchos tokens; revisar posible texto adicional"})
    verified = verified_cases or set()
    review_rows = [row for row in review_rows if (row["codigo"], row["parte"]) not in verified]
    alert_rows = [
        row for row in alert_rows
        if row.get("nivel") != "MEDIA" or (row.get("codigo", ""), row.get("parte", "")) not in verified
    ]
    conflict_rows = [
        {
            "codigo": code,
            "parte_seleccionada": select_preferred_party(parties),
            "partes_encontradas": " | ".join(sorted(parties)),
        }
        for code, parties in by_code.items()
        if code != "NO DEFINIDO" and len(parties) > 1
    ]
    return (
        pd.DataFrame(review_rows, columns=["codigo", "parte", "motivo"]),
        pd.DataFrame(conflict_rows, columns=["codigo", "parte_seleccionada", "partes_encontradas"]),
        pd.DataFrame(alert_rows, columns=["codigo", "parte", "nivel", "motivo"]),
    )


def _style_quality_sheet(worksheet: Worksheet) -> None:
    """Aplica un formato legible a una hoja de control de calidad."""
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11, color="000000")
    border = Border(*(Side(style="thin", color="D9E2F3") for _ in range(4)))
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    worksheet.row_dimensions[1].height = 25
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for row in worksheet.iter_rows(min_row=2):
        worksheet.row_dimensions[row[0].row].height = 20
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = border
    for column_cells in worksheet.columns:
        column_letter = get_column_letter(column_cells[0].column)
        maximum = max(len(normalize_text(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[column_letter].width = min(max(maximum + 5, 12), 80)


def quality_path_for(output_dir: Path, stem: str) -> Path:
    """Construye la ruta estable del reporte para conservar estados verificados."""
    safe_stem = re.sub(r"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9]+", "_", stem).strip("_").lower()
    return output_dir / f"control_calidad_{safe_stem}.xlsx"


def ensure_output_available(path: Path) -> None:
    """Detecta antes del procesamiento si un archivo de salida está bloqueado."""
    if not path.exists():
        return
    try:
        with path.open("r+b"):
            pass
    except PermissionError as exc:
        raise RuntimeError(
            f"No se puede actualizar '{path.name}' porque está abierto. "
            "Cierre ese archivo de Excel y vuelva a ejecutar el programa."
        ) from exc


def load_verified_cases(report_path: Path) -> set[tuple[str, str]]:
    """Lee los casos MEDIA marcados como VERIFICADO en el reporte anterior."""
    if not report_path.exists():
        return set()
    workbook = None
    try:
        workbook = load_workbook(report_path, read_only=True, data_only=True)
        worksheet = workbook["revision_manual"]
        rows = worksheet.iter_rows(values_only=True)
        headers = [normalize_text(value).lower() for value in next(rows, ())]
        positions = {header: index for index, header in enumerate(headers)}
        if not {"codigo", "parte", "nivel", "estado"}.issubset(positions):
            return set()
        return {
            (normalize_text(row[positions["codigo"]]), normalize_text(row[positions["parte"]]))
            for row in rows
            if normalize_text(row[positions["estado"]]).upper() == "VERIFICADO"
            and normalize_text(row[positions["nivel"]]).upper() == "MEDIA"
        }
    except (OSError, KeyError, StopIteration):
        return set()
    finally:
        if workbook is not None:
            workbook.close()


def export_quality_reports(
    review: pd.DataFrame,
    conflicts: pd.DataFrame,
    alerts: pd.DataFrame,
    output_dir: Path,
    stem: str,
    processed_count: int = 0,
) -> Path:
    """Guarda un libro Excel de control de calidad en la carpeta de notas."""
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = quality_path_for(output_dir, stem)
    workbook = Workbook()
    review_export = review.copy()
    if not review_export.empty:
        review_export["nivel"] = review_export["motivo"].map(
            lambda reason: "ALTA" if "vacía" in reason.lower() else "MEDIA"
        )
        review_export = review_export[["codigo", "parte", "nivel", "motivo"]]
    else:
        review_export = pd.DataFrame(columns=["codigo", "parte", "nivel", "motivo"])
    review_export = pd.concat([review_export, alerts], ignore_index=True).drop_duplicates()
    review_export["estado"] = "PENDIENTE"
    review_sheet = workbook.active
    review_sheet.title = "revision_manual"
    review_sheet.append(list(review_export.columns))
    for row in review_export.itertuples(index=False, name=None):
        review_sheet.append(list(row))
    status_column = list(review_export.columns).index("estado") + 1
    validation = DataValidation(type="list", formula1='"PENDIENTE,VERIFICADO"', allow_blank=False)
    review_sheet.add_data_validation(validation)
    if review_sheet.max_row >= 2:
        validation.add(f"{get_column_letter(status_column)}2:{get_column_letter(status_column)}{review_sheet.max_row}")
    conflict_sheet = workbook.create_sheet("conflictos")
    conflict_sheet.append(list(conflicts.columns))
    for row in conflicts.itertuples(index=False, name=None):
        conflict_sheet.append(list(row))
    summary = pd.DataFrame(
        [
            {"indicador": "Registros procesados", "cantidad": processed_count, "detalle": "Parejas codigo + parte en la salida principal"},
            {"indicador": "Conflictos", "cantidad": len(conflicts), "detalle": "Códigos asociados a más de una parte"},
            {"indicador": "Casos pendientes", "cantidad": len(alerts), "detalle": "Alertas que requieren revisión manual"},
            {"indicador": "Expedientes no definidos", "cantidad": int((alerts["codigo"] == "NO DEFINIDO").sum()) if not alerts.empty else 0, "detalle": "Etiquetas de expediente sin número"},
            {"indicador": "Nivel ALTA", "cantidad": "", "detalle": "Requiere corregir o completar la información en la fuente"},
            {"indicador": "Nivel MEDIA", "cantidad": "", "detalle": "Revisar y marcar VERIFICADO si el resultado es correcto"},
            {"indicador": "Estado VERIFICADO", "cantidad": "", "detalle": "La alerta MEDIA dejará de aparecer como pendiente en la siguiente ejecución"},
        ],
        columns=["indicador", "cantidad", "detalle"],
    )
    summary_sheet = workbook.create_sheet("resumen")
    summary_sheet.append(list(summary.columns))
    for row in summary.itertuples(index=False, name=None):
        summary_sheet.append(list(row))
    _style_quality_sheet(review_sheet)
    _style_quality_sheet(conflict_sheet)
    _style_quality_sheet(summary_sheet)
    workbook.save(report_path)
    return report_path


def style_output_sheet(worksheet: Worksheet, dataframe: pd.DataFrame) -> None:
    """Aplica el formato corporativo solicitado al resultado."""
    worksheet.sheet_view.showGridLines = True
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="Calibri", size=11, color="000000")
    border = Border(*(Side(style="thin", color="D9E2F3") for _ in range(4)))
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    worksheet.row_dimensions[1].height = 25
    for row in worksheet.iter_rows(min_row=2):
        worksheet.row_dimensions[row[0].row].height = 20
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border
    for index, column in enumerate(dataframe.columns, start=1):
        values = [normalize_text(value) for value in dataframe[column].tolist()]
        width = min(max(len(str(column)), *(len(value) for value in values), 10) + 5, 80)
        worksheet.column_dimensions[get_column_letter(index)].width = width


def export_xlsx(dataframe: pd.DataFrame, output_path: Path) -> None:
    """Exporta únicamente ``codigo`` y ``parte`` con estilos."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Expedientes ETL"
    worksheet.append(["codigo", "parte"])
    for row in dataframe[["codigo", "parte"]].itertuples(index=False, name=None):
        worksheet.append(list(row))
    style_output_sheet(worksheet, dataframe[["codigo", "parte"]])
    workbook.save(output_path)


def build_records(excel_paths: Sequence[Path], word_paths: Sequence[Path], code_column: str | None, party_column: str | None) -> pd.DataFrame:
    records: list[Record] = []
    for path in excel_paths:
        records.extend(process_excel(path, code_column, party_column))
    for path in word_paths:
        if path.suffix.lower() == ".docx":
            records.extend(process_word(path))
        elif path.suffix.lower() == ".doc":
            temporary_docx = convert_legacy_doc(path)
            try:
                records.extend(process_word(temporary_docx))
            finally:
                temporary_docx.unlink(missing_ok=True)
        else:
            raise ValueError(f"Formato Word no soportado: {path.name}. Use .docx o .doc.")
    return deduplicate(records)


def process_single_file(path: Path, code_column: str | None, party_column: str | None) -> pd.DataFrame:
    """Procesa una única fuente y devuelve su resultado independiente."""
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        records = process_excel(path, code_column, party_column)
    elif suffix == ".docx":
        records = process_word(path)
    elif suffix == ".doc":
        temporary_docx = convert_legacy_doc(path)
        try:
            records = process_word(temporary_docx)
        finally:
            temporary_docx.unlink(missing_ok=True)
    else:
        raise ValueError(f"Formato no soportado: {path.name}. Use .xlsx, .docx o .doc.")
    return deduplicate(records)


def output_path_for(input_path: Path, output_dir: Path) -> Path:
    """Construye el nombre de salida conservando el nombre de cada documento."""
    return output_dir / f"{input_path.stem}_procesado.xlsx"


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Limpia registros judiciales para Expedientes ETL.")
    parser.add_argument("--excel", nargs="*", type=Path, default=None, help="Archivos .xlsx de entrada. Sin este argumento, se buscan en la carpeta del script.")
    parser.add_argument("--word", nargs="*", type=Path, default=None, help="Archivos .docx de entrada. Sin este argumento, se buscan en la carpeta del script.")
    parser.add_argument("--salida", type=Path, default=None, help=f"Libro .xlsx de salida. Por defecto: {DEFAULT_OUTPUT}")
    parser.add_argument("--columna-codigo", help="Nombre exacto de la columna de identificador del expediente en el Excel de origen.")
    parser.add_argument("--columna-parte", help="Nombre exacto de la columna de parte procesal o litigante en el Excel de origen.")
    return parser.parse_args(argv)


def discover_inputs(args: argparse.Namespace) -> tuple[list[Path], list[Path], Path]:
    """Resuelve entradas y salida; permite ejecutar el script sin argumentos."""
    excel_paths = args.excel if args.excel is not None else _discover_files("*.xlsx")
    word_paths = args.word if args.word is not None else [*_discover_files("*.docx"), *_discover_files("*.doc")]
    if not excel_paths and not word_paths:
        raise ValueError(f"No se encontraron archivos .xlsx, .docx o .doc en {INPUT_DIRS[0]}.")
    return list(excel_paths), list(word_paths), args.salida or DEFAULT_OUTPUT


def _discover_files(pattern: str) -> list[Path]:
    """Busca primero en entrada y usa proyectos como compatibilidad temporal."""
    found: dict[str, Path] = {}
    for directory_index, directory in enumerate(INPUT_DIRS):
        if directory_index > 0 and any(INPUT_DIRS[0].glob(pattern)):
            continue
        for path in directory.glob(pattern):
            if path.name.startswith("~$"):
                continue
            found[str(path.resolve()).lower()] = path
    return sorted(found.values(), key=lambda path: str(path).lower())


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        excel_paths, word_paths, output_path = discover_inputs(args)
        input_paths = [*excel_paths, *word_paths]
        output_dir = output_path if output_path.suffix.lower() != ".xlsx" else output_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        total = 0
        for input_path in input_paths:
            quality_issues: list[dict[str, str]] = []
            if input_path.suffix.lower() == ".xlsx":
                records = process_excel(input_path, args.columna_codigo, args.columna_parte, quality_issues)
            elif input_path.suffix.lower() == ".docx":
                records = process_word(input_path, quality_issues)
            else:
                temporary_docx = convert_legacy_doc(input_path)
                try:
                    records = process_word(temporary_docx, quality_issues)
                finally:
                    temporary_docx.unlink(missing_ok=True)
            dataframe = deduplicate(records)
            destination = output_path_for(input_path, output_dir)
            ensure_output_available(destination)
            ensure_output_available(quality_path_for(QUALITY_DIR, input_path.stem))
            export_xlsx(dataframe, destination)
            quality_path = quality_path_for(QUALITY_DIR, input_path.stem)
            verified_cases = load_verified_cases(quality_path)
            review, conflicts, alerts = build_quality_reports(records, quality_issues, verified_cases)
            quality_path = export_quality_reports(review, conflicts, alerts, QUALITY_DIR, input_path.stem, len(dataframe))
            total += len(dataframe)
            print(f"OK: {len(dataframe)} registros de {input_path.name} -> {destination}")
            print(f"  Control de calidad: {len(review)} revisión(es), {len(conflicts)} conflicto(s) -> {quality_path}")
            print(f"  Alertas de calidad: {len(alerts)}")
    except (OSError, ValueError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Proceso terminado: {len(input_paths)} archivo(s), {total} registros en total.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
