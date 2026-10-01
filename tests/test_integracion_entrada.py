"""Pruebas de integración contra los documentos actuales de C:\\Work\\entrada."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
WORKSPACE_DIR = PROJECT_DIR
sys.path.insert(0, str(PROJECT_DIR))

import expedientes_etl as etl  # noqa: E402


class RealInputIntegrationTests(unittest.TestCase):
    def test_ignores_blank_and_section_rows(self) -> None:
        self.assertTrue(etl._is_section_label("SEDE LIMA SUR"))
        self.assertTrue(etl._is_section_label("EDIFICIO ALZAMORA - LIMA"))
        self.assertFalse(etl._is_section_label("CONSTRUCTORA SIN EXPEDIENTE S.A.C."))

    def test_excel_real_has_records_and_normalized_codes(self) -> None:
        files = sorted(
            path for path in (WORKSPACE_DIR / "entrada").glob("*.xlsx")
            if not path.name.startswith("~$")
        )
        if not files:
            self.skipTest("No hay archivos .xlsx locales para esta prueba de integración")
        records = [record for path in files for record in etl.process_excel(path)]
        self.assertGreater(len(records), 0)
        self.assertIn("00355-2015-1-3207-JR-CI-01", {record.codigo for record in records})
        self.assertTrue(all(etl.COURT_CODE_RE.fullmatch(record.codigo) for record in records))
        self.assertTrue(all(record.codigo.split("-", 1)[0].isdigit() and len(record.codigo.split("-", 1)[0]) == 5 for record in records))

    def test_word_real_has_records_and_inherits_party(self) -> None:
        files = sorted((WORKSPACE_DIR / "entrada").glob("*.docx"))
        if not files:
            self.skipTest("No hay archivos .docx locales para esta prueba de integración")
        source_issues: list[dict[str, str]] = []
        records = etl.process_word(files[0], source_issues)
        self.assertGreater(len(records), 0)
        self.assertTrue(all(record.codigo for record in records))
        self.assertTrue(all(record.parte for record in records))
        self.assertIn("Armando Carlini", {record.parte for record in records})

        _, conflicts, alerts = etl.build_quality_reports(records, source_issues)
        self.assertGreaterEqual(len(conflicts), 2)
        self.assertNotIn("NO DEFINIDO", set(conflicts["codigo"]))
        self.assertGreaterEqual(len(alerts), 7)
        self.assertTrue(
            {
                "Felisa Caballero",
                "Guillermo Espinoza",
                "Wilfredo Tenorio",
                "CONSTRUCTORA INARCO S.A.C",
                "Luciano Silva",
                "T&C FLORIDA'S REALTORS SAC",
            }.issubset(set(alerts["parte"]))
        )
        self.assertTrue(
            ((alerts["codigo"] == "NO DEFINIDO") & (alerts["parte"] == "CONSTRUCTORA INARCO S.A.C")).any()
        )
        self.assertIn(
            ("38438-2013-0-1801-JR-CI-03", "INVERSIONES PLATINUM (DIEGO FARAH)"),
            set(map(tuple, etl.deduplicate(records).to_records(index=False))),
        )
        self.assertNotIn("De Los Heros Ballen", set(alerts["parte"]))


if __name__ == "__main__":
    unittest.main()
