"""Pruebas de la capa modular pública del proyecto."""

import unittest
from pathlib import Path

from core.calidad import build_quality_reports
from core.exportacion import export_xlsx
from core.lectores import process_excel, process_word
from core.modelos import Record
from core.normalizacion import normalize_code
from core.servicio import process_file_to_outputs


class ModularArchitectureTests(unittest.TestCase):
    def test_public_modules_expose_expected_layers(self):
        self.assertEqual(normalize_code("27-2025-0-1801-JR-CI-01"), "00027-2025-0-1801-JR-CI-01")
        self.assertTrue(callable(process_excel))
        self.assertTrue(callable(process_word))
        self.assertTrue(callable(build_quality_reports))
        self.assertTrue(callable(export_xlsx))
        self.assertTrue(callable(process_file_to_outputs))

    def test_record_model_is_immutable_and_explicit(self):
        record = Record("00001-2025-0-1801-JR-CI-01", "Ana Rivera")
        self.assertEqual(record.codigo, "00001-2025-0-1801-JR-CI-01")
        self.assertEqual(record.parte, "Ana Rivera")
        with self.assertRaises(AttributeError):
            record.parte = "Otro nombre"


if __name__ == "__main__":
    unittest.main()
