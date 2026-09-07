"""Pruebas de regresión para las reglas críticas de Expedientes ETL."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))

import expedientes_etl as etl  # noqa: E402


class CodeTests(unittest.TestCase):
    def test_normalizes_correlative_to_five_digits(self) -> None:
        self.assertEqual(
            etl.normalize_code("18-2025-0-0207-JR-CI-01"),
            "00018-2025-0-0207-JR-CI-01",
        )

    def test_extracts_multiple_codes(self) -> None:
        value = "Exp. N° 13-2019-0-3006-JP-CI-01 y 18-2025-0-0207-JR-CI-01"
        self.assertEqual(
            list(etl.iter_codes(value)),
            ["00013-2019-0-3006-JP-CI-01", "00018-2025-0-0207-JR-CI-01"],
        )

    def test_expands_slash_separated_subcases(self) -> None:
        value = "00355-2015-0/1/2-3207-JR-CI-01"
        self.assertEqual(
            list(etl.iter_codes(value)),
            [
                "00355-2015-0-3207-JR-CI-01",
                "00355-2015-1-3207-JR-CI-01",
                "00355-2015-2-3207-JR-CI-01",
            ],
        )

    def test_partial_case_requires_exp_label(self) -> None:
        self.assertEqual(etl.PARTIAL_CASE_RE.search("Dirección 134-140"), None)
        self.assertEqual(
            etl.PARTIAL_CASE_RE.search("Exp. N° 27445-2025").group("code"),
            "27445-2025",
        )

    def test_preserves_source_variants_when_they_are_valid_for_the_source(self) -> None:
        self.assertEqual(
            list(etl.iter_codes("Exp. N° 27445-2025")),
            ["27445-2025"],
        )
        self.assertEqual(
            list(etl.iter_codes("02448-2024-0-18A5-JR-CI-20")),
            ["02448-2024-0-18A5-JR-CI-20"],
        )
        self.assertEqual(
            list(etl.iter_codes("15068-11-0-1801-JE-LA-17")),
            ["15068-11-0-1801-JE-LA-17"],
        )


class PartyTests(unittest.TestCase):
    def test_word_entity_keeps_parenthetical_designation(self) -> None:
        self.assertEqual(
            etl._block_party("INVERSIONES PLATINUM (DIEGO FARAH)\nCon la Municipalidad"),
            "INVERSIONES PLATINUM (DIEGO FARAH)",
        )

    def test_keeps_two_surnames_and_discards_given_names(self) -> None:
        self.assertEqual(etl.choose_party("Carlini Livelli, Armando Victorio"), "Carlini Livelli")
        self.assertEqual(etl.choose_party("Espinoza, Guillermo"), "Espinoza")

    def test_keeps_compound_surname_with_de(self) -> None:
        self.assertEqual(
            etl.choose_party("Rivera de Gonzales, Margarita Teresa"),
            "Rivera De Gonzales",
        )
        self.assertEqual(
            etl.choose_party("Gagliuffi de Castagnino, Blanca Teresa"),
            "Gagliuffi De Castagnino",
        )
        self.assertEqual(
            etl.choose_party("DE LOS HEROS BALLEN DE VAN WALLEGHEM, Rosa Maria"),
            "De Los Heros Ballen",
        )

    def test_removes_civil_status_abbreviation_from_surname(self) -> None:
        self.assertEqual(etl.choose_party("CABALLERO VDA"), "Caballero")
        self.assertEqual(
            etl.choose_party("CABALLERO VDA. DE GALVEZ, Felisa Clementina"),
            "Caballero",
        )

    def test_succession_is_preserved_as_designation(self) -> None:
        self.assertEqual(
            etl.choose_party("SUCESION DINO ARENAS LOZADA"),
            "Sucesión Dino Arenas Lozada",
        )

    def test_entities_are_not_truncated(self) -> None:
        self.assertEqual(
            etl.choose_party("INSTITUTO PERUANO DE ACCION EMPRESARIAL – IPAE"),
            "INSTITUTO PERUANO DE ACCION EMPRESARIAL – IPAE",
        )
        self.assertEqual(etl.choose_party("COPAMO, Playa Farallones"), "COPAMO")


class ConflictTests(unittest.TestCase):
    def test_only_exact_code_party_duplicates_are_removed(self) -> None:
        records = [
            etl.Record("00001-2020-0-0001-JR-CI-01", "Rivera"),
            etl.Record("00001-2020-0-0001-JR-CI-01", "Rivera"),
            etl.Record("00001-2020-0-0001-JR-CI-01", "Gonzales"),
        ]
        result = etl.deduplicate(records)
        self.assertEqual(len(result), 2)
        self.assertEqual(set(map(tuple, result.to_records(index=False))), {
            ("00001-2020-0-0001-JR-CI-01", "Rivera"),
            ("00001-2020-0-0001-JR-CI-01", "Gonzales"),
        })

    def test_conflict_report_keeps_distinct_parts(self) -> None:
        records = [
            etl.Record("00001-2020-0-0001-JR-CI-01", "Rivera"),
            etl.Record("00001-2020-0-0001-JR-CI-01", "Gonzales"),
        ]
        _, conflicts, alerts = etl.build_quality_reports(records)
        self.assertEqual(len(conflicts), 1)
        self.assertEqual(len(alerts), 2)
        self.assertEqual(set(alerts["parte"]), {"Rivera", "Gonzales"})
        self.assertIn("Rivera", conflicts.iloc[0]["partes_encontradas"])
        self.assertIn("Gonzales", conflicts.iloc[0]["partes_encontradas"])

    def test_undefined_code_is_not_a_false_conflict(self) -> None:
        records = [
            etl.Record("NO DEFINIDO", "Constructora Inarco S.A.C"),
            etl.Record("NO DEFINIDO", "Silva Alarco"),
        ]
        _, conflicts, alerts = etl.build_quality_reports(records)
        self.assertTrue(conflicts.empty)
        self.assertEqual(len(alerts), 0)

    def test_verified_media_alert_is_hidden_but_not_alta(self) -> None:
        records = [
            etl.Record("00001-2020-0-0001-JR-CI-01", "Rivera"),
            etl.Record("NO DEFINIDO", "Silva Alarco"),
        ]
        source_issues = [{"codigo": "NO DEFINIDO", "parte": "Silva Alarco", "nivel": "ALTA", "motivo": "Falta expediente"}]
        _, _, alerts = etl.build_quality_reports(
            records,
            source_issues,
            {("00001-2020-0-0001-JR-CI-01", "Rivera"), ("NO DEFINIDO", "Silva Alarco")},
        )
        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts.iloc[0]["codigo"], "NO DEFINIDO")

    def test_source_issues_are_added_to_quality_alerts(self) -> None:
        _, _, alerts = etl.build_quality_reports(
            [],
            [{"codigo": "", "parte": "", "nivel": "ALTA", "motivo": "Fila sin código judicial detectable"}],
        )
        self.assertEqual(len(alerts), 1)
        self.assertEqual(alerts.iloc[0]["nivel"], "ALTA")


if __name__ == "__main__":
    unittest.main()
