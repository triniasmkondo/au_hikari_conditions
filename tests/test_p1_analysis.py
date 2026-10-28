from __future__ import annotations

import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.analysis import build_p1_report
from au_hikari_conditions.comparison import compare_upper_campaign_notices
from au_hikari_conditions.document_reader import read_document
from au_hikari_conditions.models import DELIVERABLE_ITEM_ID_DICTIONARY
from au_hikari_conditions.normalization import normalize_blocks


class P1SourceAnalysisTest(unittest.TestCase):
    def test_reads_pdf_excel_and_word_with_evidence_locations(self) -> None:
        pdf_path = next((PROJECT_ROOT / "01_上位原本" / "2026-10").glob("*キャンペーン通知書.pdf"))
        xlsx_path = next((PROJECT_ROOT / "03_条件表" / "2026-10").glob("*条件表通知表.xlsx"))
        docx_path = next((PROJECT_ROOT / "03_条件表" / "2026-10").glob("*.docx"))

        pdf_blocks = read_document(pdf_path, project_root=PROJECT_ROOT)
        xlsx_blocks = read_document(xlsx_path, project_root=PROJECT_ROOT)
        docx_blocks = read_document(docx_path, project_root=PROJECT_ROOT)

        self.assertTrue(any(block.location == "page:1" for block in pdf_blocks))
        self.assertTrue(any("F19" in block.location and "G19" in block.location for block in xlsx_blocks))
        self.assertTrue(any(block.location.startswith("paragraph:") for block in docx_blocks))

    def test_normalizes_expected_periods_amounts_and_ladder_evidence(self) -> None:
        october_path = next((PROJECT_ROOT / "01_上位原本" / "2026-10").glob("*_20261002修正.pdf"))
        november_path = next((PROJECT_ROOT / "01_上位原本" / "2026-11").glob("*_20261028追加.pdf"))
        workbook_path = next((PROJECT_ROOT / "03_条件表" / "2026-10").glob("*条件表通知表.xlsx"))
        contract_path = next((PROJECT_ROOT / "01_上位原本" / "2026-10").glob("*個別契約書*.pdf"))

        october_fields = normalize_blocks(read_document(october_path, project_root=PROJECT_ROOT))
        november_fields = normalize_blocks(read_document(november_path, project_root=PROJECT_ROOT))
        workbook_fields = normalize_blocks(read_document(workbook_path, project_root=PROJECT_ROOT))
        contract_fields = normalize_blocks(read_document(contract_path, project_root=PROJECT_ROOT))

        october_periods = {
            value for field in october_fields if field.item_id == "p1.applicable_period"
            for value in field.normalized_values
        }
        november_periods = {
            value for field in november_fields if field.item_id == "p1.applicable_period"
            for value in field.normalized_values
        }
        self.assertIn("2026-10-01/2026-10-31", october_periods)
        self.assertIn("2026-11-01/2026-11-30", november_periods)
        self.assertTrue(
            any(
                "18000/1ID" in field.normalized_values
                and "F19" in field.location
                and "G19" in field.location
                for field in workbook_fields
                if field.item_id == "p1.commission_amount"
            )
        )
        self.assertTrue(
            any(
                field.item_id == "p1.ladder"
                and field.location == "page:18"
                and "回線獲得件数" in field.raw_value
                for field in contract_fields
            )
        )

    def test_upper_comparison_keeps_candidate_deltas_on_hold(self) -> None:
        comparison = compare_upper_campaign_notices(PROJECT_ROOT)

        self.assertEqual(comparison.status, "compared")
        self.assertEqual(len(comparison.differences), 6)
        self.assertTrue(all(item.status.endswith("_candidate") for item in comparison.differences))
        self.assertTrue(all(item.hold_reason for item in comparison.differences))
        self.assertTrue(
            all(item.previous_evidence or item.current_evidence for item in comparison.differences)
        )

    def test_repository_report_does_not_pass_gate_with_unapproved_ids(self) -> None:
        report = build_p1_report(PROJECT_ROOT)

        self.assertEqual(report["source_document_count"], 37)
        self.assertEqual(
            report["source_document_count_by_folder"],
            {"01_上位原本": 5, "03_条件表": 32},
        )
        self.assertEqual(
            report["source_document_count_by_format"],
            {".docx": 11, ".pdf": 5, ".xlsx": 21},
        )
        self.assertEqual(report["pending_item_ids"], sorted(DELIVERABLE_ITEM_ID_DICTIONARY))
        self.assertEqual(report["agency_to_upper_matching"]["status"], "held")
        self.assertEqual(report["agency_to_upper_matching"]["agency_document_count"], 32)
        self.assertEqual(report["agency_to_upper_matching"]["matched_field_count"], 0)
        self.assertEqual(report["unrecognized_documents"], [])
        self.assertEqual(report["unreadable_documents"], [])
        self.assertEqual(report["missing_item_candidates"], [])
        self.assertEqual(report["p1_gate"], "not_met")
        self.assertTrue(report["blockers"])


if __name__ == "__main__":
    unittest.main()