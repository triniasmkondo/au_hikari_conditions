from __future__ import annotations

import sys
import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.analysis import build_p1_report
from au_hikari_conditions.comparison import compare_upper_campaign_notices
from au_hikari_conditions.document_reader import EvidenceBlock, read_document
from au_hikari_conditions.models import DELIVERABLE_ITEM_ID_DICTIONARY
from au_hikari_conditions.normalization import normalize_blocks


class P1SourceAnalysisTest(unittest.TestCase):
    def test_extracts_condition_dimensions_as_held_candidates(self) -> None:
        blocks = [
            EvidenceBlock(
                source_path="source.pdf",
                location="page:2",
                text="auひかり プラス① BIGLOBE ホーム 1G 事業者変更 開通 報奨金額 18,000円/件",
            )
        ]

        fields = normalize_blocks(blocks)
        amount_field = next(field for field in fields if field.item_id == "p1.commission_amount")
        dimensions = {
            dimension.name: dimension.values
            for dimension in amount_field.condition_dimensions
        }

        self.assertEqual(dimensions["product_family"], ("auひかりプラス",))
        self.assertEqual(dimensions["plus_type"], ("プラス1",))
        self.assertEqual(dimensions["isp"], ("BIGLOBE",))
        self.assertEqual(dimensions["dwelling_type"], ("ホーム",))
        self.assertEqual(dimensions["speed"], ("1G",))
        self.assertEqual(dimensions["trigger"], ("開通",))
        self.assertEqual(dimensions["access_type"], ("事業者変更",))
        self.assertEqual(amount_field.source_path, "source.pdf")
        self.assertEqual(amount_field.location, "page:2")
        self.assertEqual(amount_field.status, "hold")

    def test_period_context_does_not_cross_pdf_pages(self) -> None:
        blocks = [
            EvidenceBlock("source.pdf", "page:1", "対象期間"),
            EvidenceBlock("source.pdf", "page:2", "2026年11月1日～2026年11月30日"),
        ]

        fields = normalize_blocks(blocks)

        self.assertFalse(any(field.item_id == "p1.applicable_period" for field in fields))

    def test_period_context_stays_within_an_excel_sheet(self) -> None:
        blocks = [
            EvidenceBlock("source.xlsx", "sheet:条件;cells:A1", "対象期間"),
            EvidenceBlock(
                "source.xlsx",
                "sheet:条件;cells:A2",
                "2026年11月1日～2026年11月30日",
            ),
            EvidenceBlock(
                "source.xlsx",
                "sheet:別条件;cells:A1",
                "2026年12月1日～2026年12月31日",
            ),
        ]

        fields = normalize_blocks(blocks)
        periods = [
            field
            for field in fields
            if field.item_id == "p1.applicable_period"
        ]

        self.assertEqual(len(periods), 1)
        self.assertEqual(
            periods[0].normalized_values,
            ("2026-11-01/2026-11-30",),
        )

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

    def test_p1_gate_can_pass_while_unapproved_decisions_remain_held(self) -> None:
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
        self.assertTrue(report["condition_dimension_candidate_counts"]["isp"])
        self.assertTrue(report["condition_dimension_candidate_counts"]["speed"])
        self.assertEqual(report["agency_to_upper_matching"]["status"], "held")
        self.assertEqual(report["agency_to_upper_matching"]["agency_document_count"], 32)
        self.assertEqual(report["agency_to_upper_matching"]["matched_field_count"], 0)
        self.assertEqual(report["unrecognized_documents"], [])
        self.assertEqual(report["unreadable_documents"], [])
        self.assertEqual(report["missing_item_candidates"], [])
        self.assertEqual(report["held_candidate_count"], report["normalized_candidate_count"])
        self.assertEqual(report["p1_gate"], "met")
        self.assertEqual(report["blockers"], [])
        self.assertTrue(report["review_holds"])
        self.assertEqual(report["agency_to_upper_matching"]["status"], "held")
        self.assertEqual(
            report["p1_gate_scope"],
            "document_analysis_and_candidate_differences",
        )
        self.assertTrue(
            any(path.endswith("au_hikari_conditions_system_spec_v1.docx")
                for path in report["design_spec_files"])
        )


if __name__ == "__main__":
    unittest.main()