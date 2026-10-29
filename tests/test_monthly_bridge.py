from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.monthly_bridge import (
    MonthlyBridgeError,
    build_monthly_comparison_workbook,
    validate_monthly_comparison_workbook,
)


class MonthlyBridgeTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name) / "fictional-project"
        self.output = Path(self.temporary.name) / "comparison.xlsx"
        self.previous_upper = self.root / "01_上位原本" / "2026-10"
        self.previous_agency = self.root / "03_条件表" / "2026-10"
        self.current_upper = self.root / "01_上位原本" / "2026-11"
        for folder in (self.previous_upper, self.previous_agency, self.current_upper):
            folder.mkdir(parents=True)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def _write_docx(self, path: Path, text: str) -> None:
        document = Document()
        document.add_paragraph(text)
        document.save(path)

    def test_builds_held_source_and_dimension_comparison_for_all_three_inputs(self) -> None:
        self._write_docx(
            self.previous_upper / "fictional-previous-upper.docx",
            "auひかりプラス BIGLOBE ホーム 1G 新規 開通 報奨金額 100円/件",
        )
        self._write_docx(
            self.previous_agency / "fictional-agency-conditions.docx",
            "auひかりプラス BIGLOBE ホーム 1G 新規 開通 報奨金額 80円/件",
        )
        self._write_docx(
            self.current_upper / "fictional-current-upper.docx",
            "auひかりプラス BIGLOBE ホーム 1G 新規 開通 報奨金額 120円/件",
        )

        report = build_monthly_comparison_workbook(
            project_root=self.root,
            previous_period="2026-10",
            target_period="2026-11",
            output_path=self.output,
        )

        self.assertEqual(report.source_count, 3)
        self.assertGreaterEqual(report.candidate_count, 3)
        self.assertGreater(report.upper_difference_count, 0)
        self.assertGreater(report.unresolved_count, 0)
        self.assertEqual(report.status, "REVIEW_REQUIRED")
        workbook = load_workbook(self.output, read_only=True, data_only=True)
        try:
            self.assertEqual(
                set(workbook.sheetnames),
                {
                    "使い方",
                    "比較情報",
                    "資料一覧",
                    "抽出根拠候補",
                    "上位差分候補",
                    "P2項目対応",
                    "不足・保留",
                },
            )
            candidates = workbook["抽出根拠候補"]
            headers = [cell.value for cell in candidates[1]]
            rows = [
                dict(zip(headers, row))
                for row in candidates.iter_rows(min_row=2, values_only=True)
            ]
            agency_rows = [row for row in rows if row["source_type"] == "PREVIOUS_AGENCY"]
            self.assertTrue(agency_rows)
            self.assertTrue(
                all(row["agency_identity_status"] == "MANUAL_AGENCY_ID_REQUIRED" for row in agency_rows)
            )
            self.assertTrue(all(row["approval_status"] == "UNAPPROVED" for row in rows))

            differences = workbook["上位差分候補"]
            diff_headers = [cell.value for cell in differences[1]]
            diff_rows = [
                dict(zip(diff_headers, row))
                for row in differences.iter_rows(min_row=2, values_only=True)
            ]
            amount_differences = [
                row for row in diff_rows if row["item_id"] == "p1.commission_amount"
            ]
            self.assertEqual(len(amount_differences), 1)
            self.assertEqual(amount_differences[0]["candidate_status"], "CHANGED_CANDIDATE")
            self.assertEqual(amount_differences[0]["approval_status"], "UNAPPROVED")

            crosswalk = workbook["P2項目対応"]
            self.assertTrue(
                all(row[4] == "NO" for row in crosswalk.iter_rows(min_row=2, values_only=True))
            )
        finally:
            workbook.close()

    def test_refuses_output_inside_project_and_existing_output(self) -> None:
        with self.assertRaisesRegex(MonthlyBridgeError, "outside the Git repository"):
            build_monthly_comparison_workbook(
                project_root=self.root,
                previous_period="2026-10",
                target_period="2026-11",
                output_path=self.root / "unsafe.xlsx",
            )

        self._write_docx(
            self.previous_upper / "fictional-previous-upper.docx",
            "auひかりプラス BIGLOBE ホーム 1G 開通 報奨金額 100円",
        )
        self._write_docx(
            self.previous_agency / "fictional-agency.docx",
            "auひかりプラス BIGLOBE ホーム 1G 開通",
        )
        self._write_docx(
            self.current_upper / "fictional-current-upper.docx",
            "auひかりプラス BIGLOBE ホーム 1G 開通 報奨金額 110円",
        )
        self.output.write_bytes(b"existing")
        with self.assertRaisesRegex(MonthlyBridgeError, "Refusing to overwrite"):
            build_monthly_comparison_workbook(
                project_root=self.root,
                previous_period="2026-10",
                target_period="2026-11",
                output_path=self.output,
            )

    def test_comparison_validation_binds_periods_and_source_hashes(self) -> None:
        source_paths = (
            self.previous_upper / "fictional-previous-upper.docx",
            self.previous_agency / "fictional-agency-conditions.docx",
            self.current_upper / "fictional-current-upper.docx",
        )
        for index, path in enumerate(source_paths):
            self._write_docx(
                path,
                f"auひかりプラス BIGLOBE ホーム 1G 開通 報奨金額 {100 + index}円",
            )
        build_monthly_comparison_workbook(
            project_root=self.root,
            previous_period="2026-10",
            target_period="2026-11",
            output_path=self.output,
        )

        validation = validate_monthly_comparison_workbook(
            self.output,
            project_root=self.root,
            previous_period="2026-10",
            target_period="2026-11",
        )
        self.assertEqual(validation.source_count, 3)
        with self.assertRaisesRegex(MonthlyBridgeError, "periods or candidate status"):
            validate_monthly_comparison_workbook(
                self.output,
                project_root=self.root,
                previous_period="2026-09",
                target_period="2026-11",
            )

        source_paths[0].write_bytes(source_paths[0].read_bytes() + b"changed")
        with self.assertRaisesRegex(MonthlyBridgeError, "changed after comparison"):
            validate_monthly_comparison_workbook(
                self.output,
                project_root=self.root,
                previous_period="2026-10",
                target_period="2026-11",
            )

    def test_comparison_candidates_cannot_be_marked_approved(self) -> None:
        for index, folder in enumerate(
            (self.previous_upper, self.previous_agency, self.current_upper)
        ):
            self._write_docx(
                folder / f"fictional-source-{index}.docx",
                f"auひかりプラス BIGLOBE ホーム 1G 開通 報奨金額 {100 + index}円",
            )
        build_monthly_comparison_workbook(
            project_root=self.root,
            previous_period="2026-10",
            target_period="2026-11",
            output_path=self.output,
        )
        workbook = load_workbook(self.output)
        worksheet = workbook["抽出根拠候補"]
        headers = [cell.value for cell in worksheet[1]]
        worksheet.cell(2, headers.index("approval_status") + 1, "APPROVED")
        workbook.save(self.output)
        workbook.close()

        with self.assertRaisesRegex(MonthlyBridgeError, "cannot be approved"):
            validate_monthly_comparison_workbook(
                self.output,
                project_root=self.root,
                previous_period="2026-10",
                target_period="2026-11",
            )

    def test_changed_condition_dimension_is_held_as_mismatch(self) -> None:
        self._write_docx(
            self.previous_upper / "fictional-previous-upper.docx",
            "auひかり プラス BIGLOBE ホーム 1G 新規 開通 報奨金額 100円/件",
        )
        self._write_docx(
            self.previous_agency / "fictional-agency-conditions.docx",
            "auひかり プラス BIGLOBE ホーム 1G 新規 開通 報奨金額 80円/件",
        )
        self._write_docx(
            self.current_upper / "fictional-current-upper.docx",
            "auひかり プラス So-net ホーム 1G 新規 開通 報奨金額 120円/件",
        )

        build_monthly_comparison_workbook(
            project_root=self.root,
            previous_period="2026-10",
            target_period="2026-11",
            output_path=self.output,
        )
        workbook = load_workbook(self.output, read_only=True, data_only=True)
        try:
            sheet = workbook["上位差分候補"]
            headers = [cell.value for cell in sheet[1]]
            differences = [
                dict(zip(headers, row))
                for row in sheet.iter_rows(min_row=2, values_only=True)
                if row[0] is not None
            ]
            amount_differences = [
                row
                for row in differences
                if row["item_id"] == "p1.commission_amount"
            ]
            self.assertEqual(len(amount_differences), 1)
            self.assertEqual(
                amount_differences[0]["candidate_status"],
                "CONDITION_MISMATCH_REQUIRES_REVIEW",
            )
            self.assertTrue(amount_differences[0]["previous_values_json"])
            self.assertTrue(amount_differences[0]["current_values_json"])
            self.assertEqual(amount_differences[0]["approval_status"], "UNAPPROVED")
        finally:
            workbook.close()


if __name__ == "__main__":
    unittest.main()
