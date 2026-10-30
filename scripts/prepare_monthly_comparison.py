from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.monthly_bridge import (
    MonthlyBridgeError,
    build_monthly_comparison_workbook,
    discover_agency_document_set,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Create an unapproved, evidence-backed comparison workbook from "
            "previous-period and current-period source documents."
        )
    )
    parser.add_argument("--previous-period", required=True, help="Previous month in YYYY-MM format")
    parser.add_argument("--target-period", required=True, help="Target month in YYYY-MM format")
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="New comparison workbook path outside the Git repository",
    )
    agency_selection = parser.add_mutually_exclusive_group()
    agency_selection.add_argument(
        "--select-agency-interactively",
        action="store_true",
        help=(
            "Choose one agency document in a local file dialog; sibling condition "
            "documents are grouped by agency-key and type."
        ),
    )
    arguments = parser.parse_args(argv)
    try:
        agency_document_set = None
        if arguments.select_agency_interactively:
            selected_document = _choose_agency_document(
                PROJECT_ROOT / "03_条件表" / arguments.previous_period
            )
            agency_document_set = discover_agency_document_set(
                project_root=PROJECT_ROOT,
                agency_directory=(
                    PROJECT_ROOT / "03_条件表" / arguments.previous_period
                ),
                selected_document=selected_document,
            )
        report = build_monthly_comparison_workbook(
            project_root=PROJECT_ROOT,
            previous_period=arguments.previous_period,
            target_period=arguments.target_period,
            output_path=arguments.output,
            agency_document_set=agency_document_set,
        )
    except (MonthlyBridgeError, OSError) as error:
        parser.error(str(error))
    print(
        json.dumps(
            {
                "status": report.status,
                "previous_period": report.previous_period,
                "target_period": report.target_period,
                "source_count": report.source_count,
                "candidate_count": report.candidate_count,
                "upper_difference_count": report.upper_difference_count,
                "unresolved_count": report.unresolved_count,
                "output": str(report.output_path),
            },
            ensure_ascii=False,
        )
    )
    return 0 if report.status == "CANDIDATES_READY_FOR_HUMAN_REVIEW" else 2


def _choose_agency_document(initial_directory: Path) -> Path:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError as error:
        raise MonthlyBridgeError(
            "Local graphical agency selection is unavailable in this Python installation."
        ) from error

    root = tk.Tk()
    root.withdraw()
    try:
        selected = filedialog.askopenfilename(
            title="Select one condition document for the agency",
            initialdir=str(initial_directory),
            filetypes=(
                ("Supported condition documents", "*.pdf *.docx *.xlsx"),
                ("All files", "*.*"),
            ),
        )
    finally:
        root.destroy()
    if not selected:
        raise MonthlyBridgeError("Agency document selection was cancelled.")
    return Path(selected)


if __name__ == "__main__":
    raise SystemExit(main())
