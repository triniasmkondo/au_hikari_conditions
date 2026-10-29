from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.p2 import P2InputError, calculate_workbook, write_report


def _is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
    except ValueError:
        return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Calculate an approved P2 workbook without creating delivery artifacts."
    )
    parser.add_argument("--input", type=Path, required=True, help="Approved input workbook outside the repository")
    parser.add_argument("--target-period", required=True, help="Calculation month in YYYY-MM format")
    parser.add_argument("--output", type=Path, required=True, help="Result workbook path outside the repository")
    arguments = parser.parse_args()

    input_path = arguments.input.resolve()
    output_path = arguments.output.resolve()
    if _is_within(input_path, PROJECT_ROOT):
        parser.error("Operational input data must be stored outside the Git repository.")
    if _is_within(output_path, PROJECT_ROOT):
        parser.error("P2 result workbooks must be stored outside the Git repository.")

    try:
        report = calculate_workbook(input_path, arguments.target_period)
        write_report(report, output_path)
    except (OSError, P2InputError, ValueError) as error:
        parser.error(str(error))
    print(
        json.dumps(
            {
                "target_period": report.target_period,
                "result_count": len(report.results),
                "proposed_count": report.proposed_count,
                "hold_count": report.hold_count,
                "output": str(output_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0 if report.hold_count == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
