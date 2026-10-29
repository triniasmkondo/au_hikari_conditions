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
    arguments = parser.parse_args(argv)
    try:
        report = build_monthly_comparison_workbook(
            project_root=PROJECT_ROOT,
            previous_period=arguments.previous_period,
            target_period=arguments.target_period,
            output_path=arguments.output,
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


if __name__ == "__main__":
    raise SystemExit(main())
