from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.p3 import P3GenerationError, generate_agency_documents


def _is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
    except ValueError:
        return False
    return True


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate internally reviewed agency forms from verified P2 results."
    )
    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="Company-approved P2 input workbook outside the repository",
    )
    parser.add_argument(
        "--p2-result",
        type=Path,
        required=True,
        help="P2 result workbook outside the repository",
    )
    parser.add_argument("--target-period", required=True, help="Target month in YYYY-MM format")
    parser.add_argument(
        "--output-directory",
        type=Path,
        required=True,
        help="Output directory outside the repository",
    )
    arguments = parser.parse_args()

    input_path = arguments.input.resolve()
    result_path = arguments.p2_result.resolve()
    output_directory = arguments.output_directory.resolve()
    for label, path in (
        ("Input", input_path),
        ("P2 result", result_path),
        ("Output directory", output_directory),
    ):
        if _is_within(path, PROJECT_ROOT):
            parser.error(f"{label} must be stored outside the Git repository.")

    try:
        report = generate_agency_documents(
            approved_input_path=input_path,
            p2_result_path=result_path,
            target_period=arguments.target_period,
            output_directory=output_directory,
        )
    except (OSError, P3GenerationError, ValueError) as error:
        parser.error(str(error))

    print(
        json.dumps(
            {
                "target_period": report.target_period,
                "status": report.status,
                "agency_count": len(report.artifacts),
                "artifacts": [
                    {
                        "agency_id": artifact.agency_id,
                        "condition_count": artifact.condition_count,
                        "excel": str(artifact.spreadsheet_path),
                        "word": str(artifact.word_path),
                    }
                    for artifact in report.artifacts
                ],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
