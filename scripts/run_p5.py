from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.p4b import DraftHistoryStore
from au_hikari_conditions.p5 import (
    P5Error,
    create_approved_zoho_drafts,
    finalize_approved_monthly_batch,
    prepare_monthly_batch,
)
from au_hikari_conditions.zoho_mail import (
    DpapiSecretProvider,
    ZohoMailClient,
    ZohoMailConfig,
)


def _add_common_inputs(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--target-period", required=True, help="Target month in YYYY-MM format")
    parser.add_argument("--batch-id", required=True, help="Unique, non-sensitive monthly batch ID")
    parser.add_argument(
        "--approved-input",
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
    parser.add_argument(
        "--monthly-comparison",
        type=Path,
        required=True,
        help="Unapproved previous/current source comparison workbook outside the repository",
    )


def _add_finalize_paths(parser: argparse.ArgumentParser) -> None:
    _add_common_inputs(parser)
    parser.add_argument(
        "--delivery-workbook",
        type=Path,
        required=True,
        help="Current company-approved recipient/template workbook outside the repository",
    )
    parser.add_argument(
        "--approval-workbook",
        type=Path,
        required=True,
        help="Reviewed and approved monthly workbook outside the repository",
    )
    parser.add_argument(
        "--candidate-directory",
        type=Path,
        required=True,
        help="P3 approval candidates outside the repository",
    )
    parser.add_argument(
        "--final-directory",
        type=Path,
        required=True,
        help="Final distribution artifacts outside the repository",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run gated P1-P5 monthly condition-table preparation."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    prepare_parser = subparsers.add_parser(
        "prepare", help="Run P1-P3 and write monthly approval requests."
    )
    _add_common_inputs(prepare_parser)
    prepare_parser.add_argument(
        "--delivery-workbook",
        type=Path,
        required=True,
        help="Company-approved recipient/template workbook outside the repository",
    )
    prepare_parser.add_argument(
        "--approval-workbook",
        type=Path,
        required=True,
        help="New approval request workbook path outside the repository",
    )
    prepare_parser.add_argument(
        "--candidate-directory",
        type=Path,
        required=True,
        help="New P3 review-candidate directory outside the repository",
    )
    prepare_parser.add_argument("--requested-by", required=True)

    finalize_parser = subparsers.add_parser(
        "finalize", help="Verify approval, finalize artifacts, and optionally save drafts."
    )
    _add_finalize_paths(finalize_parser)
    finalize_parser.add_argument(
        "--save-drafts",
        action="store_true",
        help="Create Zoho drafts only with the explicit external company-approval gate.",
    )

    arguments = parser.parse_args(argv)
    try:
        if arguments.command == "prepare":
            result = prepare_monthly_batch(
                project_root=PROJECT_ROOT,
                approved_input_path=arguments.approved_input,
                p2_result_path=arguments.p2_result,
                monthly_comparison_path=arguments.monthly_comparison,
                target_period=arguments.target_period,
                batch_id=arguments.batch_id,
                delivery_workbook_path=arguments.delivery_workbook,
                approval_workbook_path=arguments.approval_workbook,
                candidate_directory=arguments.candidate_directory,
                requested_by=arguments.requested_by,
            )
            print(
                json.dumps(
                    {
                        "phase": "AWAITING_HUMAN_QA_AND_APPROVAL",
                        "p1_status": result.p1_status,
                        "target_period": result.target_period,
                        "agency_count": result.agency_count,
                    },
                    ensure_ascii=False,
                )
            )
            return 0

        packages = finalize_approved_monthly_batch(
            project_root=PROJECT_ROOT,
            approved_input_path=arguments.approved_input,
            p2_result_path=arguments.p2_result,
            monthly_comparison_path=arguments.monthly_comparison,
            target_period=arguments.target_period,
            batch_id=arguments.batch_id,
            delivery_workbook_path=arguments.delivery_workbook,
            approval_workbook_path=arguments.approval_workbook,
            candidate_directory=arguments.candidate_directory,
            final_directory=arguments.final_directory,
        )
        if arguments.save_drafts:
            if os.environ.get("P5_ZOHO_DRAFTS_APPROVED") != "YES":
                raise P5Error("Explicit company approval for Zoho draft creation is missing.")
            history_path = os.environ.get("P5_DRAFT_HISTORY_DB")
            if not history_path:
                raise P5Error("External P5_DRAFT_HISTORY_DB configuration is required.")
            history = DraftHistoryStore(Path(history_path))
            client = ZohoMailClient(
                ZohoMailConfig.from_secret_provider(DpapiSecretProvider())
            )
            statuses = create_approved_zoho_drafts(
                packages=packages,
                mail_client=client,
                history=history,
                company_authorized=True,
            )
            phase = "DRAFTS_CREATED"
            draft_count = len(statuses)
        else:
            phase = "ARTIFACTS_FINALIZED_DRAFTS_NOT_CREATED"
            draft_count = 0
        print(
            json.dumps(
                {
                    "phase": phase,
                    "target_period": arguments.target_period,
                    "agency_count": len(packages),
                    "draft_count": draft_count,
                },
                ensure_ascii=False,
            )
        )
        return 0
    except Exception as error:
        print(
            json.dumps(
                {
                    "phase": "BLOCKED",
                    "error_category": type(error).__name__,
                    "details": "No email was sent. Review the approved inputs and resolve the gate.",
                },
                ensure_ascii=False,
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
