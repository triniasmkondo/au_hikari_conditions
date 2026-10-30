"""Create unapproved campaign candidates from explicit source files."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from au_hikari_conditions.campaign_import import build_campaign_candidates, write_campaign_candidate_workbook


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="Existing v4 input (never modified)")
    parser.add_argument("--upstream", type=Path, required=True)
    parser.add_argument("--agency", type=Path, required=True)
    parser.add_argument("--previous-agency", type=Path)
    parser.add_argument("--agency-id", required=True)
    parser.add_argument("--target-period", required=True)
    parser.add_argument("--output", type=Path, required=True, help="New workbook outside Git")
    args = parser.parse_args()
    candidates = build_campaign_candidates(args.upstream, args.agency,
        agency_id=args.agency_id, target_period=args.target_period,
        previous_agency_path=args.previous_agency)
    write_campaign_candidate_workbook(args.input, args.output, candidates)
    print(json.dumps(dict(status="UNREVIEWED", tier_count=len(candidates.rows),
        deduction_count=len(candidates.deductions), hold_count=len(candidates.holds))))
    return 2  # Candidate extraction is never payment approval.


if __name__ == "__main__":
    raise SystemExit(main())
