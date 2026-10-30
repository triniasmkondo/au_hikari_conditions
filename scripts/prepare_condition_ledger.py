"""Generate an unapproved Phase 1 agency condition ledger without network access."""
import argparse
import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from au_hikari_conditions.condition_ledger import build_ledger


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--period", default="2026-10")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--agency", help="First verify one agency before expansion")
    args = parser.parse_args()
    result = build_ledger(Path(__file__).resolve().parents[1], args.output, args.period, agency_filter=args.agency)
    print(json.dumps(result, ensure_ascii=False))
    return 2 if result["read_errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
