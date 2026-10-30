"""Create or recompute a Phase 2 review workbook. Never generates notices or mail."""
import argparse
import json
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from au_hikari_conditions.phase2_review import build_phase2


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--ledger',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--review-input',type=Path)
    parser.add_argument('--agency')
    parser.add_argument('--previous-period',default='2026-10')
    parser.add_argument('--target-period',default='2026-11')
    args=parser.parse_args()
    stats=build_phase2(Path(__file__).resolve().parents[1],args.ledger,args.output,
        args.previous_period,args.target_period,args.agency,args.review_input)
    print(json.dumps(stats,ensure_ascii=False))
    return 0


if __name__=='__main__':raise SystemExit(main())
