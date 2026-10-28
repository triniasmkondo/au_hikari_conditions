from __future__ import annotations

import json
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.analysis import build_p1_report


def main() -> int:
    report = build_p1_report(PROJECT_ROOT)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["p1_gate"] == "met" else 1


if __name__ == "__main__":
    raise SystemExit(main())