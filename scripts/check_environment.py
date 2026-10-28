from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
REQUIRED_DIRECTORIES = (
    "01_上位原本",
    "02_取得データ",
    "03_条件表",
    "04_成果物",
    "05_配信管理",
    "06_作業記録",
    "07_監査ログ",
    "src/au_hikari_conditions",
    "tests",
)


def main() -> int:
    errors: list[str] = []
    if sys.version_info < (3, 10):
        errors.append("Python 3.10 以上が必要です。")

    for directory in REQUIRED_DIRECTORIES:
        if not (PROJECT_ROOT / directory).is_dir():
            errors.append(f"必要なディレクトリーがありません: {directory}")

    sys.path.insert(0, str(PROJECT_ROOT / "src"))
    try:
        from au_hikari_conditions import Agency, ConditionItem, SourceDocument  # noqa: F401
    except ImportError as error:
        errors.append(f"データモデルを読み込めません: {error}")

    if errors:
        print("環境確認: NG")
        for error in errors:
            print(f"- {error}")
        return 1

    print(f"環境確認: OK (Python {sys.version.split()[0]})")
    print(f"プロジェクト: {PROJECT_ROOT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
