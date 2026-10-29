from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal

from .document_reader import EvidenceBlock
from .models import DELIVERABLE_ITEM_ID_DICTIONARY


DATE_RANGE_PATTERN = re.compile(
    r"(?P<start_year>20\d{2})\s*年\s*(?P<start_month>\d{1,2})\s*月\s*(?P<start_day>\d{1,2})\s*日?"
    r"\s*[～〜~\-―−]+\s*"
    r"(?P<end_year>20\d{2})\s*年\s*(?P<end_month>\d{1,2})\s*月\s*(?P<end_day>\d{1,2})\s*日?"
)
MONEY_PATTERN = re.compile(
    r"(?P<amount>(?:¥\s*)?\d[\d,]*(?:\.\d+)?)"
    r"\s*(?:\|\s*[A-Z]{1,3}\d+\s*=\s*)?円"
    r"(?:\s*/\s*(?P<unit>1件|1ID|件|ID|回|月|台))?"
)
PERIOD_LABEL_PATTERN = re.compile(r"対象期間|設定期間|業務実施対象期間|適用期間")
ISP_PATTERN = re.compile(r"So-net|BIGLOBE|@nifty|ISP|プロバイダ", re.IGNORECASE)
ACHIEVEMENT_PATTERN = re.compile(
    r"成果条件|対象申込|支払対象|登録申込|申込登録|課金開始|開通条件|契約申込"
)
LADDER_PATTERN = re.compile(
    r"ラダー|段階|(?:件数|獲得件数).{0,50}\d[\d,]*\s*(?:～|〜|~|-|以上|未満)"
    r"|\d[\d,]*\s*件\s*(?:以上|未満|～|〜|~|-|から)"
)
CLAWBACK_PATTERN = re.compile(r"戻入|返還|返金|相殺|取戻")


@dataclass(frozen=True)
class ConditionDimension:
    name: str
    values: tuple[str, ...]


@dataclass(frozen=True)
class NormalizedField:
    item_id: str
    raw_value: str
    normalized_values: tuple[str, ...]
    source_path: str
    location: str
    status: str
    hold_reasons: tuple[str, ...]
    condition_dimensions: tuple[ConditionDimension, ...] = ()


def normalize_blocks(blocks: list[EvidenceBlock]) -> list[NormalizedField]:
    normalized_blocks = [
        " ".join(unicodedata.normalize("NFKC", block.text).split()) for block in blocks
    ]
    fields: list[NormalizedField] = []

    for index, (block, text) in enumerate(zip(blocks, normalized_blocks)):
        region = _context_region(block)
        contexts = [
            normalized_blocks[context_index]
            for context_index in range(max(0, index - 1), min(len(blocks), index + 2))
            if _context_region(blocks[context_index]) == region
        ]
        context = " ".join(contexts)
        condition_dimensions = _extract_condition_dimensions(text)

        if PERIOD_LABEL_PATTERN.search(context):
            matches = list(DATE_RANGE_PATTERN.finditer(text))
            if matches:
                values = tuple(
                    f"{match['start_year']}-{int(match['start_month']):02d}-{int(match['start_day']):02d}/"
                    f"{match['end_year']}-{int(match['end_month']):02d}-{int(match['end_day']):02d}"
                    for match in matches
                )
                fields.append(
                    _field(
                        "p1.applicable_period",
                        block,
                        text,
                        values,
                        "日付範囲の候補です。期間の優先関係は自動判断しません。",
                        condition_dimensions,
                    )
                )

        if ISP_PATTERN.search(text):
            providers = tuple(
                sorted(
                    {
                        match.group(0)
                        for match in re.finditer(r"So-net|BIGLOBE|@nifty", text, re.IGNORECASE)
                    }
                )
            )
            fields.append(
                _field(
                    "p1.isp",
                    block,
                    text,
                    providers,
                    "ISP表記候補です。ISP別の条件適用は自動判断しません。",
                    condition_dimensions,
                )
            )

        if re.search(r"金額|手数料|報奨|支援金|円\s*/", text):
            amounts = tuple(
                    format(
                        Decimal(match["amount"].replace(",", "").replace("¥", "").strip()),
                        "f",
                    )
                + (f"/{match['unit']}" if match['unit'] else "")
                for match in MONEY_PATTERN.finditer(text)
            )
            if amounts:
                fields.append(
                    _field(
                        "p1.commission_amount",
                        block,
                        text,
                        amounts,
                        "金額候補です。対象サービス・税区分・適用条件は要照合です。",
                        condition_dimensions,
                    )
                )

        if ACHIEVEMENT_PATTERN.search(text):
            fields.append(
                _field(
                    "p1.achievement_condition",
                    block,
                    text,
                    (text.strip(),),
                    "成果条件文の候補です。条件の充足判定は行いません。",
                    condition_dimensions,
                )
            )

        if LADDER_PATTERN.search(text):
            fields.append(
                _field(
                    "p1.ladder",
                    block,
                    text,
                    (text.strip(),),
                    "ラダー候補です。閾値と報酬の組み合わせは要照合です。",
                    condition_dimensions,
                )
            )

        if CLAWBACK_PATTERN.search(text):
            fields.append(
                _field(
                    "p1.clawback_condition",
                    block,
                    text,
                    (text.strip(),),
                    "戻入候補です。対象・期間・上限は自動判断しません。",
                    condition_dimensions,
                )
            )

    return fields


def _field(
    item_id: str,
    block: EvidenceBlock,
    raw_value: str,
    normalized_values: tuple[str, ...],
    reason: str,
    condition_dimensions: tuple[ConditionDimension, ...] = (),
) -> NormalizedField:
    definition = DELIVERABLE_ITEM_ID_DICTIONARY[item_id]
    reasons = [reason]
    if not definition.approved:
        reasons.append("成果項目IDは暫定で、設計書承認待ちです。")
    if len(normalized_values) > 1:
        reasons.append("同一証拠位置に複数値があり、選択を保留しています。")
    return NormalizedField(
        item_id=item_id,
        raw_value=block.text,
        normalized_values=normalized_values,
        source_path=block.source_path,
        location=block.location,
        status="hold",
        hold_reasons=tuple(reasons),
        condition_dimensions=condition_dimensions,
    )


def _context_region(block: EvidenceBlock) -> str:
    location = block.location
    if location.startswith("page:"):
        return location
    if location.startswith("sheet:"):
        return location.split(";cells:", maxsplit=1)[0]
    if location.startswith("table:"):
        return location.split(";row:", maxsplit=1)[0]
    if location.startswith("paragraph:"):
        return "paragraphs"
    return location


def _extract_condition_dimensions(text: str) -> tuple[ConditionDimension, ...]:
    dimension_patterns: tuple[tuple[str, tuple[tuple[str, str], ...]], ...] = (
        (
            "product_family",
            (
                (r"au\s*ひかり\s*プラス", "auひかりプラス"),
                (r"au\s*ひかり(?!\s*プラス)", "auひかり"),
            ),
        ),
        (
            "plus_type",
            tuple((rf"プラス\s*{number}", f"プラス{number}") for number in ("1", "2", "3"))
            + tuple((f"プラス{number}", f"プラス{number}") for number in ("①", "②", "③")),
        ),
        (
            "isp",
            (
                (r"So-net", "So-net"),
                (r"BIGLOBE", "BIGLOBE"),
                (r"@nifty", "@nifty"),
            ),
        ),
        ("dwelling_type", ((r"ホーム", "ホーム"), (r"マンション", "マンション"))),
        ("speed", ((r"(?<!\d)10\s*G", "10G"), (r"(?<!\d)1\s*G", "1G"))),
        (
            "trigger",
            (
                (r"申込", "申込"),
                (r"開通", "開通"),
                (r"5G", "5G"),
                (r"FTTH", "FTTH"),
            ),
        ),
        (
            "access_type",
            ((r"事業者変更", "事業者変更"), (r"新規", "新規")),
        ),
    )
    dimensions: list[ConditionDimension] = []
    for name, patterns in dimension_patterns:
        values = tuple(
            value for pattern, value in patterns if re.search(pattern, text, re.IGNORECASE)
        )
        if values:
            dimensions.append(ConditionDimension(name=name, values=values))
    return tuple(dimensions)