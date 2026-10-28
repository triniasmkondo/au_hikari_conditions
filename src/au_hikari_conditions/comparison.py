from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from .document_reader import read_document
from .normalization import NormalizedField, normalize_blocks


@dataclass(frozen=True)
class CandidateDifference:
    item_id: str
    status: str
    previous_values: tuple[str, ...]
    current_values: tuple[str, ...]
    previous_evidence: tuple[str, ...]
    current_evidence: tuple[str, ...]
    hold_reason: str


@dataclass(frozen=True)
class DocumentComparison:
    status: str
    previous_path: str | None
    current_path: str | None
    differences: tuple[CandidateDifference, ...]
    hold_reason: str | None = None


def compare_fields(
    previous: list[NormalizedField],
    current: list[NormalizedField],
    *,
    previous_path: str,
    current_path: str,
) -> DocumentComparison:
    previous_by_id = _group_fields(previous)
    current_by_id = _group_fields(current)
    differences: list[CandidateDifference] = []

    for item_id in sorted(previous_by_id.keys() | current_by_id.keys()):
        old_fields = previous_by_id.get(item_id, [])
        new_fields = current_by_id.get(item_id, [])
        old_values = _candidate_values(old_fields)
        new_values = _candidate_values(new_fields)
        if not old_values:
            status = "added_candidate"
        elif not new_values:
            status = "removed_candidate"
        elif old_values == new_values:
            status = "unchanged_candidate"
        else:
            status = "changed_candidate"

        reasons = {
            reason
            for field in old_fields + new_fields
            for reason in field.hold_reasons
        }
        reasons.add("候補差分であり、正式ID・業務上の意味・適用版は未承認です。")
        differences.append(
            CandidateDifference(
                item_id=item_id,
                status=status,
                previous_values=old_values,
                current_values=new_values,
                previous_evidence=tuple(
                    sorted({f"{field.source_path}#{field.location}" for field in old_fields})
                ),
                current_evidence=tuple(
                    sorted({f"{field.source_path}#{field.location}" for field in new_fields})
                ),
                hold_reason="; ".join(sorted(reasons)),
            )
        )

    return DocumentComparison(
        status="compared",
        previous_path=previous_path,
        current_path=current_path,
        differences=tuple(differences),
    )


def compare_upper_campaign_notices(project_root: Path) -> DocumentComparison:
    upper_root = project_root / "01_上位原本"
    previous = list((upper_root / "2026-10").glob("*_キャンペーン通知書_20261002修正.pdf"))
    current = list((upper_root / "2026-11").glob("*_キャンペーン通知書_20261028追加.pdf"))
    if len(previous) != 1 or len(current) != 1:
        return DocumentComparison(
            status="unverified",
            previous_path=previous[0].relative_to(project_root).as_posix() if previous else None,
            current_path=current[0].relative_to(project_root).as_posix() if current else None,
            differences=(),
            hold_reason="比較対象の10月修正版または11月追加版を一意に特定できません。",
        )

    previous_blocks = read_document(previous[0], project_root=project_root)
    current_blocks = read_document(current[0], project_root=project_root)
    return compare_fields(
        normalize_blocks(previous_blocks),
        normalize_blocks(current_blocks),
        previous_path=previous[0].relative_to(project_root).as_posix(),
        current_path=current[0].relative_to(project_root).as_posix(),
    )


def _group_fields(fields: list[NormalizedField]) -> dict[str, list[NormalizedField]]:
    grouped: dict[str, list[NormalizedField]] = defaultdict(list)
    for field in fields:
        grouped[field.item_id].append(field)
    return grouped


def _candidate_values(fields: list[NormalizedField]) -> tuple[str, ...]:
    return tuple(sorted({value for field in fields for value in field.normalized_values}))