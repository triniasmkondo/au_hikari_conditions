from __future__ import annotations

from collections import Counter
from pathlib import Path

from .comparison import compare_upper_campaign_notices
from .document_reader import read_document
from .models import DELIVERABLE_ITEM_ID_DICTIONARY
from .normalization import normalize_blocks


SUPPORTED_SUFFIXES = {".pdf", ".xlsx", ".docx"}


def build_p1_report(project_root: Path) -> dict[str, object]:
    source_paths = sorted(
        path
        for folder in ("01_上位原本", "03_条件表")
        for path in (project_root / folder).rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES
    )
    source_count_by_folder = Counter(path.relative_to(project_root).parts[0] for path in source_paths)
    source_count_by_format = Counter(path.suffix.lower() for path in source_paths)
    field_counts: Counter[str] = Counter()
    unrecognized_documents: list[str] = []
    unreadable_documents: list[dict[str, str]] = []
    ambiguous_fields = 0

    for path in source_paths:
        relative_path = path.relative_to(project_root).as_posix()
        try:
            fields = normalize_blocks(read_document(path, project_root=project_root))
        except Exception as error:
            unreadable_documents.append(
                {"path": relative_path, "error": f"{type(error).__name__}: {error}"}
            )
            continue
        if not fields:
            unrecognized_documents.append(relative_path)
        for field in fields:
            field_counts[field.item_id] += 1
            if len(field.normalized_values) > 1:
                ambiguous_fields += 1

    pending_item_ids = sorted(
        item_id
        for item_id, definition in DELIVERABLE_ITEM_ID_DICTIONARY.items()
        if not definition.approved
    )
    missing_item_candidates = sorted(
        item_id for item_id in DELIVERABLE_ITEM_ID_DICTIONARY if not field_counts[item_id]
    )
    october_notices = sorted(
        path.relative_to(project_root).as_posix()
        for path in (project_root / "01_上位原本" / "2026-10").glob("*キャンペーン通知書*.pdf")
    )
    spec_files = sorted(
        path.relative_to(project_root).as_posix()
        for path in (project_root / "docs").rglob("*")
        if path.is_file() and path.suffix.lower() in {".docx", ".pdf", ".md"}
    )
    comparison = compare_upper_campaign_notices(project_root)

    blockers: list[str] = []
    if not spec_files:
        blockers.append("docs/ にP1完了ゲートを定義した設計書がありません。")
    if pending_item_ids:
        blockers.append("成果項目IDが未承認です: " + ", ".join(pending_item_ids))
        blockers.append("正式な項目対応表がないため、代理店資料と上位原本の実値照合は保留です。")
    if missing_item_candidates:
        blockers.append("資料から候補を抽出できない項目があります: " + ", ".join(missing_item_candidates))
    if unrecognized_documents:
        blockers.append(f"項目候補を抽出できない資料が{len(unrecognized_documents)}件あります。")
    if unreadable_documents:
        blockers.append(f"解析できない資料が{len(unreadable_documents)}件あります。")
    if len(october_notices) > 1:
        blockers.append("10月上位原本に複数のキャンペーン通知書があり、正式版の優先関係が未確定です。")
    if comparison.status != "compared":
        blockers.append("10月・11月の上位条件比較を完了できませんでした。")

    return {
        "p1_gate": "not_met" if blockers else "met",
        "source_document_count": len(source_paths),
        "source_document_count_by_folder": dict(sorted(source_count_by_folder.items())),
        "source_document_count_by_format": dict(sorted(source_count_by_format.items())),
        "normalized_candidate_count": sum(field_counts.values()),
        "candidate_count_by_item_id": dict(sorted(field_counts.items())),
        "held_candidate_count": sum(field_counts.values()),
        "ambiguous_field_count": ambiguous_fields,
        "pending_item_ids": pending_item_ids,
        "agency_to_upper_matching": {
            "status": "held" if pending_item_ids else "ready",
            "agency_document_count": source_count_by_folder.get("03_条件表", 0),
            "matched_field_count": 0 if pending_item_ids else None,
            "hold_reason": (
                "正式な成果項目IDと上位原本との対応規則が未承認です。"
                if pending_item_ids
                else None
            ),
        },
        "missing_item_candidates": missing_item_candidates,
        "unrecognized_documents": unrecognized_documents,
        "unreadable_documents": unreadable_documents,
        "october_campaign_notice_versions": october_notices,
        "design_spec_files": spec_files,
        "upper_condition_comparison": {
            "status": comparison.status,
            "previous_path": comparison.previous_path,
            "current_path": comparison.current_path,
            "candidate_difference_count": len(comparison.differences),
            "candidate_status_counts": dict(
                sorted(Counter(item.status for item in comparison.differences).items())
            ),
            "evidence_reference_count": sum(
                len(item.previous_evidence) + len(item.current_evidence)
                for item in comparison.differences
            ),
            "hold_reason": comparison.hold_reason,
        },
        "blockers": blockers,
    }