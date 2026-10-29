from __future__ import annotations

import hashlib
import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from au_hikari_conditions.document_reader import read_document
from au_hikari_conditions.normalization import normalize_blocks


SOURCE_SUFFIXES = {".pdf", ".xlsx", ".docx"}
APPROVAL_STATUSES = ("UNREVIEWED", "PENDING", "APPROVED", "REJECTED")
TAX_BASES = ("UNKNOWN", "EXCLUSIVE", "INCLUSIVE", "EXEMPT")
AMOUNT_STATES = ("UNKNOWN", "NOT_PUBLISHED", "DASH", "ZERO", "VALUE")

INPUT_TEMPLATE_NAME = "承認済み業務データ入力テンプレート_v4.xlsx"
EXTRACTION_NAME = "抽出候補_2026-10代理店32件_11月上位原本.xlsx"
DELIVERY_TEMPLATE_NAME = "配信管理入力テンプレート_架空サンプル_v3.xlsx"

INPUT_SHEETS: dict[str, tuple[str, ...]] = {
    "使い方": (
        "このブックは承認済み業務データの入力用テンプレートです。",
        "空欄・不明値は空欄のままにし、0円・未掲載・「－」で代用しないでください。",
        "金額・条件対応・掲載可否は、根拠資料、承認者、承認日時が記録されるまで確定計算に使いません。",
        "候補抽出ブックは別ファイルです。すべて未承認候補であり、このテンプレートに承認済みとして転記しないでください。",
        "P2計算は前回条件・承認管理を照合します。承認日時はタイムゾーン付きISO-8601、record_hashは業務レコードのSHA-256を入力してください。",
        "粗利原価は対象月の全額を1行（または共通条件が一致する原価内訳行）で入力します。複数行の原価は合算します。",
        "条件対応には前回月・対象月それぞれの適用期間を登録し、成果条件・期間ルールを承認済み同等として確認してください。",
        "実データのブックは会社承認済みのGit外保管先で管理してください。リポジトリには空テンプレートのみ配置します。",
        "11月代理店支払額の確定・帳票生成は、必要項目が承認済みとなり、条件・例外・掲載可否のQAが完了するまで禁止です。",
        "承認状態の値: UNREVIEWED / PENDING / APPROVED / REJECTED",
        "税区分の値: UNKNOWN / EXCLUSIVE / INCLUSIVE / EXEMPT。UNKNOWNは計算不可です。",
    ),
    "粗利原価": (
        "record_id",
        "agency_id",
        "agency_name",
        "condition_item_id",
        "condition_version",
        "mapping_id",
        "upstream_amount",
        "upstream_amount_state",
        "upstream_tax_basis",
        "approved_target_gross_margin",
        "gross_margin_state",
        "gross_margin_tax_basis",
        "display_policy",
        "direct_cost_component",
        "direct_cost_description",
        "cost_allocation_basis",
        "attributable_direct_cost",
        "direct_cost_state",
        "direct_cost_tax_basis",
        "minimum_gross_margin",
        "valid_from",
        "valid_until",
        "source_document",
        "evidence_location",
        "approval_status",
        "approved_by",
        "approved_at",
        "approval_reason",
    ),
    "条件対応": (
        "mapping_id",
        "source_document",
        "source_location",
        "source_candidate_id",
        "upstream_item_id",
        "agency_id",
        "agency_item_id",
        "product_family",
        "plus_type",
        "isp",
        "dwelling_type",
        "speed",
        "access_type",
        "trigger",
        "option_code",
        "campaign_code",
        "effective_from",
        "effective_until",
        "period_rule",
        "payment_timing",
        "exclusions",
        "reversal_terms",
        "equivalence_review_status",
        "approval_status",
        "approved_by",
        "approved_at",
        "review_reason",
    ),
    "代理店別例外": (
        "exception_id",
        "agency_id",
        "condition_item_id_or_rule_group",
        "priority",
        "rule_type",
        "rule_value",
        "display_policy",
        "reason",
        "source_document",
        "evidence_location",
        "approval_status",
        "approved_by",
        "approved_at",
        "valid_from",
        "valid_until",
    ),
    "前回条件": (
        "record_id",
        "target_period",
        "agency_id",
        "condition_item_id",
        "mapping_id",
        "condition_version",
        "previous_agency_reward",
        "amount_state",
        "tax_basis",
        "source_document",
        "evidence_location",
        "approval_status",
        "approved_by",
        "approved_at",
        "approval_reason",
    ),
    "承認管理": (
        "approval_id",
        "batch_id",
        "agency_id",
        "record_type",
        "record_id",
        "record_hash",
        "approval_status",
        "requested_by",
        "requested_at",
        "approved_by",
        "approved_at",
        "decision_reason",
        "supersedes_approval_id",
    ),
}

DELIVERY_SHEETS: dict[str, tuple[str, ...]] = {
    "使い方": (
        "このブックの会社名・メールアドレスは架空のサンプルで、配信には使用できません。",
        "実際の送付先・送信元・本文・履歴は会社承認済みGit外保管先に保存してください。",
        "外部送信はP4テスト環境・会社別照合・添付検証・承認スナップショット一致・P5明示承認後に限ります。",
        "宛先、宛名、TO/CC/送信元、本文、添付、テンプレート、条件マスター、適用期間が変われば承認を失効し再QA・再承認します。",
        "冪等性キーは batch_id + agency_id + artifact_hashes に基づきます。同一キーの二重送信は禁止です。",
        "タイムアウト等で送信結果が不明な場合、履歴照会を先に行い、盲目的に再送しません。",
        "ここにはZoho認証情報や実宛先を記録しないでください。",
    ),
    "送付先マスター": (
        "agency_id",
        "agency_name",
        "recipient_status",
        "to_addresses",
        "cc_addresses",
        "bcc_addresses",
        "sender_address",
        "reply_to_address",
        "salutation",
        "recipient_source",
        "approved_by",
        "approved_at",
        "valid_from",
        "valid_until",
        "last_reviewed_at",
        "notes",
    ),
    "本文テンプレート": (
        "template_id",
        "template_version",
        "doc_type",
        "subject_template",
        "body_template",
        "allowed_placeholders",
        "template_hash",
        "approval_status",
        "approved_by",
        "approved_at",
        "valid_from",
        "valid_until",
        "notes",
    ),
    "月次承認": (
        "batch_id",
        "target_period",
        "agency_id",
        "approval_status",
        "artifact_hashes",
        "recipient_snapshot_hash",
        "message_hash",
        "template_hash",
        "condition_master_hash",
        "qa_status",
        "requested_by",
        "requested_at",
        "approved_by",
        "approved_at",
        "approval_reason",
        "invalidated_at",
        "invalidation_reason",
    ),
    "送信履歴": (
        "delivery_id",
        "batch_id",
        "target_period",
        "agency_id",
        "idempotency_key",
        "delivery_status",
        "artifact_hashes",
        "recipient_snapshot_hash",
        "message_hash",
        "sender_address",
        "attempted_by",
        "attempted_at",
        "provider_message_id",
        "completed_at",
        "error_code",
        "error_detail",
        "result_checked_at",
        "audit_reference",
    ),
    "承認失効ルール": (
        "change_type",
        "approval_effect",
        "required_action",
        "automated_send_allowed",
    ),
}

DELIVERY_RULES = (
    (
        "recipient_or_salutation",
        "INVALIDATED",
        "Recipient snapshot refresh, company match, QA, and reapproval",
        "NO",
    ),
    (
        "to_cc_bcc_sender_reply_to",
        "INVALIDATED",
        "Verify approved recipient master, refresh snapshot, and reapprove",
        "NO",
    ),
    (
        "subject_or_body",
        "INVALIDATED",
        "Recalculate message hash, run leak checks, and reapprove",
        "NO",
    ),
    (
        "artifact_or_attachment",
        "INVALIDATED",
        "Revalidate hashes and company match, rerun QA, and reapprove",
        "NO",
    ),
    (
        "template_or_condition_master",
        "INVALIDATED",
        "Regenerate artifact, rerun QA, and reapprove",
        "NO",
    ),
    (
        "target_period_or_approval",
        "INVALIDATED",
        "Revalidate batch period and obtain new approval",
        "NO",
    ),
    (
        "provider_timeout_or_unknown_result",
        "RESULT_UNKNOWN",
        "Query delivery history before any retry",
        "NO",
    ),
    (
        "duplicate_idempotency_key",
        "BLOCKED",
        "Do not resend the same batch, agency, and artifact hashes",
        "NO",
    ),
)


def _style_sheet(worksheet, headers: tuple[str, ...] | None = None) -> None:
    worksheet.freeze_panes = "A2" if headers else "A1"
    worksheet.sheet_view.showGridLines = False
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    if headers:
        worksheet.auto_filter.ref = f"A1:{worksheet.cell(1, len(headers)).column_letter}1"
        worksheet.row_dimensions[1].height = 36
        for column_number, header in enumerate(headers, 1):
            worksheet.column_dimensions[
                worksheet.cell(1, column_number).column_letter
            ].width = min(max(len(header) + 3, 15), 34)


def _add_table(worksheet, name: str, headers: tuple[str, ...]) -> None:
    for column, header in enumerate(headers, 1):
        worksheet.cell(1, column, header)
    worksheet.append([None] * len(headers))
    table = Table(
        displayName=name,
        ref=f"A1:{worksheet.cell(2, len(headers)).column_letter}2",
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    worksheet.add_table(table)
    _style_sheet(worksheet, headers)


def _add_list_validation(worksheet, column: str, values: Iterable[str]) -> None:
    validation = DataValidation(
        type="list",
        formula1=f'"{",".join(values)}"',
        allow_blank=True,
    )
    validation.errorTitle = "選択値を確認してください"
    validation.error = "リストにある値を選択してください。"
    worksheet.add_data_validation(validation)
    validation.add(f"{column}2:{column}1000")


def _column_letter(headers: tuple[str, ...], header: str) -> str:
    return get_column_letter(headers.index(header) + 1)


def _add_instructions_sheet(workbook: Workbook, rows: tuple[str, ...]) -> None:
    worksheet = workbook.active
    worksheet.title = "使い方"
    worksheet.append(("区分", "説明"))
    for index, text in enumerate(rows, 1):
        worksheet.append((f"NOTE-{index:02d}", text))
    worksheet.column_dimensions["A"].width = 16
    worksheet.column_dimensions["B"].width = 110
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    for cell in worksheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(color="FFFFFF", bold=True)
    for row in worksheet.iter_rows(min_row=2):
        row[1].alignment = Alignment(wrap_text=True, vertical="top")
        worksheet.row_dimensions[row[0].row].height = 34


def create_approved_input_template(path: Path) -> None:
    workbook = Workbook()
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    _add_instructions_sheet(workbook, INPUT_SHEETS["使い方"])
    for index, (sheet_name, headers) in enumerate(
        (item for item in INPUT_SHEETS.items() if item[0] != "使い方"),
        start=1,
    ):
        worksheet = workbook.create_sheet(sheet_name)
        _add_table(worksheet, f"InputTable{index}", headers)
        _add_list_validation(
            worksheet,
            _column_letter(headers, "approval_status"),
            APPROVAL_STATUSES,
        )
        if sheet_name == "粗利原価":
            for name in (
                "upstream_tax_basis",
                "gross_margin_tax_basis",
                "direct_cost_tax_basis",
            ):
                _add_list_validation(worksheet, _column_letter(headers, name), TAX_BASES)
            for name in (
                "upstream_amount_state",
                "gross_margin_state",
                "direct_cost_state",
            ):
                _add_list_validation(
                    worksheet,
                    _column_letter(headers, name),
                    AMOUNT_STATES,
                )
            _add_list_validation(
                worksheet,
                _column_letter(headers, "display_policy"),
                ("PRESENT", "WITHHOLD", "UNDECIDED"),
            )
        elif sheet_name == "条件対応":
            _add_list_validation(
                worksheet,
                _column_letter(headers, "equivalence_review_status"),
                ("NOT_REVIEWED", "REVIEWED_EQUIVALENT", "DIFFERENT", "REVIEW_REQUIRED"),
            )
        elif sheet_name == "前回条件":
            _add_list_validation(
                worksheet,
                _column_letter(headers, "amount_state"),
                AMOUNT_STATES,
            )
            _add_list_validation(
                worksheet,
                _column_letter(headers, "tax_basis"),
                TAX_BASES,
            )
        elif sheet_name == "代理店別例外":
            _add_list_validation(
                worksheet,
                _column_letter(headers, "display_policy"),
                ("PRESENT", "WITHHOLD", "UNDECIDED"),
            )
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def create_delivery_template(path: Path) -> None:
    workbook = Workbook()
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    _add_instructions_sheet(workbook, DELIVERY_SHEETS["使い方"])
    for index, (sheet_name, headers) in enumerate(
        (item for item in DELIVERY_SHEETS.items() if item[0] != "使い方"),
        start=1,
    ):
        worksheet = workbook.create_sheet(sheet_name)
        _add_table(worksheet, f"DeliveryTable{index}", headers)
        if sheet_name == "送付先マスター":
            _add_list_validation(worksheet, "C", ("SAMPLE_ONLY_NOT_APPROVED", "PENDING", "APPROVED", "REJECTED"))
            # Reserved example.invalid addresses cannot route to a real mailbox.
            worksheet.cell(2, 1, "sample-agency-001")
            worksheet.cell(2, 2, "架空代理店サンプル合同会社")
            worksheet.cell(2, 3, "SAMPLE_ONLY_NOT_APPROVED")
            worksheet.cell(2, 4, "sample-recipient@example.invalid")
            worksheet.cell(2, 5, "sample-cc@example.invalid")
            worksheet.cell(2, 6, "sample-bcc@example.invalid")
            worksheet.cell(2, 7, "sample-sender@example.invalid")
            worksheet.cell(2, 8, "sample-reply@example.invalid")
            worksheet.cell(2, 9, "架空代理店サンプル合同会社 ご担当者様")
            worksheet.cell(2, 10, "fictional_sample_only")
            worksheet.cell(2, 16, "この行は架空サンプルです。送信禁止。")
        elif sheet_name == "本文テンプレート":
            _add_list_validation(worksheet, "H", ("SAMPLE_ONLY_NOT_APPROVED", "PENDING", "APPROVED", "REJECTED"))
            sample = (
                "sample-template-001",
                "0.1",
                "conditions",
                "【サンプル】{target_period} 条件表のご案内",
                "架空代理店サンプル合同会社 ご担当者様\n\n"
                "{target_period}分の条件表（架空サンプル）をお送りします。\n"
                "本メールと添付はテスト用であり、実際の条件を示しません。\n\n"
                "架空運用チーム\n送信禁止",
                "target_period,agency_name",
                "SAMPLE_TEMPLATE_HASH_NOT_APPROVED",
                "SAMPLE_ONLY_NOT_APPROVED",
                "sample-approver",
                "2026-10-29",
                "2026-11-01",
                "2026-11-30",
                "架空サンプル本文。送信禁止。",
            )
            for column, value in enumerate(sample, 1):
                worksheet.cell(2, column, value)
        elif sheet_name == "月次承認":
            _add_list_validation(worksheet, "D", ("SAMPLE_ONLY_NOT_APPROVED", "PENDING", "APPROVED", "REJECTED", "HOLD", "RETURNED"))
            _add_list_validation(worksheet, "J", ("NOT_RUN", "PENDING", "PASSED", "FAILED"))
            sample = (
                "sample-batch-2026-11",
                "2026-11",
                "sample-agency-001",
                "SAMPLE_ONLY_NOT_APPROVED",
                "SAMPLE_ARTIFACT_HASH",
                "SAMPLE_RECIPIENT_HASH",
                "SAMPLE_MESSAGE_HASH",
                "SAMPLE_TEMPLATE_HASH",
                "SAMPLE_MASTER_HASH",
                "NOT_RUN",
                "sample-requester",
                "2026-10-29",
                None,
                None,
                "サンプルのため承認なし",
                None,
                None,
            )
            for column, value in enumerate(sample, 1):
                worksheet.cell(2, column, value)
        elif sheet_name == "送信履歴":
            _add_list_validation(worksheet, "F", ("NOT_SENT_SAMPLE_ONLY", "PENDING", "SENT", "FAILED", "RESULT_UNKNOWN", "BLOCKED"))
            sample = (
                "sample-delivery-001",
                "sample-batch-2026-11",
                "2026-11",
                "sample-agency-001",
                "SAMPLE_IDEMPOTENCY_KEY",
                "NOT_SENT_SAMPLE_ONLY",
                "SAMPLE_ARTIFACT_HASH",
                "SAMPLE_RECIPIENT_HASH",
                "SAMPLE_MESSAGE_HASH",
                "sample-sender@example.invalid",
                "sample-operator",
                "2026-10-29",
                None,
                None,
                None,
                None,
                None,
                "架空データ。実送信なし。",
            )
            for column, value in enumerate(sample, 1):
                worksheet.cell(2, column, value)
        elif sheet_name == "承認失効ルール":
            for row in DELIVERY_RULES:
                worksheet.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _agency_name(path: Path, project_root: Path) -> str:
    if path.relative_to(project_root).parts[0] != "03_条件表":
        return "UPSTREAM"
    match = re.search(r"【202610】\s*【(.+?)】", path.name)
    if not match:
        return "REVIEW_REQUIRED"
    name = re.sub(r"\s*(御中|様)\s*$", "", match.group(1)).strip()
    return name or "REVIEW_REQUIRED"


def _source_type(path: Path, project_root: Path) -> str:
    top_level = path.relative_to(project_root).parts[0]
    return "agency_prior_approved_candidate" if top_level == "03_条件表" else "upstream_source"


def _source_paths(project_root: Path) -> list[Path]:
    roots = (
        project_root / "03_条件表" / "2026-10",
        project_root / "01_上位原本" / "2026-11",
    )
    return sorted(
        path
        for root in roots
        for path in root.rglob("*")
        if path.is_file() and path.suffix.lower() in SOURCE_SUFFIXES
    )


def _append_rows(worksheet, rows: list[tuple[object, ...]]) -> None:
    for row in rows:
        worksheet.append(row)


def create_extraction_workbook(project_root: Path, path: Path) -> dict[str, int]:
    source_paths = _source_paths(project_root)
    candidates: list[tuple[object, ...]] = []
    inventory: list[tuple[object, ...]] = []
    agency_documents: Counter[str] = Counter()
    extraction_errors: list[tuple[str, str]] = []

    for source_path in source_paths:
        relative_path = source_path.relative_to(project_root).as_posix()
        agency = _agency_name(source_path, project_root)
        inventory.append(
            (
                source_path.stem,
                agency,
                _source_type(source_path, project_root),
                source_path.suffix.lower(),
                relative_path,
                _sha256(source_path),
                "RECEIVED_NOT_APPROVED",
                "未承認候補の抽出のみ。承認状態や条件対応を推定しない。",
            )
        )
        if agency not in {"UPSTREAM", "REVIEW_REQUIRED"}:
            agency_documents[agency] += 1
        try:
            blocks = read_document(source_path, project_root=project_root)
            fields = normalize_blocks(blocks)
        except Exception as error:
            extraction_errors.append(
                (relative_path, f"{type(error).__name__}: {error}")
            )
            continue
        block_text = {
            (block.source_path, block.location): block.text for block in blocks
        }
        for number, normalized_field in enumerate(fields, 1):
            candidate_id = f"{_sha256(source_path)[:12]}-{number:04d}"
            values = json.dumps(
                normalized_field.normalized_values,
                ensure_ascii=False,
                separators=(",", ":"),
            )
            dimensions = json.dumps(
                {
                    dimension.name: dimension.values
                    for dimension in normalized_field.condition_dimensions
                },
                ensure_ascii=False,
                separators=(",", ":"),
            )
            excerpt = block_text.get(
                (normalized_field.source_path, normalized_field.location),
                normalized_field.raw_value,
            )
            candidates.append(
                (
                    candidate_id,
                    agency,
                    _source_type(source_path, project_root),
                    source_path.stem,
                    relative_path,
                    _sha256(source_path),
                    normalized_field.item_id,
                    values,
                    dimensions,
                    normalized_field.location,
                    excerpt,
                    "NOT_REVIEWED",
                    "UNAPPROVED",
                    None,
                    None,
                    "; ".join(normalized_field.hold_reasons),
                )
            )

    unknown_rows: list[tuple[object, ...]] = []
    for agency in sorted(agency_documents):
        for item, reason in (
            ("approved_target_gross_margin", "承認済み代理店別・成果項目別粗利データが未提供です。"),
            ("attributable_direct_cost", "項目別直接原価の金額・内訳・承認根拠が未提供です。"),
            ("tax_basis", "上位報酬・粗利・直接原価の税区分が確認できません。"),
            ("condition_equivalence", "前回承認条件と11月上位条件の対応・同等性が未レビューです。"),
            ("agency_exception_and_display", "代理店例外、掲載可否、告知義務の承認台帳が未提供です。"),
        ):
            unknown_rows.append(
                (
                    agency,
                    item,
                    "NEEDS_REVIEW",
                    reason,
                    "推測値・0円で補完禁止。承認まで11月確定計算不可。",
                    "UNAPPROVED",
                )
            )
    unknown_rows.extend(
        (
            ("ALL_AGENCIES", "upstream_tax_basis", "NEEDS_REVIEW", "11月上位2原本の金額税区分を確認してください。", "税抜換算・税区分一致確認まで計算保留。", "UNAPPROVED"),
            ("ALL_AGENCIES", "approved_previous_terms", "NEEDS_REVIEW", "10月上位条件原本がなく、10月→11月の上位改定差分は未検証です。", "変更なしと扱わず、改定差分の確定は保留。", "UNAPPROVED"),
            ("ALL_AGENCIES", "recipient_and_delivery_approval", "NEEDS_REVIEW", "正式宛先・宛名・CC・送信元の承認マスターが未提供です。", "P4/P5承認完了まで外部送信不可。", "UNAPPROVED"),
        )
    )

    workbook = Workbook()
    instructions = workbook.active
    instructions.title = "判定ルール"
    _append_rows(
        instructions,
        [
            ("区分", "説明"),
            ("候補状態", "抽出候補は全件 UNAPPROVED / NOT_REVIEWED。資料に書かれた値は承認済み業務データではありません。"),
            ("根拠", "各候補に相対パス、SHA-256、ページ／セル位置、原文抜粋を付与。内容と対応付けは人が確認してください。"),
            ("粗利・原価", "値・税区分は候補入力せず空欄。根拠付きの承認入力が必要です。0円を既定値にしません。"),
            ("上位・代理店対応", "類似の項目名や抽出軸だけで対応確定しません。レビューと承認まで未対応のままです。"),
            ("対象資料", "2026-10代理店条件表32件と2026-11上位原本2件。10月上位原本は別管理であり、今回の上位改定比較は未完了です。"),
            ("保存", "この抽出結果には社内報酬等の実データが含まれます。Git外の会社承認済み保管先に保管してください。"),
        ],
    )
    sources_sheet = workbook.create_sheet("資料一覧")
    source_headers = (
        "source_document_id",
        "agency_name_or_scope",
        "source_type",
        "format",
        "source_path",
        "sha256",
        "record_status",
        "review_note",
    )
    for column, header in enumerate(source_headers, 1):
        sources_sheet.cell(1, column, header)
    _append_rows(sources_sheet, inventory)
    candidates_sheet = workbook.create_sheet("抽出候補")
    candidate_headers = (
        "candidate_id",
        "agency_name_or_scope",
        "source_type",
        "source_document_id",
        "source_path",
        "source_sha256",
        "candidate_item_id",
        "candidate_values_json",
        "condition_dimensions_json",
        "evidence_location",
        "evidence_excerpt",
        "mapping_review_status",
        "approval_status",
        "approved_by",
        "approved_at",
        "hold_reason",
    )
    for column, header in enumerate(candidate_headers, 1):
        candidates_sheet.cell(1, column, header)
    _append_rows(candidates_sheet, candidates)
    unknown_sheet = workbook.create_sheet("判断待ち一覧")
    unknown_headers = (
        "agency_or_scope",
        "unknown_item",
        "status",
        "why_unknown",
        "decision_gate",
        "approval_status",
    )
    for column, header in enumerate(unknown_headers, 1):
        unknown_sheet.cell(1, column, header)
    _append_rows(unknown_sheet, unknown_rows)
    errors_sheet = workbook.create_sheet("読取エラー")
    errors_sheet.append(("source_path", "error"))
    _append_rows(errors_sheet, extraction_errors)

    for index, worksheet in enumerate(workbook.worksheets, 1):
        headers = (
            source_headers
            if worksheet.title == "資料一覧"
            else candidate_headers
            if worksheet.title == "抽出候補"
            else unknown_headers
            if worksheet.title == "判断待ち一覧"
            else None
        )
        if headers:
            table_name = f"ExtractionTable{index}"
            end_row = max(worksheet.max_row, 2)
            table = Table(
                displayName=table_name,
                ref=f"A1:{worksheet.cell(1, len(headers)).column_letter}{end_row}",
            )
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showRowStripes=True,
                showColumnStripes=False,
            )
            worksheet.add_table(table)
        _style_sheet(worksheet, headers)
        if worksheet.title in {"抽出候補", "読取エラー"}:
            worksheet.freeze_panes = "A2"
            worksheet.column_dimensions["K"].width = 90
            worksheet.column_dimensions["P"].width = 80
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
        if worksheet.title == "判定ルール":
            worksheet.column_dimensions["A"].width = 20
            worksheet.column_dimensions["B"].width = 110
            for row in worksheet.iter_rows(min_row=2):
                row[1].alignment = Alignment(wrap_text=True, vertical="top")
    candidates_sheet.auto_filter.ref = (
        f"A1:P{max(candidates_sheet.max_row, 1)}"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)

    return {
        "source_document_count": len(source_paths),
        "agency_document_count": sum(agency_documents.values()),
        "agency_count": len(agency_documents),
        "candidate_count": len(candidates),
        "unknown_item_count": len(unknown_rows),
        "read_error_count": len(extraction_errors),
    }


def _verify_template(path: Path, required_sheets: set[str]) -> None:
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        if not required_sheets.issubset(workbook.sheetnames):
            raise ValueError(f"{path.name} is missing required sheets")
        if not workbook.calculation.fullCalcOnLoad:
            raise ValueError(f"{path.name} must recalculate on open")
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser()
    output_modes = parser.add_mutually_exclusive_group()
    output_modes.add_argument(
        "--templates-only",
        action="store_true",
        help="Create only blank input and fictional delivery templates; do not extract documents.",
    )
    output_modes.add_argument(
        "--input-template-only",
        action="store_true",
        help="Create only the new blank P2 input template without touching other outputs.",
    )
    arguments = parser.parse_args()
    input_template = PROJECT_ROOT / "02_取得データ" / INPUT_TEMPLATE_NAME
    extraction_path = PROJECT_ROOT / "02_取得データ" / EXTRACTION_NAME
    delivery_template = PROJECT_ROOT / "05_配信管理" / DELIVERY_TEMPLATE_NAME
    if arguments.input_template_only:
        outputs = (input_template,)
    elif arguments.templates_only:
        outputs = (input_template, delivery_template)
    else:
        outputs = (input_template, extraction_path, delivery_template)
    existing_outputs = [str(path) for path in outputs if path.exists()]
    if existing_outputs:
        print("既存ファイルを保護するため、上書きせず停止します:")
        for path in existing_outputs:
            print(f"- {path}")
        return 1

    if not input_template.exists():
        create_approved_input_template(input_template)
    if not delivery_template.exists():
        create_delivery_template(delivery_template)
    _verify_template(
        input_template,
        {"使い方", "粗利原価", "条件対応", "代理店別例外", "前回条件", "承認管理"},
    )
    _verify_template(
        delivery_template,
        {"使い方", "送付先マスター", "本文テンプレート", "月次承認", "送信履歴", "承認失効ルール"},
    )
    if arguments.input_template_only:
        print(
            json.dumps(
                {"outputs": [str(input_template)]},
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    if arguments.templates_only:
        print(
            json.dumps(
                {"outputs": [str(input_template), str(delivery_template)]},
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    counts = create_extraction_workbook(PROJECT_ROOT, extraction_path)
    print(json.dumps({"outputs": [str(path) for path in outputs], **counts}, ensure_ascii=False, indent=2))
    return 0 if counts["read_error_count"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
