"""Optional P2 v4 campaign tiers and fail-closed reward-engine adapters.

Confirmed current prices are distinct from previous prices. Neither extraction
nor evaluation grants payment approval or selects a tier for a real opening.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
from decimal import Decimal, InvalidOperation
from datetime import timedelta
from pathlib import Path
import hashlib

from openpyxl import load_workbook

from .calculation import LadderMode
from .p2 import (
    APPROVAL_REQUIRED_HEADERS, P2InputError, _approval_evidence,
    _parse_date, _parse_target_period, _read_amount, _read_sheet, _text,
)
from .reward_design import (
    AgencyRewardProfile, CampaignSchedule, CampaignTierInput,
    RewardDesignResult, VolumeScenario, design_agency_rewards,
)

CAMPAIGN_SHEET = "キャンペーン階層"
CAMPAIGN_HEADERS = tuple("""
record_id agency_id campaign_code campaign_name condition_item_id
threshold_value threshold_values threshold_metric threshold_unit counting_scope payout_scope
eligible_opening_classes threshold_application valid_from valid_until
eligibility_terms exclusions payment_timing
upstream_unit_amount upstream_amount_state upstream_tax_basis
agency_unit_amount agency_amount_state agency_tax_basis
previous_agency_unit_amount previous_amount_state previous_tax_basis previous_period
attributable_direct_cost direct_cost_state direct_cost_tax_basis
source_document source_sha256 evidence_location
agency_source_document agency_source_sha256 agency_evidence_location
previous_source_document previous_source_sha256 previous_evidence_location
cost_source_document cost_evidence_location
approval_status approved_by approved_at approval_reason
""".split())

_GROUP_FIELDS = tuple("""
campaign_code campaign_name threshold_values threshold_metric threshold_unit counting_scope payout_scope
eligible_opening_classes threshold_application valid_from valid_until
eligibility_terms exclusions payment_timing
""".split())


@dataclass(frozen=True)
class CampaignLoadReport:
    agency_id: str
    target_period: str
    rate_basis: str
    campaigns: tuple[CampaignSchedule, ...]
    hold_reasons: tuple[str, ...]
    record_count: int

    @property
    def status(self) -> str:
        return "hold" if self.hold_reasons else "ready_for_review"


def _number(value: object) -> int | None:
    if isinstance(value, bool):
        return None
    try:
        number = Decimal(str(value))
        if not number.is_finite() or number <= 0 or number != number.to_integral_value():
            return None
        return int(number)
    except (InvalidOperation, ValueError, TypeError):
        return None


def _source_reasons(row: dict, prefix: str, base: Path) -> list[str]:
    name = _text(row.get(prefix + "source_document"))
    digest = _text(row.get(prefix + "source_sha256")).lower()
    location = _text(row.get(prefix + "evidence_location"))
    if not name or not location or len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
        return [f"{prefix or 'upstream_'}原本のパス・SHA-256・根拠位置が不足／不正です。"]
    source = Path(name)
    if not source.is_absolute():
        source = base / source
    try:
        actual = hashlib.sha256(source.read_bytes()).hexdigest()
    except OSError:
        return [f"{prefix or 'upstream_'}原本を再検証できません。"]
    return [] if actual == digest else [f"{prefix or 'upstream_'}原本のSHA-256が一致しません。"]


def load_campaign_schedules(
    input_path: Path, agency_id: str, target_period: str, *, rate_basis: str = "confirmed",
) -> CampaignLoadReport:
    """Read approved tier rows. Any invalid target row withholds ALL schedules.

    confirmed: current agency prices for fixed-price margin verification.
    previous: actual previous-month prices for least-change design.
    Source paths may be absolute or relative to the input workbook.
    """
    if rate_basis not in {"confirmed", "previous"}:
        raise ValueError("rate_basis must be confirmed or previous")
    if not agency_id.strip():
        raise P2InputError("代理店IDがありません。")
    start, end = _parse_target_period(target_period)
    previous_period_text = (start - timedelta(days=1)).strftime("%Y-%m")
    workbook = load_workbook(input_path, read_only=True, data_only=False)
    try:
        rows = _read_sheet(workbook, CAMPAIGN_SHEET, CAMPAIGN_HEADERS)
        approvals = _read_sheet(workbook, "承認管理", APPROVAL_REQUIRED_HEADERS)
    finally:
        workbook.close()
    reasons: list[str] = []
    selected: list[dict] = []
    for row in rows:
        if not _text(row.get("agency_id")):
            reasons.append("キャンペーン階層に代理店IDがない行があります。")
            continue
        if _text(row.get("agency_id")) != agency_id:
            continue
        first, last = _parse_date(row.get("valid_from")), _parse_date(row.get("valid_until"))
        if first is None or last is None or first > last:
            reasons.append("キャンペーン階層の適用期間が不足／不正です。")
            continue
        if last < start or first > end:
            continue
        selected.append(row)
        if first > start or last < end:
            reasons.append("キャンペーン階層は対象月全期間をカバーしていません。月途中の条件変更は保留です。")
    if not selected:
        reasons.append("対象代理店・対象月のキャンペーン階層がありません。")
    ids = [_text(row.get("record_id")) for row in rows]
    if any(not value for value in ids) or len(ids) != len(set(ids)):
        reasons.append("キャンペーン階層のrecord_idが欠落／重複しています。")
    item_ids = [_text(row.get("condition_item_id")) for row in selected]
    if any(not value for value in item_ids) or len(item_ids) != len(set(item_ids)):
        reasons.append("キャンペーン階層の正式成果項目IDが欠落／重複しています。")
    groups: dict[str, list[tuple[dict, CampaignTierInput]]] = {}
    for row in selected:
        label = _text(row.get("record_id")) or "(ID missing)"
        row_reasons: list[str] = []
        for key in _GROUP_FIELDS:
            if not _text(row.get(key)):
                row_reasons.append(f"{key}がありません。非適用も明示してください。")
        if any(isinstance(value, str) and value.startswith("=") for value in row.values()):
            row_reasons.append("数式セルはキャンペーン入力として利用できません。")
        if _text(row.get("threshold_unit")) != "件":
            row_reasons.append("閾値の単位は件のみ対応しています。")
        threshold = _number(row.get("threshold_value"))
        if threshold is None:
            row_reasons.append("閾値は正の整数が必要です。")
        try:
            mode = LadderMode(_text(row.get("threshold_application")))
        except ValueError:
            mode = LadderMode.UNKNOWN
        if mode == LadderMode.UNKNOWN:
            row_reasons.append("閾値の適用方式が未確定です。")
        classes = tuple(part.strip() for part in _text(row.get("eligible_opening_classes")).split(";") if part.strip())
        if not classes or len(classes) != len(set(classes)):
            row_reasons.append("対象開通区分が欠落／重複しています。")
        evidence, approval_reasons = _approval_evidence(CAMPAIGN_SHEET, row, approvals)
        row_reasons.extend(approval_reasons)
        for prefix in ("", "agency_"):
            row_reasons.extend(_source_reasons(row, prefix, input_path.parent))
        if not _text(row.get("cost_source_document")) or not _text(row.get("cost_evidence_location")):
            row_reasons.append("直接原価の根拠がありません。0円も承認済み根拠が必要です。")
        amounts: dict[str, Decimal | None] = {}
        for label_key, value_key, state_key, tax_key in (
            ("upstream", "upstream_unit_amount", "upstream_amount_state", "upstream_tax_basis"),
            ("agency", "agency_unit_amount", "agency_amount_state", "agency_tax_basis"),
            ("cost", "attributable_direct_cost", "direct_cost_state", "direct_cost_tax_basis"),
        ):
            value, errors = _read_amount(row.get(value_key), row.get(state_key), label_key)
            amounts[label_key] = value
            row_reasons.extend(errors)
            if _text(row.get(tax_key)).upper() != "EXCLUSIVE":
                row_reasons.append(f"{label_key}は税別の承認値が必要です。")
        baseline = amounts["agency"]
        if rate_basis == "previous":
            baseline, errors = _read_amount(row.get("previous_agency_unit_amount"), row.get("previous_amount_state"), "previous")
            row_reasons.extend(errors)
            if _text(row.get("previous_tax_basis")).upper() != "EXCLUSIVE":
                row_reasons.append("前月単価は税別が必要です。")
            if _text(row.get("previous_period")) != previous_period_text:
                row_reasons.append("前月単価の対象期間が一致しません。")
            row_reasons.extend(_source_reasons(row, "previous_", input_path.parent))
        reasons.extend(f"{label}: {reason}" for reason in row_reasons)
        if row_reasons:
            continue
        tier = CampaignTierInput(
            threshold=threshold,
            p2_condition_item_id=_text(row["condition_item_id"]),
            upstream_unit_reward=amounts["upstream"],
            previous_agency_unit_reward=baseline,
            direct_cost_per_unit=amounts["cost"],
            upstream_approval=evidence, previous_approval=evidence,
            direct_cost_approval=evidence,
        )
        groups.setdefault(_text(row["campaign_code"]), []).append((row, tier))
    campaigns: list[CampaignSchedule] = []
    for code, entries in groups.items():
        first = entries[0][0]
        if any(any(_text(row.get(key)) != _text(first.get(key)) for key in _GROUP_FIELDS) for row, _ in entries):
            reasons.append(f"{code}: 同一キャンペーン内で計数範囲・支払対象・期間等が矛盾しています。")
        thresholds = [tier.threshold for _, tier in entries]
        expected_thresholds = [_number(value.strip()) for value in _text(first["threshold_values"]).split(";")]
        if None in expected_thresholds or expected_thresholds != sorted(set(thresholds)):
            reasons.append(f"{code}: 承認された閾値一覧と実際の階層行が一致しません。")
        if len(thresholds) != len(set(thresholds)):
            reasons.append(f"{code}: 閾値が重複しています。")
        campaigns.append(CampaignSchedule(
            campaign_code=code, counting_scope=_text(first["counting_scope"]),
            payout_scope=_text(first["payout_scope"]),
            eligible_opening_classes=tuple(part.strip() for part in _text(first["eligible_opening_classes"]).split(";") if part.strip()),
            mode=LadderMode(_text(first["threshold_application"])),
            tiers=tuple(sorted((tier for _, tier in entries), key=lambda tier: tier.threshold)),
        ))
    return CampaignLoadReport(
        agency_id, target_period, rate_basis,
        () if reasons else tuple(campaigns), tuple(dict.fromkeys(reasons)), len(selected),
    )


def evaluate_campaign_workbook(
    input_path: Path, profile: AgencyRewardProfile, *, mode: str = "verify_confirmed",
    volume_scenarios: tuple[VolumeScenario, ...] = (),
) -> RewardDesignResult:
    """Use the existing engine, without generating or approving payment data.

    verify_confirmed fixes current prices: a margin breach must hold, never
    silently reduce the confirmed amount. design uses previous-month prices.
    """
    if mode not in {"verify_confirmed", "design"}:
        raise ValueError("mode must be verify_confirmed or design")
    loaded = load_campaign_schedules(
        input_path, profile.agency_id, profile.target_period,
        rate_basis="confirmed" if mode == "verify_confirmed" else "previous",
    )
    if loaded.hold_reasons:
        return RewardDesignResult("hold", profile.agency_id, profile.target_period,
                                  (), (), (), (), loaded.hold_reasons)
    checked_profile = replace(profile, allow_rate_changes=False) if mode == "verify_confirmed" else profile
    return design_agency_rewards(checked_profile, (), loaded.campaigns, volume_scenarios=volume_scenarios)
