from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from enum import Enum


class TaxBasis(str, Enum):
    EXCLUSIVE = "exclusive"
    INCLUSIVE = "inclusive"
    EXEMPT = "exempt"
    UNKNOWN = "unknown"


class DisplayPolicy(str, Enum):
    PRESENT = "present"
    WITHHOLD = "withhold"
    UNDECIDED = "undecided"


class LadderMode(str, Enum):
    PROGRESSIVE = "progressive"
    REACHED_TIER = "reached_tier"
    UNKNOWN = "unknown"


class ApprovalStatus(str, Enum):
    UNAPPROVED = "UNAPPROVED"
    PENDING = "PENDING"
    APPROVED = "APPROVED"
    REJECTED = "REJECTED"


@dataclass(frozen=True)
class ApprovalEvidence:
    status: ApprovalStatus = ApprovalStatus.UNAPPROVED
    approval_id: str | None = None
    source_document: str | None = None
    evidence_location: str | None = None
    approved_by: str | None = None
    approved_at: datetime | None = None

    def is_complete_approval(self) -> bool:
        return (
            self.status == ApprovalStatus.APPROVED
            and bool(self.approval_id and self.approval_id.strip())
            and bool(self.source_document and self.source_document.strip())
            and bool(self.evidence_location and self.evidence_location.strip())
            and bool(self.approved_by and self.approved_by.strip())
            and isinstance(self.approved_at, datetime)
            and self.approved_at.tzinfo is not None
            and self.approved_at.utcoffset() is not None
        )


@dataclass(frozen=True)
class ConditionSignature:
    item_key: str
    attributes: tuple[tuple[str, str], ...]
    effective_period: str
    period_rule: str
    payment_timing: str
    payout_trigger: str
    exclusions: tuple[str, ...]
    reversal_terms: tuple[str, ...]

    def matches_terms(self, other: ConditionSignature) -> bool:
        return (
            self.item_key == other.item_key
            and self.attributes == other.attributes
            and self.period_rule == other.period_rule
            and self.payment_timing == other.payment_timing
            and self.payout_trigger == other.payout_trigger
            and self.exclusions == other.exclusions
            and self.reversal_terms == other.reversal_terms
        )


@dataclass(frozen=True)
class ApprovedMarginBaseline:
    condition_signature: ConditionSignature
    tax_basis: TaxBasis
    target_gross_margin: Decimal
    display_policy: DisplayPolicy
    approval: ApprovalEvidence = ApprovalEvidence()


@dataclass(frozen=True)
class CalculationCase:
    condition_signature: ConditionSignature
    upstream_tax_basis: TaxBasis
    direct_cost_tax_basis: TaxBasis
    upstream_reward: Decimal
    direct_cost: Decimal
    baseline: ApprovedMarginBaseline
    upstream_approval: ApprovalEvidence = ApprovalEvidence()
    direct_cost_approval: ApprovalEvidence = ApprovalEvidence()
    condition_equivalence_approval: ApprovalEvidence = ApprovalEvidence()
    minimum_gross_margin: Decimal | None = None
    exception_reasons: tuple[str, ...] = ()
    contract_prohibits_payment: bool = False
    mandatory_disclosure: bool = False


@dataclass(frozen=True)
class CalculationResult:
    status: str
    proposed_agency_reward: Decimal | None
    retained_gross_margin: Decimal | None
    display_policy: DisplayPolicy
    hold_reasons: tuple[str, ...]


def calculate_agency_reward(case: CalculationCase) -> CalculationResult:
    hold_reasons: list[str] = []
    baseline = case.baseline

    if not _signature_is_complete(case.condition_signature) or not _signature_is_complete(
        baseline.condition_signature
    ):
        hold_reasons.append("成果項目・適用期間・期間ルール・支払条件の情報が不足しています。")
    if not case.condition_signature.matches_terms(baseline.condition_signature):
        hold_reasons.append("成果項目・対象条件・期間・支払・除外・戻入の同等性を確認できません。")
    for label, approval in (
        ("11月上位報酬", case.upstream_approval),
        ("直接原価", case.direct_cost_approval),
        ("前回承認済み粗利", baseline.approval),
        ("条件対応・同等性レビュー", case.condition_equivalence_approval),
    ):
        if not approval.is_complete_approval():
            hold_reasons.append(
                f"{label}は根拠・承認者・承認日時が揃った承認済みレコードではありません。"
            )
    tax_bases = (
        case.upstream_tax_basis,
        case.direct_cost_tax_basis,
        baseline.tax_basis,
    )
    if any(tax_basis == TaxBasis.UNKNOWN for tax_basis in tax_bases):
        hold_reasons.append("金額の税区分が不明です。")
    elif any(tax_basis != TaxBasis.EXCLUSIVE for tax_basis in tax_bases):
        hold_reasons.append("税抜金額への換算が確認できないため、計算を保留します。")
    if baseline.display_policy == DisplayPolicy.UNDECIDED:
        hold_reasons.append("代理店向け掲載可否が未承認です。")
    if baseline.display_policy == DisplayPolicy.WITHHOLD and case.mandatory_disclosure:
        hold_reasons.append("契約上必要な告知事項を非提示にできません。")
    if case.contract_prohibits_payment:
        hold_reasons.append("契約上の支払禁止条件があります。")
    hold_reasons.extend(case.exception_reasons)

    monetary_values = (
        ("上位報酬", case.upstream_reward),
        ("直接原価", case.direct_cost),
        ("承認済み目標粗利", baseline.target_gross_margin),
    )
    for label, amount in monetary_values:
        if not isinstance(amount, Decimal) or not amount.is_finite() or amount < 0:
            hold_reasons.append(f"{label}が有限の0以上の金額ではありません。")

    if case.minimum_gross_margin is not None:
        if (
            not isinstance(case.minimum_gross_margin, Decimal)
            or not case.minimum_gross_margin.is_finite()
            or case.minimum_gross_margin < 0
        ):
            hold_reasons.append("最低粗利が有限の0以上の金額ではありません。")
        elif (
            isinstance(baseline.target_gross_margin, Decimal)
            and baseline.target_gross_margin.is_finite()
            and baseline.target_gross_margin < case.minimum_gross_margin
        ):
            hold_reasons.append("承認済み目標粗利が最低粗利を下回ります。")

    if hold_reasons:
        return CalculationResult(
            status="hold",
            proposed_agency_reward=None,
            retained_gross_margin=None,
            display_policy=baseline.display_policy,
            hold_reasons=tuple(hold_reasons),
        )

    proposed_reward = (
        case.upstream_reward - baseline.target_gross_margin - case.direct_cost
    )
    if proposed_reward < 0:
        return CalculationResult(
            status="hold",
            proposed_agency_reward=None,
            retained_gross_margin=None,
            display_policy=baseline.display_policy,
            hold_reasons=("上位報酬から目標粗利と直接原価を控除すると負額になります。",),
        )

    return CalculationResult(
        status="proposed",
        proposed_agency_reward=proposed_reward,
        retained_gross_margin=(
            case.upstream_reward - proposed_reward - case.direct_cost
        ),
        display_policy=baseline.display_policy,
        hold_reasons=(),
    )


def _signature_is_complete(signature: ConditionSignature) -> bool:
    return all(
        (
            signature.item_key.strip(),
            signature.effective_period.strip(),
            signature.period_rule.strip(),
            signature.payment_timing.strip(),
            signature.payout_trigger.strip(),
        )
    )


@dataclass(frozen=True)
class LadderTier:
    threshold: int
    payout_amount: Decimal


@dataclass(frozen=True)
class LadderDefinition:
    counting_scope: str
    mode: LadderMode
    tiers: tuple[LadderTier, ...]


@dataclass(frozen=True)
class LadderScenario:
    counting_scope: str | None
    completed_count: int | None


@dataclass(frozen=True)
class LadderResult:
    status: str
    reached_thresholds: tuple[int, ...]
    payout_amount: Decimal | None
    is_simulation: bool
    hold_reasons: tuple[str, ...]


def simulate_ladder(
    definition: LadderDefinition,
    scenario: LadderScenario,
) -> LadderResult:
    hold_reasons: list[str] = []
    if not definition.counting_scope:
        hold_reasons.append("ラダーの計数主体が定義されていません。")
    if scenario.counting_scope is None:
        hold_reasons.append("シミュレーションの計数主体が不明です。")
    elif scenario.counting_scope != definition.counting_scope:
        hold_reasons.append("シミュレーションの計数主体が上位条件と一致しません。")
    if scenario.completed_count is None:
        hold_reasons.append("達成件数が不明です。")
    elif (
        isinstance(scenario.completed_count, bool)
        or not isinstance(scenario.completed_count, int)
    ):
        hold_reasons.append("達成件数が整数ではありません。")
    elif scenario.completed_count < 0:
        hold_reasons.append("達成件数が負数です。")
    if definition.mode == LadderMode.UNKNOWN:
        hold_reasons.append("ラダーの累進加算／到達段階適用が未確定です。")
    if not definition.tiers:
        hold_reasons.append("ラダー閾値がありません。")

    thresholds = [tier.threshold for tier in definition.tiers]
    valid_thresholds = all(
        isinstance(threshold, int)
        and not isinstance(threshold, bool)
        and threshold > 0
        for threshold in thresholds
    )
    if not valid_thresholds:
        hold_reasons.append("ラダー閾値は正数である必要があります。")
    elif len(thresholds) != len(set(thresholds)):
        hold_reasons.append("ラダー閾値が重複しています。")
    elif thresholds != sorted(thresholds):
        hold_reasons.append("ラダー閾値は昇順で指定してください。")
    if any(
        not isinstance(tier.payout_amount, Decimal)
        or not tier.payout_amount.is_finite()
        or tier.payout_amount < 0
        for tier in definition.tiers
    ):
        hold_reasons.append("ラダー報酬は有限の0以上の金額である必要があります。")

    if hold_reasons:
        return LadderResult(
            status="hold",
            reached_thresholds=(),
            payout_amount=None,
            is_simulation=True,
            hold_reasons=tuple(hold_reasons),
        )

    if scenario.completed_count is None:
        return LadderResult(
            status="hold",
            reached_thresholds=(),
            payout_amount=None,
            is_simulation=True,
            hold_reasons=("達成件数が不明です。",),
        )
    reached_tiers = tuple(
        tier for tier in definition.tiers
        if scenario.completed_count >= tier.threshold
    )
    if definition.mode == LadderMode.PROGRESSIVE:
        payout = sum(
            (tier.payout_amount for tier in reached_tiers),
            start=Decimal("0"),
        )
    elif reached_tiers:
        payout = reached_tiers[-1].payout_amount
    else:
        payout = Decimal("0")

    return LadderResult(
        status="simulated",
        reached_thresholds=tuple(tier.threshold for tier in reached_tiers),
        payout_amount=payout,
        is_simulation=True,
        hold_reasons=(),
    )
