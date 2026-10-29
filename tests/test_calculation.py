from __future__ import annotations

import sys
import unittest
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from au_hikari_conditions.calculation import (
    ApprovalEvidence,
    ApprovalStatus,
    ApprovedMarginBaseline,
    CalculationCase,
    ConditionSignature,
    DisplayPolicy,
    LadderDefinition,
    LadderMode,
    LadderScenario,
    LadderTier,
    TaxBasis,
    calculate_agency_reward,
    simulate_ladder,
)


def approved_evidence(reference: str) -> ApprovalEvidence:
    return ApprovalEvidence(
        status=ApprovalStatus.APPROVED,
        approval_id=f"approval-{reference}",
        source_document=f"approved-source-{reference}.pdf",
        evidence_location="page:1",
        approved_by="approver-1",
        approved_at=datetime(2026, 10, 29, tzinfo=timezone.utc),
    )


def condition_signature(
    *,
    item_key: str = "au.hikari.plus1.biglobe.home.1g.new",
    effective_period: str = "2026-11-01/2026-11-30",
    period_rule: str = "monthly",
    payout_trigger: str = "activation",
    reversal_terms: tuple[str, ...] = ("early_cancellation",),
) -> ConditionSignature:
    return ConditionSignature(
        item_key=item_key,
        attributes=(("isp", "BIGLOBE"), ("speed", "1G")),
        effective_period=effective_period,
        period_rule=period_rule,
        payment_timing="following_month",
        payout_trigger=payout_trigger,
        exclusions=("duplicate_application",),
        reversal_terms=reversal_terms,
    )


def calculation_case(
    *,
    condition: ConditionSignature | None = None,
    upstream_tax_basis: TaxBasis = TaxBasis.EXCLUSIVE,
    direct_cost_tax_basis: TaxBasis = TaxBasis.EXCLUSIVE,
    upstream_reward: Decimal = Decimal("100"),
    direct_cost: Decimal = Decimal("10"),
    baseline: ApprovedMarginBaseline | None = None,
    upstream_approval: ApprovalEvidence | None = None,
    direct_cost_approval: ApprovalEvidence | None = None,
    condition_equivalence_approval: ApprovalEvidence | None = None,
    minimum_gross_margin: Decimal | None = None,
    exception_reasons: tuple[str, ...] = (),
    contract_prohibits_payment: bool = False,
    mandatory_disclosure: bool = False,
) -> CalculationCase:
    signature = condition_signature()
    approved_baseline = baseline or ApprovedMarginBaseline(
        condition_signature=signature,
        tax_basis=TaxBasis.EXCLUSIVE,
        target_gross_margin=Decimal("30"),
        display_policy=DisplayPolicy.PRESENT,
        approval=approved_evidence("margin"),
    )
    return CalculationCase(
        condition_signature=condition or signature,
        upstream_tax_basis=upstream_tax_basis,
        direct_cost_tax_basis=direct_cost_tax_basis,
        upstream_reward=upstream_reward,
        direct_cost=direct_cost,
        baseline=approved_baseline,
        upstream_approval=upstream_approval or approved_evidence("upstream"),
        direct_cost_approval=direct_cost_approval or approved_evidence("cost"),
        condition_equivalence_approval=(
            condition_equivalence_approval or approved_evidence("mapping")
        ),
        minimum_gross_margin=minimum_gross_margin,
        exception_reasons=exception_reasons,
        contract_prohibits_payment=contract_prohibits_payment,
        mandatory_disclosure=mandatory_disclosure,
    )


class AgencyRewardCalculationTest(unittest.TestCase):
    def test_preserves_approved_margin_after_attributed_cost(self) -> None:
        result = calculate_agency_reward(calculation_case())

        self.assertEqual(result.status, "proposed")
        self.assertEqual(result.proposed_agency_reward, Decimal("60"))
        self.assertEqual(result.retained_gross_margin, Decimal("30"))
        self.assertEqual(result.display_policy, DisplayPolicy.PRESENT)
        self.assertEqual(result.hold_reasons, ())

    def test_holds_when_condition_signature_differs(self) -> None:
        result = calculate_agency_reward(
            calculation_case(condition=condition_signature(payout_trigger="application"))
        )

        self.assertEqual(result.status, "hold")
        self.assertIsNone(result.proposed_agency_reward)
        self.assertTrue(result.hold_reasons)

    def test_unapproved_or_incomplete_amount_and_mapping_records_hold(self) -> None:
        cases = (
            calculation_case(upstream_approval=ApprovalEvidence()),
            calculation_case(direct_cost_approval=ApprovalEvidence()),
            calculation_case(
                baseline=ApprovedMarginBaseline(
                    condition_signature=condition_signature(),
                    tax_basis=TaxBasis.EXCLUSIVE,
                    target_gross_margin=Decimal("30"),
                    display_policy=DisplayPolicy.PRESENT,
                    approval=ApprovalEvidence(),
                )
            ),
            calculation_case(
                condition_equivalence_approval=ApprovalEvidence(
                    status=ApprovalStatus.APPROVED,
                    approval_id="approval-missing-evidence",
                    approved_by="approver-1",
                    approved_at=datetime(2026, 10, 29, tzinfo=timezone.utc),
                )
            ),
        )

        for case in cases:
            with self.subTest(case=case):
                result = calculate_agency_reward(case)
                self.assertEqual(result.status, "hold")
                self.assertIsNone(result.proposed_agency_reward)

    def test_monthly_period_rollover_requires_verified_equivalence(self) -> None:
        previous_signature = condition_signature(
            effective_period="2026-10-01/2026-10-31",
        )
        baseline = ApprovedMarginBaseline(
            condition_signature=previous_signature,
            tax_basis=TaxBasis.EXCLUSIVE,
            target_gross_margin=Decimal("30"),
            display_policy=DisplayPolicy.PRESENT,
            approval=approved_evidence("margin"),
        )
        unverified = calculate_agency_reward(
            calculation_case(
                baseline=baseline,
                condition_equivalence_approval=ApprovalEvidence(),
            )
        )
        verified = calculate_agency_reward(calculation_case(baseline=baseline))

        self.assertEqual(unverified.status, "hold")
        self.assertTrue(any("同等性レビュー" in reason for reason in unverified.hold_reasons))
        self.assertEqual(verified.status, "proposed")
        self.assertEqual(verified.proposed_agency_reward, Decimal("60"))

    def test_holds_when_tax_basis_is_unknown_or_mismatched(self) -> None:
        unknown = calculate_agency_reward(
            calculation_case(upstream_tax_basis=TaxBasis.UNKNOWN)
        )
        mismatch = calculate_agency_reward(
            calculation_case(upstream_tax_basis=TaxBasis.INCLUSIVE)
        )
        direct_cost_mismatch = calculate_agency_reward(
            calculation_case(direct_cost_tax_basis=TaxBasis.INCLUSIVE)
        )
        matching_inclusive = calculate_agency_reward(
            calculation_case(
                upstream_tax_basis=TaxBasis.INCLUSIVE,
                direct_cost_tax_basis=TaxBasis.INCLUSIVE,
                baseline=ApprovedMarginBaseline(
                    condition_signature=condition_signature(),
                    tax_basis=TaxBasis.INCLUSIVE,
                    target_gross_margin=Decimal("30"),
                    display_policy=DisplayPolicy.PRESENT,
                    approval=approved_evidence("margin-inclusive"),
                ),
            )
        )

        self.assertEqual(unknown.status, "hold")
        self.assertEqual(mismatch.status, "hold")
        self.assertEqual(direct_cost_mismatch.status, "hold")
        self.assertEqual(matching_inclusive.status, "hold")
        self.assertTrue(any("税区分" in reason for reason in unknown.hold_reasons))
        self.assertTrue(any("税抜" in reason for reason in mismatch.hold_reasons))
        self.assertTrue(any("税抜" in reason for reason in direct_cost_mismatch.hold_reasons))
        self.assertTrue(any("税抜" in reason for reason in matching_inclusive.hold_reasons))

    def test_holds_negative_reward_instead_of_clamping_to_zero(self) -> None:
        result = calculate_agency_reward(
            calculation_case(upstream_reward=Decimal("20"))
        )

        self.assertEqual(result.status, "hold")
        self.assertIsNone(result.proposed_agency_reward)
        self.assertTrue(any("負額" in reason for reason in result.hold_reasons))

    def test_holds_exceptions_and_minimum_margin_violations(self) -> None:
        exception = calculate_agency_reward(
            calculation_case(exception_reasons=("固定報酬との重複確認が必要です。",))
        )
        minimum = calculate_agency_reward(
            calculation_case(minimum_gross_margin=Decimal("35"))
        )
        prohibited = calculate_agency_reward(
            calculation_case(contract_prohibits_payment=True)
        )

        self.assertEqual(exception.status, "hold")
        self.assertEqual(minimum.status, "hold")
        self.assertEqual(prohibited.status, "hold")

    def test_withheld_items_stay_internal_and_mandatory_disclosure_holds(self) -> None:
        baseline = ApprovedMarginBaseline(
            condition_signature=condition_signature(),
            tax_basis=TaxBasis.EXCLUSIVE,
            target_gross_margin=Decimal("30"),
            display_policy=DisplayPolicy.WITHHOLD,
            approval=approved_evidence("margin-withheld"),
        )
        withheld = calculate_agency_reward(calculation_case(baseline=baseline))
        mandatory = calculate_agency_reward(
            calculation_case(baseline=baseline, mandatory_disclosure=True)
        )

        self.assertEqual(withheld.status, "proposed")
        self.assertEqual(withheld.display_policy, DisplayPolicy.WITHHOLD)
        self.assertEqual(mandatory.status, "hold")
        self.assertTrue(any("告知" in reason for reason in mandatory.hold_reasons))

    def test_undecided_display_policy_holds(self) -> None:
        baseline = ApprovedMarginBaseline(
            condition_signature=condition_signature(),
            tax_basis=TaxBasis.EXCLUSIVE,
            target_gross_margin=Decimal("30"),
            display_policy=DisplayPolicy.UNDECIDED,
            approval=approved_evidence("margin-undecided"),
        )

        result = calculate_agency_reward(calculation_case(baseline=baseline))

        self.assertEqual(result.status, "hold")
        self.assertTrue(any("掲載可否" in reason for reason in result.hold_reasons))


class LadderSimulationTest(unittest.TestCase):
    def setUp(self) -> None:
        self.definition = LadderDefinition(
            counting_scope="upstream_all_agencies",
            mode=LadderMode.REACHED_TIER,
            tiers=(
                LadderTier(500, Decimal("6500")),
                LadderTier(750, Decimal("9000")),
                LadderTier(1000, Decimal("11500")),
                LadderTier(1500, Decimal("14000")),
                LadderTier(2000, Decimal("16500")),
                LadderTier(2500, Decimal("17500")),
            ),
        )

    def test_threshold_boundaries_apply_one_reached_tier_only(self) -> None:
        expected = (
            (499, (), "0"),
            (500, (500,), "6500"),
            (749, (500,), "6500"),
            (750, (500, 750), "9000"),
            (999, (500, 750), "9000"),
            (1000, (500, 750, 1000), "11500"),
            (1499, (500, 750, 1000), "11500"),
            (1500, (500, 750, 1000, 1500), "14000"),
            (1999, (500, 750, 1000, 1500), "14000"),
            (2000, (500, 750, 1000, 1500, 2000), "16500"),
            (2499, (500, 750, 1000, 1500, 2000), "16500"),
            (2500, (500, 750, 1000, 1500, 2000, 2500), "17500"),
        )
        for count, thresholds, payout in expected:
            with self.subTest(count=count):
                result = simulate_ladder(
                    self.definition,
                    LadderScenario("upstream_all_agencies", count),
                )
                self.assertEqual(result.status, "simulated")
                self.assertEqual(result.reached_thresholds, thresholds)
                self.assertEqual(result.payout_amount, Decimal(payout))

    def test_progressive_ladder_sums_reached_tiers(self) -> None:
        definition = LadderDefinition(
            counting_scope="upstream_all_agencies",
            mode=LadderMode.PROGRESSIVE,
            tiers=self.definition.tiers,
        )

        result = simulate_ladder(
            definition,
            LadderScenario("upstream_all_agencies", 750),
        )

        self.assertEqual(result.payout_amount, Decimal("15500"))
        self.assertTrue(result.is_simulation)

    def test_plus_two_three_ladder_uses_its_own_thresholds(self) -> None:
        definition = LadderDefinition(
            counting_scope="upstream_all_agencies",
            mode=LadderMode.REACHED_TIER,
            tiers=(
                LadderTier(220, Decimal("3000")),
                LadderTier(440, Decimal("5000")),
            ),
        )

        below = simulate_ladder(definition, LadderScenario("upstream_all_agencies", 219))
        first = simulate_ladder(definition, LadderScenario("upstream_all_agencies", 220))
        second = simulate_ladder(definition, LadderScenario("upstream_all_agencies", 440))

        self.assertEqual(below.payout_amount, Decimal("0"))
        self.assertEqual(first.payout_amount, Decimal("3000"))
        self.assertEqual(second.payout_amount, Decimal("5000"))

    def test_unknown_count_scope_and_ambiguous_mode_hold(self) -> None:
        unknown_count = simulate_ladder(
            self.definition,
            LadderScenario(None, 500),
        )
        wrong_scope = simulate_ladder(
            self.definition,
            LadderScenario("agency_only", 500),
        )
        unknown_mode = simulate_ladder(
            LadderDefinition(
                self.definition.counting_scope,
                LadderMode.UNKNOWN,
                self.definition.tiers,
            ),
            LadderScenario("upstream_all_agencies", 500),
        )

        self.assertEqual(unknown_count.status, "hold")
        self.assertEqual(wrong_scope.status, "hold")
        self.assertEqual(unknown_mode.status, "hold")

    def test_invalid_or_duplicate_thresholds_hold(self) -> None:
        definition = LadderDefinition(
            counting_scope="upstream_all_agencies",
            mode=LadderMode.REACHED_TIER,
            tiers=(LadderTier(500, Decimal("6500")), LadderTier(500, Decimal("9000"))),
        )

        result = simulate_ladder(
            definition,
            LadderScenario("upstream_all_agencies", 750),
        )

        self.assertEqual(result.status, "hold")
        self.assertIsNone(result.payout_amount)

if __name__ == "__main__":
    unittest.main()
