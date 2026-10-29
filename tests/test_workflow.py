from __future__ import annotations

import sys
import unittest
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from au_hikari_conditions.workflow import (
    AgencyWorkflow,
    BatchStatus,
    DeliverySnapshot,
    delivery_idempotency_key,
)


def delivery_snapshot(
    *,
    batch_id: str = "batch-2026-11",
    agency_id: str = "agency-1",
    target_period: str = "2026-11",
    source_master_hash: str = "master-sha256",
    template_hash: str = "template-sha256",
    artifact_hashes: tuple[str, ...] = ("xlsx-sha256", "docx-sha256"),
    recipient_snapshot_hash: str = "recipient-sha256",
    message_hash: str = "message-sha256",
) -> DeliverySnapshot:
    return DeliverySnapshot(
        batch_id=batch_id,
        agency_id=agency_id,
        target_period=target_period,
        source_master_hash=source_master_hash,
        template_hash=template_hash,
        artifact_hashes=artifact_hashes,
        recipient_snapshot_hash=recipient_snapshot_hash,
        message_hash=message_hash,
    )


def workflow_at_pending_approval() -> AgencyWorkflow:
    workflow = AgencyWorkflow("batch-2026-11", "agency-1")
    for status in (
        BatchStatus.PARSED,
        BatchStatus.VALIDATED,
        BatchStatus.PROPOSED,
        BatchStatus.GENERATED,
        BatchStatus.QA_PASSED,
        BatchStatus.PENDING_APPROVAL,
    ):
        workflow.transition(status, actor_id="staff-1")
    return workflow


class AgencyWorkflowTest(unittest.TestCase):
    def test_requires_full_review_path_and_snapshot_bound_approval(self) -> None:
        workflow = workflow_at_pending_approval()
        snapshot = delivery_snapshot()

        approved = workflow.transition(
            BatchStatus.APPROVED,
            actor_id="approver-1",
            snapshot=snapshot,
        )
        ready = workflow.transition(
            BatchStatus.READY_TO_SEND,
            actor_id="staff-1",
            snapshot=snapshot,
        )
        sent = workflow.transition(
            BatchStatus.SENT,
            actor_id="system",
            snapshot=snapshot,
        )

        self.assertEqual(workflow.status, BatchStatus.SENT)
        self.assertEqual(workflow.approver_id, "approver-1")
        self.assertEqual(approved.snapshot_fingerprint, snapshot.fingerprint())
        self.assertEqual(ready.snapshot_fingerprint, approved.snapshot_fingerprint)
        self.assertEqual(sent.previous_status, BatchStatus.READY_TO_SEND)
        self.assertEqual(len(workflow.events), 9)

    def test_blocks_approval_without_required_review_state_or_snapshot(self) -> None:
        workflow = AgencyWorkflow("batch-2026-11", "agency-1")

        with self.assertRaises(ValueError):
            workflow.transition(
                BatchStatus.APPROVED,
                actor_id="approver-1",
                snapshot=delivery_snapshot(),
            )

        workflow = workflow_at_pending_approval()
        with self.assertRaises(ValueError):
            workflow.transition(BatchStatus.APPROVED, actor_id="approver-1")

    def test_snapshot_changes_block_send_and_require_reapproval(self) -> None:
        workflow = workflow_at_pending_approval()
        approved_snapshot = delivery_snapshot()
        workflow.transition(
            BatchStatus.APPROVED,
            actor_id="approver-1",
            snapshot=approved_snapshot,
        )
        changed_snapshot = replace(
            approved_snapshot,
            recipient_snapshot_hash="different-recipient-sha256",
        )

        with self.assertRaises(ValueError):
            workflow.transition(
                BatchStatus.READY_TO_SEND,
                actor_id="staff-1",
                snapshot=changed_snapshot,
            )

        event = workflow.invalidate_approval(
            changed_snapshot,
            actor_id="staff-1",
            reason="宛先の変更を再確認するため",
        )

        self.assertEqual(workflow.status, BatchStatus.GENERATED)
        self.assertIsNone(workflow.approval_fingerprint)
        self.assertIsNone(workflow.approver_id)
        self.assertEqual(event.next_status, BatchStatus.GENERATED)
        self.assertEqual(event.reason, "宛先の変更を再確認するため")

    def test_snapshot_tracks_content_template_recipient_and_period(self) -> None:
        original = delivery_snapshot()
        variations = (
            replace(original, artifact_hashes=("changed-artifact-sha256",)),
            replace(original, message_hash="changed-message-sha256"),
            replace(original, template_hash="changed-template-sha256"),
            replace(original, recipient_snapshot_hash="changed-recipient-sha256"),
            replace(original, source_master_hash="changed-master-sha256"),
            replace(original, target_period="2026-12"),
        )

        self.assertTrue(all(original.fingerprint() != item.fingerprint() for item in variations))

    def test_delivery_key_is_order_independent_and_scoped(self) -> None:
        original = delivery_snapshot()
        reordered = replace(
            original,
            artifact_hashes=tuple(reversed(original.artifact_hashes)),
        )
        another_agency = replace(original, agency_id="agency-2")

        self.assertEqual(
            delivery_idempotency_key(original),
            delivery_idempotency_key(reordered),
        )
        self.assertNotEqual(
            delivery_idempotency_key(original),
            delivery_idempotency_key(another_agency),
        )

    def test_snapshot_must_match_workflow_owner(self) -> None:
        workflow = workflow_at_pending_approval()

        with self.assertRaises(ValueError):
            workflow.transition(
                BatchStatus.APPROVED,
                actor_id="approver-1",
                snapshot=delivery_snapshot(agency_id="agency-2"),
            )

    def test_events_have_actor_and_utc_timestamp(self) -> None:
        workflow = AgencyWorkflow("batch-2026-11", "agency-1")
        recorded_at = datetime(2026, 10, 29, tzinfo=timezone.utc)

        event = workflow.transition(
            BatchStatus.PARSED,
            actor_id="staff-1",
            occurred_at=recorded_at,
        )

        self.assertEqual(event.actor_id, "staff-1")
        self.assertEqual(event.occurred_at, recorded_at)
        self.assertEqual(event.next_status, BatchStatus.PARSED)

    def test_terminal_sent_state_cannot_be_reopened(self) -> None:
        workflow = workflow_at_pending_approval()
        snapshot = delivery_snapshot()
        workflow.transition(BatchStatus.APPROVED, actor_id="approver-1", snapshot=snapshot)
        workflow.transition(BatchStatus.READY_TO_SEND, actor_id="staff-1", snapshot=snapshot)
        workflow.transition(BatchStatus.SENT, actor_id="system", snapshot=snapshot)

        with self.assertRaises(ValueError):
            workflow.transition(BatchStatus.REVIEW_REQUIRED, actor_id="staff-1")


if __name__ == "__main__":
    unittest.main()
