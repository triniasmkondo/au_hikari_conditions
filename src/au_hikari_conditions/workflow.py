from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum


class BatchStatus(str, Enum):
    RECEIVED = "RECEIVED"
    PARSED = "PARSED"
    VALIDATED = "VALIDATED"
    PROPOSED = "PROPOSED"
    GENERATED = "GENERATED"
    QA_PASSED = "QA_PASSED"
    PENDING_APPROVAL = "PENDING_APPROVAL"
    APPROVED = "APPROVED"
    READY_TO_SEND = "READY_TO_SEND"
    SENT = "SENT"
    FAILED = "FAILED"
    REVIEW_REQUIRED = "REVIEW_REQUIRED"
    HOLD = "HOLD"
    RETURNED = "RETURNED"


@dataclass(frozen=True)
class DeliverySnapshot:
    batch_id: str
    agency_id: str
    target_period: str
    source_master_hash: str
    template_hash: str
    artifact_hashes: tuple[str, ...]
    recipient_snapshot_hash: str
    message_hash: str

    def fingerprint(self) -> str:
        _validate_snapshot(self)
        payload = {
            "batch_id": self.batch_id,
            "agency_id": self.agency_id,
            "target_period": self.target_period,
            "source_master_hash": self.source_master_hash,
            "template_hash": self.template_hash,
            "artifact_hashes": sorted(self.artifact_hashes),
            "recipient_snapshot_hash": self.recipient_snapshot_hash,
            "message_hash": self.message_hash,
        }
        serialized = json.dumps(payload, sort_keys=True, separators=(",", ":"))
        return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class WorkflowEvent:
    batch_id: str
    agency_id: str
    previous_status: BatchStatus
    next_status: BatchStatus
    actor_id: str
    occurred_at: datetime
    reason: str | None
    snapshot_fingerprint: str | None


_ALLOWED_TRANSITIONS: dict[BatchStatus, frozenset[BatchStatus]] = {
    BatchStatus.RECEIVED: frozenset(
        {BatchStatus.PARSED, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.PARSED: frozenset(
        {BatchStatus.VALIDATED, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.VALIDATED: frozenset(
        {BatchStatus.PROPOSED, BatchStatus.HOLD, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.PROPOSED: frozenset(
        {BatchStatus.GENERATED, BatchStatus.HOLD, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.GENERATED: frozenset(
        {BatchStatus.QA_PASSED, BatchStatus.HOLD, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.QA_PASSED: frozenset(
        {BatchStatus.PENDING_APPROVAL, BatchStatus.HOLD, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.PENDING_APPROVAL: frozenset(
        {
            BatchStatus.APPROVED,
            BatchStatus.HOLD,
            BatchStatus.RETURNED,
            BatchStatus.REVIEW_REQUIRED,
        }
    ),
    BatchStatus.APPROVED: frozenset(
        {BatchStatus.READY_TO_SEND, BatchStatus.HOLD, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.READY_TO_SEND: frozenset(
        {BatchStatus.SENT, BatchStatus.FAILED, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.RETURNED: frozenset(
        {BatchStatus.GENERATED, BatchStatus.REVIEW_REQUIRED}
    ),
    BatchStatus.HOLD: frozenset(
        {
            BatchStatus.VALIDATED,
            BatchStatus.PROPOSED,
            BatchStatus.GENERATED,
            BatchStatus.REVIEW_REQUIRED,
        }
    ),
    BatchStatus.REVIEW_REQUIRED: frozenset(
        {
            BatchStatus.PARSED,
            BatchStatus.VALIDATED,
            BatchStatus.PROPOSED,
            BatchStatus.GENERATED,
            BatchStatus.HOLD,
        }
    ),
    BatchStatus.FAILED: frozenset({BatchStatus.REVIEW_REQUIRED}),
    BatchStatus.SENT: frozenset(),
}


@dataclass
class AgencyWorkflow:
    batch_id: str
    agency_id: str
    status: BatchStatus = BatchStatus.RECEIVED
    approval_fingerprint: str | None = None
    approver_id: str | None = None
    events: list[WorkflowEvent] = field(default_factory=list)

    def transition(
        self,
        next_status: BatchStatus,
        *,
        actor_id: str,
        snapshot: DeliverySnapshot | None = None,
        reason: str | None = None,
        occurred_at: datetime | None = None,
    ) -> WorkflowEvent:
        if not actor_id.strip():
            raise ValueError("actor_id is required for every workflow transition")
        if next_status not in _ALLOWED_TRANSITIONS[self.status]:
            raise ValueError(
                f"Invalid transition: {self.status.value} -> {next_status.value}"
            )
        if self.batch_id == "" or self.agency_id == "":
            raise ValueError("batch_id and agency_id are required")
        if snapshot is not None and (
            snapshot.batch_id != self.batch_id or snapshot.agency_id != self.agency_id
        ):
            raise ValueError("Snapshot does not belong to this batch and agency")

        fingerprint: str | None = None
        if next_status in {
            BatchStatus.APPROVED,
            BatchStatus.READY_TO_SEND,
            BatchStatus.SENT,
            BatchStatus.FAILED,
        }:
            if snapshot is None:
                raise ValueError(f"{next_status.value} requires a delivery snapshot")
            fingerprint = snapshot.fingerprint()
        if next_status == BatchStatus.APPROVED:
            self.approval_fingerprint = fingerprint
            self.approver_id = actor_id
        elif next_status in {BatchStatus.READY_TO_SEND, BatchStatus.SENT}:
            if self.approval_fingerprint is None or fingerprint != self.approval_fingerprint:
                raise ValueError("Current snapshot does not match the approved snapshot")
        elif next_status in {BatchStatus.HOLD, BatchStatus.RETURNED, BatchStatus.REVIEW_REQUIRED}:
            self.approval_fingerprint = None
            self.approver_id = None
        elif next_status == BatchStatus.GENERATED and self.status == BatchStatus.RETURNED:
            self.approval_fingerprint = None
            self.approver_id = None

        event = WorkflowEvent(
            batch_id=self.batch_id,
            agency_id=self.agency_id,
            previous_status=self.status,
            next_status=next_status,
            actor_id=actor_id,
            occurred_at=occurred_at or datetime.now(timezone.utc),
            reason=reason,
            snapshot_fingerprint=fingerprint,
        )
        self.status = next_status
        self.events.append(event)
        return event

    def invalidate_approval(
        self,
        changed_snapshot: DeliverySnapshot,
        *,
        actor_id: str,
        reason: str,
        occurred_at: datetime | None = None,
    ) -> WorkflowEvent:
        if self.status not in {BatchStatus.APPROVED, BatchStatus.READY_TO_SEND}:
            raise ValueError("Only an approved or send-ready artifact can be invalidated")
        if self.approval_fingerprint is None:
            raise ValueError("There is no active approval to invalidate")
        if changed_snapshot.batch_id != self.batch_id or changed_snapshot.agency_id != self.agency_id:
            raise ValueError("Snapshot does not belong to this batch and agency")
        if changed_snapshot.fingerprint() == self.approval_fingerprint:
            raise ValueError("Changed snapshot is identical to the approved snapshot")
        if not reason.strip():
            raise ValueError("An approval invalidation reason is required")
        if not actor_id.strip():
            raise ValueError("actor_id is required for approval invalidation")

        previous_status = self.status
        self.status = BatchStatus.GENERATED
        self.approval_fingerprint = None
        self.approver_id = None
        event = WorkflowEvent(
            batch_id=self.batch_id,
            agency_id=self.agency_id,
            previous_status=previous_status,
            next_status=BatchStatus.GENERATED,
            actor_id=actor_id,
            occurred_at=occurred_at or datetime.now(timezone.utc),
            reason=reason,
            snapshot_fingerprint=changed_snapshot.fingerprint(),
        )
        self.events.append(event)
        return event


def delivery_idempotency_key(snapshot: DeliverySnapshot) -> str:
    _validate_snapshot(snapshot)
    if not snapshot.artifact_hashes:
        raise ValueError("At least one artifact is required for delivery")
    key_material = {
        "batch_id": snapshot.batch_id,
        "agency_id": snapshot.agency_id,
        "artifact_hashes": sorted(snapshot.artifact_hashes),
    }
    serialized = json.dumps(key_material, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def _validate_snapshot(snapshot: DeliverySnapshot) -> None:
    required_values = (
        snapshot.batch_id,
        snapshot.agency_id,
        snapshot.target_period,
        snapshot.source_master_hash,
        snapshot.template_hash,
        snapshot.recipient_snapshot_hash,
        snapshot.message_hash,
    )
    if any(not value.strip() for value in required_values):
        raise ValueError("Delivery snapshot is incomplete")
    if any(not artifact_hash.strip() for artifact_hash in snapshot.artifact_hashes):
        raise ValueError("Artifact hashes must not be empty")
