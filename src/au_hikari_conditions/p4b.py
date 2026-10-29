from __future__ import annotations

import hashlib
import json
import mimetypes
import re
import sqlite3
import string
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Generator, Mapping
from zipfile import BadZipFile

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from docx.opc.exceptions import PackageNotFoundError

from .p3 import GeneratedAgencyArtifacts
from .workflow import AgencyWorkflow, BatchStatus, DeliverySnapshot
from .zoho_mail import ZohoMailClient, ZohoMailError


_REPOSITORY_ROOT = Path(__file__).resolve().parents[2]
_UNAPPROVED_MARKERS = (
    "内部確認用",
    "未承認",
    "配信用成果物ではありません",
    "INTERNAL_REVIEW_ONLY",
)
_EMAIL_PATTERN = re.compile(r"^[^@\s,;<>]+@[^@\s,;<>]+\.[^@\s,;<>]+$")
_DATABASE_LOCKS: dict[str, threading.Lock] = {}
_DATABASE_LOCKS_GUARD = threading.Lock()


class DraftPreparationError(ValueError):
    """Raised when an agency draft is not safe to prepare or persist."""


class DuplicateDraftError(DraftPreparationError):
    pass


@dataclass(frozen=True)
class AgencyMailSettings:
    agency_id: str = field(repr=False)
    from_address: str = field(repr=False)
    to_addresses: tuple[str, ...] = field(repr=False)
    subject_template: str = field(repr=False)
    body_template: str = field(repr=False)
    cc_addresses: tuple[str, ...] = field(default=(), repr=False)
    bcc_addresses: tuple[str, ...] = field(default=(), repr=False)
    agency_name: str = field(default="", repr=False)
    salutation: str = field(default="", repr=False)
    template_hash: str | None = field(default=None, repr=False)


@dataclass(frozen=True)
class _Attachment:
    path: Path = field(repr=False)
    file_name: str
    sha256: str


@dataclass(frozen=True)
class DraftPreview:
    agency_id: str = field(repr=False)
    agency_name: str = field(repr=False)
    target_period: str
    from_address: str = field(repr=False)
    to_addresses: tuple[str, ...] = field(repr=False)
    cc_addresses: tuple[str, ...] = field(repr=False)
    bcc_addresses: tuple[str, ...] = field(repr=False)
    subject: str = field(repr=False)
    body: str = field(repr=False)
    attachments: tuple[_Attachment, ...] = field(repr=False)
    snapshot: DeliverySnapshot = field(repr=False)
    workflow: AgencyWorkflow = field(repr=False)

    def __repr__(self) -> str:
        return (
            "DraftPreview("
            f"target_period={self.target_period!r}, "
            f"attachment_count={len(self.attachments)})"
        )


@dataclass(frozen=True)
class DraftSaveResult:
    status: str
    attachment_count: int


class DraftHistoryStore:
    """SQLite idempotency ledger that stores no recipient, message, or raw agency data."""

    def __init__(self, path: Path) -> None:
        self.path = path.resolve()
        try:
            self.path.relative_to(_REPOSITORY_ROOT)
        except ValueError:
            pass
        else:
            raise DraftPreparationError(
                "Draft history must be stored outside the Git repository."
            )
        self.path.parent.mkdir(parents=True, exist_ok=True)
        lock_key = str(self.path).casefold()
        with _DATABASE_LOCKS_GUARD:
            self._lock = _DATABASE_LOCKS.setdefault(lock_key, threading.Lock())
        with self._connection() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS draft_attempts (
                    scope_hash TEXT PRIMARY KEY,
                    status TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    error_category TEXT,
                    http_status INTEGER
                )
                """
            )

    def reserve(self, agency_id: str, target_period: str) -> str:
        scope_hash = _scope_hash(agency_id, target_period)
        now = datetime.now(timezone.utc).isoformat()
        with self._lock, self._connection() as connection:
            try:
                connection.execute(
                    """
                    INSERT INTO draft_attempts
                        (scope_hash, status, created_at, updated_at)
                    VALUES (?, 'STARTED', ?, ?)
                    """,
                    (scope_hash, now, now),
                )
            except sqlite3.IntegrityError:
                raise DuplicateDraftError(
                    "A draft attempt already exists for this agency and period."
                ) from None
        return scope_hash

    def finish(
        self,
        scope_hash: str,
        *,
        status: str,
        error_category: str | None = None,
        http_status: int | None = None,
    ) -> None:
        if status not in {"CREATED", "FAILED", "OUTCOME_UNKNOWN"}:
            raise DraftPreparationError("Invalid draft history status.")
        now = datetime.now(timezone.utc).isoformat()
        with self._lock, self._connection() as connection:
            updated = connection.execute(
                """
                UPDATE draft_attempts
                SET status = ?, updated_at = ?, error_category = ?, http_status = ?
                WHERE scope_hash = ? AND status = 'STARTED'
                """,
                (status, now, error_category, http_status, scope_hash),
            ).rowcount
        if updated != 1:
            raise DraftPreparationError("Draft history state could not be updated.")

    def get_status(self, agency_id: str, target_period: str) -> str | None:
        scope_hash = _scope_hash(agency_id, target_period)
        with self._connection() as connection:
            row = connection.execute(
                "SELECT status FROM draft_attempts WHERE scope_hash = ?",
                (scope_hash,),
            ).fetchone()
        return str(row[0]) if row else None

    @contextmanager
    def _connection(self) -> Generator[sqlite3.Connection, None, None]:
        connection = sqlite3.connect(self.path, timeout=30, isolation_level="IMMEDIATE")
        try:
            connection.execute("PRAGMA journal_mode = WAL")
            with connection:
                yield connection
        finally:
            connection.close()


def prepare_draft_preview(
    *,
    artifacts: GeneratedAgencyArtifacts,
    target_period: str,
    settings: AgencyMailSettings,
    workflow: AgencyWorkflow,
    snapshot: DeliverySnapshot,
) -> DraftPreview:
    _validate_approval(
        artifacts=artifacts,
        target_period=target_period,
        settings=settings,
        workflow=workflow,
        snapshot=snapshot,
    )
    attachments = _validate_artifacts(artifacts, target_period)
    message = _render_message(settings, artifacts.agency_name, target_period)
    if _recipient_snapshot_hash(settings) != snapshot.recipient_snapshot_hash:
        raise DraftPreparationError(
            "Configured recipients do not match the approved snapshot."
        )
    if _message_hash(settings.from_address, message[0], message[1]) != snapshot.message_hash:
        raise DraftPreparationError(
            "Configured message does not match the approved snapshot."
        )
    if settings.template_hash and settings.template_hash != snapshot.template_hash:
        raise DraftPreparationError(
            "Configured mail template does not match the approved snapshot."
        )
    actual_hashes = tuple(sorted(attachment.sha256 for attachment in attachments))
    if actual_hashes != tuple(sorted(snapshot.artifact_hashes)):
        raise DraftPreparationError(
            "P3 attachments do not match the approved artifact snapshot."
        )
    return DraftPreview(
        agency_id=artifacts.agency_id,
        agency_name=artifacts.agency_name,
        target_period=target_period,
        from_address=settings.from_address,
        to_addresses=settings.to_addresses,
        cc_addresses=settings.cc_addresses,
        bcc_addresses=settings.bcc_addresses,
        subject=message[0],
        body=message[1],
        attachments=attachments,
        snapshot=snapshot,
        workflow=workflow,
    )


def save_draft(
    *,
    preview: DraftPreview,
    mail_client: ZohoMailClient,
    history: DraftHistoryStore,
    mode: str = "draft",
) -> DraftSaveResult:
    if mode != "draft":
        raise DraftPreparationError('Only mode="draft" is permitted.')
    _validate_live_approval(preview)
    _verify_attachment_hashes(preview.attachments)
    scope_hash = history.reserve(preview.agency_id, preview.target_period)
    uploaded: list[dict[str, str]] = []
    try:
        for attachment in preview.attachments:
            content_type = mimetypes.guess_type(attachment.file_name)[0]
            try:
                data = attachment.path.read_bytes()
            except OSError:
                raise DraftPreparationError(
                    "An approved attachment is missing or unreadable."
                ) from None
            if hashlib.sha256(data).hexdigest() != attachment.sha256:
                raise DraftPreparationError(
                    "An approved attachment changed after preview."
                )
            uploaded.append(
                _upload_attachment(mail_client, attachment, data, content_type)
            )
    except Exception as error:
        history.finish(
            scope_hash,
            status="FAILED",
            error_category=_safe_error_category(error),
        )
        if isinstance(error, DraftPreparationError):
            raise
        raise DraftPreparationError(
            "Zoho Mail attachment upload failed; no draft was saved."
        ) from None

    try:
        _validate_live_approval(preview)
        _verify_attachment_hashes(preview.attachments)
    except DraftPreparationError as error:
        history.finish(
            scope_hash,
            status="FAILED",
            error_category=_safe_error_category(error),
        )
        raise
    payload: dict[str, object] = {
        "fromAddress": preview.from_address.strip(),
        "toAddress": ",".join(address.strip() for address in preview.to_addresses),
        "subject": preview.subject,
        "content": preview.body,
        "mailFormat": "plaintext",
        "encoding": "UTF-8",
        "mode": "draft",
        "attachments": uploaded,
    }
    if preview.cc_addresses:
        payload["ccAddress"] = ",".join(
            address.strip() for address in preview.cc_addresses
        )
    if preview.bcc_addresses:
        payload["bccAddress"] = ",".join(
            address.strip() for address in preview.bcc_addresses
        )
    try:
        mail_client.save_draft(payload)
    except Exception as error:
        history.finish(
            scope_hash,
            status="OUTCOME_UNKNOWN",
            error_category=_safe_error_category(error),
        )
        raise DraftPreparationError(
            "Zoho Mail draft result is unknown; automatic retry is blocked."
        ) from None
    history.finish(scope_hash, status="CREATED")
    return DraftSaveResult(status="CREATED", attachment_count=len(uploaded))


def _validate_approval(
    *,
    artifacts: GeneratedAgencyArtifacts,
    target_period: str,
    settings: AgencyMailSettings,
    workflow: AgencyWorkflow,
    snapshot: DeliverySnapshot,
) -> None:
    if not settings.agency_id or settings.agency_id != artifacts.agency_id:
        raise DraftPreparationError("Mail settings do not match the P3 agency.")
    if settings.agency_name and settings.agency_name != artifacts.agency_name:
        raise DraftPreparationError("Mail settings agency name does not match P3.")
    if (
        snapshot.agency_id != artifacts.agency_id
        or snapshot.target_period != target_period
        or workflow.agency_id != artifacts.agency_id
        or workflow.batch_id != snapshot.batch_id
        or workflow.status != BatchStatus.APPROVED
        or not workflow.approver_id
        or workflow.approval_fingerprint != snapshot.fingerprint()
    ):
        raise DraftPreparationError(
            "A current approval for this agency, period, and snapshot is required."
        )
    if snapshot.target_period != target_period:
        raise DraftPreparationError("Approved period does not match the draft period.")


def _validate_live_approval(preview: DraftPreview) -> None:
    workflow = preview.workflow
    snapshot = preview.snapshot
    if (
        workflow.status != BatchStatus.APPROVED
        or not workflow.approver_id
        or workflow.approval_fingerprint != snapshot.fingerprint()
    ):
        raise DraftPreparationError(
            "Approval changed or expired; no Zoho request was made."
        )


def _validate_artifacts(
    artifacts: GeneratedAgencyArtifacts, target_period: str
) -> tuple[_Attachment, ...]:
    if artifacts.condition_count <= 0 or not artifacts.agency_name.strip():
        raise DraftPreparationError("P3 artifacts are incomplete.")
    paths = (artifacts.spreadsheet_path, artifacts.word_path)
    if paths[0].resolve() == paths[1].resolve():
        raise DraftPreparationError("P3 Excel and Word attachments must be distinct.")
    for path in paths:
        try:
            path.resolve().relative_to(_REPOSITORY_ROOT)
        except ValueError:
            continue
        raise DraftPreparationError(
            "Approved P3 attachments must be stored outside the Git repository."
        )
    if paths[0].suffix.lower() != ".xlsx" or paths[1].suffix.lower() != ".docx":
        raise DraftPreparationError("P3 attachments must be one Excel and one Word file.")
    _validate_workbook(paths[0], artifacts.agency_id, artifacts.agency_name, target_period)
    _validate_word(paths[1], artifacts.agency_id, artifacts.agency_name, target_period)
    return tuple(
        _Attachment(
            path=path,
            file_name=path.name,
            sha256=_file_hash(path),
        )
        for path in paths
    )


def _validate_workbook(
    path: Path, agency_id: str, agency_name: str, target_period: str
) -> None:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            if workbook.sheetnames != ["条件表"]:
                raise DraftPreparationError("P3 Excel attachment is not a condition sheet.")
            worksheet = workbook["条件表"]
            values = [
                str(worksheet.cell(row, column).value or "")
                for row in range(1, worksheet.max_row + 1)
                for column in range(1, min(worksheet.max_column, 5) + 1)
            ]
        finally:
            workbook.close()
    except (OSError, ValueError, KeyError, BadZipFile, InvalidFileException):
        raise DraftPreparationError("P3 Excel attachment could not be validated.") from None
    text = "\n".join(values)
    _ensure_approved_document(text, agency_id, agency_name, target_period)


def _validate_word(
    path: Path, agency_id: str, agency_name: str, target_period: str
) -> None:
    try:
        document = Document(path)
        values = [paragraph.text for paragraph in document.paragraphs]
        for table in document.tables:
            for row in table.rows:
                values.extend(cell.text for cell in row.cells)
        for section in document.sections:
            values.extend(paragraph.text for paragraph in section.header.paragraphs)
            values.extend(paragraph.text for paragraph in section.footer.paragraphs)
            for table in (*section.header.tables, *section.footer.tables):
                for row in table.rows:
                    values.extend(cell.text for cell in row.cells)
    except (
        OSError,
        ValueError,
        KeyError,
        BadZipFile,
        PackageNotFoundError,
    ):
        raise DraftPreparationError("P3 Word attachment could not be validated.") from None
    _ensure_approved_document("\n".join(values), agency_id, agency_name, target_period)


def _ensure_approved_document(
    text: str, agency_id: str, agency_name: str, target_period: str
) -> None:
    if any(marker.casefold() in text.casefold() for marker in _UNAPPROVED_MARKERS):
        raise DraftPreparationError(
            "P3 attachment is still marked internal or unapproved."
        )
    if agency_id not in text or agency_name not in text or target_period not in text:
        raise DraftPreparationError(
            "P3 attachment agency or period does not match the approved draft."
        )


def _render_message(
    settings: AgencyMailSettings, agency_name: str, target_period: str
) -> tuple[str, str]:
    addresses = tuple(
        address.strip()
        for address in (
            settings.from_address,
            *settings.to_addresses,
            *settings.cc_addresses,
            *settings.bcc_addresses,
        )
    )
    if any(not _EMAIL_PATTERN.fullmatch(address.strip()) for address in addresses):
        raise DraftPreparationError("A configured mail address is invalid.")
    if not settings.to_addresses:
        raise DraftPreparationError("At least one To recipient is required.")
    normalized_recipients = [address.strip().casefold() for address in addresses[1:]]
    if len(set(normalized_recipients)) != len(normalized_recipients):
        raise DraftPreparationError("Recipient addresses must be unique.")
    values = {
        "agency_name": settings.agency_name or agency_name,
        "target_period": target_period,
        "salutation": settings.salutation or f"{settings.agency_name or agency_name} 御中",
    }
    try:
        for template in (settings.subject_template, settings.body_template):
            for _, field_name, format_spec, conversion in string.Formatter().parse(
                template
            ):
                if field_name is not None and (
                    field_name not in values or format_spec or conversion
                ):
                    raise DraftPreparationError(
                        "Mail template contains an unsupported or malformed placeholder."
                    )
        subject = settings.subject_template.format_map(values).strip()
        body = settings.body_template.format_map(values).strip()
    except (KeyError, ValueError):
        raise DraftPreparationError(
            "Mail template contains an unsupported or malformed placeholder."
        ) from None
    if not subject or not body:
        raise DraftPreparationError("Rendered mail subject and body are required.")
    return subject, body


def _recipient_snapshot_hash(settings: AgencyMailSettings) -> str:
    payload = {
        "from": settings.from_address.strip().casefold(),
        "to": sorted(address.strip().casefold() for address in settings.to_addresses),
        "cc": sorted(address.strip().casefold() for address in settings.cc_addresses),
        "bcc": sorted(
            address.strip().casefold() for address in settings.bcc_addresses
        ),
    }
    return _canonical_hash(payload)


def _message_hash(from_address: str, subject: str, body: str) -> str:
    return _canonical_hash(
        {
            "from": from_address.strip().casefold(),
            "subject": subject,
            "body": body,
            "mail_format": "plaintext",
        }
    )


def _file_hash(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError:
        raise DraftPreparationError("A P3 attachment is missing or unreadable.") from None
    return digest.hexdigest()


def _verify_attachment_hashes(attachments: tuple[_Attachment, ...]) -> None:
    if any(_file_hash(item.path) != item.sha256 for item in attachments):
        raise DraftPreparationError("An approved attachment changed after preview.")


def _canonical_hash(payload: Mapping[str, object]) -> str:
    serialized = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()


def _scope_hash(agency_id: str, target_period: str) -> str:
    return _canonical_hash({"agency_id": agency_id, "target_period": target_period})


def _safe_error_category(error: Exception) -> str:
    if isinstance(error, ZohoMailError):
        return type(error).__name__
    if isinstance(error, DraftPreparationError):
        return type(error).__name__
    return "UnexpectedError"


def _upload_attachment(
    mail_client: ZohoMailClient,
    attachment: _Attachment,
    data: bytes,
    content_type: str | None,
) -> dict[str, str]:
    result = mail_client.upload_draft_attachment(
        file_name=attachment.file_name,
        content=data,
        content_type=content_type or "application/octet-stream",
    )
    if (
        result.get("attachmentName") != attachment.file_name
        or not result.get("storeName")
        or not result.get("attachmentPath")
    ):
        raise DraftPreparationError(
            "Zoho attachment response did not match the approved file."
        )
    return result
