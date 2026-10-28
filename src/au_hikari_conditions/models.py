from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from enum import Enum


class ValueType(str, Enum):
    TEXT = "text"
    INTEGER = "integer"
    DECIMAL = "decimal"
    DATE = "date"
    BOOLEAN = "boolean"


@dataclass(frozen=True)
class DeliverableItemDefinition:
    item_id: str
    label: str
    value_type: ValueType
    required: bool = False
    description: str = ""


DELIVERABLE_ITEM_ID_DICTIONARY: dict[str, DeliverableItemDefinition] = {}


@dataclass(frozen=True)
class SourceDocument:
    source_document_id: str
    file_name: str
    relative_path: str
    source_type: str
    sha256: str | None = None
    received_at: datetime | None = None


@dataclass(frozen=True)
class ConditionItem:
    condition_item_id: str
    deliverable_item_id: str
    value: str | int | Decimal | bool | None
    source_document_id: str
    agency_id: str | None = None
    valid_from: date | None = None
    valid_until: date | None = None


@dataclass(frozen=True)
class Agency:
    agency_id: str
    name: str
    external_code: str | None = None
    is_active: bool = True
