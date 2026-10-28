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
    source_labels: tuple[str, ...] = ()
    approved: bool = False


DELIVERABLE_ITEM_ID_DICTIONARY: dict[str, DeliverableItemDefinition] = {
    "p1.applicable_period": DeliverableItemDefinition(
        item_id="p1.applicable_period",
        label="適用期間",
        value_type=ValueType.TEXT,
        source_labels=("対象期間", "設定期間", "業務実施対象期間", "適用期間"),
        description="暫定ID。正式な成果項目IDとの対応は設計書確認待ち。",
    ),
    "p1.isp": DeliverableItemDefinition(
        item_id="p1.isp",
        label="ISP",
        value_type=ValueType.TEXT,
        source_labels=("ISP", "プロバイダ", "So-net", "BIGLOBE", "@nifty"),
        description="暫定ID。ISP別の条件適用規則は設計書確認待ち。",
    ),
    "p1.commission_amount": DeliverableItemDefinition(
        item_id="p1.commission_amount",
        label="手数料・報奨金額",
        value_type=ValueType.DECIMAL,
        source_labels=("手数料", "報奨金額", "金額", "支援金"),
        description="暫定ID。対象サービス・税区分・単位との対応は要確認。",
    ),
    "p1.achievement_condition": DeliverableItemDefinition(
        item_id="p1.achievement_condition",
        label="成果条件",
        value_type=ValueType.TEXT,
        source_labels=("成果条件", "対象申込", "支払対象", "条件を満たす"),
        description="暫定ID。条件文は原文を保持し、業務判断は行わない。",
    ),
    "p1.ladder": DeliverableItemDefinition(
        item_id="p1.ladder",
        label="ラダー",
        value_type=ValueType.TEXT,
        source_labels=("ラダー", "段階", "件数"),
        description="暫定ID。閾値と金額の対応関係は要確認。",
    ),
    "p1.clawback_condition": DeliverableItemDefinition(
        item_id="p1.clawback_condition",
        label="戻入条件",
        value_type=ValueType.TEXT,
        source_labels=("戻入", "返還", "返金", "相殺"),
        description="暫定ID。戻入対象・期間・上限は原文確認待ち。",
    ),
}


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
