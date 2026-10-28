import unittest
from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from au_hikari_conditions import Agency, ConditionItem, SourceDocument


class ModelSmokeTest(unittest.TestCase):
    def test_entities_can_be_instantiated(self) -> None:
        source = SourceDocument(
            source_document_id="source-1",
            file_name="conditions.xlsx",
            relative_path="01_上位原本/conditions.xlsx",
            source_type="spreadsheet",
        )
        condition = ConditionItem(
            condition_item_id="condition-1",
            deliverable_item_id="item-1",
            value="sample",
            source_document_id=source.source_document_id,
            valid_from=date(2026, 1, 1),
        )
        agency = Agency(agency_id="agency-1", name="サンプル代理店")

        self.assertEqual(condition.source_document_id, source.source_document_id)
        self.assertEqual(agency.name, "サンプル代理店")


if __name__ == "__main__":
    unittest.main()
