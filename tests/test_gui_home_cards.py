from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from autoexcel.gui import FlowSyncPage, HomePage


class GuiHomeCardsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_full_access_uses_featured_three_column_card_layout(self) -> None:
        page = HomePage()
        page.set_feature_access(True, True, True, True)

        positions = {
            index: page.grid.getItemPosition(page.grid.indexOf(card))
            for index, card in enumerate(page.cards)
        }

        self.assertEqual(positions[1], (0, 0, 1, 2))
        self.assertEqual(positions[0], (0, 2, 1, 1))
        self.assertEqual(positions[2], (1, 0, 1, 1))
        self.assertEqual(positions[3], (1, 1, 1, 1))
        self.assertEqual(positions[4], (1, 2, 1, 1))
        self.assertEqual(positions[5], (2, 0, 1, 2))
        self.assertEqual(positions[6], (3, 0, 1, 3))
        self.assertEqual(positions[7], (2, 2, 1, 1))
        self.assertEqual(page.cards[1].objectName(), "homeFeaturedCard")

    def test_flow_sync_page_contains_three_sub_features(self) -> None:
        page = FlowSyncPage()

        self.assertEqual(
            page.feature_names,
            ("TP代付同步", "TP代收同步", "钱包流水同步"),
        )
        self.assertEqual(len(page.cards), 3)
        self.assertIsNotNone(page.workbook_picker)
        self.assertIsNotNone(page.payment_orders_picker)
        self.assertIsNotNone(page.collection_workbook_picker)
        self.assertIsNotNone(page.collection_orders_first_picker)
        self.assertIsNotNone(page.collection_orders_second_picker)
        self.assertIsNotNone(page.wallet_workbook_picker)
        self.assertIsNotNone(page.wallet_flow_picker)
        self.assertEqual(page.run_button.text(), "开始同步")

        page.focus_tp_collection()

        self.assertEqual(page.active_sync_feature, "collection")
        self.assertEqual(page.sync_forms.currentIndex(), 1)
        self.assertEqual(page.run_button.text(), "开始 TP代收同步")

        page.select_sync_feature("wallet")

        self.assertEqual(page.active_sync_feature, "wallet")
        self.assertEqual(page.sync_forms.currentIndex(), 2)
        self.assertEqual(page.run_button.text(), "开始钱包流水同步")


if __name__ == "__main__":
    unittest.main()
