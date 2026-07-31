from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QLabel
from PySide6.QtGui import QMovie

from autoexcel.gui import FlowSyncPage, HomePage, build_app_style
from autoexcel.license import LicenseInfo


class GuiHomeCardsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_full_access_uses_balanced_three_column_card_layout(self) -> None:
        page = HomePage()
        page.resize(1200, 800)
        page.show()
        self.app.processEvents()
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
        self.assertEqual(positions[5], (2, 0, 1, 1))
        self.assertEqual(positions[6], (2, 1, 1, 1))
        self.assertEqual(positions[7], (2, 2, 1, 1))
        self.assertEqual(page.cards[1].objectName(), "homeFeaturedCard")

    def test_narrow_home_page_reflows_cards_to_two_columns(self) -> None:
        page = HomePage()
        page.resize(700, 650)
        page.show()
        self.app.processEvents()
        page.set_feature_access(True, True, True, True)

        positions = {
            index: page.grid.getItemPosition(page.grid.indexOf(card))
            for index, card in enumerate(page.cards)
        }

        self.assertEqual(positions[1], (0, 0, 1, 2))
        self.assertEqual(positions[0], (1, 0, 1, 1))
        self.assertEqual(positions[2], (1, 1, 1, 1))
        self.assertEqual(positions[7], (4, 0, 1, 1))

    def test_home_hero_uses_looping_rabbit_animation(self) -> None:
        page = HomePage()

        self.assertIs(page.hero_icon.movie(), page.hero_movie)
        self.assertTrue(page.hero_movie.isValid())
        self.assertEqual(page.hero_movie.state(), QMovie.MovieState.Running)
        self.assertEqual(page.hero_icon.size().toTuple(), (92, 92))
        self.assertEqual(page.hero_movie.scaledSize().toTuple(), (92, 92))

    def test_flow_sync_page_contains_one_click_and_four_sub_features(self) -> None:
        page = FlowSyncPage()

        self.assertEqual(
            page.feature_names,
            ("TP代付同步", "TP代收同步", "TP提现同步", "钱包流水同步"),
        )
        self.assertIsNotNone(page.full_sync_card)
        self.assertEqual(len(page.cards), 4)
        self.assertEqual(page.full_sync_card.badge.text(), "")
        self.assertTrue(page.full_sync_card.badge.isHidden())
        self.assertIsNotNone(page.full_sync_card.corner_icon_label)
        self.assertFalse(page.full_sync_card.corner_icon_label.pixmap().isNull())
        self.assertTrue(all(card.corner_icon_label is None for card in page.cards))
        for card in [page.full_sync_card, *page.cards]:
            self.assertEqual(card.minimumHeight(), 88)
            self.assertEqual(card.maximumHeight(), 100)
            self.assertFalse(
                any(
                    label.objectName() == "cardAction"
                    for label in card.findChildren(QLabel)
                )
            )
        self.assertIsNotNone(page.flow_sync_directory_picker)
        self.assertIsNotNone(page.workbook_picker)
        self.assertIsNotNone(page.payment_orders_picker)
        self.assertIsNotNone(page.collection_workbook_picker)
        self.assertIsNotNone(page.collection_orders_first_picker)
        self.assertIsNotNone(page.collection_orders_second_picker)
        self.assertIsNotNone(page.withdrawal_workbook_picker)
        self.assertIsNotNone(page.withdrawal_orders_picker)
        self.assertIsNotNone(page.wallet_workbook_picker)
        self.assertIsNotNone(page.wallet_flow_picker)
        self.assertEqual(page.active_sync_feature, "folder")
        self.assertEqual(page.sync_forms.currentIndex(), 4)
        self.assertEqual(page.run_button.text(), "开始一键流水同步")
        self.assertTrue(page.full_sync_card.property("syncSelected"))
        self.assertFalse(page.cards[0].property("syncSelected"))
        self.assertFalse(page.cards[1].property("syncSelected"))
        self.assertFalse(page.cards[2].property("syncSelected"))
        self.assertFalse(page.cards[3].property("syncSelected"))

        page.focus_tp_collection()

        self.assertEqual(page.active_sync_feature, "collection")
        self.assertEqual(page.sync_forms.currentIndex(), 1)
        self.assertEqual(page.run_button.text(), "开始 TP代收同步")
        self.assertFalse(page.full_sync_card.property("syncSelected"))
        self.assertFalse(page.cards[0].property("syncSelected"))
        self.assertTrue(page.cards[1].property("syncSelected"))
        self.assertFalse(page.cards[2].property("syncSelected"))
        self.assertFalse(page.cards[3].property("syncSelected"))

        page.select_sync_feature("withdrawal")

        self.assertEqual(page.active_sync_feature, "withdrawal")
        self.assertEqual(page.sync_forms.currentIndex(), 2)
        self.assertEqual(page.run_button.text(), "开始 TP提现同步")
        self.assertFalse(page.full_sync_card.property("syncSelected"))
        self.assertFalse(page.cards[0].property("syncSelected"))
        self.assertFalse(page.cards[1].property("syncSelected"))
        self.assertTrue(page.cards[2].property("syncSelected"))
        self.assertFalse(page.cards[3].property("syncSelected"))

        page.select_sync_feature("wallet")

        self.assertEqual(page.active_sync_feature, "wallet")
        self.assertEqual(page.sync_forms.currentIndex(), 3)
        self.assertEqual(page.run_button.text(), "开始钱包流水同步")
        self.assertFalse(page.full_sync_card.property("syncSelected"))
        self.assertFalse(page.cards[0].property("syncSelected"))
        self.assertFalse(page.cards[1].property("syncSelected"))
        self.assertFalse(page.cards[2].property("syncSelected"))
        self.assertTrue(page.cards[3].property("syncSelected"))

    def test_flow_sync_feature_cards_fit_without_horizontal_scroll(self) -> None:
        self.app.setStyleSheet(build_app_style("indigo"))
        page = FlowSyncPage()
        page.resize(1000, 800)
        page.show()
        self.app.processEvents()

        self.assertEqual(page.form_scroll.horizontalScrollBar().maximum(), 0)
        self.assertEqual(
            page.form_scroll.verticalScrollBarPolicy(),
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff,
        )
        self.assertLessEqual(
            page.form_panel.width(),
            page.form_scroll.viewport().width(),
        )
        self.assertTrue(all(card.width() < 240 for card in page.cards))

    def test_flow_sync_run_requires_license(self) -> None:
        page = FlowSyncPage()
        with (
            patch(
                "autoexcel.gui.load_license",
                return_value=LicenseInfo(False, "未配置密钥"),
            ),
            patch("autoexcel.gui.QMessageBox.warning") as warning,
            patch.object(page, "run_full_flow_sync") as run_full_flow_sync,
        ):
            page.run_sync()

        warning.assert_called_once()
        run_full_flow_sync.assert_not_called()


if __name__ == "__main__":
    unittest.main()
