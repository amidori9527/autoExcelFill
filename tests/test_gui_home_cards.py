from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from autoexcel.gui import HomePage


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
        self.assertEqual(positions[6], (2, 2, 1, 1))
        self.assertEqual(page.cards[1].objectName(), "homeFeaturedCard")


if __name__ == "__main__":
    unittest.main()
