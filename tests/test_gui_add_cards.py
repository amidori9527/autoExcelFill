from __future__ import annotations

import os
from pathlib import Path
import unittest
from unittest.mock import ANY, patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

from autoexcel.gui import AddCardsPage


class GuiAddCardsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_color_controls_support_selected_and_random_modes(self) -> None:
        page = AddCardsPage()

        self.assertEqual(page.sheet_color_mode.count(), 2)
        self.assertEqual(page.sheet_color_mode.currentData(), "selected")
        self.assertTrue(page.sheet_color_button.isEnabled())
        self.assertEqual(page.selected_sheet_color, "#3157D5")
        self.assertEqual(page.sheet_color_button.width(), 210)
        self.assertEqual(page.sheet_color_button.height(), 40)

        with patch(
            "autoexcel.gui.QColorDialog.getColor",
            return_value=QColor("#e11d48"),
        ):
            page.sheet_color_button.click()

        self.assertEqual(page.selected_sheet_color, "#E11D48")
        self.assertIn("#E11D48", page.sheet_color_button.text())

        page.sheet_color_mode.setCurrentIndex(1)

        self.assertEqual(page.sheet_color_mode.currentData(), "random")
        self.assertFalse(page.sheet_color_button.isEnabled())

    def test_random_mode_is_passed_to_background_task(self) -> None:
        page = AddCardsPage()
        page.cards_input.setPlainText("1234\n3121")
        page.sheet_color_mode.setCurrentIndex(1)
        workbook = Path("/tmp/cards.xlsx")

        with (
            patch("autoexcel.gui.load_license") as load_license,
            patch.object(page.path_picker, "path", return_value=workbook),
            patch.object(page, "start_task") as start_task,
        ):
            load_license.return_value.allows.return_value = True
            page.run()

        task = start_task.call_args.args[0]
        with patch("autoexcel.gui.run_add_cards_task") as run_task:
            task(lambda _message: None)

        run_task.assert_called_once_with(
            workbook,
            "1234\n3121",
            ANY,
            sheet_color=None,
            random_sheet_colors=True,
        )


if __name__ == "__main__":
    unittest.main()
