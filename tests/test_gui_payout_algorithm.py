from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QMessageBox

from autoexcel.gui import DiffPage


class GuiPayoutAlgorithmTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_payout_algorithm_switch_controls_collection_excel_field(self) -> None:
        page = DiffPage()

        self.assertEqual(page.payout_algorithm_combo.currentData(), "zee")
        self.assertTrue(page.payout_collection_upstream_field.isHidden())

        page.payout_algorithm_combo.setCurrentIndex(
            page.payout_algorithm_combo.findData("finerbit")
        )

        self.assertFalse(page.payout_collection_upstream_field.isHidden())

    def test_collection_and_payout_default_to_manual_upload(self) -> None:
        page = DiffPage()

        self.assertEqual(page.mode_combo.currentData(), "files")
        self.assertEqual(page.mode_stack.currentIndex(), 1)
        self.assertEqual(page.payout_mode_combo.currentData(), "files")
        self.assertEqual(page.payout_mode_stack.currentIndex(), 1)

    def test_finerbit_execution_is_blocked_until_calculation_is_connected(self) -> None:
        page = DiffPage()
        page.payout_algorithm_combo.setCurrentIndex(
            page.payout_algorithm_combo.findData("finerbit")
        )

        with patch.object(QMessageBox, "warning") as warning, patch.object(
            page, "start_task"
        ) as start_task:
            page.run_payout()

        warning.assert_called_once()
        start_task.assert_not_called()


if __name__ == "__main__":
    unittest.main()
