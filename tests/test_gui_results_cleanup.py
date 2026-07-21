from __future__ import annotations

from datetime import date, datetime
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QMessageBox

from autoexcel.gui import ResultsPage


class GuiResultsCleanupTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_historical_results_are_strictly_older_than_seven_days(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            result_dir = Path(temporary_directory)
            old_result = result_dir / "order_diff_group_20260713_120000.html"
            boundary_result = result_dir / "order_diff_group_20260714_120000.html"
            recent_result = result_dir / "payout_order_diff_group_20260721_120000.html"
            unrelated_file = result_dir / "notes_20260701_120000.html"
            for path in (old_result, boundary_result, recent_result, unrelated_file):
                path.touch()

            with patch("autoexcel.gui.get_result_dir", return_value=result_dir):
                page = ResultsPage()
                candidates = page.historical_result_entries(date(2026, 7, 21))

            self.assertEqual([entry[2] for entry in candidates], [old_result])

    def test_cleanup_only_deletes_after_user_confirms(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            old_result = Path(temporary_directory) / "order_diff_group_20260713_120000.html"
            old_result.touch()
            entry = (datetime(2026, 7, 13, 12), "代收", old_result)
            page = ResultsPage()

            with patch.object(
                page, "historical_result_entries", return_value=[entry]
            ), patch.object(
                QMessageBox,
                "question",
                return_value=QMessageBox.StandardButton.No,
            ) as question:
                page.clean_historical_results()

            self.assertTrue(old_result.exists())
            self.assertIn("七天前", question.call_args.args[2])

            with patch.object(
                page, "historical_result_entries", return_value=[entry]
            ), patch.object(
                QMessageBox,
                "question",
                return_value=QMessageBox.StandardButton.Yes,
            ), patch.object(QMessageBox, "information") as information, patch.object(
                page, "refresh_results"
            ) as refresh_results:
                page.clean_historical_results()

            self.assertFalse(old_result.exists())
            information.assert_called_once()
            refresh_results.assert_called_once()


if __name__ == "__main__":
    unittest.main()
