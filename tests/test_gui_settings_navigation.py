from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication, QLineEdit, QMessageBox

from autoexcel.config_editor import read_ini
from autoexcel.gui import MainWindow
from autoexcel.license import LicenseInfo


class GuiSettingsNavigationTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def create_window(self, config_path: Path) -> MainWindow:
        with (
            patch("autoexcel.gui.editable_config_path", return_value=config_path),
            patch(
                "autoexcel.gui.load_license",
                return_value=LicenseInfo(False, "未配置密钥"),
            ),
        ):
            return MainWindow()

    def test_cancel_keeps_unsaved_changes_and_stays_on_settings(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text(
                "[ui]\ntheme = indigo\nmode = light\n[fill]\nworkbook = before.xlsx\n",
                encoding="utf-8",
            )
            window = self.create_window(config_path)
            window.show_page(7)
            workbook = window.settings_page.controls[("fill", "workbook")]
            self.assertIsInstance(workbook, QLineEdit)
            workbook.setText("after.xlsx")

            with patch(
                "autoexcel.gui.QMessageBox.question",
                return_value=QMessageBox.StandardButton.Cancel,
            ):
                window.show_page(0)

            self.assertEqual(window.pages.currentIndex(), 7)
            self.assertTrue(window.settings_page.has_unsaved_changes())
            self.assertTrue(window.sidebar.buttons[-1].isChecked())
            window.close()

    def test_save_persists_changes_then_switches_page(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text(
                "[ui]\ntheme = indigo\nmode = light\n[fill]\nworkbook = before.xlsx\n",
                encoding="utf-8",
            )
            window = self.create_window(config_path)
            window.show_page(7)
            workbook = window.settings_page.controls[("fill", "workbook")]
            self.assertIsInstance(workbook, QLineEdit)
            workbook.setText("after.xlsx")

            with (
                patch(
                    "autoexcel.gui.QMessageBox.question",
                    return_value=QMessageBox.StandardButton.Save,
                ),
                patch("autoexcel.gui.QMessageBox.information"),
            ):
                window.show_page(0)

            self.assertEqual(window.pages.currentIndex(), 0)
            self.assertFalse(window.settings_page.has_unsaved_changes())
            self.assertEqual(
                read_ini(config_path).get("fill", "workbook"),
                "after.xlsx",
            )
            window.close()

    def test_discard_restores_values_and_preview_theme_then_switches(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text(
                "[ui]\ntheme = indigo\nmode = light\n[fill]\nworkbook = before.xlsx\n",
                encoding="utf-8",
            )
            window = self.create_window(config_path)
            window.show_page(7)
            workbook = window.settings_page.controls[("fill", "workbook")]
            self.assertIsInstance(workbook, QLineEdit)
            workbook.setText("after.xlsx")
            window.settings_page.theme_combo.setCurrentIndex(
                window.settings_page.theme_combo.findData("emerald")
            )
            self.assertEqual(window.theme_name, "emerald")

            with patch(
                "autoexcel.gui.QMessageBox.question",
                return_value=QMessageBox.StandardButton.Discard,
            ):
                window.show_page(0)

            self.assertEqual(window.pages.currentIndex(), 0)
            self.assertEqual(workbook.text(), "before.xlsx")
            self.assertEqual(window.theme_name, "indigo")
            self.assertFalse(window.settings_page.has_unsaved_changes())
            window.close()

    def test_unchanged_settings_switch_without_prompt(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text(
                "[ui]\ntheme = indigo\nmode = light\n",
                encoding="utf-8",
            )
            window = self.create_window(config_path)
            window.show_page(7)

            with patch("autoexcel.gui.QMessageBox.question") as question:
                window.show_page(0)

            question.assert_not_called()
            self.assertEqual(window.pages.currentIndex(), 0)
            window.close()


if __name__ == "__main__":
    unittest.main()
