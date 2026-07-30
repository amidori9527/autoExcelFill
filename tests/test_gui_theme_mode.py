from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFrame

from autoexcel.config_editor import read_ini
from autoexcel.gui import (
    MainWindow,
    NoWheelComboBox,
    build_app_style,
    normalize_ui_mode,
    ui_mode_from_config,
)
from autoexcel.license import LicenseInfo


class GuiThemeModeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_normalizes_mode_and_reads_config(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text("[ui]\nmode = dark\n", encoding="utf-8")

            self.assertEqual(normalize_ui_mode("DARK"), "dark")
            self.assertEqual(normalize_ui_mode("unknown"), "light")
            self.assertEqual(ui_mode_from_config(config_path), "dark")

    def test_dark_style_resolves_all_color_tokens(self) -> None:
        style = build_app_style("indigo", "dark")

        self.assertNotIn("__THEME_", style)
        self.assertIn("#111522", style)
        self.assertIn("#181d2a", style)
        self.assertIn("#f5f7fb", style)

    def test_combo_popup_uses_themed_item_states(self) -> None:
        light_style = build_app_style("indigo", "light")
        dark_style = build_app_style("indigo", "dark")

        self.assertIn("QComboBox QAbstractItemView::item:hover", light_style)
        self.assertIn("QComboBox QAbstractItemView::item:selected", light_style)
        self.assertIn("selection-color: #ffffff", light_style)
        self.assertIn("selection-background-color: #6673d9", light_style)
        self.assertIn("background: #181d2a", dark_style)
        self.assertIn("selection-background-color: #7c88f2", dark_style)

    def test_combo_popup_has_no_native_frame(self) -> None:
        combo = NoWheelComboBox()
        popup = combo.view().window()

        self.assertEqual(popup.objectName(), "comboPopup")
        self.assertTrue(
            popup.testAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        )
        self.assertIsInstance(popup, QFrame)
        self.assertEqual(popup.frameShape(), QFrame.Shape.NoFrame)

    def test_sidebar_button_toggles_and_persists_mode(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text("[ui]\ntheme = indigo\nmode = light\n", encoding="utf-8")
            with (
                patch("autoexcel.gui.editable_config_path", return_value=config_path),
                patch(
                    "autoexcel.gui.load_license",
                    return_value=LicenseInfo(False, "未配置密钥"),
                ),
            ):
                window = MainWindow()
                window.show()
                self.app.processEvents()

                self.assertGreater(
                    window.sidebar.mode_toggle.x(),
                    window.sidebar.footer_label.x(),
                )
                self.assertEqual(window.ui_mode, "light")
                self.assertIn("深色", window.sidebar.mode_toggle.toolTip())

                window.sidebar.mode_toggle.click()

                self.assertEqual(window.ui_mode, "dark")
                self.assertEqual(read_ini(config_path).get("ui", "mode"), "dark")
                self.assertIn("浅色", window.sidebar.mode_toggle.toolTip())
                window.close()


if __name__ == "__main__":
    unittest.main()
