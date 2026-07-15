from __future__ import annotations

import os
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from autoexcel.gui import MainWindow
from autoexcel.license import LicenseInfo


class GuiLicenseAccessTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_no_license_hides_order_features(self) -> None:
        with patch("autoexcel.gui.load_license", return_value=LicenseInfo(False, "未配置密钥")):
            window = MainWindow()

        self.assertTrue(window.sidebar.buttons[2].isHidden())
        self.assertTrue(window.sidebar.buttons[3].isHidden())
        self.assertTrue(window.home_page.cards[1].isHidden())
        self.assertTrue(window.home_page.cards[2].isHidden())

    def test_full_license_shows_order_features(self) -> None:
        with patch("autoexcel.gui.load_license", return_value=LicenseInfo(False, "未配置密钥")):
            window = MainWindow()
        info = LicenseInfo(True, "密钥有效", frozenset({"order_diff", "fetch_orders"}))

        window.apply_license(info)

        self.assertFalse(window.sidebar.buttons[2].isHidden())
        self.assertFalse(window.sidebar.buttons[3].isHidden())
        self.assertFalse(window.home_page.cards[1].isHidden())
        self.assertFalse(window.home_page.cards[2].isHidden())


if __name__ == "__main__":
    unittest.main()
