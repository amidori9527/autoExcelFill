from __future__ import annotations

import os
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from autoexcel.gui import Sidebar, build_app_style


class GuiSidebarTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])
        cls.app.setStyleSheet(build_app_style("indigo"))

    def test_brand_title_and_caption_have_enough_width(self) -> None:
        sidebar = Sidebar()
        sidebar.resize(sidebar.width(), 700)
        sidebar.show()
        self.app.processEvents()

        self.assertGreaterEqual(
            sidebar.brand_label.width(), sidebar.brand_label.sizeHint().width()
        )
        self.assertGreaterEqual(
            sidebar.brand_caption.width(), sidebar.brand_caption.sizeHint().width()
        )

    def test_flow_sync_navigation_opens_flow_sync_page(self) -> None:
        sidebar = Sidebar()

        labels = [button.text() for button in sidebar.buttons]
        flow_sync_index = labels.index("流水同步")

        self.assertEqual(sidebar.page_indexes[flow_sync_index], 8)


if __name__ == "__main__":
    unittest.main()
