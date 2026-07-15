from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from autoexcel.gui import FillPage


class GuiConfigWiringTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_fill_page_passes_configured_batch_size_to_task(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            config_path = Path(temporary_directory) / "config.ini"
            config_path.write_text("[fill]\nlimit_sheets = 100\n", encoding="utf-8")
            page = FillPage()

            with patch("autoexcel.gui.editable_config_path", return_value=config_path), patch.object(
                page, "start_task"
            ) as start_task:
                page.run()

            task = start_task.call_args.args[0]
            with patch("autoexcel.gui.run_fill_task") as run_fill_task:
                task(lambda _message: None)

            self.assertEqual(run_fill_task.call_args.kwargs["limit_sheets"], 100)


if __name__ == "__main__":
    unittest.main()
