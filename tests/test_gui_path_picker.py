from __future__ import annotations

import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QMimeData, QUrl
from PySide6.QtWidgets import QApplication

from autoexcel.gui import PathPicker


class FakeDropEvent:
    def __init__(self, path: Path) -> None:
        self.mime_data = QMimeData()
        self.mime_data.setUrls([QUrl.fromLocalFile(str(path))])
        self.accepted = False
        self.ignored = False

    def mimeData(self) -> QMimeData:
        return self.mime_data

    def acceptProposedAction(self) -> None:
        self.accepted = True

    def ignore(self) -> None:
        self.ignored = True


class GuiPathPickerTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_directory_picker_accepts_dropped_folder(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            picker = PathPicker("directory")
            drag_event = FakeDropEvent(directory)
            drop_event = FakeDropEvent(directory)

            picker.dragEnterEvent(drag_event)
            picker.dropEvent(drop_event)

            self.assertTrue(picker.acceptDrops())
            self.assertFalse(picker.edit.acceptDrops())
            self.assertTrue(drag_event.accepted)
            self.assertTrue(drop_event.accepted)
            self.assertEqual(picker.path(), directory)
            self.assertIn("拖入文件夹", picker.edit.placeholderText())


if __name__ == "__main__":
    unittest.main()
