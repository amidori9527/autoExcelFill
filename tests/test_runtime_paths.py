from __future__ import annotations

from pathlib import Path
import unittest
from unittest.mock import patch

from autoexcel.runtime_paths import application_directory


class RuntimePathsTest(unittest.TestCase):
    def test_mac_bundle_uses_directory_beside_app(self) -> None:
        executable = "/Applications/AutoExcel.app/Contents/MacOS/AutoExcel"
        with patch("autoexcel.runtime_paths.sys.platform", "darwin"), patch(
            "autoexcel.runtime_paths.sys.executable", executable
        ):
            self.assertEqual(application_directory(), Path("/Applications"))

    def test_windows_uses_executable_parent(self) -> None:
        executable = "/release/AutoExcel/AutoExcel.exe"
        with patch("autoexcel.runtime_paths.sys.platform", "win32"), patch(
            "autoexcel.runtime_paths.sys.executable", executable
        ):
            self.assertEqual(application_directory(), Path("/release/AutoExcel"))
