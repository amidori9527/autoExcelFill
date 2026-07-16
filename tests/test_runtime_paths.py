from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from autoexcel.runtime_paths import application_directory, ensure_workspace_directories


class RuntimePathsTest(unittest.TestCase):
    def test_mac_bundle_uses_directory_beside_app(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            base_directory = Path(temporary_directory).resolve()
            executable = (
                base_directory
                / "SmartSheet Desk.app"
                / "Contents"
                / "MacOS"
                / "SmartSheet Desk"
            )
            with patch("autoexcel.runtime_paths.sys.platform", "darwin"), patch(
                "autoexcel.runtime_paths.sys.executable", str(executable)
            ):
                self.assertEqual(application_directory(), base_directory)

    def test_windows_uses_executable_parent(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            application_directory_path = (
                Path(temporary_directory).resolve() / "release" / "SmartSheet Desk"
            )
            executable = application_directory_path / "SmartSheet Desk.exe"
            with patch("autoexcel.runtime_paths.sys.platform", "win32"), patch(
                "autoexcel.runtime_paths.sys.executable", str(executable)
            ):
                self.assertEqual(application_directory(), application_directory_path)

    def test_frozen_app_creates_workspace_directories_beside_executable(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            application_directory_path = (
                Path(temporary_directory).resolve() / "SmartSheet Desk"
            )
            executable = application_directory_path / "SmartSheet Desk.exe"
            with patch("autoexcel.runtime_paths.sys.frozen", True, create=True), patch(
                "autoexcel.runtime_paths.sys.platform", "win32"
            ), patch("autoexcel.runtime_paths.sys.executable", str(executable)):
                workspace = ensure_workspace_directories()

            self.assertEqual(workspace, application_directory_path / "workspace")
            self.assertTrue((workspace / "diffOrders").is_dir())
