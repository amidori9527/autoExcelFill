from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from autoexcel.config_editor import read_fill_limit_sheets, read_ini, update_ini


class ConfigEditorTest(unittest.TestCase):
    def test_update_ini_preserves_comments_and_unedited_values(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "config.ini"
            path.write_text(
                "# 全局说明\n[fill]\n# 批次数量\nlimit_sheets = 20\nworkbook = old.xlsx\n",
                encoding="utf-8",
            )

            update_ini(path, {"fill": {"limit_sheets": "35"}})

            text = path.read_text(encoding="utf-8")
            self.assertIn("# 全局说明", text)
            self.assertIn("# 批次数量", text)
            self.assertIn("limit_sheets = 35", text)
            self.assertIn("workbook = old.xlsx", text)

    def test_update_ini_creates_group_section(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "conf.ini"

            update_ini(path, {"diff_orders": {"platform": "finerbit"}})

            parser = read_ini(path)
            self.assertEqual(parser.get("diff_orders", "platform"), "finerbit")

    def test_read_fill_limit_sheets_uses_saved_value(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "config.ini"
            path.write_text("[fill]\nlimit_sheets = 100\n", encoding="utf-8")

            self.assertEqual(read_fill_limit_sheets(path), 100)

    def test_read_fill_limit_sheets_defaults_to_twenty(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "config.ini"

            self.assertEqual(read_fill_limit_sheets(path), 20)


if __name__ == "__main__":
    unittest.main()
