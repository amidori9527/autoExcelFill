from __future__ import annotations

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

from autoexcel.gui_tasks import _jobs_from_directory, run_fill_task


class GuiTasksTest(unittest.TestCase):
    def test_jobs_from_directory_matches_one_pair_for_date(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            (directory / "TranDetailReport_123_202607140000.xlsx").touch()
            (directory / "收款订单_202607140001.xlsx").touch()

            jobs = _jobs_from_directory(directory, date(2026, 7, 14))

            self.assertEqual(len(jobs), 1)
            self.assertTrue(jobs[0].upstream_path.name.startswith("TranDetailReport"))
            self.assertTrue(jobs[0].backend_path.name.startswith("收款订单"))

    def test_jobs_from_directory_rejects_empty_directory(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            with self.assertRaisesRegex(FileNotFoundError, "没有找到 Excel"):
                _jobs_from_directory(Path(temporary_directory), date(2026, 7, 14))

    def test_fill_task_rejects_non_xlsx_file_before_processing(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            workbook = Path(temporary_directory) / "orders.xls"
            workbook.touch()

            with self.assertRaisesRegex(ValueError, "xlsx"):
                run_fill_task(workbook, date(2026, 7, 14), lambda _message: None)

    def test_fill_task_updates_colored_sheet(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            workbook_path = directory / "ledger.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Merchant A"
            sheet.sheet_properties.tabColor = "3157D5"
            sheet["A1"] = date(2026, 7, 13)
            sheet["K1"] = 100
            sheet["M1"] = 120
            workbook.save(workbook_path)

            with patch(
                "autoexcel.gui_tasks.create_process_log_path",
                return_value=directory / "fill.log",
            ):
                result = run_fill_task(
                    workbook_path, date(2026, 7, 14), lambda _message: None
                )

            updated = load_workbook(workbook_path, data_only=False)
            self.assertEqual(result.title, "Excel 填充完成")
            self.assertEqual(updated["Merchant A"]["A2"].value.date(), date(2026, 7, 14))

    def test_fill_task_updates_b2b_and_income_after_colored_sheets(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            workbook_path = directory / "ledger.xlsx"
            workbook = Workbook()
            colored = workbook.active
            colored.title = "Merchant A"
            colored.sheet_properties.tabColor = "3157D5"
            colored["A1"] = date(2026, 7, 13)

            b2b = workbook.create_sheet("B2B支出")
            b2b.append([None] * 15)
            b2b.append(["日期", *[f"字段{index}" for index in range(2, 16)]])
            b2b.append([date(2026, 7, 13), *([1] * 14)])
            b2b.append(["汇总", *([None] * 14)])
            b2b_table = Table(displayName="B2BTable", ref="A2:O4")
            b2b_table.totalsRowCount = 1
            b2b.add_table(b2b_table)

            income = workbook.create_sheet("收入")
            income.append(["日期", *[f"字段{index}" for index in range(2, 14)]])
            income.append([date(2026, 7, 13), *([1] * 12)])
            income.append(["汇总", *([None] * 12)])
            table = Table(displayName="IncomeTable", ref="A1:M3")
            table.totalsRowCount = 1
            income.add_table(table)
            workbook.save(workbook_path)

            with patch(
                "autoexcel.gui_tasks.create_process_log_path",
                return_value=directory / "fill.log",
            ):
                run_fill_task(workbook_path, date(2026, 7, 14), lambda _message: None)

            updated = load_workbook(workbook_path, data_only=False)
            updated_b2b = updated["B2B支出"]
            self.assertEqual(updated_b2b["A3"].value.date(), date(2026, 7, 13))
            self.assertEqual(updated_b2b["A4"].value.date(), date(2026, 7, 14))
            self.assertEqual(updated_b2b["A5"].value, "汇总")
            self.assertEqual(updated_b2b.tables["B2BTable"].ref, "A2:O5")
            updated_income = updated["收入"]
            self.assertEqual(updated_income["A2"].value.date(), date(2026, 7, 13))
            self.assertEqual(updated_income["A3"].value.date(), date(2026, 7, 14))
            self.assertEqual(updated_income["A4"].value, "汇总")
            self.assertEqual(updated_income.tables["IncomeTable"].ref, "A1:M4")


if __name__ == "__main__":
    unittest.main()
