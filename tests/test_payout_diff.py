from __future__ import annotations

import csv
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from autoexcel import diff_orders, payout_diff


class PayoutDiffTest(unittest.TestCase):
    def write_upstream(self, path: Path) -> None:
        with path.open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(["TRANS_ID", "TRX_STATUS", "TRX_AMT", "FEE", "FED"])
            writer.writerow(["100", "COMPLETED", "100", "0.30", "0.04"])
            writer.writerow(["200", "FAILED", "75", "0.20", "0.03"])
            writer.writerow(["300", "COMPLETED", "50", "0.15", "0.02"])

    def write_backend(self, path: Path) -> None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "付款订单"
        sheet.append(
            [
                "transactionId",
                "交易状态",
                "付款金额(PKR)",
                "手续费(PKR)",
                "支付方式名称",
            ]
        )
        sheet.append(["100", "上游已打款", 100, 1.50, "jazzcash"])
        sheet.append(["300", "处理中", 50, 0.75, "jazzcash"])
        sheet.append(["400", "上游已打款", 25, 0.38, "easypaisa"])
        workbook.save(path)

    def test_reads_successful_orders_and_calculates_payout_fees(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            upstream = directory / "上游.csv"
            backend = directory / "付款订单.xlsx"
            self.write_upstream(upstream)
            self.write_backend(backend)

            result = payout_diff.read_job_diff_result(
                payout_diff.make_job(upstream, backend)
            ).result

            self.assertEqual(result.special_mode, "payout")
            self.assertEqual(result.a_count, 2)
            self.assertEqual(result.b_count, 2)
            self.assertEqual(diff_orders.unique_order_ids(result.a_only), ["300"])
            self.assertEqual(diff_orders.unique_order_ids(result.b_only), ["400"])
            self.assertEqual(result.mismatched, [])
            self.assertEqual(result.a_fee, Decimal("0.51"))
            self.assertEqual(result.b_fee, Decimal("1.88"))
            self.assertEqual(result.channel_cost, Decimal("0.52"))

    def test_directory_recognition_uses_headers(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            upstream = directory / "renamed.csv"
            backend = directory / "renamed.xlsx"
            self.write_upstream(upstream)
            self.write_backend(backend)

            jobs = payout_diff.find_jobs_in_directory(directory)

            self.assertEqual(len(jobs), 1)
            self.assertEqual(jobs[0].upstream_path, upstream)
            self.assertEqual(jobs[0].backend_path, backend)

    def test_payout_html_contains_both_profit_metrics(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            upstream = directory / "上游.csv"
            backend = directory / "付款订单.xlsx"
            self.write_upstream(upstream)
            self.write_backend(backend)
            job_result = payout_diff.read_job_diff_result(
                payout_diff.make_job(upstream, backend)
            )

            html = diff_orders.render_stats_html([job_result])

            self.assertIn("毛利润（平台手续费-上游手续费）", html)
            self.assertIn("渠道利润参考（平台手续费-渠道成本）", html)


if __name__ == "__main__":
    unittest.main()
