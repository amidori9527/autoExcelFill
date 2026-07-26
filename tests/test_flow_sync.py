from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from autoexcel.flow_sync import (
    FlowSyncFiles,
    TP_COLLECTION_HEADERS,
    TP_PAYOUT_HEADERS,
    TpCollectionSyncResult,
    TpPayoutSyncResult,
    WALLET_FLOW_HEADERS,
    _replace_tp_collection_roots,
    _replace_tp_payout_roots,
    _replace_wallet_flow_roots,
    discover_flow_sync_files,
    sync_all_flows,
)
from autoexcel.fast_xlsx import MAIN_NS


def tag(name: str) -> str:
    return f"{{{MAIN_NS}}}{name}"


def add_inline_cell(
    row: ET.Element,
    ref: str,
    value: str,
    style: str | None = None,
) -> None:
    attributes = {"r": ref, "t": "inlineStr"}
    if style is not None:
        attributes["s"] = style
    cell = ET.SubElement(row, tag("c"), attributes)
    inline = ET.SubElement(cell, tag("is"))
    ET.SubElement(inline, tag("t")).text = value


def make_row(
    row_number: int,
    values: dict[str, str],
    style: str | None = None,
) -> ET.Element:
    row = ET.Element(tag("row"), {"r": str(row_number), "ht": "25"})
    for column, value in values.items():
        add_inline_cell(row, f"{column}{row_number}", value, style)
    return row


def make_workbook_root(
    detail_ids: list[str],
    include_summary: bool,
    detail_style: str,
) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    max_row = len(detail_ids) + (4 if include_summary else 1)
    ET.SubElement(root, tag("dimension"), {"ref": f"A1:V{max_row}"})
    sheet_data = ET.SubElement(root, tag("sheetData"))
    header_values = {
        chr(ord("A") + index): value
        for index, value in enumerate(TP_PAYOUT_HEADERS)
    }
    sheet_data.append(make_row(1, header_values, "20"))
    for offset, order_id in enumerate(detail_ids, start=2):
        sheet_data.append(
            make_row(
                offset,
                {"A": order_id, "B": f"merchant-{order_id}", "V": ""},
                detail_style,
            )
        )
    if include_summary:
        detail_end = len(detail_ids) + 1
        sheet_data.append(make_row(detail_end + 1, {}))
        sheet_data.append(
            make_row(
                detail_end + 2,
                {"E": "支付通道", "F": "平台钱包ID"},
                "30",
            )
        )
        sheet_data.append(make_row(detail_end + 3, {"E": "总计", "G": "100"}, "30"))
    ET.SubElement(
        root,
        tag("autoFilter"),
        {"ref": f"A1:V{len(detail_ids) + 1}"},
    )
    return root


def row_by_number(root: ET.Element) -> dict[int, ET.Element]:
    return {
        int(row.attrib["r"]): row
        for row in root.findall(f"{tag('sheetData')}/{tag('row')}")
    }


def make_collection_root(
    detail_ids: list[str],
    status: str,
    include_summary: bool,
    source_header: bool = False,
) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    max_row = len(detail_ids) + (4 if include_summary else 1)
    ET.SubElement(root, tag("dimension"), {"ref": f"A1:Z{max_row}"})
    sheet_data = ET.SubElement(root, tag("sheetData"))
    headers = list(TP_COLLECTION_HEADERS)
    if source_header:
        headers[12] = "平台钱包ID"
    sheet_data.append(
        make_row(
            1,
            {
                chr(ord("A") + index): value
                for index, value in enumerate(headers)
            },
            "20",
        )
    )
    for offset, order_id in enumerate(detail_ids, start=2):
        sheet_data.append(
            make_row(
                offset,
                {"A": order_id, "B": f"merchant-{order_id}", "W": status},
                "10",
            )
        )
    if include_summary:
        detail_end = len(detail_ids) + 1
        sheet_data.append(make_row(detail_end + 1, {}))
        sheet_data.append(
            make_row(
                detail_end + 2,
                {
                    "D": "求和项:实收金额(৳)",
                    "E": "求和项:手续费(৳)",
                    "F": "计数项:平台订单号",
                },
                "30",
            )
        )
        sheet_data.append(
            make_row(detail_end + 3, {"B": "总计", "F": "100"}, "30")
        )
    ET.SubElement(
        root,
        tag("autoFilter"),
        {"ref": f"A1:Z{len(detail_ids) + 1}"},
    )
    return root


def make_wallet_root(
    transaction_ids: list[str],
    include_summary: bool,
    source_layout: bool,
) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    header_row = 2 if source_layout else 1
    first_detail_row = header_row + 1
    max_row = first_detail_row + len(transaction_ids) + (
        2 if include_summary else -1
    )
    ET.SubElement(root, tag("dimension"), {"ref": f"A1:O{max_row}"})
    sheet_data = ET.SubElement(root, tag("sheetData"))
    if source_layout:
        sheet_data.append(make_row(1, {"A": "平台钱包流水记录"}, "40"))
    sheet_data.append(
        make_row(
            header_row,
            {
                chr(ord("A") + index): value
                for index, value in enumerate(WALLET_FLOW_HEADERS)
            },
            "20",
        )
    )
    for offset, transaction_id in enumerate(
        transaction_ids,
        start=first_detail_row,
    ):
        completed_at = (
            "2026-07-24 00:00:31"
            if transaction_id == "cross-day"
            else "2026-07-23 12:00:00"
        )
        sheet_data.append(
            make_row(
                offset,
                {
                    "C": "正常",
                    "D": completed_at,
                    "E": "2026-07-23 23:59:46",
                    "F": transaction_id,
                    "K": "长款",
                    "M": "100",
                },
                "10",
            )
        )
    if include_summary:
        detail_end = first_detail_row + len(transaction_ids) - 1
        sheet_data.append(make_row(detail_end + 1, {}))
        sheet_data.append(
            make_row(
                detail_end + 2,
                {
                    "B": "钱包ID",
                    "C": "求和项:交易金额",
                    "D": "计数项:交易类型",
                },
                "30",
            )
        )
        sheet_data.append(
            make_row(detail_end + 3, {"B": "总计", "D": "100"}, "30")
        )
    if not source_layout:
        ET.SubElement(
            root,
            tag("autoFilter"),
            {"ref": f"A1:N{first_detail_row + len(transaction_ids) - 1}"},
        )
    return root


class TpPayoutSyncTest(unittest.TestCase):
    def test_source_shared_strings_are_converted_for_target_workbook(self) -> None:
        target = make_workbook_root(["old-1"], True, "10")
        source = make_workbook_root(["placeholder"], False, "1")
        source_cell = row_by_number(source)[2].find(tag("c"))
        source_cell.attrib["t"] = "s"
        for child in list(source_cell):
            source_cell.remove(child)
        ET.SubElement(source_cell, tag("v")).text = "42"

        _replace_tp_payout_roots(
            target,
            source,
            {},
            {42: "new-shared-order"},
        )

        target_cell = row_by_number(target)[2].find(tag("c"))
        self.assertEqual(target_cell.attrib["t"], "inlineStr")
        self.assertEqual(
            target_cell.find(f"{tag('is')}/{tag('t')}").text,
            "new-shared-order",
        )

    def test_more_source_rows_shift_summary_down(self) -> None:
        target = make_workbook_root(["old-1", "old-2"], True, "10")
        source = make_workbook_root(
            ["new-1", "new-2", "new-3", "new-4"],
            False,
            "1",
        )

        result = _replace_tp_payout_roots(target, source, {}, {})

        self.assertEqual(result.old_detail_end_row, 3)
        self.assertEqual(result.new_detail_end_row, 5)
        self.assertEqual(result.shifted_rows, 2)
        self.assertEqual(result.summary_row, 7)
        rows = row_by_number(target)
        self.assertEqual(
            rows[2].find(tag("c")).find(f"{tag('is')}/{tag('t')}").text,
            "new-1",
        )
        self.assertEqual(rows[2].find(tag("c")).attrib["s"], "10")
        self.assertEqual(rows[7].findall(tag("c"))[0].attrib["r"], "E7")
        self.assertEqual(target.find(tag("autoFilter")).attrib["ref"], "A1:V5")
        self.assertEqual(target.find(tag("dimension")).attrib["ref"], "A1:V8")

    def test_fewer_source_rows_leave_summary_in_place(self) -> None:
        target = make_workbook_root(["old-1", "old-2", "old-3"], True, "10")
        source = make_workbook_root(["new-1"], False, "1")

        result = _replace_tp_payout_roots(target, source, {}, {})

        self.assertEqual(result.old_detail_end_row, 4)
        self.assertEqual(result.new_detail_end_row, 2)
        self.assertEqual(result.shifted_rows, 0)
        self.assertEqual(result.summary_row, 6)
        rows = row_by_number(target)
        self.assertNotIn(3, rows)
        self.assertNotIn(4, rows)
        self.assertEqual(rows[6].findall(tag("c"))[0].attrib["r"], "E6")
        self.assertEqual(target.find(tag("autoFilter")).attrib["ref"], "A1:V2")


class TpCollectionSyncTest(unittest.TestCase):
    def test_sources_are_classified_and_success_rows_are_written_first(self) -> None:
        target = make_collection_root(
            ["old-1", "old-2"],
            "支付成功",
            include_summary=True,
        )
        partial_source = make_collection_root(
            ["partial-1"],
            "部分支付",
            include_summary=False,
            source_header=True,
        )
        successful_source = make_collection_root(
            ["success-1", "success-2"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )

        result = _replace_tp_collection_roots(
            target,
            partial_source,
            successful_source,
            {},
            {},
            {},
        )

        rows = row_by_number(target)
        order_ids = [
            rows[row_number].find(tag("c")).find(f"{tag('is')}/{tag('t')}").text
            for row_number in range(2, 5)
        ]
        self.assertEqual(
            order_ids,
            ["success-1", "success-2", "partial-1"],
        )
        self.assertEqual(result.successful_rows, 2)
        self.assertEqual(result.partial_rows, 1)
        self.assertEqual(result.removed_rows, 2)
        self.assertEqual(result.inserted_rows, 3)
        self.assertEqual(result.summary_row, 6)
        self.assertEqual(result.shifted_rows, 1)
        self.assertEqual(target.find(tag("autoFilter")).attrib["ref"], "A1:Z4")

    def test_duplicate_platform_order_id_is_rejected(self) -> None:
        target = make_collection_root(
            ["old-1"],
            "支付成功",
            include_summary=True,
        )
        successful_source = make_collection_root(
            ["duplicate-id"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )
        partial_source = make_collection_root(
            ["duplicate-id"],
            "部分支付",
            include_summary=False,
            source_header=True,
        )

        with self.assertRaisesRegex(ValueError, "平台订单号重复"):
            _replace_tp_collection_roots(
                target,
                successful_source,
                partial_source,
                {},
                {},
                {},
            )

    def test_two_files_with_same_status_are_rejected(self) -> None:
        target = make_collection_root(
            ["old-1"],
            "支付成功",
            include_summary=True,
        )
        first_source = make_collection_root(
            ["success-1"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )
        second_source = make_collection_root(
            ["success-2"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )

        with self.assertRaisesRegex(ValueError, "两个收款订单文件都是"):
            _replace_tp_collection_roots(
                target,
                first_source,
                second_source,
                {},
                {},
                {},
            )


class WalletFlowSyncTest(unittest.TestCase):
    def test_all_source_rows_are_written_including_cross_day_record(self) -> None:
        target = make_wallet_root(
            ["old-1", "old-2"],
            include_summary=True,
            source_layout=False,
        )
        source = make_wallet_root(
            ["cross-day", "same-day-1", "same-day-2"],
            include_summary=False,
            source_layout=True,
        )

        result = _replace_wallet_flow_roots(target, source, {}, {})

        rows = row_by_number(target)
        transaction_ids = []
        for row_number in range(2, 5):
            cell = next(
                cell
                for cell in rows[row_number].findall(tag("c"))
                if cell.attrib["r"].startswith("F")
            )
            transaction_ids.append(
                cell.find(f"{tag('is')}/{tag('t')}").text
            )
        self.assertEqual(
            transaction_ids,
            ["cross-day", "same-day-1", "same-day-2"],
        )
        self.assertEqual(result.removed_rows, 2)
        self.assertEqual(result.inserted_rows, 3)
        self.assertEqual(result.summary_row, 6)
        self.assertEqual(result.shifted_rows, 1)
        self.assertEqual(target.find(tag("autoFilter")).attrib["ref"], "A1:N4")


class FullFlowSyncTest(unittest.TestCase):
    @staticmethod
    def save_workbook(path: Path, sheet_names: list[str]) -> None:
        workbook = Workbook()
        workbook.active.title = sheet_names[0]
        for sheet_name in sheet_names[1:]:
            workbook.create_sheet(sheet_name)
        workbook.save(path)

    def test_discovers_sources_by_prefix_and_workbook_by_sheets(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            workbook = directory / "2026.7.23孟加拉对账结果.xlsx"
            self.save_workbook(
                workbook,
                ["TP代付", "TP代收", "长款(当日)", "其他"],
            )
            payment = directory / "付款订单_001.xlsx"
            first_collection = directory / "收款订单_001.xlsx"
            second_collection = directory / "收款订单_002.xlsx"
            wallet = directory / "平台钱包流水记录_001.xlsx"
            for path in (
                payment,
                first_collection,
                second_collection,
                wallet,
            ):
                self.save_workbook(path, ["Sheet1"])
            (directory / "~$付款订单_临时.xlsx").touch()

            result = discover_flow_sync_files(directory)

            self.assertEqual(result.workbook, workbook)
            self.assertEqual(result.payment_orders, payment)
            self.assertEqual(
                result.collection_orders,
                (first_collection, second_collection),
            )
            self.assertEqual(result.wallet_flow, wallet)

    def test_rejects_incorrect_source_file_count(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            self.save_workbook(directory / "付款订单_001.xlsx", ["Sheet1"])

            with self.assertRaisesRegex(ValueError, "收款订单 Excel：应有 2 个"):
                discover_flow_sync_files(directory)

    def test_failure_does_not_change_original_workbook(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            workbook = directory / "workbook.xlsx"
            workbook.write_bytes(b"original")
            files = FlowSyncFiles(
                directory=directory,
                workbook=workbook,
                payment_orders=directory / "付款订单.xlsx",
                collection_orders=(
                    directory / "收款订单1.xlsx",
                    directory / "收款订单2.xlsx",
                ),
                wallet_flow=directory / "平台钱包流水.xlsx",
            )
            payout_result = TpPayoutSyncResult(1, 2, 2, 3, None, 1)

            def change_temporary(path: Path, *_args, **_kwargs):
                path.write_bytes(b"partial")
                return payout_result

            with (
                patch(
                    "autoexcel.flow_sync.discover_flow_sync_files",
                    return_value=files,
                ),
                patch(
                    "autoexcel.flow_sync.sync_tp_payout",
                    side_effect=change_temporary,
                ),
                patch(
                    "autoexcel.flow_sync.sync_tp_collection",
                    side_effect=ValueError("collection failed"),
                ),
            ):
                with self.assertRaisesRegex(ValueError, "collection failed"):
                    sync_all_flows(directory)

            self.assertEqual(workbook.read_bytes(), b"original")
            self.assertFalse(
                list(directory.glob("*.flow-sync.*.tmp.xlsx"))
            )

    def test_success_replaces_original_after_all_three_steps(self) -> None:
        with TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            workbook = directory / "workbook.xlsx"
            workbook.write_bytes(b"original")
            files = FlowSyncFiles(
                directory=directory,
                workbook=workbook,
                payment_orders=directory / "付款订单.xlsx",
                collection_orders=(
                    directory / "收款订单1.xlsx",
                    directory / "收款订单2.xlsx",
                ),
                wallet_flow=directory / "平台钱包流水.xlsx",
            )
            payout_result = TpPayoutSyncResult(1, 2, 2, 3, None, 1)
            collection_result = TpCollectionSyncResult(
                1,
                2,
                2,
                3,
                None,
                1,
                1,
                1,
            )

            def append_marker(marker: bytes, result):
                def update(path: Path, *_args, **_kwargs):
                    path.write_bytes(path.read_bytes() + marker)
                    return result

                return update

            with (
                patch(
                    "autoexcel.flow_sync.discover_flow_sync_files",
                    return_value=files,
                ),
                patch(
                    "autoexcel.flow_sync.sync_tp_payout",
                    side_effect=append_marker(b"-payout", payout_result),
                ),
                patch(
                    "autoexcel.flow_sync.sync_tp_collection",
                    side_effect=append_marker(
                        b"-collection",
                        collection_result,
                    ),
                ),
                patch(
                    "autoexcel.flow_sync.sync_wallet_flow",
                    side_effect=append_marker(b"-wallet", payout_result),
                ),
            ):
                result = sync_all_flows(directory)

            self.assertEqual(
                workbook.read_bytes(),
                b"original-payout-collection-wallet",
            )
            self.assertEqual(result.files, files)
            self.assertFalse(
                list(directory.glob("*.flow-sync.*.tmp.xlsx"))
            )

if __name__ == "__main__":
    unittest.main()
