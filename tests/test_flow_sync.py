from __future__ import annotations

import unittest
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook

from autoexcel.flow_sync import (
    FlowSyncFiles,
    TP_COLLECTION_HEADERS,
    TP_PAYOUT_HEADERS,
    TP_WITHDRAWAL_SOURCE_HEADERS,
    TP_WITHDRAWAL_TARGET_HEADERS,
    TpCollectionSyncResult,
    TpPayoutSyncResult,
    TpWithdrawalSyncResult,
    WALLET_FLOW_HEADERS,
    _cell_text,
    _merge_source_sheet_roots,
    _replace_tp_collection_roots,
    _replace_tp_payout_roots,
    _replace_tp_withdrawal_roots,
    _replace_wallet_flow_roots,
    _validate_collection_header,
    _validate_header,
    _validate_wallet_flow_header,
    _withdrawal_pivot_replacements,
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


def make_withdrawal_target_root(
    detail_ids: list[str],
    include_summary: bool,
) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    max_row = len(detail_ids) + (4 if include_summary else 1)
    ET.SubElement(root, tag("dimension"), {"ref": f"A1:Y{max_row}"})
    sheet_data = ET.SubElement(root, tag("sheetData"))
    sheet_data.append(
        make_row(
            1,
            {
                chr(ord("A") + index): value
                for index, value in enumerate(TP_WITHDRAWAL_TARGET_HEADERS)
            },
            "20",
        )
    )
    for offset, order_id in enumerate(detail_ids, start=2):
        sheet_data.append(
            make_row(
                offset,
                {
                    "A": order_id,
                    "D": "100",
                    "K": "old-wallet",
                    "X": "已打款",
                },
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
                    "D": "平台钱包",
                    "E": "求和项:提现金额",
                    "F": "求和项:手续费",
                    "G": "计数项:平台订单号",
                    "H": "求和项:到账金额",
                },
                "30",
            )
        )
        sheet_data.append(
            make_row(detail_end + 3, {"D": "总计"}, "30")
        )
    ET.SubElement(
        root,
        tag("autoFilter"),
        {"ref": f"A1:Y{len(detail_ids) + 1}"},
    )
    return root


def make_withdrawal_source_root(
    detail_rows: list[tuple[str, str]],
) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    ET.SubElement(
        root,
        tag("dimension"),
        {"ref": f"A1:Z{len(detail_rows) + 2}"},
    )
    sheet_data = ET.SubElement(root, tag("sheetData"))
    sheet_data.append(make_row(1, {"A": "商户提现申请"}, "40"))
    sheet_data.append(
        make_row(
            2,
            {
                chr(ord("A") + index): value
                for index, value in enumerate(TP_WITHDRAWAL_SOURCE_HEADERS)
            },
            "20",
        )
    )
    for row_number, (order_id, status) in enumerate(detail_rows, start=3):
        sheet_data.append(
            make_row(
                row_number,
                {
                    "A": order_id,
                    "D": "1,234.50",
                    "E": "0",
                    "F": "9.25",
                    "G": "133.45",
                    "H": "0",
                    "K": "identity-should-be-dropped",
                    "L": "wallet-001",
                    "M": "bank-001",
                    "Q": "01300000000",
                    "Y": status,
                    "Z": "reason",
                },
                "10",
            )
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


def make_headerless_root(rows: list[dict[str, str]], max_column: str) -> ET.Element:
    root = ET.Element(tag("worksheet"))
    ET.SubElement(
        root,
        tag("dimension"),
        {"ref": f"A1:{max_column}{len(rows)}"},
    )
    sheet_data = ET.SubElement(root, tag("sheetData"))
    for row_number, values in enumerate(rows, start=1):
        sheet_data.append(make_row(row_number, values, "10"))
    return root


class TpPayoutSyncTest(unittest.TestCase):
    def test_amount_text_is_written_as_number_while_order_id_stays_text(self) -> None:
        target = make_workbook_root(["old-1"], True, "10")
        source = make_workbook_root(["001234"], False, "1")
        amount_source_cell = ET.SubElement(
            row_by_number(source)[2],
            tag("c"),
            {"r": "F2", "t": "s", "s": "1"},
        )
        ET.SubElement(amount_source_cell, tag("v")).text = "42"

        _replace_tp_payout_roots(target, source, {}, {42: "1,234.50"})

        target_row = row_by_number(target)[2]
        order_id_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "A2"
        )
        amount_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "F2"
        )
        self.assertEqual(order_id_cell.attrib["t"], "inlineStr")
        self.assertEqual(
            order_id_cell.find(f"{tag('is')}/{tag('t')}").text,
            "001234",
        )
        self.assertNotIn("t", amount_cell.attrib)
        self.assertEqual(amount_cell.find(tag("v")).text, "1234.50")

    def test_duplicate_key_across_source_sheets_is_rejected(self) -> None:
        first_source = make_workbook_root(["duplicate-id"], False, "1")
        continuation = make_headerless_root(
            [{"A": "duplicate-id"}],
            "V",
        )

        with self.assertRaisesRegex(ValueError, "关键字段重复"):
            _merge_source_sheet_roots(
                [first_source, continuation],
                ["付款订单", "Sheet1"],
                {},
                "付款订单",
                1,
                "A",
                _validate_header,
            )

    def test_repeated_header_on_continuation_sheet_is_skipped(self) -> None:
        target = make_workbook_root(["old-1"], True, "10")
        first_source = make_workbook_root(["new-1"], False, "1")
        continuation = make_workbook_root(["new-2"], False, "1")
        merged_source = _merge_source_sheet_roots(
            [first_source, continuation],
            ["付款订单", "Sheet1"],
            {},
            "付款订单",
            1,
            "A",
            _validate_header,
        )

        result = _replace_tp_payout_roots(
            target,
            merged_source,
            {},
            {},
        )

        self.assertEqual(result.inserted_rows, 2)
        self.assertEqual(
            [
                _cell_text(row_by_number(target)[row_number].find(tag("c")), {})
                for row_number in range(2, 4)
            ],
            ["new-1", "new-2"],
        )

    def test_headerless_continuation_sheet_is_appended(self) -> None:
        target = make_workbook_root(["old-1"], True, "10")
        first_source = make_workbook_root(["new-1", "new-2"], False, "1")
        continuation = make_headerless_root(
            [
                {"A": "new-3", "B": "merchant-new-3"},
                {"A": "new-4", "B": "merchant-new-4"},
            ],
            "V",
        )
        merged_source = _merge_source_sheet_roots(
            [first_source, continuation],
            ["付款订单", "Sheet1"],
            {},
            "付款订单",
            1,
            "A",
            _validate_header,
        )

        result = _replace_tp_payout_roots(
            target,
            merged_source,
            {},
            {},
        )

        self.assertEqual(result.inserted_rows, 4)
        self.assertEqual(
            [
                _cell_text(row_by_number(target)[row_number].find(tag("c")), {})
                for row_number in range(2, 6)
            ],
            ["new-1", "new-2", "new-3", "new-4"],
        )

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
    def test_amount_text_is_written_as_number_while_order_id_stays_text(self) -> None:
        target = make_collection_root(
            ["old-1"],
            "支付成功",
            include_summary=True,
        )
        successful_source = make_collection_root(
            ["000123"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )
        add_inline_cell(
            row_by_number(successful_source)[2],
            "G2",
            "88.60",
            "10",
        )
        partial_source = make_collection_root(
            ["partial-1"],
            "部分支付",
            include_summary=False,
            source_header=True,
        )

        _replace_tp_collection_roots(
            target,
            successful_source,
            partial_source,
            {},
            {},
            {},
        )

        target_row = row_by_number(target)[2]
        order_id_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "A2"
        )
        amount_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "G2"
        )
        self.assertEqual(order_id_cell.attrib["t"], "inlineStr")
        self.assertEqual(
            order_id_cell.find(f"{tag('is')}/{tag('t')}").text,
            "000123",
        )
        self.assertNotIn("t", amount_cell.attrib)
        self.assertEqual(amount_cell.find(tag("v")).text, "88.60")

    def test_headerless_continuation_sheet_is_appended(self) -> None:
        target = make_collection_root(
            ["old-1"],
            "支付成功",
            include_summary=True,
        )
        successful_first = make_collection_root(
            ["success-1"],
            "支付成功",
            include_summary=False,
            source_header=True,
        )
        successful_continuation = make_headerless_root(
            [
                {"A": "success-2", "W": "支付成功"},
                {"A": "success-3", "W": "支付成功"},
            ],
            "Z",
        )
        successful_source = _merge_source_sheet_roots(
            [successful_first, successful_continuation],
            ["收款订单", "Sheet1"],
            {},
            "收款订单文件 1",
            1,
            "A",
            _validate_collection_header,
        )
        partial_source = make_collection_root(
            ["partial-1"],
            "部分支付",
            include_summary=False,
            source_header=True,
        )

        result = _replace_tp_collection_roots(
            target,
            successful_source,
            partial_source,
            {},
            {},
            {},
        )

        self.assertEqual(result.successful_rows, 3)
        self.assertEqual(result.partial_rows, 1)
        self.assertEqual(result.inserted_rows, 4)

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


class TpWithdrawalSyncTest(unittest.TestCase):
    def test_only_paid_rows_are_mapped_and_amounts_become_numbers(self) -> None:
        target = make_withdrawal_target_root(
            ["old-1", "old-2"],
            include_summary=True,
        )
        source = make_withdrawal_source_root(
            [
                ("000123", "已打款"),
                ("pending-1", "处理中"),
            ]
        )

        result = _replace_tp_withdrawal_roots(
            target,
            source,
            {},
            {},
        )

        self.assertEqual(result.removed_rows, 2)
        self.assertEqual(result.inserted_rows, 1)
        self.assertEqual(result.skipped_rows, 1)
        rows = row_by_number(target)
        detail = rows[2]
        values = {
            cell.attrib["r"]: _cell_text(cell, {})
            for cell in detail.findall(tag("c"))
        }
        self.assertEqual(values["A2"], "000123")
        self.assertEqual(values["K2"], "wallet-001")
        self.assertEqual(values["L2"], "bank-001")
        self.assertEqual(values["P2"], "01300000000")
        self.assertEqual(values["X2"], "已打款")
        self.assertEqual(values["Y2"], "reason")
        self.assertNotIn("Z2", values)
        amount_cell = next(
            cell
            for cell in detail.findall(tag("c"))
            if cell.attrib["r"] == "D2"
        )
        self.assertNotIn("t", amount_cell.attrib)
        self.assertEqual(amount_cell.find(tag("v")).text, "1234.50")

    def test_pivot_source_and_location_are_updated(self) -> None:
        result = TpWithdrawalSyncResult(
            removed_rows=2,
            inserted_rows=4,
            old_detail_end_row=3,
            new_detail_end_row=5,
            summary_row=7,
            shifted_rows=2,
            skipped_rows=0,
        )
        worksheet_relationships = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/'
            'package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/pivotTable" '
            'Target="../pivotTables/pivotTable1.xml"/>'
            "</Relationships>"
        )
        pivot = (
            f'<pivotTableDefinition xmlns="{MAIN_NS}">'
            '<location ref="D5:H7"/>'
            "</pivotTableDefinition>"
        )
        pivot_relationships = (
            '<Relationships xmlns="http://schemas.openxmlformats.org/'
            'package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/pivotCacheDefinition" '
            'Target="../pivotCache/pivotCacheDefinition1.xml"/>'
            "</Relationships>"
        )
        cache = (
            f'<pivotCacheDefinition xmlns="{MAIN_NS}">'
            '<cacheSource type="worksheet">'
            '<worksheetSource ref="A1:Y3" sheet="TP提现"/>'
            "</cacheSource>"
            "</pivotCacheDefinition>"
        )

        with TemporaryDirectory() as temporary_directory:
            archive_path = Path(temporary_directory) / "pivot.xlsx"
            with zipfile.ZipFile(archive_path, "w") as archive:
                archive.writestr(
                    "xl/worksheets/_rels/sheet1.xml.rels",
                    worksheet_relationships,
                )
                archive.writestr("xl/pivotTables/pivotTable1.xml", pivot)
                archive.writestr(
                    "xl/pivotTables/_rels/pivotTable1.xml.rels",
                    pivot_relationships,
                )
                archive.writestr(
                    "xl/pivotCache/pivotCacheDefinition1.xml",
                    cache,
                )

            with zipfile.ZipFile(archive_path) as archive:
                replacements = _withdrawal_pivot_replacements(
                    archive,
                    "xl/worksheets/sheet1.xml",
                    result,
                )

        pivot_root = ET.fromstring(
            replacements["xl/pivotTables/pivotTable1.xml"]
        )
        self.assertEqual(
            pivot_root.find(tag("location")).attrib["ref"],
            "D7:H9",
        )
        cache_root = ET.fromstring(
            replacements[
                "xl/pivotCache/pivotCacheDefinition1.xml"
            ]
        )
        worksheet_source = cache_root.find(
            f"{tag('cacheSource')}/{tag('worksheetSource')}"
        )
        self.assertEqual(worksheet_source.attrib["ref"], "A1:Y5")
        self.assertEqual(cache_root.attrib["refreshOnLoad"], "1")
        self.assertEqual(cache_root.attrib["enableRefresh"], "1")


class WalletFlowSyncTest(unittest.TestCase):
    def test_amount_text_is_written_as_number_while_transaction_id_stays_text(
        self,
    ) -> None:
        target = make_wallet_root(
            ["old-1"],
            include_summary=True,
            source_layout=False,
        )
        source = make_wallet_root(
            ["000999"],
            include_summary=False,
            source_layout=True,
        )

        _replace_wallet_flow_roots(target, source, {}, {})

        target_row = row_by_number(target)[2]
        transaction_id_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "F2"
        )
        amount_cell = next(
            cell for cell in target_row.findall(tag("c"))
            if cell.attrib["r"] == "M2"
        )
        self.assertEqual(transaction_id_cell.attrib["t"], "inlineStr")
        self.assertEqual(
            transaction_id_cell.find(f"{tag('is')}/{tag('t')}").text,
            "000999",
        )
        self.assertNotIn("t", amount_cell.attrib)
        self.assertEqual(amount_cell.find(tag("v")).text, "100")

    def test_headerless_continuation_sheet_is_appended(self) -> None:
        target = make_wallet_root(
            ["old-1"],
            include_summary=True,
            source_layout=False,
        )
        first_source = make_wallet_root(
            ["wallet-1"],
            include_summary=False,
            source_layout=True,
        )
        continuation = make_headerless_root(
            [
                {"F": "wallet-2", "K": "长款"},
                {"F": "wallet-3", "K": "长款"},
            ],
            "O",
        )
        merged_source = _merge_source_sheet_roots(
            [first_source, continuation],
            ["平台钱包流水记录", "Sheet1"],
            {},
            "平台钱包流水记录",
            2,
            "F",
            _validate_wallet_flow_header,
        )

        result = _replace_wallet_flow_roots(
            target,
            merged_source,
            {},
            {},
        )

        self.assertEqual(result.inserted_rows, 3)

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
                ["TP代付", "TP代收", "TP提现", "长款(当日)", "其他"],
            )
            payment = directory / "付款订单_001.xlsx"
            first_collection = directory / "收款订单_001.xlsx"
            second_collection = directory / "收款订单_002.xlsx"
            withdrawal = directory / "商户提现申请_001.xlsx"
            wallet = directory / "平台钱包流水记录_001.xlsx"
            for path in (
                payment,
                first_collection,
                second_collection,
                withdrawal,
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
            self.assertEqual(result.withdrawal_orders, withdrawal)
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
                withdrawal_orders=directory / "商户提现申请.xlsx",
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

    def test_success_replaces_original_after_all_four_steps(self) -> None:
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
                withdrawal_orders=directory / "商户提现申请.xlsx",
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
            withdrawal_result = TpWithdrawalSyncResult(
                1,
                2,
                2,
                3,
                None,
                1,
                0,
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
                    "autoexcel.flow_sync.sync_tp_withdrawal",
                    side_effect=append_marker(
                        b"-withdrawal",
                        withdrawal_result,
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
                b"original-payout-collection-withdrawal-wallet",
            )
            self.assertEqual(result.files, files)
            self.assertFalse(
                list(directory.glob("*.flow-sync.*.tmp.xlsx"))
            )

if __name__ == "__main__":
    unittest.main()
