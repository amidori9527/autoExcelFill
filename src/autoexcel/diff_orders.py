from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DIR = PROJECT_ROOT / "workspace" / "diffOrders"
DEFAULT_A = DEFAULT_DIR / "上游.xlsx"
DEFAULT_B = DEFAULT_DIR / "后台.xlsx"
DEFAULT_OUTPUT = DEFAULT_DIR / "diff_result.xlsx"
HEADER_KEYWORDS = {
    "TRANSACTION REFERENCE NUMBER",
    "请求上游订单号",
}


@dataclass(frozen=True)
class OrderEntry:
    order_id: str
    row_number: int


@dataclass(frozen=True)
class DiffResult:
    a_only: list[OrderEntry]
    b_only: list[OrderEntry]
    a_count: int
    b_count: int
    common_count: int
    a_duplicate_count: int
    b_duplicate_count: int


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare upstream/backend order IDs.")
    parser.add_argument("--a", type=Path, default=DEFAULT_A, help="A workbook path. Default: 上游.xlsx")
    parser.add_argument("--b", type=Path, default=DEFAULT_B, help="B workbook path. Default: 后台.xlsx")
    parser.add_argument("--a-col", default="L", help="A order ID column. Default: L")
    parser.add_argument("--b-col", default="D", help="B order ID column. Default: D")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Output workbook path.")
    return parser.parse_args()


def normalize_order_id(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = str(value)
    text = text.strip()
    if not text or text in HEADER_KEYWORDS:
        return None
    return text


def read_order_ids(path: Path, column: str) -> list[OrderEntry]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    column_index = column_index_from_string(column)
    entries: list[OrderEntry] = []
    for row_number, row in enumerate(
        worksheet.iter_rows(min_col=column_index, max_col=column_index, values_only=True),
        start=1,
    ):
        order_id = normalize_order_id(row[0])
        if order_id is None:
            continue
        entries.append(OrderEntry(order_id=order_id, row_number=row_number))
    return entries


def diff_orders(a_entries: list[OrderEntry], b_entries: list[OrderEntry]) -> DiffResult:
    a_ids = {entry.order_id for entry in a_entries}
    b_ids = {entry.order_id for entry in b_entries}
    a_counter = Counter(entry.order_id for entry in a_entries)
    b_counter = Counter(entry.order_id for entry in b_entries)
    a_only_ids = a_ids - b_ids
    b_only_ids = b_ids - a_ids
    common_ids = a_ids & b_ids

    return DiffResult(
        a_only=[entry for entry in a_entries if entry.order_id in a_only_ids],
        b_only=[entry for entry in b_entries if entry.order_id in b_only_ids],
        a_count=len(a_ids),
        b_count=len(b_ids),
        common_count=len(common_ids),
        a_duplicate_count=sum(count - 1 for count in a_counter.values() if count > 1),
        b_duplicate_count=sum(count - 1 for count in b_counter.values() if count > 1),
    )


def write_result(output_path: Path, result: DiffResult, a_path: Path, b_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()

    summary = workbook.active
    summary.title = "汇总"
    summary.append(["项目", "数量"])
    summary.append(["A文件", str(a_path)])
    summary.append(["B文件", str(b_path)])
    summary.append(["A唯一订单数", result.a_count])
    summary.append(["B唯一订单数", result.b_count])
    summary.append(["两边都有的订单数", result.common_count])
    summary.append(["A独有订单数", len({entry.order_id for entry in result.a_only})])
    summary.append(["B独有订单数", len({entry.order_id for entry in result.b_only})])
    summary.append(["A内部重复额外行数", result.a_duplicate_count])
    summary.append(["B内部重复额外行数", result.b_duplicate_count])
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 80

    a_sheet = workbook.create_sheet("A独有")
    a_sheet.append(["订单ID", "A行号"])
    for entry in result.a_only:
        a_sheet.append([entry.order_id, entry.row_number])
    a_sheet.column_dimensions["A"].width = 28
    a_sheet.column_dimensions["B"].width = 12

    b_sheet = workbook.create_sheet("B独有")
    b_sheet.append(["订单ID", "B行号"])
    for entry in result.b_only:
        b_sheet.append([entry.order_id, entry.row_number])
    b_sheet.column_dimensions["A"].width = 28
    b_sheet.column_dimensions["B"].width = 12

    workbook.save(output_path)


def main() -> None:
    args = parse_args()
    a_entries = read_order_ids(args.a, args.a_col)
    b_entries = read_order_ids(args.b, args.b_col)
    result = diff_orders(a_entries, b_entries)
    write_result(args.output, result, args.a, args.b)

    print("订单差异比对完成")
    print("----------------------------------------")
    print(f"A: {args.a} ({result.a_count} unique)")
    print(f"B: {args.b} ({result.b_count} unique)")
    print(f"两边都有: {result.common_count}")
    print(f"A独有: {len({entry.order_id for entry in result.a_only})}")
    print(f"B独有: {len({entry.order_id for entry in result.b_only})}")
    print(f"结果文件: {args.output}")


if __name__ == "__main__":
    main()
