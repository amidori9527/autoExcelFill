from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from html import escape
import json
from pathlib import Path
import subprocess
import sys
import traceback
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DIR = PROJECT_ROOT / "workspace" / "diffOrders"
DEFAULT_A = DEFAULT_DIR / "上游.xlsx"
DEFAULT_B = DEFAULT_DIR / "后台.xlsx"
DEFAULT_TEMPLATE = PROJECT_ROOT / "template" / "order_diff.html"
DEFAULT_RESULT_DIR = PROJECT_ROOT / "result"
WORKSPACE_DIR_NAME = "workspace"
DIFF_ORDERS_DIR_NAME = "diffOrders"
TEMPLATE_FILE_NAME = "order_diff.html"
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


def is_frozen_app() -> bool:
    return getattr(sys, "frozen", False)


def get_executable_directory() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


def format_file_size(path: Path) -> str:
    size = path.stat().st_size
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{size} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def list_xlsx_files(directory: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in directory.glob("*.xlsx")
            if path.is_file() and not path.name.startswith(("~$", ".~", "._"))
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def iter_diff_orders_directories() -> list[Path]:
    directories: list[Path] = []
    if is_frozen_app():
        base_directories = (get_executable_directory(), Path.cwd(), PROJECT_ROOT)
    else:
        base_directories = (Path.cwd(), PROJECT_ROOT)
    for base_directory in base_directories:
        diff_dir = base_directory / WORKSPACE_DIR_NAME / DIFF_ORDERS_DIR_NAME
        if diff_dir not in directories:
            directories.append(diff_dir)
    return directories


def find_xlsx_directory() -> tuple[Path, list[Path]]:
    checked: list[Path] = []
    for directory in iter_diff_orders_directories():
        checked.append(directory)
        if not directory.is_dir():
            continue
        files = list_xlsx_files(directory)
        if files:
            return directory, files
    checked_text = "\n  ".join(str(path) for path in checked)
    raise FileNotFoundError(
        f"没有在 workspace/diffOrders 文件夹里找到 xlsx 文件。"
        f"请把 Excel 放入 workspace/diffOrders 后再运行。\n"
        f"已检查目录：\n  {checked_text}"
    )


def get_template_path() -> Path:
    if is_frozen_app():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidate = Path(meipass) / "template" / TEMPLATE_FILE_NAME
            if candidate.exists():
                return candidate
        candidate = get_executable_directory() / "_internal" / "template" / TEMPLATE_FILE_NAME
        if candidate.exists():
            return candidate
    return DEFAULT_TEMPLATE


def get_result_dir() -> Path:
    if is_frozen_app():
        return get_executable_directory() / "result"
    return DEFAULT_RESULT_DIR


def choose_one_workbook(files: list[Path], prompt_text: str) -> Path:
    default_index = 1
    while True:
        raw_value = input(f"{prompt_text}，直接回车默认 {default_index}：").strip()
        if not raw_value:
            return files[default_index - 1]

        unquoted = raw_value.strip("'\"")
        if unquoted.isdigit():
            selected_index = int(unquoted)
            if 1 <= selected_index <= len(files):
                return files[selected_index - 1]
            print(f"编号超出范围，请输入 1 到 {len(files)}。")
            continue

        selected_path = Path(unquoted).expanduser()
        if not selected_path.is_absolute():
            selected_path = files[0].parent / selected_path
        if selected_path.is_file() and selected_path.suffix.lower() == ".xlsx":
            return selected_path
        print("没有找到这个 xlsx 文件，请重新输入。")


def choose_diff_workbooks_interactive() -> tuple[Path, Path]:
    if not sys.stdin.isatty():
        raise RuntimeError("交互模式需要可输入的终端；或请使用 --a/--b 参数直接运行。")

    directory, files = find_xlsx_directory()

    print("订单差异比对工具")
    print("----------------------------------------")
    print(f"Excel 搜索目录：{directory}")
    print("可选 Excel 文件：")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name} ({format_file_size(path)})")
    print()

    a_path = choose_one_workbook(files, "请输入上游订单 Excel 的编号或文件名")
    print(f"  已选上游：{a_path.name}")
    b_path = choose_one_workbook(files, "请输入后台订单 Excel 的编号或文件名")
    print(f"  已选后台：{b_path.name}")
    print()
    return a_path, b_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare upstream/backend order IDs.")
    parser.add_argument("--a", type=Path, default=None, help="A (上游) workbook path.")
    parser.add_argument("--b", type=Path, default=None, help="B (后台) workbook path.")
    parser.add_argument("--a-col", default="L", help="A order ID column. Default: L")
    parser.add_argument("--b-col", default="D", help="B order ID column. Default: D")
    parser.add_argument("--template", type=Path, default=None, help="HTML template path.")
    parser.add_argument("--result-dir", type=Path, default=None, help="HTML result directory.")
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="交互模式：选择 workspace/diffOrders 里的 Excel 文件。",
    )
    args = parser.parse_args(argv)
    args.interactive = args.interactive or len(argv if argv is not None else sys.argv[1:]) == 0
    return args


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


def unique_entries(entries: list[OrderEntry]) -> list[OrderEntry]:
    seen: set[str] = set()
    unique: list[OrderEntry] = []
    for entry in entries:
        if entry.order_id in seen:
            continue
        seen.add(entry.order_id)
        unique.append(entry)
    return unique


def render_upstream_order_table(entries: list[OrderEntry]) -> str:
    if not entries:
        return '<div class="empty">无</div>'

    columns = 4
    rows = [
        "<table>",
        "<thead><tr>" + "".join("<th>上游独有订单ID</th>" for _ in range(columns)) + "</tr></thead>",
        "<tbody>",
    ]
    for start in range(0, len(entries), columns):
        cells = []
        for entry in entries[start : start + columns]:
            order_id = escape(entry.order_id, quote=True)
            cells.append(
                "<td>"
                '<label class="order-check">'
                f'<input type="checkbox" data-check-order-id="{order_id}">'
                f'<button class="copy-order" data-order-id="{order_id}" title="点击复制订单号">{order_id}</button>'
                "</label>"
                "</td>"
            )
        cells.extend("<td></td>" for _ in range(columns - len(cells)))
        rows.append("<tr>" + "".join(cells) + "</tr>")
    rows.extend(["</tbody>", "</table>"])
    return "\n".join(rows)


def render_html(template_path: Path, result: DiffResult) -> str:
    template = template_path.read_text(encoding="utf-8")
    upstream_only = unique_entries(result.a_only)
    values = {
        "title": "上游独有订单ID",
        "upstream_only_count": str(len(upstream_only)),
        "upstream_only_table": render_upstream_order_table(upstream_only),
        "upstream_only_json": json.dumps([entry.order_id for entry in upstream_only], ensure_ascii=False),
    }
    html = template
    for key, value in values.items():
        html = html.replace("{{ " + key + " }}", value)
    return html


def write_html_result(result_dir: Path, template_path: Path, result: DiffResult) -> Path:
    result_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = result_dir / f"order_diff_{timestamp}.html"
    output_path.write_text(render_html(template_path, result), encoding="utf-8")
    return output_path


def copy_to_clipboard(text: str) -> bool:
    try:
        process = subprocess.Popen(["pbcopy"], stdin=subprocess.PIPE)
        process.communicate(input=text.encode("utf-8"))
        return process.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def should_pause_before_exit() -> bool:
    return is_frozen_app() and sys.stdin.isatty()


def pause_before_exit() -> None:
    if not should_pause_before_exit():
        return
    try:
        input("\n按回车退出...")
    except EOFError:
        pass


def write_error_log(error: BaseException) -> Path:
    log_path = get_executable_directory() / "diff-orders-error.log"
    log_text = "".join(traceback.format_exception(type(error), error, error.__traceback__))
    log_path.write_text(log_text, encoding="utf-8")
    return log_path


def main() -> None:
    args = parse_args()
    template_path = args.template or get_template_path()
    result_dir = args.result_dir or get_result_dir()

    if args.interactive:
        a_path, b_path = choose_diff_workbooks_interactive()
    else:
        a_path = args.a or DEFAULT_A
        b_path = args.b or DEFAULT_B

    print("正在读取并比对订单，请稍候...")
    a_entries = read_order_ids(a_path, args.a_col)
    b_entries = read_order_ids(b_path, args.b_col)
    result = diff_orders(a_entries, b_entries)
    html_path = write_html_result(result_dir, template_path, result)

    upstream_only = unique_entries(result.a_only)
    order_ids = [entry.order_id for entry in upstream_only]

    print()
    print("订单差异比对完成")
    print("----------------------------------------")
    print(f"上游 Excel: {a_path} ({result.a_count} unique)")
    print(f"后台 Excel: {b_path} ({result.b_count} unique)")
    print(f"上游独有订单数: {len(order_ids)}")
    print(f"HTML结果文件: {html_path}")

    if not order_ids:
        print("没有可复制的订单号")
        return

    if copy_to_clipboard("\n".join(order_ids)):
        print(f"已复制 {len(order_ids)} 个上游独有订单号到剪贴板，可直接粘贴使用")
    else:
        print("复制到剪贴板失败，请手动从 HTML 结果文件中复制")


if __name__ == "__main__":
    exit_code = 0
    try:
        main()
    except KeyboardInterrupt:
        print("\n已取消。")
        exit_code = 130
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        print(f"错误：{error}")
        exit_code = 1
    except Exception as error:
        try:
            log_path = write_error_log(error)
            print(f"程序执行失败：{error}")
            print(f"详细错误已写入：{log_path}")
        except Exception:
            print("程序执行失败，并且写入错误日志时也失败：")
            traceback.print_exc()
        exit_code = 1
    finally:
        pause_before_exit()
    sys.exit(exit_code)
