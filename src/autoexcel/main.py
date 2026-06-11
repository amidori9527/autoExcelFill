from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from autoexcel.fast_xlsx import add_current_date_to_colored_sheets_fast
from autoexcel.operations import freeze_colored_sheets_next_day, freeze_next_day_row
from autoexcel.workbook_io import list_sheet_names, preview_sheet


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_FILE = PROJECT_ROOT / "data.xlsx"
EXAMPLE_SHEET = "example"
WORKSPACE_DIR_NAME = "workspace"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Inspect and automate data.xlsx.")
    parser.add_argument("--workbook", type=Path, default=None)
    parser.add_argument("--sheet", default=EXAMPLE_SHEET)
    parser.add_argument("--rows", type=int, default=10)
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Choose an xlsx file from the current directory and use fill defaults.",
    )
    parser.add_argument("--list-sheets", action="store_true")
    parser.add_argument(
        "--copy-sheet-from",
        help="Create --sheet by copying this source sheet when --sheet does not exist.",
    )
    parser.add_argument(
        "--freeze-next-day-row",
        action="store_true",
        help="Insert a value row before the last date row, then advance that last row to current date.",
    )
    parser.add_argument(
        "--colored-sheets",
        action="store_true",
        help="Apply --freeze-next-day-row to every sheet with a tab color.",
    )
    parser.add_argument("--previous-date", default="2026-03-26", help="Deprecated; kept for old commands.")
    parser.add_argument(
        "--source-date",
        default="2026-03-27",
        help="Deprecated; kept for old commands.",
    )
    parser.add_argument(
        "--current-date",
        help="Current date to add. If omitted, defaults to today's date.",
    )
    parser.add_argument(
        "--limit-sheets",
        type=int,
        help="Stop after this many sheets are successfully changed.",
    )
    parser.add_argument(
        "--fast-xml",
        action="store_true",
        help="Use direct xlsx XML editing for large files. Supports colored sheet batches.",
    )
    parser.add_argument(
        "--run-until-done",
        action="store_true",
        help="Repeat fast XML batches until no more sheets are changed.",
    )
    args = parser.parse_args(argv)
    args.interactive = args.interactive or len(argv if argv is not None else sys.argv[1:]) == 0
    if args.workbook is None:
        args.workbook = DATA_FILE
    return args


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def resolve_current_date(args: argparse.Namespace) -> date:
    if args.current_date:
        return parse_date(args.current_date)
    return date.today()


def format_file_size(path: Path) -> str:
    size = path.stat().st_size
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{size} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def list_current_xlsx_files(directory: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in directory.glob("*.xlsx")
            if path.is_file() and not path.name.startswith(("~$", ".~", "._"))
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def get_executable_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


def iter_workspace_directories() -> list[Path]:
    directories: list[Path] = []
    if getattr(sys, "frozen", False):
        base_directories = (get_executable_directory(), Path.cwd(), PROJECT_ROOT)
    else:
        base_directories = (Path.cwd(), PROJECT_ROOT)

    for base_directory in base_directories:
        workspace_directory = base_directory / WORKSPACE_DIR_NAME
        if workspace_directory not in directories:
            directories.append(workspace_directory)
    return directories


def find_workbook_directory() -> tuple[Path, list[Path]]:
    checked: list[Path] = []
    for directory in iter_workspace_directories():
        checked.append(directory)
        if not directory.is_dir():
            continue
        files = list_current_xlsx_files(directory)
        if files:
            return directory, files
    checked_text = "\n  ".join(str(path) for path in checked)
    raise FileNotFoundError(
        f"没有在 workspace 文件夹里找到 xlsx 文件。请把 Excel 放入 workspace 后再运行。\n"
        f"已检查目录：\n  {checked_text}"
    )


def choose_workbook_from_current_directory() -> Path:
    if not sys.stdin.isatty():
        raise RuntimeError("交互模式需要可输入的终端；或请使用 --workbook 等参数直接运行。")

    workbook_directory, files = find_workbook_directory()

    print("Excel 数据填报工具")
    print("----------------------------------------")
    print(f"Excel 搜索目录：{workbook_directory}")
    print("请选择要处理的 workbook：")
    for index, path in enumerate(files, start=1):
        print(f"  {index}. {path.name} ({format_file_size(path)})")

    default_index = 1
    while True:
        raw_value = input(f"输入编号或文件路径，直接回车默认 {default_index}：").strip()
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
            selected_path = workbook_directory / selected_path
        if selected_path.is_file() and selected_path.suffix.lower() == ".xlsx":
            return selected_path
        print("没有找到这个 xlsx 文件，请重新输入。")


def choose_current_date() -> str:
    default_date = date.today().strftime("%Y-%m-%d")
    while True:
        raw_value = input(f"目标日期 YYYY-MM-DD，直接回车默认今天 {default_date}：").strip()
        selected_value = raw_value or default_date
        try:
            parse_date(selected_value)
        except ValueError:
            print("日期格式不对，请输入类似 2026-06-10 的格式。")
            continue
        return selected_value


def apply_interactive_fill_defaults(args: argparse.Namespace) -> argparse.Namespace:
    args.workbook = choose_workbook_from_current_directory()
    args.current_date = choose_current_date()
    args.freeze_next_day_row = True
    args.colored_sheets = True
    args.fast_xml = True
    args.run_until_done = True
    args.limit_sheets = args.limit_sheets or 20

    print()
    print("将使用以下默认参数处理：")
    print(f"  workbook: {args.workbook}")
    print(f"  current-date: {args.current_date}")
    print("  colored-sheets: yes")
    print("  fast-xml: yes")
    print(f"  limit-sheets: {args.limit_sheets}")
    print("  run-until-done: yes")
    print()
    return args


def main() -> None:
    args = parse_args()

    if args.interactive:
        args = apply_interactive_fill_defaults(args)

    if args.list_sheets:
        for sheet_name in list_sheet_names(args.workbook):
            print(sheet_name)
        return

    if args.freeze_next_day_row:
        current_date = resolve_current_date(args)
        if args.fast_xml:
            if not args.colored_sheets:
                raise ValueError("--fast-xml currently requires --colored-sheets")
            total_changed = 0
            batch_number = 0
            while True:
                batch_number += 1
                result = add_current_date_to_colored_sheets_fast(
                    xlsx_path=args.workbook,
                    current_date=current_date,
                    limit_sheets=args.limit_sheets or 20,
                )
                changed_count = len(result.changed)
                total_changed += changed_count
                print(
                    f"Batch {batch_number}: changed {changed_count}, "
                    f"skipped {len(result.skipped)}."
                )
                for sheet_name, row in result.changed:
                    print(f"  changed: {sheet_name} row {row}")

                if not args.run_until_done or changed_count == 0:
                    if changed_count == 0:
                        print("No more matching colored sheets to change.")
                    print(f"Total changed: {total_changed}")
                    break
            return

        workbook = load_workbook(args.workbook, data_only=False)
        if args.sheet not in workbook.sheetnames and args.copy_sheet_from:
            if args.copy_sheet_from not in workbook.sheetnames:
                available = ", ".join(workbook.sheetnames)
                raise ValueError(
                    f"Source sheet '{args.copy_sheet_from}' not found. Available sheets: {available}"
                )
            copied_sheet = workbook.copy_worksheet(workbook[args.copy_sheet_from])
            copied_sheet.title = args.sheet
            workbook.save(args.workbook)

            workbook = load_workbook(args.workbook, data_only=False)

        values_workbook = load_workbook(args.workbook, data_only=True)
        if args.colored_sheets:
            changed, skipped = freeze_colored_sheets_next_day(
                workbook=workbook,
                values_workbook=values_workbook,
                current_date=current_date,
                limit_sheets=args.limit_sheets,
            )
            workbook.save(args.workbook)
            print(f"Changed {len(changed)} colored sheets.")
            for sheet_name, row in changed:
                print(f"  changed: {sheet_name} row {row}")
            print(f"Skipped {len(skipped)} colored sheets.")
            for sheet_name, reason in skipped[:50]:
                print(f"  skipped: {sheet_name} ({reason})")
            if len(skipped) > 50:
                print(f"  ... {len(skipped) - 50} more skipped")
            return

        inserted_row = freeze_next_day_row(
            workbook=workbook,
            values_workbook=values_workbook,
            sheet_name=args.sheet,
            current_date=current_date,
        )
        workbook.save(args.workbook)
        print(f"Inserted row {inserted_row} in sheet '{args.sheet}' and pasted source row as values.")
        return

    preview = preview_sheet(args.workbook, args.sheet, max_rows=args.rows)

    print(f"Workbook: {args.workbook}")
    print(f"Sheet: {preview.sheet_name}")
    print(f"Size: {preview.max_row} rows x {preview.max_column} columns")
    print("Preview:")
    for row in preview.rows:
        print(row)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n已取消。")
        sys.exit(130)
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        print(f"错误：{error}")
        sys.exit(1)
