from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from autoexcel import add_b2b, add_cards, diff_orders, fetch_orders
from autoexcel.fast_xlsx import (
    add_current_date_to_colored_sheets_fast,
    advance_summary_table_sheet_fast,
)
from autoexcel.main import FillSummary, append_log, create_process_log_path


LogCallback = Callable[[str], None]


@dataclass(frozen=True)
class TaskResult:
    title: str
    summary: str
    output_path: Path | None = None


def run_add_cards_task(workbook: Path, text: str, log: LogCallback) -> TaskResult:
    cards = add_cards.parse_card_numbers(text)
    if not cards:
        raise ValueError("请至少输入一个卡号")
    log(f"正在向 {workbook.name} 添加 {len(cards)} 张卡…")
    result = add_cards.add_cards_to_workbook(workbook, cards)
    summary = f"新增 {len(result.created)} 张卡，跳过 {len(result.skipped)} 张。"
    if result.skipped:
        details = "；".join(f"{card}（{reason}）" for card, reason in result.skipped)
        summary = f"{summary} 跳过详情：{details}"
    return TaskResult("增卡完成", summary, workbook)


def run_add_b2b_task(
    workbook: Path,
    text: str,
    mapping: add_b2b.FieldMapping,
    log: LogCallback,
) -> TaskResult:
    lines = add_b2b.parse_input_text(text)
    log(f"正在向 {workbook.name} 写入 {len(lines)} 行 B2B 数据…")
    result = add_b2b.append_b2b_to_workbook(workbook, lines, mapping)
    summary = (
        f"新增 {result.inserted_count} 行，位置："
        f"提取B2B!A{result.start_row}:I{result.end_row}。"
    )
    if result.converted_negative_count:
        summary += f" 已将 {result.converted_negative_count} 行负金额转为正数。"
    return TaskResult("提取B2B完成", summary, workbook)


def run_fill_task(
    workbook: Path,
    target_date: date,
    log: LogCallback,
    limit_sheets: int = 20,
) -> TaskResult:
    if not workbook.is_file():
        raise FileNotFoundError(f"Excel 文件不存在：{workbook}")
    if workbook.suffix.lower() != ".xlsx":
        raise ValueError("请选择 .xlsx 文件")

    log_path = create_process_log_path()
    summary = FillSummary(workbook=workbook, current_date=target_date, log_path=log_path)
    append_log(log_path, "SmartSheet Desk 处理日志")
    append_log(log_path, f"开始时间: {datetime.now():%Y-%m-%d %H:%M:%S}")
    append_log(log_path, f"Workbook: {workbook}")
    append_log(log_path, f"目标日期: {target_date:%Y-%m-%d}")

    start_index = 0
    while True:
        summary.batch_count += 1
        batch_number = summary.batch_count
        log(f"第 {batch_number} 批：正在处理，请勿在 Excel/WPS 中打开文件…")
        result = add_current_date_to_colored_sheets_fast(
            xlsx_path=workbook,
            current_date=target_date,
            limit_sheets=limit_sheets,
            start_index=start_index,
            progress=log,
        )
        start_index = result.next_start_index
        summary.changed.extend(result.changed)
        summary.skipped_count += len(result.skipped)
        append_log(
            log_path,
            f"Batch {batch_number}: changed {len(result.changed)}, skipped {len(result.skipped)}",
        )
        for sheet_name, row in result.changed:
            append_log(log_path, f"  changed: {sheet_name} row {row}")
        for sheet_name, reason in result.skipped:
            append_log(log_path, f"  skipped: {sheet_name} ({reason})")
        log(f"第 {batch_number} 批完成：修改 {len(result.changed)}，跳过 {len(result.skipped)}")

        if not result.changed or start_index >= result.total_sheets:
            break

    for sheet_name, opening_balance_increment in (
        ("B2B支出", False),
        ("收入", False),
        ("每日余额监测", True),
    ):
        log(f"正在处理{sheet_name}工作表…")
        sheet_result = advance_summary_table_sheet_fast(
            xlsx_path=workbook,
            sheet_name=sheet_name,
            current_date=target_date,
            progress=log,
            opening_balance_increment=opening_balance_increment,
        )
        if sheet_result.changed:
            summary.changed.append((sheet_name, sheet_result.inserted_row or 0))
            append_log(log_path, f"  changed: {sheet_name} row {sheet_result.inserted_row}")
            log(f"{sheet_name}工作表完成：插入第 {sheet_result.inserted_row} 行")
        else:
            summary.skipped_count += 1
            append_log(log_path, f"  skipped: {sheet_name} ({sheet_result.reason})")
            log(f"{sheet_name}工作表跳过：{sheet_result.reason}")

    append_log(log_path, f"结束时间: {datetime.now():%Y-%m-%d %H:%M:%S}")
    summary_text = (
        f"完成 {summary.batch_count} 个批次，修改 {summary.changed_count} 个工作表，"
        f"跳过 {summary.skipped_count} 个工作表。"
    )
    return TaskResult("Excel 填充完成", summary_text, workbook)


def _jobs_from_directory(directory: Path, target_date: date) -> list[diff_orders.DiffJob]:
    if not directory.is_dir():
        raise FileNotFoundError(f"订单目录不存在：{directory}")
    files = diff_orders.list_xlsx_files(directory)
    if not files:
        raise FileNotFoundError(f"目录中没有找到 Excel 文件：{directory}")

    date_text = target_date.strftime("%Y%m%d")
    backend_paths = diff_orders.matching_workbooks(
        files, diff_orders.match_backend_workbook, "后台", date_text
    )
    upstream_paths = diff_orders.matching_workbook_candidates(
        files, diff_orders.match_upstream_workbook, date_text
    )
    upstream_paths.extend(
        path
        for path in diff_orders.easypaisa_upstream_paths_for_backends(files, backend_paths)
        if path not in upstream_paths
    )
    if not upstream_paths:
        raise FileNotFoundError(f"没有找到 {date_text} 的上游 Excel")
    duplicate_paths = sorted(
        (path for path in files if diff_orders.match_duplicate_payment_workbook(path, date_text)),
        key=lambda path: path.name,
    )
    return diff_orders.pair_diff_workbooks(upstream_paths, backend_paths, duplicate_paths)


def run_diff_task(directory: Path, target_date: date, log: LogCallback) -> TaskResult:
    jobs = _jobs_from_directory(directory, target_date)
    return _run_diff_jobs(jobs, log)


def run_diff_files_task(
    upstream_path: Path,
    backend_path: Path,
    duplicate_path: Path | None,
    platform_mode: str,
    log: LogCallback,
) -> TaskResult:
    required_files = (
        ("上游 Excel", upstream_path),
        ("平台收款订单 Excel", backend_path),
    )
    for label, path in required_files:
        if not path.is_file():
            raise FileNotFoundError(f"{label} 不存在：{path}")
        if path.suffix.lower() != ".xlsx":
            raise ValueError(f"{label} 必须是 .xlsx 文件")
    if duplicate_path is not None:
        if not duplicate_path.is_file():
            raise FileNotFoundError(f"代收重复支付订单 Excel 不存在：{duplicate_path}")
        if duplicate_path.suffix.lower() != ".xlsx":
            raise ValueError("代收重复支付订单必须是 .xlsx 文件")

    resolved_mode = platform_mode or diff_orders.platform_mode_for_pair(
        upstream_path, backend_path, duplicate_path
    )
    job = diff_orders.DiffJob(
        upstream_path=upstream_path,
        backend_path=backend_path,
        duplicate_path=duplicate_path,
        platform_mode=resolved_mode,
    )
    return _run_diff_jobs([job], log)


def _run_diff_jobs(
    jobs: list[diff_orders.DiffJob], log: LogCallback
) -> TaskResult:
    log(f"已匹配 {len(jobs)} 组订单文件")
    args = argparse.Namespace(a_col="L", b_col="D")
    results: list[diff_orders.JobDiffResult] = []
    for index, job in enumerate(jobs, start=1):
        log(f"正在比对第 {index}/{len(jobs)} 组：{job.upstream_path.name}")
        results.append(diff_orders.read_job_diff_result(job, args))

    result_dir = diff_orders.get_result_dir()
    if len(results) == 1:
        html_path = diff_orders.write_html_result(
            result_dir, diff_orders.get_template_path(), results[0]
        )
    else:
        html_path = diff_orders.write_batch_html_result(result_dir, results)
    upstream_only = sum(len(diff_orders.unique_entries(item.result.a_only)) for item in results)
    return TaskResult(
        "订单比对完成",
        f"共处理 {len(results)} 组文件，上游独有订单 {upstream_only} 条。",
        html_path,
    )


def run_fetch_task(
    target_date: date,
    generate_new_report: bool,
    log: LogCallback,
) -> TaskResult:
    config = fetch_orders.load_fetch_orders_config()
    log("正在登录后台服务…")
    login_status, login_body = fetch_orders.post_login(config)
    if not fetch_orders.response_success(login_status, login_body):
        raise RuntimeError(f"登录失败，HTTP 状态码：{login_status}")
    session = fetch_orders.extract_portal_session(login_body)
    fetch_orders.update_login_config_session(session, login_status)
    log("登录成功")

    schedule_status = None
    if generate_new_report:
        log(f"正在生成 {target_date:%Y-%m-%d} 的报表…")
        schedule_status, schedule_body = fetch_orders.post_schedule_ready(
            config, session, target_date
        )
        if not fetch_orders.response_success(schedule_status, schedule_body):
            raise RuntimeError(f"生成报表失败，HTTP 状态码：{schedule_status}")

    log("正在查询已调度报表…")
    scheduled_status, scheduled_body = fetch_orders.post_scheduled_reports(config, session)
    if not fetch_orders.response_success(scheduled_status, scheduled_body):
        raise RuntimeError(f"查询报表失败，HTTP 状态码：{scheduled_status}")
    reports = fetch_orders.extract_scheduled_reports(scheduled_body)
    scheduled_count = len(reports) or fetch_orders.count_response_items(scheduled_body)
    selected_report = fetch_orders.choose_first_report_for_date(reports, target_date)
    log(f"正在下载：{selected_report.file_name}")
    download_status, download_headers, body = fetch_orders.post_download_zip(
        config, session, selected_report
    )
    content_type = download_headers.get("Content-Type", "")
    if not 200 <= download_status < 300:
        raise RuntimeError(f"下载失败，HTTP 状态码：{download_status}")
    if "json" in content_type.lower() and not body.startswith(b"PK"):
        raise RuntimeError(f"下载接口没有返回 Excel：{fetch_orders.parse_json_error(body)}")

    output_path = fetch_orders.save_downloaded_excel(config, selected_report, body)
    fetch_orders.update_login_config_session(
        session,
        login_status,
        schedule_status,
        scheduled_status,
        scheduled_count,
        target_date,
        selected_report,
        output_path,
        download_status,
    )
    return TaskResult("订单下载完成", f"文件已保存为 {output_path.name}", output_path)
