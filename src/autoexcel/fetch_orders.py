from __future__ import annotations

import argparse
import configparser
from dataclasses import dataclass
from datetime import date, datetime, timedelta
import json
from pathlib import Path
import ssl
import sys
import traceback
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin
from urllib.request import Request, urlopen


PROJECT_ROOT = Path(__file__).resolve().parents[2]
CONFIG_FILE_NAME = "config.ini"
LOGIN_CONFIG_FILE_NAME = "loginConf.ini"
DEFAULT_RESULT_DIR = PROJECT_ROOT / "result"
DEFAULT_DIFF_ORDERS_DIR = PROJECT_ROOT / "workspace" / "diffOrders"
UPSTREAM_LOGIN_SECTION = "upstream_server_login"
UPSTREAM_SESSION_SECTION = "upstream_server_session"
LEGACY_FETCH_LOGIN_SECTION = "fetch_orders_login"


@dataclass(frozen=True)
class FetchOrdersConfig:
    login_url: str
    api_key: str
    username: str
    password: str
    ip_address: str
    session_id: str
    transition_id: str
    institution_id: str
    origin: str
    referer: str
    user_agent: str
    timeout_seconds: int
    verify_ssl: bool
    report_url: str
    scheduled_reports_url: str
    download_url: str
    download_dir: Path
    report_name: str
    transaction_status: str


@dataclass(frozen=True)
class PortalSession:
    token: str
    user_id: str
    user_name: str
    user_type: str
    merchant_code: str


@dataclass(frozen=True)
class ScheduledReport:
    file_path: str
    file_name: str
    status: str
    progress_status: str
    is_download: str
    created_at: datetime | None


def is_frozen_app() -> bool:
    return getattr(sys, "frozen", False)


def get_executable_directory() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


def get_config_path() -> Path:
    return find_config_path(CONFIG_FILE_NAME)


def get_login_config_path() -> Path:
    return find_config_path(LOGIN_CONFIG_FILE_NAME)


def find_config_path(file_name: str) -> Path:
    if is_frozen_app():
        candidates = (
            get_executable_directory() / file_name,
            Path.cwd() / file_name,
            PROJECT_ROOT / file_name,
        )
    else:
        candidates = (
            Path.cwd() / file_name,
            PROJECT_ROOT / file_name,
        )
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def parse_bool(value: str, option_name: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "yes", "y", "true", "on"}:
        return True
    if normalized in {"0", "no", "n", "false", "off"}:
        return False
    raise ValueError(f"config.ini 中 {option_name} 必须是 true 或 false")


def read_config_file(path: Path) -> configparser.ConfigParser:
    config = configparser.ConfigParser(strict=False)
    if path.exists():
        config.read(path, encoding="utf-8")
    return config


def read_config() -> configparser.SectionProxy:
    config = read_config_file(get_config_path())
    if "fetch_orders" not in config:
        config["fetch_orders"] = {}
    return config["fetch_orders"]


def read_login_config() -> configparser.SectionProxy:
    login_config_path = get_login_config_path()
    config = read_config_file(login_config_path)
    if UPSTREAM_LOGIN_SECTION not in config:
        if LEGACY_FETCH_LOGIN_SECTION in config:
            config[UPSTREAM_LOGIN_SECTION] = dict(config[LEGACY_FETCH_LOGIN_SECTION])
            with login_config_path.open("w", encoding="utf-8") as file:
                config.write(file)
        else:
            config[UPSTREAM_LOGIN_SECTION] = {}
    return config[UPSTREAM_LOGIN_SECTION]


def resolve_project_path(value: str, default_path: Path) -> Path:
    raw_value = value.strip()
    if not raw_value:
        return default_path
    path = Path(raw_value).expanduser()
    if path.is_absolute():
        return path
    return PROJECT_ROOT / path


def required_value(config: configparser.SectionProxy, key: str, file_name: str, section_name: str) -> str:
    value = config.get(key, "").strip()
    if not value:
        raise ValueError(f"请先在 {file_name} 的 [{section_name}] 中填写 {key}")
    return value


def load_fetch_orders_config() -> FetchOrdersConfig:
    config = read_config()
    login_config = read_login_config()
    return FetchOrdersConfig(
        login_url=required_value(config, "login_url", CONFIG_FILE_NAME, "fetch_orders"),
        api_key=required_value(login_config, "api_key", LOGIN_CONFIG_FILE_NAME, UPSTREAM_LOGIN_SECTION),
        username=required_value(login_config, "username", LOGIN_CONFIG_FILE_NAME, UPSTREAM_LOGIN_SECTION),
        password=required_value(login_config, "password", LOGIN_CONFIG_FILE_NAME, UPSTREAM_LOGIN_SECTION),
        ip_address=login_config.get("ip_address", "").strip(),
        session_id=login_config.get("session_id", "").strip(),
        transition_id=login_config.get("transition_id", "").strip(),
        institution_id=login_config.get("institution_id", "").strip(),
        origin=config.get("origin", "https://pgw-portal.jazzcash.com.pk").strip(),
        referer=config.get("referer", "https://pgw-portal.jazzcash.com.pk/").strip(),
        user_agent=config.get(
            "user_agent",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/149.0.0.0 Safari/537.36",
        ).strip(),
        timeout_seconds=int(config.get("timeout_seconds", "30").strip() or "30"),
        verify_ssl=parse_bool(config.get("verify_ssl", "true"), "fetch_orders.verify_ssl"),
        report_url=required_value(config, "report_url", CONFIG_FILE_NAME, "fetch_orders"),
        scheduled_reports_url=required_value(config, "scheduled_reports_url", CONFIG_FILE_NAME, "fetch_orders"),
        download_url=required_value(config, "download_url", CONFIG_FILE_NAME, "fetch_orders"),
        download_dir=resolve_project_path(config.get("download_dir", ""), DEFAULT_DIFF_ORDERS_DIR),
        report_name=config.get("report_name", "TransactionDetailReport").strip() or "TransactionDetailReport",
        transaction_status=config.get("transaction_status", "Completed").strip() or "Completed",
    )


def build_headers(config: FetchOrdersConfig, token: str | None = None) -> dict[str, str]:
    return {
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Access-Control-Allow-Origin": "*",
        "Authorization": f"Bearer {token}" if token else "Bearer null",
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "Content-Type": "application/json",
        "Origin": config.origin,
        "Pragma": "no-cache",
        "Referer": config.referer,
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": config.user_agent,
        "apikey": config.api_key,
        "sec-ch-ua": '"Google Chrome";v="149", "Chromium";v="149", "Not)A;Brand";v="24"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"macOS"',
    }


def build_payload(config: FetchOrdersConfig) -> dict[str, str]:
    return {
        "Username": config.username,
        "Password": config.password,
        "IpAddress": config.ip_address,
        "SessionId": config.session_id,
        "TransitionId": config.transition_id,
        "InstitutionId": config.institution_id,
    }


def default_target_date() -> date:
    return date.today() - timedelta(days=1)


def parse_date_value(value: str) -> date:
    raw_value = value.strip()
    current_year = date.today().year

    if not raw_value:
        return default_target_date()
    if raw_value.isdigit() and len(raw_value) == 4:
        return date(current_year, int(raw_value[:2]), int(raw_value[2:]))

    normalized = raw_value.replace("/", "-").replace(".", "-")
    if "-" in normalized:
        parts = normalized.split("-")
        if len(parts) == 2:
            return date(current_year, int(parts[0]), int(parts[1]))
        if len(parts) == 3:
            return datetime.strptime(normalized, "%Y-%m-%d").date()

    raise ValueError("日期格式不对，请输入 2026-06-28、0628、06-28 或 06/28")


def choose_target_date() -> date:
    default_date = default_target_date()
    if not sys.stdin.isatty():
        return default_date

    default_text = default_date.strftime("%Y-%m-%d")
    while True:
        raw_value = input(f"请输入拉取订单日期，直接回车默认前一天 {default_text}：").strip()
        try:
            return parse_date_value(raw_value)
        except ValueError as error:
            print(error)


def resolve_target_date(raw_value: str | None) -> date:
    if raw_value:
        return parse_date_value(raw_value)
    return choose_target_date()


def format_report_datetime(target_date: date, end_of_day: bool = False) -> str:
    time_text = "23:59" if end_of_day else "00:00"
    return f"{target_date:%m/%d/%Y} {time_text}"


def build_schedule_payload(config: FetchOrdersConfig, target_date: date) -> dict[str, Any]:
    start_date = format_report_datetime(target_date)
    end_date = format_report_datetime(target_date, end_of_day=True)
    return {
        "DateFrom": start_date,
        "DateTo": end_date,
        "TRANS_TYPE": None,
        "TRANS_STATUS": config.transaction_status,
        "MERCHANT_ID": None,
        "ReportName": config.report_name,
        "StartDate": start_date,
        "EndDate": end_date,
        "TransactionType": None,
        "TransactionStatus": config.transaction_status,
        "MerchantId": None,
    }


def build_scheduled_reports_payload(config: FetchOrdersConfig) -> dict[str, Any]:
    return {
        "reportName": config.report_name,
        "user_Id": 0,
    }


def build_download_payload(report: ScheduledReport) -> dict[str, str]:
    return {
        "filePath": report.file_path,
        "fileName": report.file_name,
    }


def parse_response_body(raw_body: bytes) -> Any:
    text = raw_body.decode("utf-8", errors="replace")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {"raw": text}


def post_json(
    url: str,
    headers: dict[str, str],
    payload: dict[str, Any],
    timeout_seconds: int,
    verify_ssl: bool,
) -> tuple[int, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(url, data=data, headers=headers, method="POST")
    ssl_context = None if verify_ssl else ssl._create_unverified_context()
    try:
        with urlopen(request, timeout=timeout_seconds, context=ssl_context) as response:
            return response.status, parse_response_body(response.read())
    except HTTPError as error:
        body = parse_response_body(error.read())
        return error.code, body
    except URLError as error:
        raise RuntimeError(f"请求失败：{error}") from error


def post_binary(
    url: str,
    headers: dict[str, str],
    payload: dict[str, Any],
    timeout_seconds: int,
    verify_ssl: bool,
) -> tuple[int, dict[str, str], bytes]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = Request(url, data=data, headers=headers, method="POST")
    ssl_context = None if verify_ssl else ssl._create_unverified_context()
    try:
        with urlopen(request, timeout=timeout_seconds, context=ssl_context) as response:
            return response.status, dict(response.headers.items()), response.read()
    except HTTPError as error:
        return error.code, dict(error.headers.items()), error.read()
    except URLError as error:
        raise RuntimeError(f"请求失败：{error}") from error


def post_login(config: FetchOrdersConfig) -> tuple[int, Any]:
    return post_json(
        config.login_url,
        build_headers(config),
        build_payload(config),
        config.timeout_seconds,
        config.verify_ssl,
    )


def post_schedule_ready(config: FetchOrdersConfig, session: PortalSession, target_date: date) -> tuple[int, Any]:
    return post_json(
        config.report_url,
        build_headers(config, token=session.token),
        build_schedule_payload(config, target_date),
        config.timeout_seconds,
        config.verify_ssl,
    )


def post_scheduled_reports(config: FetchOrdersConfig, session: PortalSession) -> tuple[int, Any]:
    return post_json(
        config.scheduled_reports_url,
        build_headers(config, token=session.token),
        build_scheduled_reports_payload(config),
        config.timeout_seconds,
        config.verify_ssl,
    )


def post_download_zip(config: FetchOrdersConfig, session: PortalSession, report: ScheduledReport) -> tuple[int, dict[str, str], bytes]:
    return post_binary(
        config.download_url,
        build_headers(config, token=session.token),
        build_download_payload(report),
        config.timeout_seconds,
        config.verify_ssl,
    )


def extract_portal_session(body: Any) -> PortalSession:
    if not isinstance(body, dict):
        raise ValueError("登录响应不是 JSON 对象，无法提取 token")
    token = str(body.get("token") or "").strip()
    if not token:
        raise ValueError("登录响应中没有 token，无法继续查询报表")
    return PortalSession(
        token=token,
        user_id=str(body.get("userId") or ""),
        user_name=str(body.get("userName") or ""),
        user_type=str(body.get("userType") or ""),
        merchant_code=str(body.get("merchantCode") or ""),
    )


def parse_report_timestamp(file_name: str) -> datetime | None:
    parts = file_name.split("_")
    for part in reversed(parts):
        digits = "".join(char for char in part if char.isdigit())
        if len(digits) >= 14:
            try:
                return datetime.strptime(digits[:14], "%Y%m%d%H%M%S")
            except ValueError:
                continue
    return None


def get_case_insensitive(record: dict[str, Any], key: str) -> Any:
    target = key.lower()
    for record_key, value in record.items():
        if str(record_key).lower() == target:
            return value
    return None


def get_report_table(body: Any) -> list[Any]:
    if isinstance(body, list):
        return body
    if not isinstance(body, dict):
        return []
    data = body.get("data") or body.get("Data")
    if isinstance(data, dict):
        table = data.get("table") or data.get("Table")
        if isinstance(table, list):
            return table
    for key in ("table", "Table", "result", "Result", "items", "Items"):
        value = body.get(key)
        if isinstance(value, list):
            return value
    return []


def extract_scheduled_reports(body: Any) -> list[ScheduledReport]:
    reports: list[ScheduledReport] = []
    for item in get_report_table(body):
        if not isinstance(item, dict):
            continue
        file_path = str(
            get_case_insensitive(item, "filelink")
            or get_case_insensitive(item, "url")
            or get_case_insensitive(item, "filePath")
            or ""
        ).strip()
        file_name = str(
            get_case_insensitive(item, "reportfilename")
            or get_case_insensitive(item, "fileName")
            or Path(file_path).name
            or ""
        ).strip()
        if not file_path or not file_name:
            continue
        status = str(get_case_insensitive(item, "status") or "").strip()
        progress_status = str(get_case_insensitive(item, "progresS_STATUS") or get_case_insensitive(item, "progressStatus") or "").strip()
        is_download = str(get_case_insensitive(item, "iS_DOWNLOAD") or get_case_insensitive(item, "isDownload") or "").strip()
        reports.append(
            ScheduledReport(
                file_path=file_path,
                file_name=file_name,
                status=status,
                progress_status=progress_status,
                is_download=is_download,
                created_at=parse_report_timestamp(file_name),
            )
        )
    return reports


def is_downloadable_report(report: ScheduledReport) -> bool:
    return (
        report.status.lower() == "complete"
        and report.progress_status in {"100", "100.0", ""}
        and report.is_download in {"1", "true", "True", "TRUE", ""}
    )


def choose_latest_report(reports: list[ScheduledReport]) -> ScheduledReport:
    downloadable_reports = [report for report in reports if is_downloadable_report(report)]
    if not downloadable_reports:
        raise ValueError("已调度报表列表里没有找到可下载的 Complete Excel。")
    return max(downloadable_reports, key=lambda report: report.created_at or datetime.min)


def parse_json_error(raw_body: bytes) -> str:
    try:
        body = json.loads(raw_body.decode("utf-8", errors="replace"))
    except json.JSONDecodeError:
        return raw_body[:300].decode("utf-8", errors="replace")
    message = body.get("message") or body.get("Message") or body.get("error") or body.get("Error") if isinstance(body, dict) else ""
    return str(message or body)[:300]


def save_downloaded_excel(config: FetchOrdersConfig, report: ScheduledReport, raw_body: bytes) -> Path:
    config.download_dir.mkdir(parents=True, exist_ok=True)
    output_path = config.download_dir / report.file_name
    output_path.write_bytes(raw_body)
    return output_path


def ask_generate_new_report() -> bool:
    if not sys.stdin.isatty():
        return False
    while True:
        raw_value = input("是否需要生成最新的 Excel？输入 y 生成，直接回车或输入 n 使用最近已生成文件：").strip().lower()
        if raw_value in {"", "n", "no"}:
            return False
        if raw_value in {"y", "yes"}:
            return True
        print("请输入 y 或 n；直接回车表示不生成，使用最近已生成文件。")


def update_login_config_session(
    session: PortalSession,
    login_status_code: int,
    schedule_status_code: int | None = None,
    scheduled_reports_status_code: int | None = None,
    scheduled_reports_count: int | None = None,
    target_date: date | None = None,
    downloaded_report: ScheduledReport | None = None,
    downloaded_file_path: Path | None = None,
    download_status_code: int | None = None,
) -> Path:
    login_config_path = get_login_config_path()
    config = read_config_file(login_config_path)
    if UPSTREAM_SESSION_SECTION not in config:
        config[UPSTREAM_SESSION_SECTION] = {}

    session_section = config[UPSTREAM_SESSION_SECTION]
    session_section["updated_at"] = datetime.now().isoformat(timespec="seconds")
    session_section["token"] = session.token
    session_section["user_id"] = session.user_id
    session_section["user_name"] = session.user_name
    session_section["user_type"] = session.user_type
    session_section["merchant_code"] = session.merchant_code
    session_section["login_status_code"] = str(login_status_code)
    if schedule_status_code is not None:
        session_section["schedule_status_code"] = str(schedule_status_code)
    if scheduled_reports_status_code is not None:
        session_section["scheduled_reports_status_code"] = str(scheduled_reports_status_code)
    if scheduled_reports_count is not None:
        session_section["scheduled_reports_count"] = str(scheduled_reports_count)
    if target_date is not None:
        session_section["last_report_date"] = target_date.strftime("%Y-%m-%d")
        session_section["last_start_date"] = format_report_datetime(target_date)
        session_section["last_end_date"] = format_report_datetime(target_date, end_of_day=True)
    if downloaded_report is not None:
        session_section["last_report_file_path"] = downloaded_report.file_path
        session_section["last_report_file_name"] = downloaded_report.file_name
        if downloaded_report.created_at is not None:
            session_section["last_report_created_at"] = downloaded_report.created_at.isoformat(timespec="seconds")
    if downloaded_file_path is not None:
        session_section["last_downloaded_file"] = str(downloaded_file_path)
    if download_status_code is not None:
        session_section["download_status_code"] = str(download_status_code)

    with login_config_path.open("w", encoding="utf-8") as file:
        config.write(file)
    return login_config_path


def count_response_items(body: Any) -> int:
    if isinstance(body, list):
        return len(body)
    if isinstance(body, dict):
        for key in ("data", "Data", "result", "Result", "items", "Items"):
            value = body.get(key)
            if isinstance(value, list):
                return len(value)
            if isinstance(value, dict):
                nested_count = count_response_items(value)
                if nested_count:
                    return nested_count
    return 0


def response_success(status_code: int, body: Any) -> bool:
    if not 200 <= status_code < 300:
        return False
    if isinstance(body, dict):
        status_value = str(body.get("status") or body.get("Status") or body.get("code") or body.get("Code") or "")
        message = str(body.get("message") or body.get("Message") or "")
        if status_value and status_value not in {"0", "00", "200", "success", "SUCCESS"}:
            return False
        if "invalid" in message.lower() or "fail" in message.lower() or "error" in message.lower():
            return False
    return True


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fetch order Excel files for diffOrders.")
    parser.add_argument("--config", type=Path, help="Reserved for future use; currently reads config.ini automatically.")
    parser.add_argument("--date", help="Target report date. Defaults to yesterday. Examples: 2026-06-28, 0628, 06-28.")
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    config = load_fetch_orders_config()
    print("正在登录后台服务，准备拉取订单 Excel...")
    status_code, body = post_login(config)
    if response_success(status_code, body):
        print(f"登录请求完成，HTTP 状态码：{status_code}")
    else:
        print(f"登录可能失败，HTTP 状态码：{status_code}。请查看保存的响应文件。")
        raise SystemExit(1)

    session = extract_portal_session(body)
    login_config_path = update_login_config_session(session, status_code)
    print(f"登录会话信息已写入：{login_config_path}")

    if args.date:
        target_date = resolve_target_date(args.date)
    elif ask_generate_new_report():
        target_date = resolve_target_date(None)
    else:
        target_date = None
    schedule_status_code = None
    if target_date is not None:
        print(f"正在查询报表任务：{target_date:%Y-%m-%d} 00:00-23:59")
        schedule_status_code, schedule_body = post_schedule_ready(config, session, target_date)
        update_login_config_session(session, status_code, schedule_status_code, target_date=target_date)
        if response_success(schedule_status_code, schedule_body):
            print(f"报表查询完成，HTTP 状态码：{schedule_status_code}")
        else:
            print(f"报表查询可能失败，HTTP 状态码：{schedule_status_code}。")
            raise SystemExit(1)
    else:
        print("本次不生成新 Excel，改用最近已生成的可下载文件。")

    print("正在查询已调度报表列表...")
    scheduled_status_code, scheduled_body = post_scheduled_reports(config, session)
    scheduled_reports = extract_scheduled_reports(scheduled_body)
    scheduled_count = len(scheduled_reports) or count_response_items(scheduled_body)
    update_login_config_session(
        session,
        status_code,
        schedule_status_code,
        scheduled_status_code,
        scheduled_count,
        target_date,
    )
    if response_success(scheduled_status_code, scheduled_body):
        print(f"已调度报表查询完成，HTTP 状态码：{scheduled_status_code}，记录数：{scheduled_count}")
    else:
        print(f"已调度报表查询可能失败，HTTP 状态码：{scheduled_status_code}。")
        raise SystemExit(1)

    latest_report = choose_latest_report(scheduled_reports)
    print(f"准备下载最近生成的 Excel：{latest_report.file_name}")
    download_status_code, download_headers, download_body = post_download_zip(config, session, latest_report)
    content_type = download_headers.get("Content-Type", "")
    if not 200 <= download_status_code < 300:
        raise RuntimeError(f"下载失败，HTTP 状态码：{download_status_code}，响应：{parse_json_error(download_body)}")
    if "json" in content_type.lower() and not download_body.startswith(b"PK"):
        raise RuntimeError(f"下载接口没有返回 Excel 文件，响应：{parse_json_error(download_body)}")

    downloaded_path = save_downloaded_excel(config, latest_report, download_body)
    update_login_config_session(
        session,
        status_code,
        schedule_status_code,
        scheduled_status_code,
        scheduled_count,
        target_date,
        latest_report,
        downloaded_path,
        download_status_code,
    )
    print(f"Excel 下载完成：{downloaded_path}")


if __name__ == "__main__":
    exit_code = 0
    try:
        main()
    except (RuntimeError, ValueError) as error:
        print(f"错误：{error}")
        exit_code = 1
    except Exception:
        print("程序执行失败：")
        traceback.print_exc()
        exit_code = 1
    sys.exit(exit_code)
