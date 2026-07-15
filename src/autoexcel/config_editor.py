from __future__ import annotations

import configparser
import os
from pathlib import Path
import re
import shutil
import tempfile

from autoexcel import fetch_orders
from autoexcel.runtime_paths import application_directory, bundled_resource


def editable_config_path() -> Path:
    current_path = fetch_orders.get_config_path()
    resource_path = bundled_resource(fetch_orders.CONFIG_FILE_NAME)
    if resource_path is None or current_path != resource_path:
        return current_path

    target_path = application_directory() / fetch_orders.CONFIG_FILE_NAME
    if not target_path.exists():
        shutil.copy2(resource_path, target_path)
    return target_path


def read_ini(path: Path) -> configparser.ConfigParser:
    parser = configparser.ConfigParser(strict=False)
    if path.exists():
        parser.read(path, encoding="utf-8")
    return parser


def read_fill_limit_sheets(path: Path, default: int = 20) -> int:
    parser = read_ini(path)
    raw_value = parser.get("fill", "limit_sheets", fallback=str(default)).strip()
    value = int(raw_value or default)
    if value < 1:
        raise ValueError("config.ini 中 fill.limit_sheets 必须大于 0")
    return value


def update_ini(path: Path, updates: dict[str, dict[str, str]]) -> None:
    text = path.read_text(encoding="utf-8") if path.exists() else ""
    lines = text.splitlines()
    section_ranges = _section_ranges(lines)

    for section, values in updates.items():
        if section not in section_ranges:
            if lines and lines[-1].strip():
                lines.append("")
            lines.append(f"[{section}]")
            lines.extend(f"{key} = {value}" for key, value in values.items())
            section_ranges = _section_ranges(lines)
            continue

        start, end = section_ranges[section]
        remaining = dict(values)
        for index in range(start + 1, end):
            match = re.match(r"^(\s*)([^#;\s][^=]*?)(\s*=\s*)(.*)$", lines[index])
            if not match:
                continue
            key = match.group(2).strip().lower()
            if key not in remaining:
                continue
            lines[index] = f"{match.group(1)}{match.group(2)}{match.group(3)}{remaining.pop(key)}"
        for key, value in remaining.items():
            lines.insert(end, f"{key} = {value}")
            end += 1
        section_ranges = _section_ranges(lines)

    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary_name = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    os.close(fd)
    temporary_path = Path(temporary_name)
    try:
        temporary_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
        os.replace(temporary_path, path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _section_ranges(lines: list[str]) -> dict[str, tuple[int, int]]:
    headers: list[tuple[str, int]] = []
    for index, line in enumerate(lines):
        match = re.match(r"^\s*\[([^]]+)]\s*$", line)
        if match:
            headers.append((match.group(1).strip().lower(), index))
    return {
        name: (start, headers[index + 1][1] if index + 1 < len(headers) else len(lines))
        for index, (name, start) in enumerate(headers)
    }
