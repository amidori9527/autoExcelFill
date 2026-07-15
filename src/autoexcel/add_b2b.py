from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
import os
from pathlib import Path
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
)

from autoexcel.fast_xlsx import (
    REL_NS,
    _cell_at,
    _force_full_calculation,
    _row_number,
    _sheet_target_path,
    _tag,
)
from autoexcel.main import choose_workbook_from_current_directory, pause_before_exit


TARGET_SHEET_NAME = "提取B2B"
DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
TIME_RE = re.compile(r"\d{2}:\d{2}:\d{2}")
RANGE_RE = re.compile(r"([A-Z]+)(\d+):([A-Z]+)(\d+)")


@dataclass(frozen=True)
class InputField:
    value: str
    token_start: int
    token_count: int = 1


@dataclass(frozen=True)
class ParsedInputLine:
    line_number: int
    raw: str
    fields: tuple[InputField, ...]


@dataclass(frozen=True)
class FieldMapping:
    date_time: int
    trx_id: int
    outgoing_card: int
    amount: int


@dataclass(frozen=True)
class B2BRecord:
    raw: str
    date_time: str
    extraction_date: str
    trx_id: str
    outgoing_card: str
    amount: Decimal
    amount_was_negative: bool


@dataclass(frozen=True)
class AppendB2BResult:
    inserted_count: int
    start_row: int
    end_row: int
    converted_negative_count: int


def _parse_date_time(value: str) -> datetime:
    try:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    except ValueError as error:
        raise ValueError("日期时间必须是 YYYY-MM-DD HH:MM:SS 格式") from error


def split_input_line(raw: str, line_number: int = 1) -> ParsedInputLine:
    text = raw.strip()
    if not text:
        raise ValueError(f"第 {line_number} 行为空")

    tokens = text.split()
    date_time_starts = [
        index
        for index in range(len(tokens) - 1)
        if DATE_RE.fullmatch(tokens[index]) and TIME_RE.fullmatch(tokens[index + 1])
    ]
    if len(date_time_starts) != 1:
        raise ValueError(
            f"第 {line_number} 行必须且只能包含一个 YYYY-MM-DD HH:MM:SS 日期时间"
        )

    date_time_start = date_time_starts[0]
    fields: list[InputField] = []
    token_index = 0
    while token_index < len(tokens):
        if token_index == date_time_start:
            fields.append(
                InputField(
                    value=f"{tokens[token_index]} {tokens[token_index + 1]}",
                    token_start=token_index + 1,
                    token_count=2,
                )
            )
            token_index += 2
            continue
        fields.append(InputField(value=tokens[token_index], token_start=token_index + 1))
        token_index += 1

    return ParsedInputLine(line_number=line_number, raw=text, fields=tuple(fields))


def parse_input_text(text: str) -> list[ParsedInputLine]:
    parsed = [
        split_input_line(raw, line_number)
        for line_number, raw in enumerate(text.splitlines(), start=1)
        if raw.strip()
    ]
    if not parsed:
        raise ValueError("没有输入任何数据")
    return parsed


def _decimal_or_none(value: str) -> Decimal | None:
    try:
        number = Decimal(value)
    except InvalidOperation:
        return None
    return number if number.is_finite() else None


def guess_field_mapping(fields: tuple[InputField, ...]) -> FieldMapping:
    date_time = next(
        (index for index, field in enumerate(fields) if field.token_count == 2),
        0,
    )
    amount_candidates = [
        index for index, field in enumerate(fields) if _decimal_or_none(field.value) is not None
    ]
    amount = amount_candidates[-1] if amount_candidates else len(fields) - 1

    card_candidates = [
        index
        for index, field in enumerate(fields)
        if index != amount and field.value.isdigit()
    ]
    outgoing_card = card_candidates[0] if card_candidates else 0
    trx_candidates = [
        index
        for index, field in enumerate(fields)
        if index not in {date_time, amount, outgoing_card}
        and any(character.isalpha() for character in field.value)
        and any(character.isdigit() for character in field.value)
    ]
    trx_id = trx_candidates[0] if trx_candidates else 0
    return FieldMapping(date_time, trx_id, outgoing_card, amount)


def _mapping_indexes(mapping: FieldMapping) -> tuple[int, ...]:
    return (
        mapping.date_time,
        mapping.trx_id,
        mapping.outgoing_card,
        mapping.amount,
    )


def validate_and_build_records(
    lines: list[ParsedInputLine],
    mapping: FieldMapping,
) -> list[B2BRecord]:
    if not lines:
        raise ValueError("没有输入任何数据")

    field_count = len(lines[0].fields)
    indexes = _mapping_indexes(mapping)
    if any(index < 0 or index >= field_count for index in indexes):
        raise ValueError("字段映射超出第一行的字段范围")
    if len(set(indexes)) != len(indexes):
        raise ValueError(
            "日期时间、TRXID、转出卡号和金额不能选择同一个字段"
        )

    sample_spans = [
        (lines[0].fields[index].token_start, lines[0].fields[index].token_count)
        for index in indexes
    ]
    records: list[B2BRecord] = []
    errors: list[str] = []
    for line in lines:
        if len(line.fields) != field_count:
            errors.append(
                f"第 {line.line_number} 行：字段数为 {len(line.fields)}，"
                f"第一行为 {field_count}"
            )
            continue
        spans = [
            (line.fields[index].token_start, line.fields[index].token_count)
            for index in indexes
        ]
        if spans != sample_spans:
            errors.append(f"第 {line.line_number} 行：字段顺序与第一行不一致")
            continue

        try:
            parsed_date_time = _parse_date_time(line.fields[mapping.date_time].value)
        except ValueError as error:
            errors.append(f"第 {line.line_number} 行：{error}")
            continue

        amount_text = line.fields[mapping.amount].value
        amount = _decimal_or_none(amount_text)
        if amount is None:
            errors.append(f"第 {line.line_number} 行：金额“{amount_text}”不是数字")
            continue
        if amount == 0:
            errors.append(
                f"第 {line.line_number} 行：金额必须大于 0，当前为 {amount_text}"
            )
            continue
        amount_was_negative = amount < 0
        amount = abs(amount)

        outgoing_card = line.fields[mapping.outgoing_card].value
        if not outgoing_card.isdigit():
            errors.append(
                f"第 {line.line_number} 行：转出卡号“{outgoing_card}”不是纯数字"
            )
            continue
        records.append(
            B2BRecord(
                raw=line.raw,
                date_time=parsed_date_time.strftime("%Y-%m-%d %H:%M:%S"),
                extraction_date=parsed_date_time.strftime("%Y%m%d"),
                trx_id=line.fields[mapping.trx_id].value,
                outgoing_card=outgoing_card,
                amount=amount,
                amount_was_negative=amount_was_negative,
            )
        )

    if errors:
        shown = errors[:20]
        if len(errors) > len(shown):
            shown.append(f"另有 {len(errors) - len(shown)} 行错误未显示")
        raise ValueError("输入校验失败，未写入 Excel：\n" + "\n".join(shown))
    return records


def _format_decimal(value: Decimal) -> str:
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f")


def _field_formula(row_number: int, field: InputField, total_tokens: int) -> str:
    expression = f"TRIM(A{row_number})"
    if field.token_start > 1:
        expression = (
            f'_xlfn.TEXTAFTER({expression}," ",{field.token_start - 1})'
        )
    field_end = field.token_start + field.token_count - 1
    remaining_tokens = total_tokens - field.token_start + 1
    if field.token_count < remaining_tokens:
        if field.token_count == 1:
            expression = f'_xlfn.TEXTBEFORE({expression}," ")'
        else:
            expression = f'_xlfn.TEXTBEFORE({expression}," ",{field.token_count})'
    if field_end > total_tokens:
        raise ValueError("字段位置超出原始数据范围")
    return expression


def _apply_style(cell: ET.Element, style: str | None) -> None:
    if style is not None:
        cell.set("s", style)


def _inline_string_cell(ref: str, value: str, style: str | None) -> ET.Element:
    cell = ET.Element(_tag("c"), {"r": ref, "t": "inlineStr"})
    _apply_style(cell, style)
    inline = ET.SubElement(cell, _tag("is"))
    text = ET.SubElement(inline, _tag("t"))
    if value != value.strip() or "  " in value:
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value
    return cell


def _formula_cell(
    ref: str,
    formula: str,
    cached_value: str,
    style: str | None,
    string_result: bool = False,
) -> ET.Element:
    attributes = {"r": ref}
    if string_result:
        attributes["t"] = "str"
    cell = ET.Element(_tag("c"), attributes)
    _apply_style(cell, style)
    ET.SubElement(cell, _tag("f")).text = formula
    ET.SubElement(cell, _tag("v")).text = cached_value
    return cell


def _number_cell(ref: str, value: Decimal, style: str | None) -> ET.Element:
    cell = ET.Element(_tag("c"), {"r": ref})
    _apply_style(cell, style)
    ET.SubElement(cell, _tag("v")).text = _format_decimal(value)
    return cell


def _column_styles(rows: list[ET.Element], columns: tuple[str, ...]) -> dict[str, str | None]:
    styles: dict[str, str | None] = {}
    recent_rows = rows[-500:]
    for column in columns:
        for row in reversed(recent_rows):
            cell = _cell_at(row, column)
            if cell is not None:
                styles[column] = cell.attrib.get("s")
                break
        if column in styles:
            continue
        counts = Counter(
            cell.attrib["s"]
            for row in rows
            if (cell := _cell_at(row, column)) is not None and "s" in cell.attrib
        )
        styles[column] = counts.most_common(1)[0][0] if counts else None
    return styles


def _extend_range(reference: str, end_row: int) -> str:
    match = RANGE_RE.fullmatch(reference)
    if not match:
        return reference
    start_column, start_row, end_column, _ = match.groups()
    return f"{start_column}{start_row}:{end_column}{end_row}"


def _append_sheet_xml(
    sheet_xml: bytes,
    records: list[B2BRecord],
    sample_fields: tuple[InputField, ...],
    mapping: FieldMapping,
) -> tuple[bytes, int, int]:
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find(_tag("sheetData"))
    if sheet_data is None:
        raise ValueError(f'工作表“{TARGET_SHEET_NAME}”缺少 sheetData')
    rows = sheet_data.findall(_tag("row"))
    if not rows:
        raise ValueError(f'工作表“{TARGET_SHEET_NAME}”没有表头')

    last_row_number = max(_row_number(row) for row in rows)
    start_row = last_row_number + 1
    styles = _column_styles(rows, ("A", "B", "C", "D", "E", "F", "I"))
    total_tokens = sum(field.token_count for field in sample_fields)
    date_time_field = sample_fields[mapping.date_time]
    trx_field = sample_fields[mapping.trx_id]
    outgoing_field = sample_fields[mapping.outgoing_card]

    for offset, record in enumerate(records):
        row_number = start_row + offset
        row = ET.Element(_tag("row"), {"r": str(row_number)})
        row.append(_inline_string_cell(f"A{row_number}", record.raw, styles["A"]))
        row.append(
            _formula_cell(
                f"B{row_number}",
                _field_formula(row_number, date_time_field, total_tokens),
                record.date_time,
                styles["B"],
                string_result=True,
            )
        )
        row.append(
            _formula_cell(
                f"C{row_number}",
                f'SUBSTITUTE(LEFT(B{row_number},10),"-","")',
                record.extraction_date,
                styles["C"],
                string_result=True,
            )
        )
        row.append(
            _formula_cell(
                f"D{row_number}",
                _field_formula(row_number, outgoing_field, total_tokens),
                record.outgoing_card,
                styles["D"],
                string_result=True,
            )
        )
        row.append(_number_cell(f"E{row_number}", record.amount, styles["E"]))
        commission = record.amount * Decimal("0.004")
        row.append(
            _formula_cell(
                f"F{row_number}",
                f"E{row_number}*0.004",
                _format_decimal(commission),
                styles["F"],
            )
        )
        row.append(
            _formula_cell(
                f"I{row_number}",
                _field_formula(row_number, trx_field, total_tokens),
                record.trx_id,
                styles["I"],
                string_result=True,
            )
        )
        sheet_data.append(row)

    end_row = start_row + len(records) - 1
    dimension = root.find(_tag("dimension"))
    if dimension is not None and dimension.attrib.get("ref"):
        dimension.set("ref", _extend_range(dimension.attrib["ref"], end_row))
    auto_filter = root.find(_tag("autoFilter"))
    if auto_filter is not None and auto_filter.attrib.get("ref"):
        auto_filter.set("ref", _extend_range(auto_filter.attrib["ref"], end_row))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), start_row, end_row


def append_b2b_to_workbook(
    xlsx_path: Path,
    lines: list[ParsedInputLine],
    mapping: FieldMapping,
) -> AppendB2BResult:
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Excel 文件不存在：{xlsx_path}")
    if xlsx_path.suffix.lower() != ".xlsx":
        raise ValueError("请选择 .xlsx 文件")

    records = validate_and_build_records(lines, mapping)
    temporary_path: Path | None = None
    with zipfile.ZipFile(xlsx_path, "r") as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {relation.attrib["Id"]: relation.attrib["Target"] for relation in rels}
        sheets = workbook.find(_tag("sheets"))
        if sheets is None:
            raise ValueError("Excel 文件没有工作表")
        rel_id_attribute = f"{{{REL_NS}}}id"
        target_sheet = next(
            (
                sheet
                for sheet in sheets.findall(_tag("sheet"))
                if sheet.attrib["name"] == TARGET_SHEET_NAME
            ),
            None,
        )
        if target_sheet is None:
            raise ValueError(f'Excel 中没有“{TARGET_SHEET_NAME}”工作表')
        sheet_path = _sheet_target_path(relationships[target_sheet.attrib[rel_id_attribute]])
        updated_sheet, start_row, end_row = _append_sheet_xml(
            archive.read(sheet_path), records, lines[0].fields, mapping
        )
        updated_parts = {
            sheet_path: updated_sheet,
            "xl/workbook.xml": _force_full_calculation(archive.read("xl/workbook.xml")),
        }

        file_descriptor, temporary_name = tempfile.mkstemp(
            prefix=f"{xlsx_path.stem}.", suffix=".tmp.xlsx", dir=xlsx_path.parent
        )
        os.close(file_descriptor)
        temporary_path = Path(temporary_name)
        try:
            with zipfile.ZipFile(
                temporary_path, "w", compression=zipfile.ZIP_DEFLATED
            ) as output:
                for item in archive.infolist():
                    content = (
                        updated_parts[item.filename]
                        if item.filename in updated_parts
                        else archive.read(item.filename)
                    )
                    output.writestr(item, content)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise

    if temporary_path is None:
        raise RuntimeError("没有生成临时 Excel 文件")
    try:
        os.replace(temporary_path, xlsx_path)
    finally:
        temporary_path.unlink(missing_ok=True)
    return AppendB2BResult(
        len(records),
        start_row,
        end_row,
        sum(record.amount_was_negative for record in records),
    )


class B2BInputDialog(QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("新增 B2B 数据")
        self.setMinimumSize(760, 420)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("粘贴多行数据，按 Enter 提交："))
        self.input = QPlainTextEdit()
        self.input.setPlaceholderText(
            "2026-07-12 22:04:23  75NVDJEI  01635548053  01850801086  50000"
        )
        self.input.installEventFilter(self)
        layout.addWidget(self.input)
        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        submit = QPushButton("下一步")
        submit.clicked.connect(self.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(submit)
        layout.addLayout(buttons)

    def eventFilter(self, watched, event) -> bool:
        if (
            watched is self.input
            and event.type() == QEvent.KeyPress
            and event.key() in (Qt.Key_Return, Qt.Key_Enter)
            and not event.modifiers()
        ):
            self.accept()
            return True
        return super().eventFilter(watched, event)

    def text(self) -> str:
        return self.input.toPlainText()


class FieldMappingDialog(QDialog):
    def __init__(self, fields: tuple[InputField, ...], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("确认 B2B 字段对应关系")
        self.setMinimumWidth(680)
        layout = QVBoxLayout(self)
        layout.addWidget(
            QLabel("以下内容来自第一条数据，请确认每个值要写入的列：")
        )
        for index, field in enumerate(fields, start=1):
            layout.addWidget(QLabel(f"字段 {index}：{field.value}"))

        defaults = guess_field_mapping(fields)
        options = [f"字段 {index}：{field.value}" for index, field in enumerate(fields, start=1)]
        form = QFormLayout()
        self.selectors: dict[str, QComboBox] = {}
        labels_and_defaults = (
            ("date_time", "日期时间（B、C列）", defaults.date_time),
            ("trx_id", "TRXID（I列）", defaults.trx_id),
            ("outgoing_card", "转出卡号（D列）", defaults.outgoing_card),
            ("amount", "金额（E列）", defaults.amount),
        )
        for key, label, default in labels_and_defaults:
            selector = QComboBox()
            selector.addItems(options)
            selector.setCurrentIndex(default)
            self.selectors[key] = selector
            form.addRow(label, selector)
        layout.addLayout(form)

        buttons = QHBoxLayout()
        buttons.addStretch()
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        confirm = QPushButton("确认并写入")
        confirm.clicked.connect(self.accept)
        buttons.addWidget(cancel)
        buttons.addWidget(confirm)
        layout.addLayout(buttons)

    def mapping(self) -> FieldMapping:
        return FieldMapping(
            date_time=self.selectors["date_time"].currentIndex(),
            trx_id=self.selectors["trx_id"].currentIndex(),
            outgoing_card=self.selectors["outgoing_card"].currentIndex(),
            amount=self.selectors["amount"].currentIndex(),
        )

    def accept(self) -> None:
        if len(set(_mapping_indexes(self.mapping()))) != 4:
            QMessageBox.warning(
                self,
                "字段重复",
                "四个目标字段必须分别选择不同的示例值。",
            )
            return
        super().accept()


def main() -> int:
    try:
        print("Excel 新增 B2B 数据工具")
        print("----------------------------------------")
        selected_file = choose_workbook_from_current_directory()
        app = QApplication.instance() or QApplication(sys.argv)
        input_dialog = B2BInputDialog()
        input_dialog.input.setFocus()
        if input_dialog.exec() != QDialog.Accepted:
            return 0
        lines = parse_input_text(input_dialog.text())
        mapping_dialog = FieldMappingDialog(lines[0].fields)
        if mapping_dialog.exec() != QDialog.Accepted:
            return 0
        result = append_b2b_to_workbook(selected_file, lines, mapping_dialog.mapping())
    except Exception as error:
        print(f"新增 B2B 数据失败：{error}")
        return 1
    else:
        print(f"已写入：{selected_file}")
        print(
            f"新增 {result.inserted_count} 行，位置："
            f"{TARGET_SHEET_NAME}!A{result.start_row}:I{result.end_row}"
        )
        if result.converted_negative_count:
            print(
                f"提示：发现 {result.converted_negative_count} 行负金额，"
                "已自动转为正数。"
            )
        return 0
    finally:
        pause_before_exit()


if __name__ == "__main__":
    raise SystemExit(main())
