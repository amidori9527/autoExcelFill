from __future__ import annotations

import re
import sys
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from threading import Event
from typing import Callable

from PySide6.QtCore import QByteArray, QDate, QLocale, QObject, QRect, QRectF, QRunnable, QSize, QThreadPool, Qt, QUrl, Signal, Slot
from PySide6.QtGui import QCursor, QDesktopServices, QFont, QIcon, QMovie, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QAbstractSpinBox,
    QApplication,
    QCalendarWidget,
    QCheckBox,
    QComboBox,
    QDialog,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLayout,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from autoexcel.add_b2b import FieldMapping, guess_field_mapping, parse_input_text
from autoexcel.config_editor import (
    editable_config_path,
    read_fill_limit_sheets,
    read_ini,
    update_ini,
)
from autoexcel.diff_orders import get_result_dir
from autoexcel.fetch_orders import get_login_config_path
from autoexcel.gui_tasks import (
    TaskResult,
    run_add_b2b_task,
    run_add_cards_task,
    run_diff_files_task,
    run_diff_task,
    run_fetch_task,
    run_fill_task,
    run_full_flow_sync_task,
    run_payout_diff_files_task,
    run_payout_diff_task,
    run_tp_collection_sync_task,
    run_tp_payout_sync_task,
    run_wallet_flow_sync_task,
)
from autoexcel.license import (
    FEATURE_ADD_B2B,
    FEATURE_ADD_CARDS,
    FEATURE_FETCH_ORDERS,
    FEATURE_ORDER_DIFF,
    LicenseInfo,
    install_license,
    license_file_path,
    load_license,
    remove_license,
)
from autoexcel.main import default_target_date
from autoexcel.poetry import poetry_line_for
from autoexcel.runtime_paths import (
    bundled_resource,
    ensure_workspace_directories,
    workspace_directory,
)
from autoexcel.version import VERSION


APP_STYLE = """
QWidget {
    color: #172033;
    font-family: "Inter", "SF Pro Text", "PingFang SC", "Microsoft YaHei UI", "Microsoft YaHei", Arial;
    font-size: 13px;
}
QMainWindow, QWidget#appRoot, QScrollArea#pageScroll > QWidget > QWidget {
    background: #f6f7fb;
}
QFrame#sidebar {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #111827, stop:1 #1e293b);
    border: none;
}
QLabel#sidebarBrand {
    color: #101828;
    font-family: "Avenir Next", "SF Pro Display", "PingFang SC", "Microsoft YaHei UI";
    font-size: 17px;
    font-weight: 700;
}
QLabel#sidebarCaption { color: #7f8ca3; font-size: 10px; }
QLabel#sidebarBrandCaption {
    color: #7f8ca3;
    font-family: "Songti SC", "STSong", "Noto Serif CJK SC", "SimSun", "PingFang SC";
    font-size: 11px;
    font-weight: 500;
    padding: 0 2px;
}
QLabel#navSection {
    color: #7f8ca3;
    font-size: 9px;
    font-weight: 700;
    padding: 4px 10px 2px 10px;
}
QPushButton#navButton {
    min-height: 40px;
    padding: 0 13px;
    text-align: left;
    color: #aeb8ca;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 10px;
}
QPushButton#navButton:hover {
    color: #101828;
    background: #1e293b;
    border-color: #1e293b;
}
QPushButton#navButton:checked {
    color: #fefefe;
    background: #2f5bea;
    border-color: #2f5bea;
    font-weight: 700;
}
QLabel#pageEyebrow {
    color: #2f5bea;
    font-size: 10px;
    font-weight: 700;
}
QLabel#pageTitle, QLabel#heroTitle {
    color: #101828;
    font-family: "Avenir Next", "SF Pro Display", "PingFang SC", "Microsoft YaHei UI";
    font-weight: 700;
}
QLabel#pageTitle { font-size: 27px; }
QLabel#heroTitle { font-size: 30px; }
QLabel#heroDescription { color: #667085; font-size: 13px; }
QLabel#heroBrandIcon {
    background: transparent;
    border: none;
    padding: 0;
}
QLabel#sectionTitle {
    color: #101828;
    font-family: "Avenir Next", "SF Pro Display", "PingFang SC", "Microsoft YaHei UI";
    font-size: 16px;
    font-weight: 700;
}
QLabel#cardTitle, QLabel#featuredCardTitle {
    color: #101828;
    font-family: "Avenir Next", "SF Pro Display", "PingFang SC", "Microsoft YaHei UI";
    font-weight: 700;
}
QLabel#cardTitle { font-size: 16px; }
QLabel#featuredCardTitle { font-size: 18px; }
QLabel#cardDescription { color: #7a8699; font-size: 12px; }
QLabel#cardAction {
    color: #3157d5;
    background: #edf2ff;
    border-radius: 8px;
    padding: 4px 9px;
    font-size: 10px;
    font-weight: 700;
}
QLabel#fieldLabel { color: #344054; font-size: 12px; font-weight: 700; }
QLabel#muted, QLabel#fieldHint { color: #7a8699; }
QLabel#fieldHint { font-size: 11px; }
QLabel#statusPill {
    color: #3157d5;
    background: #edf2ff;
    border-radius: 9px;
    padding: 3px 8px;
    font-size: 9px;
    font-weight: 700;
}
QFrame#panel, QFrame#homeCard, QFrame#homeFeaturedCard, QFrame#settingSection {
    background: #ffffff;
    border: 1px solid #e5e8ef;
    border-radius: 16px;
}
QFrame#homeHero {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f8faff, stop:1 #edf2ff);
    border: 1px solid #cdd7fb;
    border-radius: 20px;
}
QFrame#homeFeaturedCard {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f4f7ff, stop:1 #ffffff);
    border-color: #cdd7fb;
}
QFrame#licensePanel {
    background: #f8faff;
    border: 1px solid #cdd7fb;
    border-radius: 16px;
}
QLabel#licenseValid { color: #067647; font-weight: 700; }
QLabel#licenseInvalid { color: #b42318; font-weight: 700; }
QFrame#homeCard:hover { border-color: #aebcf3; background: #fbfcff; }
QFrame#homeFeaturedCard:hover { border-color: #9eafe9; background: #f4f7ff; }
QFrame#resultRow {
    background: #fbfcff;
    border: 1px solid #e5e8ef;
    border-radius: 12px;
}
QFrame#resultRow:hover { background: #f7f9ff; border-color: #b8c5f4; }
QFrame#emptyResult {
    background: #fbfcff;
    border: 1px dashed #d7dce5;
    border-radius: 12px;
}
QLabel#resultTitle { color: #172033; font-size: 14px; font-weight: 700; }
QLabel#resultBadgeCollection {
    color: #067647; background: #ecfdf3; border-radius: 9px;
    padding: 4px 10px; font-size: 11px; font-weight: 700;
}
QLabel#resultBadgePayout {
    color: #3157d5; background: #edf2ff; border-radius: 9px;
    padding: 4px 10px; font-size: 11px; font-weight: 700;
}
QPushButton#resultViewButton {
    color: #3157d5; background: #f4f7ff; border-color: #cdd7fb; font-weight: 700;
}
QPushButton#resultViewButton:hover { color: #244ac7; background: #e9efff; border-color: #9eafe9; }
QScrollArea#resultScroll, QWidget#resultList { background: transparent; }
QLabel#cardIcon {
    color: #2f5bea;
    background: #edf2ff;
    border: 1px solid #cdd7fb;
    border-radius: 12px;
    font-size: 21px; font-weight: 700; qproperty-alignment: AlignCenter;
}
QPushButton {
    min-height: 36px;
    border-radius: 9px;
    padding: 0 15px;
    border: 1px solid #d7dce5;
    background: #ffffff;
    color: #344054;
}
QPushButton:hover { background: #f8faff; border-color: #9eafe9; }
QPushButton#primary { background: #2f5bea; color: #fefefe; border: none; font-weight: 700; }
QPushButton#primary:hover { background: #244ac7; }
QPushButton#ghost { background: transparent; border: none; color: #667085; }
QPushButton#ghost:hover { background: #eef1f6; color: #344054; }
QPushButton#danger { color: #c4322b; border-color: #f0b9b5; background: #fffafa; }
QPushButton:disabled { color: #98a2b3; background: #f0f2f5; border-color: #e5e7eb; }
QPushButton#modeToggle {
    min-width: 30px;
    max-width: 30px;
    min-height: 30px;
    max-height: 30px;
    padding: 0;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 9px;
}
QPushButton#modeToggle:hover {
    background: #eef1f6;
    border-color: #d7dce5;
}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
    min-height: 38px;
    border: 1px solid #d7dce5;
    border-radius: 9px;
    padding: 0 11px;
    background: #ffffff;
    selection-background-color: #2f5bea;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {
    border-color: #2f5bea;
    background: #fbfcff;
}
QComboBox::drop-down { border: none; width: 28px; }
QCheckBox { spacing: 9px; color: #344054; }
QCheckBox::indicator { width: 34px; height: 18px; border-radius: 9px; background: #cbd2de; }
QCheckBox::indicator:checked { background: #2f5bea; image: none; }
QPlainTextEdit {
    background: #0f172a; color: #cbd5e1; border: none; border-radius: 8px;
    padding: 10px; font-family: Menlo, Consolas, monospace; font-size: 11px;
}
QProgressBar { border: none; background: #e8ebf1; border-radius: 3px; height: 6px; }
QProgressBar::chunk { background: #2f5bea; border-radius: 3px; }
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical {
    width: 9px;
    margin: 4px 2px 4px 2px;
    background: transparent;
}
QScrollBar::handle:vertical {
    min-height: 36px;
    background: #cbd2de;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover { background: #aeb8ca; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
QDialog#calendarDialog { background: #ffffff; }
QCalendarWidget QWidget { alternate-background-color: #ffffff; }
QCalendarWidget QTableView {
    background: #ffffff; border: none; selection-background-color: #2f5bea;
    selection-color: #fefefe; outline: none;
}
"""


DEFAULT_THEME = "indigo"
DEFAULT_MODE = "light"
THEME_LABELS = {
    "emerald": "翡翠绿",
    "indigo": "靛青蓝",
}
THEME_PALETTES = {
    "emerald": {
        "text": "#24312c",
        "canvas": "#f4f6f3",
        "sidebar": "#eef4f1",
        "sidebar_muted": "#768d82",
        "muted_dark": "#697870",
        "sidebar_text": "#4f655b",
        "sidebar_hover": "#dde9e2",
        "primary": "#3f8b6c",
        "title": "#18231f",
        "body": "#394a43",
        "muted": "#75837d",
        "primary_text": "#34775d",
        "primary_soft": "#eaf4ef",
        "border": "#dfe6e2",
        "soft_surface": "#f7faf8",
        "soft_border": "#cadfd5",
        "hover_border": "#a8ccbc",
        "hover_surface": "#fbfdfc",
        "result_hover": "#f3f8f5",
        "result_hover_border": "#adccbf",
        "control_border": "#d5ddd9",
        "primary_button_soft": "#f0f7f3",
        "primary_hover": "#307456",
        "primary_soft_hover": "#e2f0e9",
        "control_hover_border": "#9fc4b4",
        "ghost_hover": "#eaf0ed",
        "disabled_text": "#9ba7a1",
        "disabled_bg": "#eef1ef",
        "check_off": "#c9d2ce",
        "log_text": "#cbd7d2",
        "log_bg": "#17251f",
        "progress_bg": "#e2e8e5",
    },
    "indigo": {
        "text": "#242936",
        "canvas": "#f4f5f7",
        "sidebar": "#f1f3f8",
        "sidebar_muted": "#7c869c",
        "muted_dark": "#697184",
        "sidebar_text": "#596277",
        "sidebar_hover": "#e2e6f0",
        "primary": "#6673d9",
        "title": "#181b24",
        "body": "#3d4351",
        "muted": "#7c8495",
        "primary_text": "#5865c7",
        "primary_soft": "#eef0ff",
        "border": "#e1e4e9",
        "soft_surface": "#f8f8fc",
        "soft_border": "#d7dcf3",
        "hover_border": "#bac3ec",
        "hover_surface": "#fcfcfe",
        "result_hover": "#f6f7fc",
        "result_hover_border": "#c1c8e9",
        "control_border": "#d9dde5",
        "primary_button_soft": "#f3f4fc",
        "primary_hover": "#5561c5",
        "primary_soft_hover": "#e7eaff",
        "control_hover_border": "#abb4df",
        "ghost_hover": "#eceef3",
        "disabled_text": "#9ca3b1",
        "disabled_bg": "#eff1f4",
        "check_off": "#cbd0da",
        "log_text": "#cdd3df",
        "log_bg": "#202638",
        "progress_bg": "#e5e7ec",
    },
}
LIGHT_COMMON_COLORS = {
    "surface": "#ffffff",
    "on_primary": "#ffffff",
    "success_text": "#067647",
    "success_surface": "#ecfdf3",
    "error_text": "#b42318",
    "danger_text": "#c4322b",
    "danger_border": "#f0b9b5",
    "danger_surface": "#fffafa",
    "disabled_border": "#e5e7eb",
}
for palette in THEME_PALETTES.values():
    palette.update(LIGHT_COMMON_COLORS)

DARK_THEME_PALETTES = {
    "emerald": {
        "text": "#e4ece8",
        "canvas": "#101713",
        "sidebar": "#151f1a",
        "sidebar_muted": "#8ca096",
        "muted_dark": "#9aaba3",
        "sidebar_text": "#c0cec7",
        "sidebar_hover": "#223129",
        "primary": "#56b58c",
        "title": "#f1f7f4",
        "body": "#d2ddd7",
        "muted": "#94a69d",
        "primary_text": "#86d3b0",
        "primary_soft": "#1f3a2e",
        "border": "#293a32",
        "soft_surface": "#17231d",
        "soft_border": "#355244",
        "hover_border": "#4f725f",
        "hover_surface": "#1b2821",
        "result_hover": "#1d2c24",
        "result_hover_border": "#486a58",
        "control_border": "#384a41",
        "primary_button_soft": "#1d3127",
        "primary_hover": "#45a27a",
        "primary_soft_hover": "#284737",
        "control_hover_border": "#5a7c69",
        "ghost_hover": "#243129",
        "disabled_text": "#718078",
        "disabled_bg": "#222c27",
        "check_off": "#4c5d54",
        "log_text": "#d1ddd7",
        "log_bg": "#090f0c",
        "progress_bg": "#2a3931",
        "surface": "#17201c",
        "on_primary": "#ffffff",
        "success_text": "#68d5a0",
        "success_surface": "#173629",
        "error_text": "#ff938a",
        "danger_text": "#ff938a",
        "danger_border": "#714047",
        "danger_surface": "#301f23",
        "disabled_border": "#35413b",
    },
    "indigo": {
        "text": "#e7eaf2",
        "canvas": "#111522",
        "sidebar": "#171c2a",
        "sidebar_muted": "#8993aa",
        "muted_dark": "#9ca5b8",
        "sidebar_text": "#c5cbd9",
        "sidebar_hover": "#232a3d",
        "primary": "#7c88f2",
        "title": "#f5f7fb",
        "body": "#d5dae5",
        "muted": "#9aa4b8",
        "primary_text": "#aeb7ff",
        "primary_soft": "#282f50",
        "border": "#2d3548",
        "soft_surface": "#181d2b",
        "soft_border": "#3b4668",
        "hover_border": "#56658d",
        "hover_surface": "#1d2332",
        "result_hover": "#202637",
        "result_hover_border": "#4d5b80",
        "control_border": "#3a4357",
        "primary_button_soft": "#202641",
        "primary_hover": "#6976e2",
        "primary_soft_hover": "#30395f",
        "control_hover_border": "#59698f",
        "ghost_hover": "#252c3c",
        "disabled_text": "#717b90",
        "disabled_bg": "#252a36",
        "check_off": "#505a6e",
        "log_text": "#d3d9e7",
        "log_bg": "#0b0f18",
        "progress_bg": "#2c3445",
        "surface": "#181d2a",
        "on_primary": "#ffffff",
        "success_text": "#5ed39b",
        "success_surface": "#173b2c",
        "error_text": "#ff8a80",
        "danger_text": "#ff8a80",
        "danger_border": "#6e3b42",
        "danger_surface": "#2f1d22",
        "disabled_border": "#343b4a",
    },
}
STYLE_COLOR_TOKENS = {
    "#172033": "text",
    "#f6f7fb": "canvas",
    "#111827": "sidebar",
    "#7f8ca3": "sidebar_muted",
    "#667085": "muted_dark",
    "#aeb8ca": "sidebar_text",
    "#1e293b": "sidebar_hover",
    "#2f5bea": "primary",
    "#101828": "title",
    "#344054": "body",
    "#7a8699": "muted",
    "#3157d5": "primary_text",
    "#edf2ff": "primary_soft",
    "#e5e8ef": "border",
    "#f8faff": "soft_surface",
    "#cdd7fb": "soft_border",
    "#aebcf3": "hover_border",
    "#fbfcff": "hover_surface",
    "#f7f9ff": "result_hover",
    "#b8c5f4": "result_hover_border",
    "#d7dce5": "control_border",
    "#f4f7ff": "primary_button_soft",
    "#244ac7": "primary_hover",
    "#e9efff": "primary_soft_hover",
    "#9eafe9": "control_hover_border",
    "#eef1f6": "ghost_hover",
    "#98a2b3": "disabled_text",
    "#f0f2f5": "disabled_bg",
    "#cbd2de": "check_off",
    "#cbd5e1": "log_text",
    "#0f172a": "log_bg",
    "#e8ebf1": "progress_bg",
    "#ffffff": "surface",
    "#fefefe": "on_primary",
    "#067647": "success_text",
    "#ecfdf3": "success_surface",
    "#b42318": "error_text",
    "#c4322b": "danger_text",
    "#f0b9b5": "danger_border",
    "#fffafa": "danger_surface",
    "#e5e7eb": "disabled_border",
}


def normalize_theme_name(value: str) -> str:
    normalized = value.strip().lower()
    return normalized if normalized in THEME_PALETTES else DEFAULT_THEME


def normalize_ui_mode(value: str) -> str:
    normalized = value.strip().lower()
    return normalized if normalized in {"light", "dark"} else DEFAULT_MODE


def theme_palette(theme_name: str, mode: str = DEFAULT_MODE) -> dict[str, str]:
    palettes = DARK_THEME_PALETTES if normalize_ui_mode(mode) == "dark" else THEME_PALETTES
    return palettes[normalize_theme_name(theme_name)]


def build_app_style(theme_name: str, mode: str = DEFAULT_MODE) -> str:
    style = APP_STYLE
    palette = theme_palette(theme_name, mode)
    for color, token in STYLE_COLOR_TOKENS.items():
        style = style.replace(color, f"__THEME_{token.upper()}__")
    for token, color in palette.items():
        style = style.replace(f"__THEME_{token.upper()}__", color)
    return style


def theme_name_from_config(path: Path) -> str:
    parser = read_ini(path)
    return normalize_theme_name(parser.get("ui", "theme", fallback=DEFAULT_THEME))


def ui_mode_from_config(path: Path) -> str:
    parser = read_ini(path)
    return normalize_ui_mode(parser.get("ui", "mode", fallback=DEFAULT_MODE))


def icon_resource_path(name: str) -> Path:
    resource_name = f"icon/sidebar/{name}.svg"
    path = bundled_resource(resource_name)
    if path is not None:
        return path
    return Path(__file__).resolve().parents[2] / resource_name


def brand_icon_path() -> Path:
    resource_name = "icon/cover-v5.png"
    path = bundled_resource(resource_name)
    if path is not None:
        return path
    return Path(__file__).resolve().parents[2] / resource_name


def hero_animation_path() -> Path:
    resource_name = "icon/baby-rabbit.webp"
    path = bundled_resource(resource_name)
    if path is not None:
        return path
    return Path(__file__).resolve().parents[2] / resource_name


def themed_svg_pixmap(name: str, color: str, size: QSize) -> QPixmap:
    path = icon_resource_path(name)
    if not path.is_file():
        return QPixmap()
    svg_text = path.read_text(encoding="utf-8")
    svg_text = re.sub(
        r'stroke="#[0-9A-Fa-f]{6}"',
        f'stroke="{color}"',
        svg_text,
    )
    svg_text = re.sub(r'stroke-width="[^"]+"', 'stroke-width="2"', svg_text)
    renderer = QSvgRenderer(QByteArray(svg_text.encode("utf-8")))
    if not renderer.isValid():
        return QPixmap()

    app = QApplication.instance()
    screen = app.primaryScreen() if app is not None else None
    device_pixel_ratio = max(1.0, float(screen.devicePixelRatio()) if screen else 1.0)
    pixel_size = QSize(
        max(1, round(size.width() * device_pixel_ratio)),
        max(1, round(size.height() * device_pixel_ratio)),
    )
    pixmap = QPixmap(pixel_size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(
        painter,
        QRectF(0, 0, pixel_size.width(), pixel_size.height()),
    )
    painter.end()
    pixmap.setDevicePixelRatio(device_pixel_ratio)
    return pixmap


def themed_card_icon(name: str, palette: dict[str, str], size: QSize) -> QIcon:
    return QIcon(themed_svg_pixmap(name, palette["primary"], size))


def themed_navigation_icon(name: str, palette: dict[str, str]) -> QIcon:
    icon = QIcon()
    size = QSize(18, 18)
    default_pixmap = themed_svg_pixmap(name, palette["sidebar_text"], size)
    active_pixmap = themed_svg_pixmap(name, "#ffffff", size)
    icon.addPixmap(default_pixmap, QIcon.Mode.Normal, QIcon.State.Off)
    icon.addPixmap(active_pixmap, QIcon.Mode.Normal, QIcon.State.On)
    icon.addPixmap(active_pixmap, QIcon.Mode.Active, QIcon.State.Off)
    icon.addPixmap(active_pixmap, QIcon.Mode.Active, QIcon.State.On)
    return icon


class TaskCancelled(Exception):
    pass


class WorkerSignals(QObject):
    log = Signal(str)
    result = Signal(object)
    error = Signal(str)
    cancelled = Signal()
    finished = Signal()


class NoWheelComboBox(QComboBox):
    def wheelEvent(self, event) -> None:
        event.ignore()


class NoWheelSpinBox(QSpinBox):
    def wheelEvent(self, event) -> None:
        event.ignore()


class NoWheelDoubleSpinBox(QDoubleSpinBox):
    def wheelEvent(self, event) -> None:
        event.ignore()


class TaskWorker(QRunnable):
    def __init__(self, task: Callable[[Callable[[str], None]], TaskResult]) -> None:
        super().__init__()
        self.task = task
        self.signals = WorkerSignals()
        self.cancel_event = Event()

    def cancel(self) -> None:
        self.cancel_event.set()

    def _log(self, message: str) -> None:
        if self.cancel_event.is_set():
            raise TaskCancelled
        self.signals.log.emit(message)

    @Slot()
    def run(self) -> None:
        try:
            result = self.task(self._log)
            if self.cancel_event.is_set():
                raise TaskCancelled
            self.signals.result.emit(result)
        except TaskCancelled:
            self.signals.cancelled.emit()
        except Exception:
            self.signals.error.emit(traceback.format_exc())
        finally:
            self.signals.finished.emit()


class CalendarDialog(QDialog):
    def __init__(self, selected_date: date, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("calendarDialog")
        self.setWindowTitle("选择日期")
        self.setModal(True)
        self.setFixedSize(390, 430)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        header = QHBoxLayout()
        previous_button = QPushButton("‹")
        previous_button.setObjectName("ghost")
        previous_button.setFixedWidth(38)
        next_button = QPushButton("›")
        next_button.setObjectName("ghost")
        next_button.setFixedWidth(38)
        self.month_label = QLabel()
        self.month_label.setObjectName("sectionTitle")
        self.month_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.addWidget(previous_button)
        header.addWidget(self.month_label, 1)
        header.addWidget(next_button)
        layout.addLayout(header)

        self.calendar = QCalendarWidget()
        self.calendar.setLocale(QLocale(QLocale.Language.Chinese, QLocale.Country.China))
        self.calendar.setNavigationBarVisible(False)
        self.calendar.setGridVisible(False)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        self.calendar.setHorizontalHeaderFormat(QCalendarWidget.HorizontalHeaderFormat.SingleLetterDayNames)
        self.calendar.setSelectedDate(QDate(selected_date.year, selected_date.month, selected_date.day))
        self.calendar.currentPageChanged.connect(self.update_month_label)
        self.calendar.activated.connect(lambda _date: self.accept())
        previous_button.clicked.connect(self.calendar.showPreviousMonth)
        next_button.clicked.connect(self.calendar.showNextMonth)
        layout.addWidget(self.calendar, 1)

        footer = QHBoxLayout()
        yesterday_button = QPushButton("昨天")
        today_button = QPushButton("今天")
        cancel_button = QPushButton("取消")
        confirm_button = QPushButton("确定")
        confirm_button.setObjectName("primary")
        yesterday_button.clicked.connect(
            lambda: self.calendar.setSelectedDate(QDate.currentDate().addDays(-1))
        )
        today_button.clicked.connect(lambda: self.calendar.setSelectedDate(QDate.currentDate()))
        cancel_button.clicked.connect(self.reject)
        confirm_button.clicked.connect(self.accept)
        footer.addWidget(yesterday_button)
        footer.addWidget(today_button)
        footer.addStretch()
        footer.addWidget(cancel_button)
        footer.addWidget(confirm_button)
        layout.addLayout(footer)
        self.update_month_label(self.calendar.yearShown(), self.calendar.monthShown())

    def update_month_label(self, year: int, month: int) -> None:
        self.month_label.setText(f"{year} 年 {month} 月")

    def selected_date(self) -> date:
        selected = self.calendar.selectedDate()
        return date(selected.year(), selected.month(), selected.day())


class DatePicker(QFrame):
    date_changed = Signal(object)

    def __init__(self, selected_date: date) -> None:
        super().__init__()
        self._date = selected_date
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)
        self.text = QLineEdit()
        self.text.setReadOnly(True)
        self.button = QPushButton("选择日期")
        self.button.clicked.connect(self.open_calendar)
        layout.addWidget(self.text, 1)
        layout.addWidget(self.button)
        self._refresh()

    def value(self) -> date:
        return self._date

    def set_value(self, value: date) -> None:
        self._date = value
        self._refresh()
        self.date_changed.emit(value)

    def open_calendar(self) -> None:
        dialog = CalendarDialog(self._date, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.set_value(dialog.selected_date())

    def _refresh(self) -> None:
        weekday = "一二三四五六日"[self._date.weekday()]
        self.text.setText(f"{self._date:%Y-%m-%d}   星期{weekday}")


class PageHeader(QWidget):
    def __init__(self, eyebrow: str, title: str, description: str) -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        eyebrow_label = QLabel(eyebrow.upper())
        eyebrow_label.setObjectName("pageEyebrow")
        title_label = QLabel(title)
        title_label.setObjectName("pageTitle")
        description_label = QLabel(description)
        description_label.setObjectName("muted")
        description_label.setWordWrap(True)
        layout.addWidget(eyebrow_label)
        layout.addWidget(title_label)
        layout.addWidget(description_label)


class HomeCard(QFrame):
    clicked = Signal()

    def __init__(
        self,
        icon_name: str,
        title: str,
        description: str,
        accent: str,
        featured: bool = False,
    ) -> None:
        super().__init__()
        self.icon_name = icon_name
        self.featured = featured
        self.setObjectName("homeFeaturedCard" if featured else "homeCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(126)
        self.setMaximumHeight(140)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(19, 18, 19, 17)
        layout.setSpacing(14)
        self.icon_label = QLabel()
        self.icon_label.setObjectName("cardIcon")
        self.icon_label.setFixedSize(48, 48)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.icon_label, 0, Qt.AlignmentFlag.AlignTop)

        content = QVBoxLayout()
        content.setContentsMargins(0, 0, 0, 0)
        content.setSpacing(6)
        heading = QHBoxLayout()
        heading.setSpacing(8)
        title_label = QLabel(title)
        title_label.setObjectName("featuredCardTitle" if featured else "cardTitle")
        badge = QLabel(accent)
        badge.setObjectName("statusPill")
        badge.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        heading.addWidget(title_label)
        heading.addWidget(badge)
        heading.addStretch()
        description_label = QLabel(description)
        description_label.setObjectName("cardDescription")
        description_label.setWordWrap(True)
        action_label = QLabel("打开功能  →")
        action_label.setObjectName("cardAction")
        action_label.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        content.addLayout(heading)
        content.addWidget(description_label)
        content.addStretch()
        content.addWidget(action_label)
        layout.addLayout(content, 1)
        self.apply_theme(theme_palette(DEFAULT_THEME))

    def apply_theme(self, palette: dict[str, str]) -> None:
        icon = themed_card_icon(self.icon_name, palette, QSize(22, 22))
        self.icon_label.setPixmap(icon.pixmap(QSize(22, 22)))

    def mousePressEvent(self, event) -> None:  # type: ignore[no-untyped-def]
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class HomePage(QScrollArea):
    page_requested = Signal(int)

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("pageScroll")
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.content = QWidget()
        layout = QVBoxLayout(self.content)
        layout.setContentsMargins(36, 32, 36, 34)
        layout.setSpacing(20)

        hero = QFrame()
        hero.setObjectName("homeHero")
        hero_layout = QHBoxLayout(hero)
        hero_layout.setContentsMargins(26, 23, 24, 23)
        hero_layout.setSpacing(24)
        hero_copy = QVBoxLayout()
        hero_copy.setContentsMargins(0, 0, 0, 0)
        hero_copy.setSpacing(7)
        eyebrow = QLabel("SMARTSHEET DESK  ·  WORKSPACE")
        eyebrow.setObjectName("pageEyebrow")
        title = QLabel("今天要处理什么？")
        title.setObjectName("heroTitle")
        subtitle = QLabel("行到水穷处，坐看云起时。")
        subtitle.setObjectName("heroDescription")
        subtitle.setWordWrap(True)
        for label in (eyebrow, title, subtitle):
            label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        hero_copy.addWidget(eyebrow)
        hero_copy.addWidget(title)
        hero_copy.addWidget(subtitle)
        hero_layout.addLayout(hero_copy, 1)
        self.hero_icon = QLabel()
        self.hero_icon.setObjectName("heroBrandIcon")
        self.hero_icon.setFixedSize(92, 92)
        self.hero_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        animation_path = hero_animation_path()
        self.hero_movie = QMovie(str(animation_path))
        self.hero_movie.setScaledSize(QSize(92, 92))
        if self.hero_movie.isValid():
            self.hero_icon.setMovie(self.hero_movie)
            self.hero_movie.start()
        else:
            icon_path = brand_icon_path()
            if icon_path.is_file():
                self.hero_icon.setPixmap(QIcon(str(icon_path)).pixmap(QSize(62, 62)))
        hero_layout.addWidget(self.hero_icon, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(hero)

        self.grid = QGridLayout()
        self.grid.setHorizontalSpacing(13)
        self.grid.setVerticalSpacing(13)
        for column in range(3):
            self.grid.setColumnStretch(column, 1)
        self.cards = [
            HomeCard("sheet-add", "Excel 增行", "批量处理带颜色标签的工作表。", "批量增行"),
            HomeCard("compare", "订单差异比对", "分别执行代收与代付对账并生成 HTML 汇总。", "自动对账", featured=True),
            HomeCard("download", "订单报表下载", "登录后台并下载指定日期的订单报表。", "在线获取"),
            HomeCard("card-add", "增卡", "根据模板批量创建卡号工作表。", "批量建卡"),
            HomeCard("extract", "提取B2B", "批量提取并写入 B2B 交易数据。", "交易提取"),
            HomeCard("results", "对账结果", "按生成日期查看代收与代付对账报告。", "报告归档"),
            HomeCard("flow-sync", "流水同步", "整合 TP 代付、TP 代收和钱包流水同步。", "多源同步"),
            HomeCard("settings", "配置管理", "可视化维护全局配置和功能参数。", "系统设置"),
        ]
        self.card_page_indexes = (1, 2, 3, 4, 5, 6, 8, 7)
        for card, page_index in zip(self.cards, self.card_page_indexes):
            card.clicked.connect(
                lambda checked=False, page=page_index: self.page_requested.emit(page)
            )
        self._feature_visibility = (True, False, False, False, False, False, False, True)
        self._column_count = 0
        self.set_feature_access(False, False, False, False)
        layout.addLayout(self.grid)
        layout.addStretch()
        self.setWidget(self.content)

    def set_feature_access(
        self,
        order_diff: bool,
        fetch_orders: bool,
        add_cards: bool,
        add_b2b: bool,
    ) -> None:
        self._feature_visibility = (
            True,
            order_diff,
            fetch_orders,
            add_cards,
            add_b2b,
            order_diff,
            order_diff,
            True,
        )
        self._relayout_cards()

    def _relayout_cards(self) -> None:
        visible_cards = [
            card
            for card, visible in zip(self.cards, self._feature_visibility)
            if visible
        ]
        for card in self.cards:
            self.grid.removeWidget(card)
            card.setVisible(card in visible_cards)

        column_count = 3 if self.viewport().width() >= 900 else 2
        self._column_count = column_count
        order_diff = self._feature_visibility[1]
        if order_diff and column_count == 3:
            self.grid.addWidget(self.cards[1], 0, 0, 1, 2)
            self.grid.addWidget(self.cards[0], 0, 2)
            remaining_cards = [
                card
                for card in self.cards
                if card in visible_cards and card not in (self.cards[0], self.cards[1])
            ]
            for index, card in enumerate(remaining_cards):
                self.grid.addWidget(card, 1 + index // 3, index % 3)
            return

        if order_diff:
            self.grid.addWidget(self.cards[1], 0, 0, 1, column_count)
            remaining_cards = [
                card
                for card in self.cards
                if card in visible_cards and card is not self.cards[1]
            ]
            for index, card in enumerate(remaining_cards):
                self.grid.addWidget(card, 1 + index // column_count, index % column_count)
            return

        for index, card in enumerate(visible_cards):
            self.grid.addWidget(
                card,
                index // column_count,
                index % column_count,
            )

    def resizeEvent(self, event) -> None:  # type: ignore[no-untyped-def]
        super().resizeEvent(event)
        column_count = 3 if self.viewport().width() >= 900 else 2
        if column_count != self._column_count:
            self._relayout_cards()

    def apply_theme(self, palette: dict[str, str]) -> None:
        for card in self.cards:
            card.apply_theme(palette)


class FieldBlock(QWidget):
    def __init__(self, title: str, control: QWidget, hint: str = "") -> None:
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)
        label = QLabel(title)
        label.setObjectName("fieldLabel")
        layout.addWidget(label)
        layout.addWidget(control)
        if hint:
            hint_label = QLabel(hint)
            hint_label.setObjectName("fieldHint")
            hint_label.setWordWrap(True)
            layout.addWidget(hint_label)


class TaskPage(QWidget):
    def __init__(
        self,
        eyebrow: str,
        title: str,
        description: str,
        scroll_form: bool = False,
    ) -> None:
        super().__init__()
        self.worker: TaskWorker | None = None
        self.output_path: Path | None = None
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 30)
        outer.setSpacing(14)
        outer.addWidget(PageHeader(eyebrow, title, description))

        self.form_panel = QFrame()
        self.form_panel.setObjectName("panel")
        self.form = QVBoxLayout(self.form_panel)
        self.form.setContentsMargins(22, 20, 22, 20)
        self.form.setSpacing(14)
        if scroll_form:
            self.form.setSizeConstraint(QLayout.SizeConstraint.SetMinimumSize)
            self.form_scroll = QScrollArea()
            self.form_scroll.setWidgetResizable(True)
            self.form_scroll.setHorizontalScrollBarPolicy(
                Qt.ScrollBarPolicy.ScrollBarAlwaysOff
            )
            self.form_scroll.setWidget(self.form_panel)
            outer.addWidget(self.form_scroll, 1)
        else:
            self.form_scroll = None
            outer.addWidget(self.form_panel)

        actions = QHBoxLayout()
        actions.setSpacing(9)
        self.run_button = QPushButton("开始执行")
        self.run_button.setObjectName("primary")
        self.cancel_button = QPushButton("停止")
        self.cancel_button.setObjectName("danger")
        self.cancel_button.setEnabled(False)
        self.cancel_button.clicked.connect(self.cancel_task)
        self.open_button = QPushButton("打开结果")
        self.open_button.setVisible(False)
        self.open_button.clicked.connect(self.open_result)
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("muted")
        actions.addWidget(self.run_button)
        actions.addWidget(self.cancel_button)
        actions.addWidget(self.open_button)
        actions.addStretch()
        actions.addWidget(self.status_label)
        outer.addLayout(actions)

        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        outer.addWidget(self.progress)

        self.log_toggle = QPushButton("执行日志  ▾")
        self.log_toggle.setObjectName("ghost")
        self.log_toggle.setCheckable(True)
        self.log_toggle.setChecked(False)
        self.log_toggle.setFixedWidth(110)
        self.log_toggle.clicked.connect(self.toggle_log)
        outer.addWidget(self.log_toggle)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMinimumHeight(96)
        self.log_view.setMaximumHeight(120)
        self.log_view.setVisible(False)
        outer.addWidget(self.log_view)
        if not scroll_form:
            outer.addStretch()

    def add_field(self, title: str, control: QWidget, hint: str = "") -> None:
        self.form.addWidget(FieldBlock(title, control, hint))

    def start_task(self, task: Callable[[Callable[[str], None]], TaskResult]) -> None:
        self.log_view.clear()
        self.output_path = None
        self.open_button.setVisible(False)
        self.run_button.setEnabled(False)
        self.cancel_button.setEnabled(True)
        self.progress.setRange(0, 0)
        self.status_label.setText("正在执行…")
        self.log_view.appendPlainText("任务已开始")
        worker = TaskWorker(task)
        self.worker = worker
        worker.signals.log.connect(self.log_view.appendPlainText)
        worker.signals.result.connect(self.task_succeeded)
        worker.signals.error.connect(self.task_failed)
        worker.signals.cancelled.connect(self.task_cancelled)
        worker.signals.finished.connect(self.task_finished)
        QThreadPool.globalInstance().start(worker)

    def toggle_log(self, visible: bool) -> None:
        self.log_view.setVisible(visible)
        self.log_toggle.setText("执行日志  ▴" if visible else "执行日志  ▾")

    def cancel_task(self) -> None:
        if self.worker:
            self.cancel_button.setEnabled(False)
            self.status_label.setText("正在安全停止…")
            self.worker.cancel()

    @Slot(object)
    def task_succeeded(self, result: TaskResult) -> None:
        self.output_path = result.output_path
        self.status_label.setText(result.summary)
        self.log_view.appendPlainText(result.summary)
        self.open_button.setVisible(result.output_path is not None)
        QMessageBox.information(self, result.title, result.summary)

    @Slot(str)
    def task_failed(self, details: str) -> None:
        self.log_view.appendPlainText(details)
        self.status_label.setText("执行失败，请查看日志")
        self.log_toggle.setChecked(True)
        self.toggle_log(True)
        last_line = next((line for line in reversed(details.splitlines()) if line.strip()), "执行失败")
        QMessageBox.critical(self, "执行失败", last_line)

    @Slot()
    def task_cancelled(self) -> None:
        self.status_label.setText("任务已停止")
        self.log_view.appendPlainText("任务已停止")

    @Slot()
    def task_finished(self) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.run_button.setEnabled(True)
        self.cancel_button.setEnabled(False)
        self.worker = None

    def open_result(self) -> None:
        if self.output_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output_path)))


class PathPicker(QWidget):
    def __init__(
        self,
        mode: str,
        initial: str = "",
        file_filter: str = "",
        allowed_suffixes: tuple[str, ...] = (".xlsx",),
    ) -> None:
        super().__init__()
        self.mode = mode
        self.file_filter = file_filter
        self.allowed_suffixes = tuple(suffix.lower() for suffix in allowed_suffixes)
        self.setAcceptDrops(mode == "file")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)
        self.edit = QLineEdit(initial)
        suffix_hint = " / ".join(self.allowed_suffixes)
        self.edit.setPlaceholderText(
            f"可点击浏览或拖入 {suffix_hint} 文件"
            if mode == "file"
            else "请选择路径"
        )
        if mode == "file":
            self.edit.setAcceptDrops(False)
        button = QPushButton("浏览…")
        button.clicked.connect(self.choose)
        layout.addWidget(self.edit, 1)
        layout.addWidget(button)

    def path(self) -> Path:
        return Path(self.edit.text().strip()).expanduser()

    def choose(self) -> None:
        start = self.edit.text().strip() or str(Path.cwd())
        if self.mode == "directory":
            selected = QFileDialog.getExistingDirectory(self, "选择目录", start)
        else:
            selected, _ = QFileDialog.getOpenFileName(self, "选择文件", start, self.file_filter)
        if selected:
            self.edit.setText(selected)

    def dragEnterEvent(self, event) -> None:
        if self.mode == "file" and event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event) -> None:
        paths = [
            Path(url.toLocalFile())
            for url in event.mimeData().urls()
            if url.isLocalFile()
        ]
        if len(paths) != 1:
            QMessageBox.warning(self, "无法添加文件", "请一次拖入一个文件。")
            event.ignore()
            return
        path = paths[0]
        if not path.is_file() or path.suffix.lower() not in self.allowed_suffixes:
            suffixes = " / ".join(self.allowed_suffixes)
            QMessageBox.warning(
                self, "文件格式不支持", f"请拖入有效的 {suffixes} 文件。"
            )
            event.ignore()
            return
        self.edit.setText(str(path))
        event.acceptProposedAction()


class FlowSyncPage(TaskPage):
    def __init__(self) -> None:
        super().__init__(
            "Transaction sync",
            "流水同步",
            "选择流水类型并上传对应文件。",
            scroll_form=True,
        )
        self.feature_names = ("TP代付同步", "TP代收同步", "钱包流水同步")
        self.full_sync_card = HomeCard(
            "flow-sync",
            "一键流水同步",
            "选择一个文件夹，自动识别工作簿、付款订单、两份收款订单和平台钱包流水。",
            "推荐",
            featured=True,
        )
        self.full_sync_card.clicked.connect(
            lambda: self.select_sync_feature("folder")
        )
        self.form.addWidget(self.full_sync_card)
        descriptions = (
            "替换工作簿中的 TP代付 明细。",
            "同步 TP 代收业务流水。",
            "同步钱包账户流水。",
        )
        accents = ("可用", "可用", "可用")
        card_container = QWidget()
        self.grid = QGridLayout(card_container)
        self.grid.setContentsMargins(0, 0, 0, 4)
        self.grid.setHorizontalSpacing(14)
        self.grid.setVerticalSpacing(14)
        for column in range(3):
            self.grid.setColumnStretch(column, 1)
        self.cards: list[HomeCard] = []
        for column, (name, description, accent) in enumerate(
            zip(self.feature_names, descriptions, accents)
        ):
            card = HomeCard("flow-sync", name, description, accent)
            if column == 0:
                card.clicked.connect(lambda: self.select_sync_feature("payout"))
            elif column == 1:
                card.clicked.connect(lambda: self.select_sync_feature("collection"))
            else:
                card.clicked.connect(lambda: self.select_sync_feature("wallet"))
            self.grid.addWidget(card, 0, column)
            self.cards.append(card)
        self.form.addWidget(card_container)

        self.sync_forms = QStackedWidget()

        payout_form = QWidget()
        payout_layout = QVBoxLayout(payout_form)
        payout_layout.setContentsMargins(0, 0, 0, 0)
        payout_layout.setSpacing(14)
        self.workbook_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.payment_orders_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        payout_layout.addWidget(
            FieldBlock(
                "工作簿",
                self.workbook_picker,
                "工作簿中必须包含名为“TP代付”的工作表；处理前请关闭 Excel/WPS。",
            )
        )
        payout_layout.addWidget(
            FieldBlock(
                "付款订单 Excel",
                self.payment_orders_picker,
                "读取第一个工作表，并按 A 列平台订单号动态识别明细范围。",
            )
        )

        collection_form = QWidget()
        collection_layout = QVBoxLayout(collection_form)
        collection_layout.setContentsMargins(0, 0, 0, 0)
        collection_layout.setSpacing(14)
        self.collection_workbook_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        self.collection_orders_first_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        self.collection_orders_second_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        collection_layout.addWidget(
            FieldBlock(
                "工作簿",
                self.collection_workbook_picker,
                "工作簿中必须包含名为“TP代收”的工作表；处理前请关闭 Excel/WPS。",
            )
        )
        collection_layout.addWidget(
            FieldBlock(
                "收款订单 Excel 1",
                self.collection_orders_first_picker,
                "可上传支付成功或部分支付文件，程序会根据 W 列自动识别。",
            )
        )
        collection_layout.addWidget(
            FieldBlock(
                "收款订单 Excel 2",
                self.collection_orders_second_picker,
                "与文件 1 顺序无关；两个文件的平台订单号不能重复。",
            )
        )

        wallet_form = QWidget()
        wallet_layout = QVBoxLayout(wallet_form)
        wallet_layout.setContentsMargins(0, 0, 0, 0)
        wallet_layout.setSpacing(14)
        self.wallet_workbook_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        self.wallet_flow_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        wallet_layout.addWidget(
            FieldBlock(
                "工作簿",
                self.wallet_workbook_picker,
                "工作簿中必须包含名为“长款(当日)”的工作表；处理前请关闭 Excel/WPS。",
            )
        )
        wallet_layout.addWidget(
            FieldBlock(
                "平台钱包流水记录 Excel",
                self.wallet_flow_picker,
                "读取第一个工作表，第 3 行开始的全部流水都会写入。",
            )
        )

        folder_form = QWidget()
        folder_layout = QVBoxLayout(folder_form)
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(14)
        self.flow_sync_directory_picker = PathPicker("directory")
        folder_layout.addWidget(
            FieldBlock(
                "同步文件夹",
                self.flow_sync_directory_picker,
                "文件夹中需包含工作簿、1 个付款订单、2 个收款订单和 1 个平台钱包流水 Excel。",
            )
        )
        self.sync_forms.addWidget(payout_form)
        self.sync_forms.addWidget(collection_form)
        self.sync_forms.addWidget(wallet_form)
        self.sync_forms.addWidget(folder_form)
        self.form.addWidget(self.sync_forms)

        self.active_sync_feature = "folder"
        self.sync_forms.setCurrentIndex(3)
        self.run_button.setText("开始一键流水同步")
        self.run_button.clicked.connect(self.run_sync)

    def focus_tp_payout(self) -> None:
        self.select_sync_feature("payout")

    def select_sync_feature(self, feature: str) -> None:
        self.active_sync_feature = feature
        feature_indexes = {
            "payout": 0,
            "collection": 1,
            "wallet": 2,
            "folder": 3,
        }
        button_labels = {
            "payout": "开始 TP代付同步",
            "collection": "开始 TP代收同步",
            "wallet": "开始钱包流水同步",
            "folder": "开始一键流水同步",
        }
        self.sync_forms.setCurrentIndex(feature_indexes[feature])
        self.run_button.setText(button_labels[feature])
        if feature == "payout":
            self.workbook_picker.edit.setFocus()
        elif feature == "collection":
            self.collection_workbook_picker.edit.setFocus()
        elif feature == "wallet":
            self.wallet_workbook_picker.edit.setFocus()
        else:
            self.flow_sync_directory_picker.edit.setFocus()

    def run_sync(self) -> None:
        if not load_license().allows(FEATURE_ORDER_DIFF):
            QMessageBox.warning(
                self,
                "功能未授权",
                "请先在配置管理中验证包含订单比对权限的密钥。",
            )
            return
        if self.active_sync_feature == "folder":
            self.run_full_flow_sync()
        elif self.active_sync_feature == "collection":
            self.run_tp_collection()
        elif self.active_sync_feature == "wallet":
            self.run_wallet_flow()
        else:
            self.run_tp_payout()

    def focus_tp_collection(self) -> None:
        self.select_sync_feature("collection")

    def run_full_flow_sync(self) -> None:
        directory_text = self.flow_sync_directory_picker.edit.text().strip()
        if not directory_text:
            QMessageBox.warning(
                self,
                "文件夹未选择",
                "请选择包含工作簿和四个来源 Excel 的文件夹。",
            )
            return
        directory = Path(directory_text).expanduser()
        self.start_task(
            lambda log: run_full_flow_sync_task(
                directory,
                log,
            )
        )

    def run_tp_payout(self) -> None:
        workbook_text = self.workbook_picker.edit.text().strip()
        payment_orders_text = self.payment_orders_picker.edit.text().strip()
        if not workbook_text or not payment_orders_text:
            QMessageBox.warning(
                self,
                "文件未选择",
                "请选择工作簿和付款订单 Excel 文件。",
            )
            return
        workbook = Path(workbook_text).expanduser()
        payment_orders = Path(payment_orders_text).expanduser()
        self.start_task(
            lambda log: run_tp_payout_sync_task(
                workbook,
                payment_orders,
                log,
            )
        )

    def run_tp_collection(self) -> None:
        workbook_text = self.collection_workbook_picker.edit.text().strip()
        first_orders_text = self.collection_orders_first_picker.edit.text().strip()
        second_orders_text = self.collection_orders_second_picker.edit.text().strip()
        if not workbook_text or not first_orders_text or not second_orders_text:
            QMessageBox.warning(
                self,
                "文件未选择",
                "请选择工作簿和两个收款订单 Excel 文件。",
            )
            return
        workbook = Path(workbook_text).expanduser()
        first_orders = Path(first_orders_text).expanduser()
        second_orders = Path(second_orders_text).expanduser()
        self.start_task(
            lambda log: run_tp_collection_sync_task(
                workbook,
                first_orders,
                second_orders,
                log,
            )
        )

    def run_wallet_flow(self) -> None:
        workbook_text = self.wallet_workbook_picker.edit.text().strip()
        wallet_flow_text = self.wallet_flow_picker.edit.text().strip()
        if not workbook_text or not wallet_flow_text:
            QMessageBox.warning(
                self,
                "文件未选择",
                "请选择工作簿和平台钱包流水记录 Excel 文件。",
            )
            return
        workbook = Path(workbook_text).expanduser()
        wallet_flow = Path(wallet_flow_text).expanduser()
        self.start_task(
            lambda log: run_wallet_flow_sync_task(
                workbook,
                wallet_flow,
                log,
            )
        )

    def apply_theme(self, palette: dict[str, str]) -> None:
        self.full_sync_card.apply_theme(palette)
        for card in self.cards:
            card.apply_theme(palette)


class FillPage(TaskPage):
    def __init__(self) -> None:
        super().__init__("Local workbook", "Excel 增行", "选择工作簿和目标日期，处理过程完全在本机完成。")
        self.path_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.date_picker = DatePicker(default_target_date())
        self.add_field("工作簿", self.path_picker, "处理前请关闭 Excel/WPS 中正在打开的目标文件。")
        self.add_field("目标日期", self.date_picker)
        self.run_button.clicked.connect(self.run)

    def run(self) -> None:
        limit_sheets = read_fill_limit_sheets(editable_config_path())
        self.start_task(
            lambda log: run_fill_task(
                self.path_picker.path(),
                self.date_picker.value(),
                log,
                limit_sheets=limit_sheets,
            )
        )


class DiffPage(TaskPage):
    def __init__(self) -> None:
        super().__init__(
            "Reconciliation",
            "订单差异比对",
            "分别执行代收或代付订单对账，两种业务使用独立规则。",
            scroll_form=True,
        )
        self.business_combo = NoWheelComboBox()
        self.business_combo.addItem("代收订单比对", "collection")
        self.business_combo.addItem("代付订单比对", "payout")

        self.business_stack = QStackedWidget()
        collection_panel = QWidget()
        collection_layout = QVBoxLayout(collection_panel)
        collection_layout.setContentsMargins(0, 0, 0, 0)
        collection_layout.setSpacing(14)

        self.mode_combo = NoWheelComboBox()
        self.mode_combo.addItem("按订单目录处理", "directory")
        self.mode_combo.addItem("手动上传 Excel", "files")

        self.mode_stack = QStackedWidget()
        directory_panel = QWidget()
        directory_layout = QVBoxLayout(directory_panel)
        directory_layout.setContentsMargins(0, 0, 0, 0)
        directory_layout.setSpacing(14)
        self.path_picker = PathPicker(
            "directory", str(workspace_directory() / "diffOrders")
        )
        self.path_picker.edit.textChanged.connect(self.load_group_config)
        self.date_picker = DatePicker(date.today())
        self.platform_combo = NoWheelComboBox()
        self.platform_combo.addItem("自动 / 默认算法", "")
        self.platform_combo.addItem("finerBit", "finerbit")
        self.platform_combo.addItem("EasyPaisa", "easypaisa")
        save_group_button = QPushButton("保存分组配置")
        save_group_button.clicked.connect(self.save_group_config)
        platform_row = QWidget()
        platform_layout = QHBoxLayout(platform_row)
        platform_layout.setContentsMargins(0, 0, 0, 0)
        platform_layout.setSpacing(7)
        platform_layout.addWidget(self.platform_combo, 1)
        platform_layout.addWidget(save_group_button)
        directory_layout.addWidget(
            FieldBlock(
                "订单目录",
                self.path_picker,
                "目录可以是 diffOrders，也可以是包含 conf.ini 的具体分组目录。",
            )
        )
        directory_layout.addWidget(FieldBlock("订单日期", self.date_picker))
        directory_layout.addWidget(
            FieldBlock(
                "分组算法（conf.ini）",
                platform_row,
                "保存到所选目录的 conf.ini；留空时沿用程序默认判断。",
            )
        )

        manual_panel = QWidget()
        manual_layout = QVBoxLayout(manual_panel)
        manual_layout.setContentsMargins(0, 0, 0, 0)
        manual_layout.setSpacing(14)
        self.upstream_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.backend_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.duplicate_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.manual_platform_combo = NoWheelComboBox()
        self.manual_platform_combo.addItem("自动 / 默认算法", "")
        self.manual_platform_combo.addItem("finerBit", "finerbit")
        self.manual_platform_combo.addItem("EasyPaisa", "easypaisa")
        manual_layout.addWidget(
            FieldBlock("上游 Excel（必传）", self.upstream_picker)
        )
        manual_layout.addWidget(
            FieldBlock("平台收款订单 Excel（必传）", self.backend_picker)
        )
        manual_layout.addWidget(
            FieldBlock(
                "代收重复支付订单（选传）",
                self.duplicate_picker,
                "没有重复支付订单时保持为空。",
            )
        )
        manual_layout.addWidget(
            FieldBlock(
                "分组算法",
                self.manual_platform_combo,
                "自动模式根据文件名和重复支付文件判断；也可以手动指定算法。",
            )
        )

        self.mode_stack.addWidget(directory_panel)
        self.mode_stack.addWidget(manual_panel)
        collection_layout.addWidget(FieldBlock("处理方式", self.mode_combo))
        collection_layout.addWidget(self.mode_stack)
        self.mode_combo.currentIndexChanged.connect(self.mode_stack.setCurrentIndex)
        self.mode_combo.setCurrentIndex(self.mode_combo.findData("files"))

        payout_panel = QWidget()
        payout_layout = QVBoxLayout(payout_panel)
        payout_layout.setContentsMargins(0, 0, 0, 0)
        payout_layout.setSpacing(14)
        self.payout_mode_combo = NoWheelComboBox()
        self.payout_mode_combo.addItem("按订单目录处理", "directory")
        self.payout_mode_combo.addItem("手动上传账单", "files")
        self.payout_algorithm_combo = NoWheelComboBox()
        self.payout_algorithm_combo.addItem("zee", "zee")
        self.payout_algorithm_combo.addItem("finerBit", "finerbit")
        self.payout_mode_stack = QStackedWidget()

        payout_directory_panel = QWidget()
        payout_directory_layout = QVBoxLayout(payout_directory_panel)
        payout_directory_layout.setContentsMargins(0, 0, 0, 0)
        self.payout_path_picker = PathPicker(
            "directory", str(workspace_directory() / "diffOrders")
        )
        payout_directory_layout.addWidget(
            FieldBlock(
                "代付订单目录",
                self.payout_path_picker,
                "按表头自动识别上游 .csv/.xlsx 和我方付款订单 .xlsx。",
            )
        )

        payout_manual_panel = QWidget()
        payout_manual_layout = QVBoxLayout(payout_manual_panel)
        payout_manual_layout.setContentsMargins(0, 0, 0, 0)
        payout_manual_layout.setSpacing(14)
        self.payout_upstream_picker = PathPicker(
            "file",
            file_filter="上游账单 (*.csv *.xlsx)",
            allowed_suffixes=(".csv", ".xlsx"),
        )
        self.payout_backend_picker = PathPicker(
            "file", file_filter="Excel (*.xlsx)"
        )
        self.payout_collection_upstream_picker = PathPicker(
            "file",
            file_filter="finerBit上游代收账单 (*.csv *.xlsx)",
            allowed_suffixes=(".csv", ".xlsx"),
        )
        payout_manual_layout.addWidget(
            FieldBlock(
                "代付上游账单（必传）",
                self.payout_upstream_picker,
                "支持 .csv 或 .xlsx，需包含 TRANS_ID、TRX_STATUS、TRX_AMT、FEE、FED。",
            )
        )
        payout_manual_layout.addWidget(
            FieldBlock(
                "我方付款订单（必传）",
                self.payout_backend_picker,
                "需包含 transactionId、交易状态、付款金额、手续费和支付方式名称。",
            )
        )
        self.payout_collection_upstream_field = FieldBlock(
            "上游代收账单（必传）",
            self.payout_collection_upstream_picker,
            "Transaction Details 文件；支持拖拽或点击浏览上传 .csv/.xlsx。",
        )
        payout_manual_layout.addWidget(self.payout_collection_upstream_field)

        self.payout_mode_stack.addWidget(payout_directory_panel)
        self.payout_mode_stack.addWidget(payout_manual_panel)
        payout_layout.addWidget(FieldBlock("处理方式", self.payout_mode_combo))
        payout_layout.addWidget(
            FieldBlock(
                "分组算法",
                self.payout_algorithm_combo,
                "zee 使用现有代付规则；finerBit 使用独立费率配置。",
            )
        )
        payout_layout.addWidget(self.payout_mode_stack)
        self.payout_mode_combo.currentIndexChanged.connect(
            self.payout_mode_stack.setCurrentIndex
        )
        self.payout_mode_combo.setCurrentIndex(
            self.payout_mode_combo.findData("files")
        )
        self.payout_algorithm_combo.currentIndexChanged.connect(
            self.update_payout_algorithm_fields
        )
        self.update_payout_algorithm_fields()

        self.business_stack.addWidget(collection_panel)
        self.business_stack.addWidget(payout_panel)
        self.add_field("对账类型", self.business_combo)
        self.form.addWidget(self.business_stack)
        self.business_combo.currentIndexChanged.connect(
            self.business_stack.setCurrentIndex
        )
        self.run_button.clicked.connect(self.run)
        self.load_group_config()

    def run(self) -> None:
        if not load_license().allows(FEATURE_ORDER_DIFF):
            QMessageBox.warning(self, "功能未授权", "请先在配置管理中验证包含订单比对权限的密钥。")
            return
        if self.business_combo.currentData() == "payout":
            self.run_payout()
            return
        if self.mode_combo.currentData() == "directory":
            self.start_task(
                lambda log: run_diff_task(
                    self.path_picker.path(), self.date_picker.value(), log
                )
            )
            return

        upstream_text = self.upstream_picker.edit.text().strip()
        backend_text = self.backend_picker.edit.text().strip()
        if not upstream_text or not backend_text:
            QMessageBox.warning(
                self,
                "文件未选择",
                "请选择上游 Excel 和平台收款订单 Excel。",
            )
            return
        duplicate_text = self.duplicate_picker.edit.text().strip()
        upstream_path = Path(upstream_text).expanduser()
        backend_path = Path(backend_text).expanduser()
        duplicate_path = Path(duplicate_text).expanduser() if duplicate_text else None
        platform_mode = str(self.manual_platform_combo.currentData())
        self.start_task(
            lambda log: run_diff_files_task(
                upstream_path,
                backend_path,
                duplicate_path,
                platform_mode,
                log,
            )
        )

    def run_payout(self) -> None:
        algorithm = str(self.payout_algorithm_combo.currentData())
        if algorithm == "finerbit" and self.payout_mode_combo.currentData() == "directory":
            QMessageBox.warning(
                self,
                "finerBit 请使用手动上传",
                "finerBit 需要额外上传 Transaction Details 上游代收账单。",
            )
            return
        if self.payout_mode_combo.currentData() == "directory":
            self.start_task(
                lambda log: run_payout_diff_task(
                    self.payout_path_picker.path(), log
                )
            )
            return

        upstream_text = self.payout_upstream_picker.edit.text().strip()
        backend_text = self.payout_backend_picker.edit.text().strip()
        collection_text = self.payout_collection_upstream_picker.edit.text().strip()
        if not upstream_text or not backend_text or (algorithm == "finerbit" and not collection_text):
            QMessageBox.warning(
                self,
                "文件未选择",
                "请选择代付上游账单、我方付款订单"
                + ("和上游代收账单。" if algorithm == "finerbit" else "。"),
            )
            return
        upstream_path = Path(upstream_text).expanduser()
        backend_path = Path(backend_text).expanduser()
        collection_path = Path(collection_text).expanduser() if collection_text else None
        self.start_task(
            lambda log: run_payout_diff_files_task(
                upstream_path,
                backend_path,
                algorithm,
                collection_path,
                log,
            )
        )

    def update_payout_algorithm_fields(self, _index: int = -1) -> None:
        self.payout_collection_upstream_field.setVisible(
            self.payout_algorithm_combo.currentData() == "finerbit"
        )

    def load_group_config(self) -> None:
        config_path = self.path_picker.path() / "conf.ini"
        parser = read_ini(config_path)
        value = parser.get("diff_orders", "platform", fallback="").strip().lower()
        index = self.platform_combo.findData(value)
        self.platform_combo.setCurrentIndex(max(index, 0))

    def save_group_config(self) -> None:
        directory = self.path_picker.path()
        if not directory.is_dir():
            QMessageBox.warning(self, "无法保存", "请先选择存在的订单分组目录。")
            return
        update_ini(
            directory / "conf.ini",
            {"diff_orders": {"platform": str(self.platform_combo.currentData())}},
        )
        QMessageBox.information(self, "配置已保存", f"已更新 {directory / 'conf.ini'}")


class FetchPage(TaskPage):
    def __init__(self) -> None:
        super().__init__("Remote report", "订单报表下载", "使用本地登录配置查询并下载指定日期的订单报表。")
        self.date_picker = DatePicker(default_target_date())
        self.generate_checkbox = QCheckBox("先生成该日期的最新报表")
        self.generate_checkbox.setChecked(True)
        login_hint = QLabel(f"登录配置：{get_login_config_path()}")
        login_hint.setObjectName("fieldHint")
        login_hint.setWordWrap(True)
        self.add_field("报表日期", self.date_picker)
        self.form.addWidget(self.generate_checkbox)
        self.form.addWidget(login_hint)
        self.run_button.clicked.connect(self.run)

    def run(self) -> None:
        if not load_license().allows(FEATURE_FETCH_ORDERS):
            QMessageBox.warning(self, "功能未授权", "请先在配置管理中验证包含订单下载权限的密钥。")
            return
        self.start_task(
            lambda log: run_fetch_task(
                self.date_picker.value(), self.generate_checkbox.isChecked(), log
            )
        )


class AddCardsPage(TaskPage):
    def __init__(self) -> None:
        super().__init__("Workbook tools", "增卡", "选择工作簿并粘贴卡号，每行一个卡号。")
        self.path_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.cards_input = QPlainTextEdit()
        self.cards_input.setMinimumHeight(170)
        self.cards_input.setPlaceholderText("1234\n3121\n1341")
        self.add_field("工作簿", self.path_picker, "处理前请关闭 Excel/WPS 中的目标文件。")
        self.add_field("卡号", self.cards_input, "每行一个卡号；重复或已存在的卡号会自动跳过。")
        self.run_button.clicked.connect(self.run)

    def run(self) -> None:
        if not load_license().allows(FEATURE_ADD_CARDS):
            QMessageBox.warning(self, "功能未授权", "请先验证包含增卡权限的密钥。")
            return
        workbook = self.path_picker.path()
        cards_text = self.cards_input.toPlainText()
        self.start_task(
            lambda log: run_add_cards_task(workbook, cards_text, log)
        )


class AddB2BPage(TaskPage):
    def __init__(self) -> None:
        super().__init__(
            "Workbook tools",
            "提取B2B",
            "粘贴多行 B2B 数据，确认自动识别的字段后写入提取B2B工作表。",
            scroll_form=True,
        )
        self.path_picker = PathPicker("file", file_filter="Excel (*.xlsx)")
        self.data_input = QPlainTextEdit()
        self.data_input.setMinimumHeight(100)
        self.data_input.setPlaceholderText(
            "2026-07-12 22:04:23 75NVDJEI 01635548053 01850801086 -50000"
        )
        self.date_time_combo = NoWheelComboBox()
        self.trx_id_combo = NoWheelComboBox()
        self.outgoing_card_combo = NoWheelComboBox()
        self.amount_combo = NoWheelComboBox()
        self.mapping_hint = QLabel("粘贴数据后自动识别第一行字段。")
        self.mapping_hint.setObjectName("fieldHint")
        self.add_field("工作簿", self.path_picker, "处理前请关闭 Excel/WPS 中的目标文件。")
        self.add_field("B2B 数据", self.data_input, "每行一条；负金额会自动转为正数。")
        self.add_field("日期时间字段", self.date_time_combo)
        self.add_field("TRXID 字段", self.trx_id_combo)
        self.add_field("转出卡号字段", self.outgoing_card_combo)
        self.add_field("金额字段", self.amount_combo)
        self.form.addWidget(self.mapping_hint)
        self.data_input.textChanged.connect(self.refresh_mapping)
        self.run_button.clicked.connect(self.run)

    def refresh_mapping(self) -> None:
        combos = (
            self.date_time_combo,
            self.trx_id_combo,
            self.outgoing_card_combo,
            self.amount_combo,
        )
        try:
            fields = parse_input_text(self.data_input.toPlainText())[0].fields
            defaults = guess_field_mapping(fields)
        except ValueError as error:
            for combo in combos:
                combo.clear()
            self.mapping_hint.setText(str(error))
            return

        default_indexes = (
            defaults.date_time,
            defaults.trx_id,
            defaults.outgoing_card,
            defaults.amount,
        )
        for combo, default_index in zip(combos, default_indexes):
            combo.clear()
            for index, field in enumerate(fields):
                combo.addItem(f"{index + 1}. {field.value}", index)
            combo.setCurrentIndex(default_index)
        self.mapping_hint.setText("已自动识别；如有需要，可手动调整字段。")

    def run(self) -> None:
        if not load_license().allows(FEATURE_ADD_B2B):
            QMessageBox.warning(self, "功能未授权", "请先验证包含提取B2B权限的密钥。")
            return
        combos = (
            self.date_time_combo,
            self.trx_id_combo,
            self.outgoing_card_combo,
            self.amount_combo,
        )
        if any(combo.currentData() is None for combo in combos):
            QMessageBox.warning(self, "字段未识别", "请先粘贴有效的 B2B 数据。")
            return
        mapping = FieldMapping(*(int(combo.currentData()) for combo in combos))
        workbook = self.path_picker.path()
        input_text = self.data_input.toPlainText()
        self.start_task(
            lambda log: run_add_b2b_task(workbook, input_text, mapping, log)
        )


class ResultRow(QFrame):
    view_requested = Signal(object)

    def __init__(
        self,
        generated_at: datetime,
        result_type: str,
        display_name: str,
        path: Path,
        palette: dict[str, str],
    ) -> None:
        super().__init__()
        self.setObjectName("resultRow")
        self.setToolTip(path.name)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 14, 16, 14)
        layout.setSpacing(15)

        badge = QLabel(result_type)
        badge.setObjectName(
            "resultBadgePayout" if result_type == "代付" else "resultBadgeCollection"
        )
        badge.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        layout.addWidget(badge)

        details = QVBoxLayout()
        details.setSpacing(4)
        title = QLabel(display_name)
        title.setObjectName("resultTitle")
        generated_label = QLabel(f"生成时间  {generated_at:%Y-%m-%d %H:%M:%S}")
        generated_label.setObjectName("muted")
        details.addWidget(title)
        details.addWidget(generated_label)
        layout.addLayout(details, 1)

        view_button = QPushButton("查看")
        view_button.setObjectName("resultViewButton")
        view_button.setIcon(themed_card_icon("results", palette, QSize(16, 16)))
        view_button.setIconSize(QSize(16, 16))
        view_button.clicked.connect(lambda: self.view_requested.emit(path))
        layout.addWidget(view_button)


class ResultsPage(QWidget):
    RESULT_TIMESTAMP_PATTERN = re.compile(r"_(\d{8})_(\d{6})$")

    def __init__(self) -> None:
        super().__init__()
        self.palette = theme_palette(DEFAULT_THEME)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 32, 40, 30)
        outer.setSpacing(18)
        outer.addWidget(
            PageHeader(
                "Reconciliation results",
                "对账结果",
                "按生成日期查看代收与代付订单比对结果。",
            )
        )

        filter_panel = QFrame()
        filter_panel.setObjectName("panel")
        filter_layout = QHBoxLayout(filter_panel)
        filter_layout.setContentsMargins(20, 16, 20, 16)
        filter_layout.setSpacing(12)
        filter_label = QLabel("日期筛选")
        filter_label.setObjectName("fieldLabel")
        filter_layout.addWidget(filter_label)
        self.date_picker = DatePicker(date.today())
        self.date_picker.setFixedWidth(280)
        filter_layout.addWidget(self.date_picker)
        search_button = QPushButton("搜索")
        search_button.setObjectName("primary")
        search_button.clicked.connect(self.refresh_results)
        filter_layout.addWidget(search_button)
        self.cleanup_button = QPushButton("清理历史结果")
        self.cleanup_button.clicked.connect(self.clean_historical_results)
        filter_layout.addWidget(self.cleanup_button)
        filter_layout.addStretch()
        outer.addWidget(filter_panel)

        results_panel = QFrame()
        results_panel.setObjectName("panel")
        results_layout = QVBoxLayout(results_panel)
        results_layout.setContentsMargins(20, 18, 20, 20)
        results_layout.setSpacing(14)
        results_header = QHBoxLayout()
        results_title = QLabel("搜索结果")
        results_title.setObjectName("sectionTitle")
        self.count_label = QLabel()
        self.count_label.setObjectName("muted")
        results_header.addWidget(results_title)
        results_header.addStretch()
        results_header.addWidget(self.count_label)
        results_layout.addLayout(results_header)

        result_scroll = QScrollArea()
        result_scroll.setObjectName("resultScroll")
        result_scroll.setWidgetResizable(True)
        result_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff
        )
        result_list = QWidget()
        result_list.setObjectName("resultList")
        self.result_list_layout = QVBoxLayout(result_list)
        self.result_list_layout.setContentsMargins(0, 0, 0, 0)
        self.result_list_layout.setSpacing(10)
        self.result_list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        result_scroll.setWidget(result_list)
        results_layout.addWidget(result_scroll, 1)
        outer.addWidget(results_panel, 1)

        self.refresh_results()

    def result_entries(self) -> list[tuple[datetime, str, Path]]:
        result_dir = get_result_dir()
        if not result_dir.is_dir():
            return []
        entries: list[tuple[datetime, str, Path]] = []
        for path in result_dir.glob("*.html"):
            if path.name.startswith("payout_order_diff_"):
                result_type = "代付"
            elif path.name.startswith("order_diff_"):
                result_type = "代收"
            else:
                continue
            match = self.RESULT_TIMESTAMP_PATTERN.search(path.stem)
            if match is None:
                continue
            generated_at = datetime.strptime("".join(match.groups()), "%Y%m%d%H%M%S")
            entries.append((generated_at, result_type, path))
        entries.sort(key=lambda entry: entry[0], reverse=True)
        return entries

    def historical_result_entries(
        self, reference_date: date | None = None
    ) -> list[tuple[datetime, str, Path]]:
        cutoff_date = (reference_date or date.today()) - timedelta(days=7)
        return [
            entry
            for entry in self.result_entries()
            if entry[0].date() < cutoff_date
        ]

    def clean_historical_results(self) -> None:
        entries = self.historical_result_entries()
        if not entries:
            QMessageBox.information(
                self,
                "无需清理",
                "暂无七天前的对账结果需要清理。",
            )
            return

        cutoff_date = date.today() - timedelta(days=7)
        answer = QMessageBox.question(
            self,
            "确认清理历史结果",
            f"将清理七天前的对账结果（{cutoff_date:%Y-%m-%d} 之前），"
            f"共 {len(entries)} 个文件。\n\n删除后无法恢复，是否确定清理？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        deleted_count = 0
        failed_paths: list[Path] = []
        for _generated_at, _result_type, path in entries:
            try:
                path.unlink()
                deleted_count += 1
            except OSError:
                failed_paths.append(path)
        self.refresh_results()

        if failed_paths:
            QMessageBox.warning(
                self,
                "部分结果清理失败",
                f"已清理 {deleted_count} 个结果，{len(failed_paths)} 个文件清理失败。",
            )
            return
        QMessageBox.information(
            self,
            "清理完成",
            f"已清理 {deleted_count} 个七天前的对账结果。",
        )

    def refresh_results(self) -> None:
        selected_date = self.date_picker.value()
        entries = [
            entry
            for entry in self.result_entries()
            if entry[0].date() == selected_date
        ]
        while self.result_list_layout.count():
            item = self.result_list_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for generated_at, result_type, path in entries:
            row = ResultRow(
                generated_at,
                result_type,
                self.display_name(path, result_type),
                path,
                self.palette,
            )
            row.view_requested.connect(self.open_path)
            self.result_list_layout.addWidget(row)
        if not entries:
            empty_state = QFrame()
            empty_state.setObjectName("emptyResult")
            empty_state.setMinimumHeight(150)
            empty_layout = QVBoxLayout(empty_state)
            empty_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty_title = QLabel("当前日期暂无对账结果")
            empty_title.setObjectName("resultTitle")
            empty_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty_hint = QLabel("请选择其他日期后点击搜索")
            empty_hint.setObjectName("muted")
            empty_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty_layout.addWidget(empty_title)
            empty_layout.addWidget(empty_hint)
            self.result_list_layout.addWidget(empty_state)
        self.count_label.setText(f"找到 {len(entries)} 条结果")

    def display_name(self, path: Path, result_type: str) -> str:
        prefix = "payout_order_diff_" if result_type == "代付" else "order_diff_"
        name = path.stem.removeprefix(prefix)
        name = self.RESULT_TIMESTAMP_PATTERN.sub("", name).removeprefix("batch_")
        group_name = name.strip("_- ").replace("_", " · ") or "未命名"
        return f"{group_name} 对账结果"

    def apply_theme(self, palette: dict[str, str]) -> None:
        self.palette = palette
        self.refresh_results()

    def open_path(self, path: Path) -> None:
        if not path.is_file():
            QMessageBox.warning(self, "结果不存在", "该结果文件已被移动或删除。")
            self.refresh_results()
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


class SettingSection(QFrame):
    def __init__(self, title: str, description: str) -> None:
        super().__init__()
        self.setObjectName("settingSection")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(20, 18, 20, 18)
        self.layout.setSpacing(13)
        title_label = QLabel(title)
        title_label.setObjectName("sectionTitle")
        description_label = QLabel(description)
        description_label.setObjectName("fieldHint")
        description_label.setWordWrap(True)
        self.layout.addWidget(title_label)
        self.layout.addWidget(description_label)

    def add_setting(self, label: str, control: QWidget) -> None:
        row = QHBoxLayout()
        label_widget = QLabel(label)
        label_widget.setMinimumWidth(185)
        row.addWidget(label_widget)
        row.addWidget(control, 1)
        self.layout.addLayout(row)


class SettingsPage(QWidget):
    license_changed = Signal(object)
    theme_changed = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.config_path = editable_config_path()
        self.controls: dict[tuple[str, str], QWidget] = {}
        outer = QVBoxLayout(self)
        outer.setContentsMargins(40, 30, 40, 24)
        outer.setSpacing(14)
        header_row = QHBoxLayout()
        header_row.addWidget(PageHeader("Application settings", "配置管理", "可视化维护全局 config.ini；账号、密码和会话信息不会在这里显示。"), 1)
        save_button = QPushButton("保存配置")
        save_button.setObjectName("primary")
        save_button.clicked.connect(self.save_values)
        header_row.addWidget(save_button)
        outer.addLayout(header_row)

        license_panel = QFrame()
        license_panel.setObjectName("licensePanel")
        license_layout = QVBoxLayout(license_panel)
        license_layout.setContentsMargins(20, 17, 20, 17)
        license_layout.setSpacing(9)
        license_header = QHBoxLayout()
        license_title = QLabel("功能授权密钥")
        license_title.setObjectName("sectionTitle")
        self.license_status = QLabel()
        license_header.addWidget(license_title)
        license_header.addStretch()
        license_header.addWidget(self.license_status)
        self.license_input = QLineEdit()
        self.license_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.license_input.setPlaceholderText("粘贴 AX1 开头的授权密钥")
        license_actions = QHBoxLayout()
        validate_button = QPushButton("验证并启用")
        validate_button.setObjectName("primary")
        validate_button.clicked.connect(self.validate_and_install_license)
        remove_button = QPushButton("移除密钥")
        remove_button.clicked.connect(self.clear_license)
        license_actions.addWidget(self.license_input, 1)
        license_actions.addWidget(validate_button)
        license_actions.addWidget(remove_button)
        self.license_detail = QLabel()
        self.license_detail.setObjectName("fieldHint")
        self.license_detail.setWordWrap(True)
        license_layout.addLayout(license_header)
        license_layout.addLayout(license_actions)
        license_layout.addWidget(self.license_detail)
        outer.addWidget(license_panel)

        path_label = QLabel(f"配置文件：{self.config_path}")
        path_label.setObjectName("fieldHint")
        outer.addWidget(path_label)
        scroll = QScrollArea()
        scroll.setObjectName("pageScroll")
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 8, 0)
        content_layout.setSpacing(14)

        appearance_section = SettingSection(
            "外观设置",
            "切换后立即预览；点击保存配置后，下次启动会继续使用当前主题。",
        )
        self.theme_combo = NoWheelComboBox()
        for theme_name, label in THEME_LABELS.items():
            self.theme_combo.addItem(label, theme_name)
        self.theme_combo.currentIndexChanged.connect(self.emit_theme_change)
        appearance_section.add_setting("界面主题", self.theme_combo)
        self.controls[("ui", "theme")] = self.theme_combo

        fill_section = SettingSection("Excel 增行", "控制批处理方式和默认工作簿。")
        self.add_text(fill_section, "目标日期", "fill", "target_date", "留空时使用前一天")
        self.add_number(fill_section, "每批工作表数量", "fill", "limit_sheets", 1, 500)
        self.add_bool(fill_section, "仅处理彩色标签", "fill", "colored_sheets")
        self.add_bool(fill_section, "快速 XML 模式", "fill", "fast_xml")
        self.add_bool(fill_section, "持续运行直至完成", "fill", "run_until_done")
        self.add_bool(fill_section, "运行前选择工作簿", "fill", "select_workbook")
        self.add_text(fill_section, "默认工作簿", "fill", "workbook", "可填写文件名或绝对路径")

        diff_section = SettingSection(
            "订单比对",
            "finerbit 费率填写百分值，例如填写 4 表示 4%，填写 2.3 表示 2.3%。",
        )
        self.add_bool(diff_section, "完成后打开 HTML", "diff_orders", "auto_open_html")
        self.add_decimal(
            diff_section,
            "Easypaisa 费率",
            "diff_orders",
            "easypaisa_rate_percent",
            0,
            100,
            4,
        )
        self.add_decimal(
            diff_section,
            "JazzCash 费率",
            "diff_orders",
            "jazzcash_rate_percent",
            0,
            100,
            2.3,
        )

        payout_finerbit_section = SettingSection(
            "代付 finerBit 费率",
            "费率填写百分值，例如填写 1.3 表示 1.3%。",
        )
        self.add_decimal(
            payout_finerbit_section,
            "Easypaisa 费率",
            "diff_orders",
            "payout_finerbit_easypaisa_rate_percent",
            0,
            100,
            1.3,
        )
        self.add_decimal(
            payout_finerbit_section,
            "JazzCash 费率",
            "diff_orders",
            "payout_finerbit_jazzcash_rate_percent",
            0,
            100,
            1.25,
        )

        fetch_section = SettingSection("订单下载", "接口与下载设置。登录凭据继续保存在 loginConf.ini。")
        self.add_text(fetch_section, "登录接口", "fetch_orders", "login_url")
        self.add_text(fetch_section, "生成报表接口", "fetch_orders", "report_url")
        self.add_text(fetch_section, "报表列表接口", "fetch_orders", "scheduled_reports_url")
        self.add_text(fetch_section, "下载接口", "fetch_orders", "download_url")
        self.add_text(fetch_section, "报表名称", "fetch_orders", "report_name")
        self.add_text(fetch_section, "交易状态", "fetch_orders", "transaction_status")
        self.add_text(fetch_section, "下载目录", "fetch_orders", "download_dir")
        self.add_number(fetch_section, "请求超时（秒）", "fetch_orders", "timeout_seconds", 1, 300)
        self.add_bool(fetch_section, "验证 SSL 证书", "fetch_orders", "verify_ssl")
        self.add_text(fetch_section, "Origin", "fetch_orders", "origin")
        self.add_text(fetch_section, "Referer", "fetch_orders", "referer")
        self.add_text(fetch_section, "User-Agent", "fetch_orders", "user_agent")

        content_layout.addWidget(appearance_section)
        content_layout.addWidget(fill_section)
        content_layout.addWidget(diff_section)
        content_layout.addWidget(payout_finerbit_section)
        content_layout.addWidget(fetch_section)
        content_layout.addStretch()
        scroll.setWidget(content)
        outer.addWidget(scroll, 1)
        self.load_values()
        self.refresh_license_status()

    def add_text(
        self,
        section_widget: SettingSection,
        label: str,
        section: str,
        option: str,
        placeholder: str = "",
    ) -> None:
        control = QLineEdit()
        control.setPlaceholderText(placeholder)
        section_widget.add_setting(label, control)
        self.controls[(section, option)] = control

    def add_number(
        self,
        section_widget: SettingSection,
        label: str,
        section: str,
        option: str,
        minimum: int,
        maximum: int,
    ) -> None:
        control = NoWheelSpinBox()
        control.setRange(minimum, maximum)
        control.setButtonSymbols(QAbstractSpinBox.ButtonSymbols.PlusMinus)
        section_widget.add_setting(label, control)
        self.controls[(section, option)] = control

    def add_bool(
        self, section_widget: SettingSection, label: str, section: str, option: str
    ) -> None:
        control = QCheckBox("启用")
        section_widget.add_setting(label, control)
        self.controls[(section, option)] = control

    def add_decimal(
        self,
        section_widget: SettingSection,
        label: str,
        section: str,
        option: str,
        minimum: float,
        maximum: float,
        default: float = 0,
    ) -> None:
        control = NoWheelDoubleSpinBox()
        control.setRange(minimum, maximum)
        control.setValue(default)
        control.setDecimals(4)
        control.setSingleStep(0.1)
        control.setSuffix(" %")
        section_widget.add_setting(label, control)
        self.controls[(section, option)] = control

    def load_values(self) -> None:
        parser = read_ini(self.config_path)
        for (section, option), control in self.controls.items():
            value = parser.get(section, option, fallback="")
            if isinstance(control, QLineEdit):
                control.setText(value)
            elif isinstance(control, QDoubleSpinBox):
                control.setValue(float(value or control.value()))
            elif isinstance(control, QSpinBox):
                control.setValue(int(value or control.minimum()))
            elif isinstance(control, QCheckBox):
                control.setChecked(value.strip().lower() in {"1", "yes", "true", "on", "y"})
            elif isinstance(control, QComboBox):
                selected_index = control.findData(normalize_theme_name(value))
                control.setCurrentIndex(max(selected_index, 0))

    def save_values(self) -> None:
        updates: dict[str, dict[str, str]] = {}
        for (section, option), control in self.controls.items():
            if isinstance(control, QLineEdit):
                value = control.text().strip()
            elif isinstance(control, QDoubleSpinBox):
                value = f"{control.value():.4f}".rstrip("0").rstrip(".")
            elif isinstance(control, QSpinBox):
                value = str(control.value())
            elif isinstance(control, QComboBox):
                value = str(control.currentData())
            else:
                value = "true" if isinstance(control, QCheckBox) and control.isChecked() else "false"
            updates.setdefault(section, {})[option] = value
        update_ini(self.config_path, updates)
        QMessageBox.information(self, "配置已保存", f"已更新 {self.config_path}")

    def emit_theme_change(self, _index: int = -1) -> None:
        self.theme_changed.emit(str(self.theme_combo.currentData()))

    def validate_and_install_license(self) -> None:
        info = install_license(self.license_input.text())
        if not info.valid:
            self.refresh_license_status(info)
            QMessageBox.warning(self, "密钥验证失败", info.message)
            return
        self.license_input.clear()
        self.refresh_license_status(info)
        self.license_changed.emit(info)
        QMessageBox.information(self, "密钥验证成功", "付费功能已根据许可证权限开放。")

    def clear_license(self) -> None:
        remove_license()
        self.license_input.clear()
        info = load_license()
        self.refresh_license_status(info)
        self.license_changed.emit(info)

    def refresh_license_status(self, info: LicenseInfo | None = None) -> None:
        current = info or load_license()
        self.license_status.setText("已授权" if current.valid else "未授权")
        self.license_status.setObjectName("licenseValid" if current.valid else "licenseInvalid")
        self.license_status.style().unpolish(self.license_status)
        self.license_status.style().polish(self.license_status)
        if not current.valid:
            self.license_detail.setText(f"{current.message} · 密钥文件：{license_file_path()}")
            return
        expiration = (
            current.expires_at.strftime("%Y-%m-%d %H:%M UTC")
            if current.expires_at
            else "永久"
        )
        self.license_detail.setText(f"到期时间：{expiration}")


class Sidebar(QFrame):
    page_requested = Signal(int)
    mode_toggle_requested = Signal()

    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("sidebar")
        self.setFixedWidth(236)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 24, 16, 18)
        layout.setSpacing(6)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        self.brand_icon = QLabel()
        self.brand_icon.setFixedSize(40, 40)
        icon_path = brand_icon_path()
        if icon_path.is_file():
            self.brand_icon.setPixmap(QIcon(str(icon_path)).pixmap(QSize(40, 40)))
        self.brand_label = QLabel("SmartSheet Desk")
        self.brand_label.setObjectName("sidebarBrand")
        poetry = poetry_line_for()
        self.brand_caption = QLabel(poetry.replace("，", "，\n", 1))
        self.brand_caption.setObjectName("sidebarBrandCaption")
        self.brand_caption.setWordWrap(True)
        self.brand_caption.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        brand_text = QVBoxLayout()
        brand_text.setContentsMargins(0, 0, 0, 0)
        brand_text.setSpacing(4)
        brand_text.addWidget(self.brand_label)
        brand_text.addWidget(self.brand_caption)
        brand_row.addWidget(self.brand_icon, 0, Qt.AlignmentFlag.AlignTop)
        brand_row.addLayout(brand_text, 1)
        layout.addLayout(brand_row)
        layout.addSpacing(24)
        section = QLabel("常用工具")
        section.setObjectName("navSection")
        layout.addWidget(section)
        navigation = [
            ("工作台", "dashboard", 0),
            ("Excel 增行", "sheet-add", 1),
            ("订单比对", "compare", 2),
            ("订单下载", "download", 3),
            ("增卡", "card-add", 4),
            ("提取B2B", "extract", 5),
            ("流水同步", "flow-sync", 8),
            ("对账结果", "results", 6),
            ("配置管理", "settings", 7),
        ]
        self.icon_names = [icon_name for _label, icon_name, _page in navigation]
        self.page_indexes = [page for _label, _icon, page in navigation]
        self.buttons: list[QPushButton] = []
        for label, icon_name, page_index in navigation:
            if page_index == 6:
                layout.addSpacing(10)
                management_section = QLabel("数据与设置")
                management_section.setObjectName("navSection")
                layout.addWidget(management_section)
            button = QPushButton(label)
            button.setObjectName("navButton")
            button.setCheckable(True)
            button.setIcon(
                themed_navigation_icon(icon_name, theme_palette(DEFAULT_THEME))
            )
            button.setIconSize(QSize(18, 18))
            button.clicked.connect(
                lambda checked=False, page=page_index: self.page_requested.emit(page)
            )
            layout.addWidget(button)
            self.buttons.append(button)
        layout.addStretch()
        footer_row = QHBoxLayout()
        footer_row.setContentsMargins(0, 0, 0, 0)
        footer_row.setSpacing(8)
        self.footer_label = QLabel(f"SmartSheet Desk  ·  v{VERSION}")
        self.footer_label.setObjectName("sidebarCaption")
        footer_row.addWidget(self.footer_label)
        footer_row.addStretch()
        self.mode_toggle = QPushButton()
        self.mode_toggle.setObjectName("modeToggle")
        self.mode_toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.mode_toggle.setIconSize(QSize(17, 17))
        self.mode_toggle.clicked.connect(self.mode_toggle_requested.emit)
        footer_row.addWidget(self.mode_toggle)
        layout.addLayout(footer_row)
        self.apply_theme(theme_palette(DEFAULT_THEME), DEFAULT_MODE)
        self.select(0)

    def apply_theme(self, palette: dict[str, str], mode: str = DEFAULT_MODE) -> None:
        for button, icon_name in zip(self.buttons, self.icon_names):
            button.setIcon(themed_navigation_icon(icon_name, palette))
        dark_mode = normalize_ui_mode(mode) == "dark"
        icon_name = "sun" if dark_mode else "moon"
        self.mode_toggle.setIcon(
            QIcon(themed_svg_pixmap(icon_name, palette["sidebar_text"], QSize(17, 17)))
        )
        target_mode_label = "浅色" if dark_mode else "深色"
        self.mode_toggle.setToolTip(f"切换到{target_mode_label}模式")
        self.mode_toggle.setAccessibleName(f"切换到{target_mode_label}模式")

    def select(self, index: int) -> None:
        for page_index, button in zip(self.page_indexes, self.buttons):
            button.setChecked(page_index == index)

    def set_feature_access(
        self,
        order_diff: bool,
        fetch_orders: bool,
        add_cards: bool,
        add_b2b: bool,
    ) -> None:
        self.buttons[2].setVisible(order_diff)
        self.buttons[3].setVisible(fetch_orders)
        self.buttons[4].setVisible(add_cards)
        self.buttons[5].setVisible(add_b2b)
        self.buttons[6].setVisible(order_diff)
        self.buttons[7].setVisible(order_diff)


def centered_window_geometry(window_size: QSize, available_geometry: QRect) -> QRect:
    width = min(window_size.width(), available_geometry.width())
    height = min(window_size.height(), available_geometry.height())
    left = available_geometry.left() + (available_geometry.width() - width) // 2
    top = available_geometry.top() + (available_geometry.height() - height) // 2
    return QRect(left, top, width, height)


def center_window_on_current_screen(window: QWidget) -> None:
    app = QApplication.instance()
    if app is None:
        return
    screen = app.screenAt(QCursor.pos()) or app.primaryScreen()
    if screen is None:
        return
    target = centered_window_geometry(window.size(), screen.availableGeometry())
    if target.width() < window.minimumWidth() or target.height() < window.minimumHeight():
        window.setMinimumSize(
            min(window.minimumWidth(), target.width()),
            min(window.minimumHeight(), target.height()),
        )
    window.resize(target.size())
    window.move(target.topLeft())


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("SmartSheet Desk")
        self.resize(1280, 820)
        self.setMinimumSize(940, 650)
        root = QWidget()
        root.setObjectName("appRoot")
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.sidebar = Sidebar()
        self.pages = QStackedWidget()
        self.home_page = HomePage()
        self.fill_page = FillPage()
        self.diff_page = DiffPage()
        self.fetch_page = FetchPage()
        self.add_cards_page = AddCardsPage()
        self.add_b2b_page = AddB2BPage()
        self.results_page = ResultsPage()
        self.settings_page = SettingsPage()
        self.flow_sync_page = FlowSyncPage()
        self.pages.addWidget(self.home_page)
        self.pages.addWidget(self.fill_page)
        self.pages.addWidget(self.diff_page)
        self.pages.addWidget(self.fetch_page)
        self.pages.addWidget(self.add_cards_page)
        self.pages.addWidget(self.add_b2b_page)
        self.pages.addWidget(self.results_page)
        self.pages.addWidget(self.settings_page)
        self.pages.addWidget(self.flow_sync_page)
        layout.addWidget(self.sidebar)
        layout.addWidget(self.pages, 1)
        self.setCentralWidget(root)
        self.sidebar.page_requested.connect(self.show_page)
        self.sidebar.mode_toggle_requested.connect(self.toggle_ui_mode)
        self.home_page.page_requested.connect(self.show_page)
        self.settings_page.license_changed.connect(self.apply_license)
        self.settings_page.theme_changed.connect(self.apply_theme)
        self.theme_name = theme_name_from_config(self.settings_page.config_path)
        self.ui_mode = ui_mode_from_config(self.settings_page.config_path)
        self.apply_theme(self.theme_name)
        self.apply_license(load_license())

    @Slot(str)
    def apply_theme(self, theme_name: str) -> None:
        self.theme_name = normalize_theme_name(theme_name)
        palette = theme_palette(self.theme_name, self.ui_mode)
        app = QApplication.instance()
        if app is not None:
            app.setStyleSheet(build_app_style(self.theme_name, self.ui_mode))
        self.sidebar.apply_theme(palette, self.ui_mode)
        self.home_page.apply_theme(palette)
        self.results_page.apply_theme(palette)
        self.flow_sync_page.apply_theme(palette)

    @Slot()
    def toggle_ui_mode(self) -> None:
        self.ui_mode = "light" if self.ui_mode == "dark" else "dark"
        update_ini(self.settings_page.config_path, {"ui": {"mode": self.ui_mode}})
        self.apply_theme(self.theme_name)

    def show_page(self, index: int) -> None:
        current_license = load_license()
        if current_license != self.license_info:
            self.apply_license(current_license)
        required_feature = {
            2: FEATURE_ORDER_DIFF,
            3: FEATURE_FETCH_ORDERS,
            4: FEATURE_ADD_CARDS,
            5: FEATURE_ADD_B2B,
            6: FEATURE_ORDER_DIFF,
            8: FEATURE_ORDER_DIFF,
        }.get(index)
        if required_feature and not self.license_info.allows(required_feature):
            return
        self.pages.setCurrentIndex(index)
        self.sidebar.select(index)
        if index == 6:
            self.results_page.refresh_results()
        elif index == 7:
            self.settings_page.load_values()
            self.settings_page.refresh_license_status()

    @Slot(object)
    def apply_license(self, info: LicenseInfo) -> None:
        self.license_info = info
        order_diff = info.allows(FEATURE_ORDER_DIFF)
        fetch_orders = info.allows(FEATURE_FETCH_ORDERS)
        add_cards = info.allows(FEATURE_ADD_CARDS)
        add_b2b = info.allows(FEATURE_ADD_B2B)
        self.sidebar.set_feature_access(order_diff, fetch_orders, add_cards, add_b2b)
        self.home_page.set_feature_access(order_diff, fetch_orders, add_cards, add_b2b)
        required_feature = {
            2: FEATURE_ORDER_DIFF,
            3: FEATURE_FETCH_ORDERS,
            4: FEATURE_ADD_CARDS,
            5: FEATURE_ADD_B2B,
            6: FEATURE_ORDER_DIFF,
            8: FEATURE_ORDER_DIFF,
        }.get(self.pages.currentIndex())
        if required_feature and not info.allows(required_feature):
            self.show_page(0)


def main() -> None:
    ensure_workspace_directories()
    app = QApplication(sys.argv)
    app.setApplicationName("SmartSheet Desk")
    icon_path = bundled_resource("icon/cover-v5.png")
    if icon_path is None:
        icon_path = Path(__file__).resolve().parents[2] / "icon" / "cover-v5.png"
    if icon_path.is_file():
        app.setWindowIcon(QIcon(str(icon_path)))
    app.setStyle("Fusion")
    config_path = editable_config_path()
    app.setStyleSheet(
        build_app_style(
            theme_name_from_config(config_path),
            ui_mode_from_config(config_path),
        )
    )
    app.setFont(QFont("", 13))
    window = MainWindow()
    center_window_on_current_screen(window)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
