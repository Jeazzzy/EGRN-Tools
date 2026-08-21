"""Светлая и тёмная темы приложения K Tools."""

from __future__ import annotations

import ctypes
import sys
from enum import Enum
from string import Template

from PySide6.QtCore import QObject, Qt, Signal, Slot
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication, QTableWidget, QWidget


class Theme(str, Enum):
    LIGHT = "light"
    DARK = "dark"


THEME_COLORS = {
    Theme.LIGHT: {
        "text": "#182230",
        "muted": "#64748b",
        "placeholder": "#94a3b8",
        "background": "#f4f7fb",
        "surface": "#ffffff",
        "surface_alt": "#f8fafc",
        "input": "#ffffff",
        "border": "#d7e0eb",
        "border_strong": "#b8c5d4",
        "hover": "#eef3f8",
        "pressed": "#e2e8f0",
        "disabled_bg": "#eef2f6",
        "disabled_text": "#94a3b8",
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "accent_soft": "#eff6ff",
        "accent_soft_hover": "#dbeafe",
        "danger": "#b91c1c",
        "sidebar": "#ffffff",
        "sidebar_border": "#dfe7f0",
        "sidebar_text": "#475569",
        "sidebar_muted": "#8492a6",
        "sidebar_hover": "#eef3f8",
        "brand": "#111827",
        "warning_bg": "#fff7db",
        "warning_text": "#7a5510",
        "warning_border": "#f1d58a",
        "success_row": "#dcfce7",
        "warning_row": "#fef3c7",
        "error_row": "#fee2e2",
    },
    Theme.DARK: {
        "text": "#e5e7eb",
        "muted": "#9aa8bc",
        "placeholder": "#748198",
        "background": "#0f172a",
        "surface": "#172033",
        "surface_alt": "#1c2940",
        "input": "#111a2b",
        "border": "#334155",
        "border_strong": "#475569",
        "hover": "#243249",
        "pressed": "#2d3d56",
        "disabled_bg": "#1a2538",
        "disabled_text": "#68758a",
        "accent": "#3b82f6",
        "accent_hover": "#60a5fa",
        "accent_soft": "#172c4f",
        "accent_soft_hover": "#1d3d6d",
        "danger": "#f87171",
        "sidebar": "#0a1120",
        "sidebar_border": "#243044",
        "sidebar_text": "#cbd5e1",
        "sidebar_muted": "#7f8da3",
        "sidebar_hover": "#1b273a",
        "brand": "#f8fafc",
        "warning_bg": "#3d3118",
        "warning_text": "#f6d782",
        "warning_border": "#725c25",
        "success_row": "#183d2c",
        "warning_row": "#493817",
        "error_row": "#4a2228",
    },
}


_STYLE_TEMPLATE = Template(
    """
* {
    font-family: "Segoe UI";
    font-size: 14px;
    color: $text;
}
QMainWindow, QDialog, QWidget#appRoot, QScrollArea,
QScrollArea > QWidget > QWidget {
    background: $background;
}
QFrame#sidebar {
    background: $sidebar;
    border: none;
    border-right: 1px solid $sidebar_border;
}
QLabel#brand {
    color: $brand;
    font-size: 22px;
    font-weight: 700;
}
QLabel#brandSub {
    color: $sidebar_muted;
    font-size: 12px;
}
QPushButton#navButton {
    background: transparent;
    color: $sidebar_text;
    border: none;
    border-radius: 8px;
    text-align: left;
    padding: 11px 14px;
    font-weight: 500;
}
QPushButton#navButton:hover {
    background: $sidebar_hover;
    color: $brand;
}
QPushButton#navButton:checked {
    background: $accent;
    color: #ffffff;
}
QPushButton#themeButton {
    background: transparent;
    color: $sidebar_text;
    border: 1px solid $sidebar_border;
    text-align: left;
    padding: 9px 12px;
}
QPushButton#themeButton:hover {
    background: $sidebar_hover;
    border-color: $border_strong;
    color: $brand;
}
QLabel#pageTitle {
    color: $brand;
    font-size: 27px;
    font-weight: 700;
}
QLabel#pageSubtitle {
    color: $muted;
    font-size: 14px;
}
QFrame#card {
    background: $surface;
    border: 1px solid $border;
    border-radius: 12px;
}
QLabel#cardTitle {
    color: $brand;
    font-size: 16px;
    font-weight: 650;
}
QLabel#toolTitle {
    color: $brand;
    font-size: 19px;
    font-weight: 700;
}
QLabel#toolSummary {
    color: $text;
    font-size: 15px;
    padding-bottom: 4px;
}
QLabel#helpSectionTitle {
    color: $accent;
    font-size: 12px;
    font-weight: 700;
    padding-top: 5px;
}
QLabel#helpSectionText {
    color: $muted;
}
QLabel#warningBanner {
    background: $warning_bg;
    color: $warning_text;
    border: 1px solid $warning_border;
    border-radius: 8px;
    padding: 12px;
}
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QListWidget,
QTableWidget, QPlainTextEdit {
    background: $input;
    color: $text;
    border: 1px solid $border_strong;
    border-radius: 7px;
    padding: 7px;
    selection-background-color: $accent;
    selection-color: #ffffff;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QListWidget:focus, QTableWidget:focus, QPlainTextEdit:focus {
    border: 1px solid $accent;
}
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled,
QDoubleSpinBox:disabled, QListWidget:disabled, QPlainTextEdit:disabled {
    background: $disabled_bg;
    color: $disabled_text;
}
QLineEdit, QPlainTextEdit {
    selection-background-color: $accent;
    selection-color: #ffffff;
}
QComboBox::drop-down {
    width: 24px;
    border: none;
}
QComboBox QAbstractItemView {
    background: $surface;
    color: $text;
    border: 1px solid $border_strong;
    selection-background-color: $accent;
    selection-color: #ffffff;
    outline: none;
}
QSplitter#mdbContentSplitter::handle:vertical {
    background: $border_strong;
    height: 8px;
    margin: 2px 36px;
    border-radius: 2px;
}
QSplitter#mdbContentSplitter::handle:vertical:hover {
    background: $accent;
}
QPushButton {
    background: $surface;
    color: $text;
    border: 1px solid $border_strong;
    border-radius: 7px;
    padding: 8px 14px;
    font-weight: 550;
}
QPushButton:hover {
    background: $hover;
    border-color: $border_strong;
}
QPushButton:pressed {
    background: $pressed;
}
QPushButton:disabled {
    color: $disabled_text;
    background: $disabled_bg;
    border-color: $border;
}
QPushButton[primary="true"] {
    background: $accent;
    color: #ffffff;
    border-color: $accent;
}
QPushButton[primary="true"]:hover {
    background: $accent_hover;
    border-color: $accent_hover;
}
QPushButton[primary="true"]:disabled {
    background: $disabled_bg;
    color: $disabled_text;
    border-color: $border;
}
QPushButton[danger="true"] {
    color: $danger;
}
QFrame#dropZone {
    background: $accent_soft;
    border: 2px dashed $accent;
    border-radius: 11px;
}
QFrame#dropZone[dragActive="true"] {
    background: $accent_soft_hover;
    border-color: $accent_hover;
}
QLabel#dropIcon {
    color: $accent;
    font-size: 28px;
    font-weight: 400;
}
QLabel#dropText {
    color: $muted;
}
QProgressBar {
    background: $pressed;
    color: $text;
    border: none;
    border-radius: 5px;
    height: 10px;
    text-align: center;
}
QProgressBar::chunk {
    background: $accent;
    border-radius: 5px;
}
QGroupBox {
    background: $surface;
    border: 1px solid $border;
    border-radius: 10px;
    margin-top: 12px;
    padding: 12px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 5px;
    color: $text;
    font-weight: 650;
}
QHeaderView::section {
    background: $surface_alt;
    color: $text;
    border: none;
    border-right: 1px solid $border;
    border-bottom: 1px solid $border;
    padding: 9px;
    font-weight: 650;
}
QTableCornerButton::section {
    background: $surface_alt;
    border: none;
    border-right: 1px solid $border;
    border-bottom: 1px solid $border;
}
QTableWidget {
    gridline-color: $border;
}
QTableWidget::item:selected, QListWidget::item:selected {
    background: $accent;
    color: #ffffff;
}
QTabWidget::pane {
    background: $surface;
    border: 1px solid $border;
    border-radius: 8px;
}
QTabBar::tab {
    background: transparent;
    color: $muted;
    border: none;
    border-bottom: 2px solid transparent;
    padding: 9px 14px;
}
QTabBar::tab:hover {
    color: $text;
}
QTabBar::tab:selected {
    color: $accent;
    border-bottom-color: $accent;
    font-weight: 650;
}
QMenu {
    background: $surface;
    color: $text;
    border: 1px solid $border_strong;
    padding: 5px;
}
QMenu::item {
    border-radius: 5px;
    padding: 7px 24px 7px 10px;
}
QMenu::item:selected {
    background: $accent;
    color: #ffffff;
}
QMenu::item:disabled {
    color: $disabled_text;
}
QMenu::separator {
    background: $border;
    height: 1px;
    margin: 5px 8px;
}
QToolTip {
    background: $surface;
    color: $text;
    border: 1px solid $border_strong;
    padding: 5px;
}
QScrollBar:vertical {
    background: $background;
    width: 12px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: $border_strong;
    min-height: 28px;
    border-radius: 6px;
    margin: 2px;
}
QScrollBar::handle:vertical:hover {
    background: $muted;
}
QScrollBar:horizontal {
    background: $background;
    height: 12px;
    margin: 0;
}
QScrollBar::handle:horizontal {
    background: $border_strong;
    min-width: 28px;
    border-radius: 6px;
    margin: 2px;
}
QScrollBar::handle:horizontal:hover {
    background: $muted;
}
QScrollBar::add-line, QScrollBar::sub-line,
QScrollBar::add-page, QScrollBar::sub-page {
    background: none;
    border: none;
}
"""
)


SEMANTIC_BACKGROUND_ROLE = int(Qt.ItemDataRole.UserRole) + 73
_SEMANTIC_COLOR_KEYS = {
    "success": "success_row",
    "warning": "warning_row",
    "error": "error_row",
}


def stylesheet_for(theme: Theme) -> str:
    """Возвращает таблицу стилей для выбранной темы."""
    return _STYLE_TEMPLATE.substitute(THEME_COLORS[Theme(theme)])


def palette_for(theme: Theme) -> QPalette:
    """Создаёт Qt-палитру для нативных и не покрытых QSS элементов."""
    colors = THEME_COLORS[Theme(theme)]
    palette = QPalette()
    roles = {
        QPalette.ColorRole.Window: "background",
        QPalette.ColorRole.WindowText: "text",
        QPalette.ColorRole.Base: "input",
        QPalette.ColorRole.AlternateBase: "surface_alt",
        QPalette.ColorRole.ToolTipBase: "surface",
        QPalette.ColorRole.ToolTipText: "text",
        QPalette.ColorRole.Text: "text",
        QPalette.ColorRole.Button: "surface",
        QPalette.ColorRole.ButtonText: "text",
        QPalette.ColorRole.BrightText: "brand",
        QPalette.ColorRole.Link: "accent",
        QPalette.ColorRole.Highlight: "accent",
    }
    for role, key in roles.items():
        palette.setColor(role, QColor(colors[key]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
    palette.setColor(
        QPalette.ColorRole.PlaceholderText,
        QColor(colors["placeholder"]),
    )
    for role in (
        QPalette.ColorRole.WindowText,
        QPalette.ColorRole.Text,
        QPalette.ColorRole.ButtonText,
        QPalette.ColorRole.PlaceholderText,
    ):
        palette.setColor(
            QPalette.ColorGroup.Disabled,
            role,
            QColor(colors["disabled_text"]),
        )
    palette.setColor(
        QPalette.ColorGroup.Disabled,
        QPalette.ColorRole.Base,
        QColor(colors["disabled_bg"]),
    )
    palette.setColor(
        QPalette.ColorGroup.Disabled,
        QPalette.ColorRole.Button,
        QColor(colors["disabled_bg"]),
    )
    return palette


def _read_windows_apps_use_light_theme() -> int | None:
    if sys.platform != "win32":
        return None
    try:
        import winreg

        path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return int(value)
    except (ImportError, OSError, TypeError, ValueError):
        return None


def theme_from_color_scheme(scheme) -> Theme | None:
    if scheme == Qt.ColorScheme.Dark:
        return Theme.DARK
    if scheme == Qt.ColorScheme.Light:
        return Theme.LIGHT
    return None


def detect_system_theme(app: QApplication | None = None) -> Theme:
    """Определяет тему приложений Windows с безопасным Qt-fallback."""
    windows_value = _read_windows_apps_use_light_theme()
    if windows_value is not None:
        return Theme.LIGHT if windows_value else Theme.DARK

    application = app or QApplication.instance()
    if application is not None:
        detected = theme_from_color_scheme(application.styleHints().colorScheme())
        if detected is not None:
            return detected
        window_color = application.palette().color(QPalette.ColorRole.Window)
        if window_color.isValid():
            return Theme.DARK if window_color.lightness() < 128 else Theme.LIGHT
    return Theme.LIGHT


def active_theme() -> Theme:
    app = QApplication.instance()
    if app is None:
        return Theme.LIGHT
    value = app.property("kToolsTheme")
    try:
        return Theme(value)
    except (TypeError, ValueError):
        return Theme.LIGHT


def set_semantic_background(item, semantic: str) -> None:
    """Помечает фон ячейки так, чтобы он обновлялся при смене темы."""
    if semantic not in _SEMANTIC_COLOR_KEYS:
        raise ValueError(f"Неизвестный семантический цвет: {semantic}")
    item.setData(SEMANTIC_BACKGROUND_ROLE, semantic)
    item.setBackground(_semantic_color(semantic, active_theme()))


def refresh_semantic_backgrounds(parent: QWidget) -> None:
    tables = []
    if isinstance(parent, QTableWidget):
        tables.append(parent)
    tables.extend(parent.findChildren(QTableWidget))
    for table in tables:
        for row in range(table.rowCount()):
            for column in range(table.columnCount()):
                item = table.item(row, column)
                if item is None:
                    continue
                semantic = item.data(SEMANTIC_BACKGROUND_ROLE)
                if semantic in _SEMANTIC_COLOR_KEYS:
                    item.setBackground(_semantic_color(semantic, active_theme()))


def _semantic_color(semantic: str, theme: Theme) -> QColor:
    color_key = _SEMANTIC_COLOR_KEYS[semantic]
    return QColor(THEME_COLORS[theme][color_key])


def _set_windows_title_bar(widget: QWidget, theme: Theme) -> None:
    if sys.platform != "win32" or not widget.isWindow():
        return
    try:
        enabled = ctypes.c_int(theme == Theme.DARK)
        result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            int(widget.winId()),
            20,
            ctypes.byref(enabled),
            ctypes.sizeof(enabled),
        )
        if result != 0:
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                int(widget.winId()),
                19,
                ctypes.byref(enabled),
                ctypes.sizeof(enabled),
            )
    except (AttributeError, OSError, TypeError, ValueError):
        pass


class ThemeManager(QObject):
    """Применяет тему, следует за Windows и учитывает ручной выбор."""

    theme_changed = Signal(str)

    def __init__(self, app: QApplication):
        super().__init__(app)
        self._app = app
        self._theme = detect_system_theme(app)
        self._manual_override = False
        self._app.styleHints().colorSchemeChanged.connect(
            self._on_system_color_scheme_changed
        )
        self._apply()

    @property
    def theme(self) -> Theme:
        return self._theme

    @property
    def follows_system(self) -> bool:
        return not self._manual_override

    def toggle(self) -> None:
        self._manual_override = True
        next_theme = Theme.DARK if self._theme == Theme.LIGHT else Theme.LIGHT
        self.set_theme(next_theme)

    def set_theme(self, theme: Theme) -> None:
        theme = Theme(theme)
        if theme == self._theme:
            return
        self._theme = theme
        self._apply()

    def refresh_widgets(self) -> None:
        for widget in self._app.topLevelWidgets():
            refresh_semantic_backgrounds(widget)
            _set_windows_title_bar(widget, self._theme)

    @Slot(object)
    def _on_system_color_scheme_changed(self, scheme) -> None:
        if self._manual_override:
            return
        theme = theme_from_color_scheme(scheme) or detect_system_theme(self._app)
        self.set_theme(theme)

    def _apply(self) -> None:
        self._app.setProperty("kToolsTheme", self._theme.value)
        self._app.setPalette(palette_for(self._theme))
        self._app.setStyleSheet(stylesheet_for(self._theme))
        self.refresh_widgets()
        self.theme_changed.emit(self._theme.value)
