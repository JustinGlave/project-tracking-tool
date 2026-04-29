"""
app_gui.py — Main UI for <AppName>

Rename this file and update the app name / window title throughout.
See CLAUDE.md for the full design guide.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QSettings, QThread, Signal
from PySide6.QtGui import QAction, QColor, QIcon, QPalette
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from version import __version__
import updater


# ── Path helpers ──────────────────────────────────────────────────────────────

def _resource_path(filename: str) -> str:
    """Path to a bundled asset (icon, image). Works from source and exe."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", ""))
    else:
        base = Path(__file__).parent
    return str(base / filename)


def _app_data_path(filename: str) -> str:
    """Path to user data in %APPDATA%\\ATS Inc\\<AppName>\\."""
    base = Path(os.environ.get("APPDATA", Path.home())) / "ATS Inc" / "YOUR APP NAME"
    base.mkdir(parents=True, exist_ok=True)
    return str(base / filename)


# ── Update worker ─────────────────────────────────────────────────────────────

class _UpdateChecker(QThread):
    found = Signal(object)  # emits UpdateInfo when an update is available

    def run(self) -> None:
        info = updater.check_for_update()
        if info:
            self.found.emit(info)


# ── Main window ───────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    APP_NAME = "Your App Name"  # change this

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{self.APP_NAME} — v{__version__}")
        self.resize(1100, 700)

        # App icon
        icon_path = _resource_path("PTT_Normal.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self._update_info: updater.UpdateInfo | None = None

        self._build_menu()
        self._build_ui()
        self._restore_settings()
        self._check_for_updates()

    # ── Menu bar ──────────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        mb = self.menuBar()

        # File
        file_menu = mb.addMenu("File")
        exit_act = QAction("Exit", self)
        exit_act.triggered.connect(self.close)
        file_menu.addAction(exit_act)

        # View
        view_menu = mb.addMenu("View")
        self._dark_mode_action = QAction("Dark Mode", self)
        self._dark_mode_action.setCheckable(True)
        settings = QSettings("ATS Inc", self.APP_NAME)
        dark_on = settings.value("darkMode", "true") != "false"
        self._dark_mode_action.setChecked(dark_on)
        self._dark_mode_action.triggered.connect(self._toggle_dark_mode)
        view_menu.addAction(self._dark_mode_action)

        # Help
        help_menu = mb.addMenu("Help")
        about_act = QAction(f"About {self.APP_NAME}", self)
        about_act.triggered.connect(self._show_about)
        help_menu.addAction(about_act)

    # ── Central UI ────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(8)
        self.setCentralWidget(root)

        # ── Sidebar ──────────────────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setFixedWidth(220)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(6)

        new_btn = QPushButton("+ New")
        new_btn.clicked.connect(self._on_new)
        sidebar_layout.addWidget(new_btn)

        self._list = QListWidget()
        self._list.currentRowChanged.connect(self._on_selection_changed)
        sidebar_layout.addWidget(self._list)

        root_layout.addWidget(sidebar)

        # ── Main area ─────────────────────────────────────────────────────
        main_area = QWidget()
        main_layout = QVBoxLayout(main_area)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(8)

        # Header
        header = self._build_header()
        main_layout.addWidget(header)

        # Content (replace with your actual content widget)
        content = self._build_content()
        main_layout.addWidget(content, stretch=1)

        # Update banner (hidden by default)
        self._update_banner = self._build_update_banner()
        self._update_banner.setVisible(False)
        main_layout.addWidget(self._update_banner)

        root_layout.addWidget(main_area, stretch=1)

    def _build_header(self) -> QWidget:
        header = QWidget()
        header.setObjectName("Panel")
        layout = QVBoxLayout(header)
        layout.setContentsMargins(16, 12, 16, 12)

        self._title_label = QLabel("Select an item")
        self._title_label.setObjectName("ProjectTitle")
        layout.addWidget(self._title_label)

        self._subtitle_label = QLabel("")
        self._subtitle_label.setObjectName("ProjectSubtitle")
        layout.addWidget(self._subtitle_label)

        return header

    def _build_content(self) -> QWidget:
        # Replace this placeholder with your actual content
        content = QWidget()
        content.setObjectName("Panel")
        layout = QVBoxLayout(content)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        placeholder = QLabel("Select an item from the sidebar\nor click + New to get started.")
        placeholder.setObjectName("ProjectSubtitle")
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(placeholder)

        return content

    def _build_update_banner(self) -> QWidget:
        banner = QWidget()
        banner.setObjectName("UpdateBanner")
        layout = QHBoxLayout(banner)
        layout.setContentsMargins(12, 6, 12, 6)

        self._update_msg = QLabel("")
        self._update_msg.setObjectName("UpdateMsg")
        layout.addWidget(self._update_msg)
        layout.addStretch()

        whats_new_btn = QPushButton("What's New?")
        whats_new_btn.clicked.connect(self._show_whats_new)
        layout.addWidget(whats_new_btn)

        install_btn = QPushButton("Install & Restart")
        install_btn.setObjectName("InstallBtn")
        install_btn.clicked.connect(self._install_update)
        layout.addWidget(install_btn)

        dismiss_btn = QPushButton("✕")
        dismiss_btn.setFixedWidth(32)
        dismiss_btn.clicked.connect(lambda: banner.setVisible(False))
        layout.addWidget(dismiss_btn)

        return banner

    # ── Slots ─────────────────────────────────────────────────────────────────

    def _on_new(self) -> None:
        # TODO: open a dialog to create a new item, add to self._list
        pass

    def _on_selection_changed(self, row: int) -> None:
        # TODO: load and display the selected item
        if row < 0:
            self._title_label.setText("Select an item")
            self._subtitle_label.setText("")
            return
        item = self._list.item(row)
        if item:
            self._title_label.setText(item.text())

    def _show_about(self) -> None:
        QMessageBox.information(
            self,
            f"About {self.APP_NAME}",
            f"{self.APP_NAME}\nVersion {__version__}\n\nBuilt for ATS Inc.",
        )

    # ── Dark / light mode ─────────────────────────────────────────────────────

    def _toggle_dark_mode(self) -> None:
        dark = self._dark_mode_action.isChecked()
        app = QApplication.instance()
        if app:
            if dark:
                apply_dark_theme(app)
            else:
                apply_light_theme(app)
        settings = QSettings("ATS Inc", self.APP_NAME)
        settings.setValue("darkMode", "true" if dark else "false")

    # ── Settings persistence ───────────────────────────────────────────────────

    def _restore_settings(self) -> None:
        settings = QSettings("ATS Inc", self.APP_NAME)
        dark_on = settings.value("darkMode", "true") != "false"
        app = QApplication.instance()
        if app:
            if dark_on:
                apply_dark_theme(app)
            else:
                apply_light_theme(app)

    def closeEvent(self, event) -> None:
        settings = QSettings("ATS Inc", self.APP_NAME)
        settings.setValue("geometry", self.saveGeometry())
        super().closeEvent(event)

    # ── Auto-updater ──────────────────────────────────────────────────────────

    def _check_for_updates(self) -> None:
        self._update_checker = _UpdateChecker()
        self._update_checker.found.connect(self._on_update_found)
        self._update_checker.start()

    def _on_update_found(self, info: updater.UpdateInfo) -> None:
        self._update_info = info
        self._update_msg.setText(
            f"Update available — v{info.latest_version} is ready. You're on v{info.current_version}."
        )
        self._update_banner.setVisible(True)

    def _show_whats_new(self) -> None:
        if not self._update_info:
            return
        QMessageBox.information(
            self,
            f"What's New in v{self._update_info.latest_version}",
            self._update_info.release_notes or "No release notes provided.",
        )

    def _install_update(self) -> None:
        if not self._update_info:
            return
        try:
            updater.download_and_apply(self._update_info)
        except RuntimeError as exc:
            QMessageBox.critical(self, "Update Failed", str(exc))


# ── Theme functions ───────────────────────────────────────────────────────────

def apply_light_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()
    color_roles = [
        (QPalette.ColorRole.Window,          QColor(210, 212, 218)),
        (QPalette.ColorRole.WindowText,      QColor(25, 25, 25)),
        (QPalette.ColorRole.Base,            QColor(225, 227, 232)),
        (QPalette.ColorRole.AlternateBase,   QColor(200, 202, 208)),
        (QPalette.ColorRole.ToolTipBase,     QColor(255, 255, 220)),
        (QPalette.ColorRole.ToolTipText,     QColor(20, 20, 20)),
        (QPalette.ColorRole.Text,            QColor(25, 25, 25)),
        (QPalette.ColorRole.Button,          QColor(195, 198, 206)),
        (QPalette.ColorRole.ButtonText,      QColor(25, 25, 25)),
        (QPalette.ColorRole.BrightText,      QColor(180, 0, 0)),
        (QPalette.ColorRole.Highlight,       QColor(72, 124, 255)),
        (QPalette.ColorRole.HighlightedText, QColor(255, 255, 255)),
        (QPalette.ColorRole.Link,            QColor(0, 90, 200)),
    ]
    for role, color in color_roles:
        palette.setColor(role, color)
    app.setPalette(palette)
    app.setStyleSheet("""
        QWidget {
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 11pt;
        }
        QMainWindow, QMenuBar, QMenu, QStatusBar {
            background: #d2d4da;
            color: #191919;
        }
        QMenuBar::item:selected, QMenu::item:selected {
            background: #487cff;
            color: white;
        }
        #Panel, #StatCard {
            background: rgba(220, 222, 228, 200);
            border: 1px solid #b0b4be;
            border-radius: 14px;
        }
        QLabel#ProjectTitle {
            font-size: 14pt;
            font-weight: 700;
            color: #111111;
        }
        QLabel#ProjectSubtitle {
            color: #555b66;
            font-size: 10pt;
        }
        QLabel#SectionTitle {
            font-size: 12pt;
            font-weight: 600;
            color: #191919;
        }
        QLabel#StatTitle {
            color: #555b66;
            font-size: 7pt;
        }
        QLabel#StatValue {
            font-size: 10pt;
            font-weight: 700;
            color: #191919;
        }
        QPushButton, QToolButton {
            background: #c3c6ce;
            border: 1px solid #a8acb8;
            border-radius: 10px;
            padding: 6px 16px;
            color: #191919;
        }
        QPushButton:hover, QToolButton:hover {
            background: #b2b6c2;
        }
        QPushButton:pressed, QToolButton:pressed {
            background: #a0a4b0;
        }
        QLineEdit, QPlainTextEdit, QComboBox, QDateEdit {
            background: #e1e3e8;
            border: 1px solid #a8acb8;
            border-radius: 10px;
            padding: 8px;
            color: #191919;
            selection-background-color: #487cff;
        }
        QListWidget, QTableWidget {
            background: transparent;
            border: 1px solid #a8acb8;
            border-radius: 10px;
            padding: 8px;
            color: #191919;
            selection-background-color: #487cff;
        }
        QTableWidget::item {
            background: rgba(215, 217, 223, 180);
            border: none;
            padding: 4px 8px;
        }
        QTableWidget::item:alternate {
            background: rgba(200, 202, 208, 180);
        }
        QTableWidget::item:selected {
            background: rgba(72, 124, 255, 180);
            color: white;
        }
        QHeaderView::section {
            background: rgba(195, 198, 206, 220);
            color: #191919;
            padding: 8px;
            border: none;
            border-right: 1px solid #a8acb8;
            border-bottom: 1px solid #a8acb8;
            font-weight: 600;
        }
        QGroupBox {
            border: 1px solid #a8acb8;
            border-radius: 12px;
            margin-top: 10px;
            padding-top: 12px;
            font-weight: 600;
            color: #191919;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 4px;
        }
        QLabel#MetaCaption {
            color: #555b66;
            font-size: 9pt;
            font-weight: 600;
        }
        QLabel#MetaValue {
            color: #252d3a;
            font-size: 9pt;
        }
        QFrame#ResizeHandle {
            background: #a8acb8;
            border: none;
        }
        QFrame#ResizeHandle:hover {
            background: #487cff;
        }
        QListWidget::item {
            border-radius: 10px;
            padding: 10px;
            margin: 2px 0;
            color: #191919;
        }
        QListWidget::item:selected {
            background: #487cff;
            color: white;
        }
        #UpdateBanner {
            background: rgba(195, 228, 205, 220);
            border-top: 1px solid #5cb87a;
        }
        #UpdateBanner QLabel#UpdateMsg {
            color: #1a6830;
            font-weight: 600;
        }
        #InstallBtn {
            background: #2d8a4a;
            border: 1px solid #3daa5a;
            color: white;
            font-weight: 700;
        }
        #InstallBtn:hover {
            background: #3daa5a;
        }
    """)


def apply_dark_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()
    color_roles = [
        (QPalette.ColorRole.Window,          QColor(28, 28, 28)),
        (QPalette.ColorRole.WindowText,      QColor(230, 230, 230)),
        (QPalette.ColorRole.Base,            QColor(18, 18, 18)),
        (QPalette.ColorRole.AlternateBase,   QColor(35, 35, 35)),
        (QPalette.ColorRole.ToolTipBase,     QColor(240, 240, 240)),
        (QPalette.ColorRole.ToolTipText,     QColor(20, 20, 20)),
        (QPalette.ColorRole.Text,            QColor(230, 230, 230)),
        (QPalette.ColorRole.Button,          QColor(45, 45, 45)),
        (QPalette.ColorRole.ButtonText,      QColor(235, 235, 235)),
        (QPalette.ColorRole.BrightText,      QColor(255, 90, 90)),
        (QPalette.ColorRole.Highlight,       QColor(72, 124, 255)),
        (QPalette.ColorRole.HighlightedText, QColor(255, 255, 255)),
        (QPalette.ColorRole.Link,            QColor(102, 169, 255)),
    ]
    for role, color in color_roles:
        palette.setColor(role, color)
    app.setPalette(palette)
    app.setStyleSheet("""
        QWidget {
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 11pt;
        }
        QMainWindow, QMenuBar, QMenu, QStatusBar {
            background: #1c1c1c;
            color: #e8e8e8;
        }
        QMenuBar::item:selected, QMenu::item:selected {
            background: #487cff;
            color: white;
        }
        #Panel, #StatCard {
            background: rgba(38, 38, 38, 160);
            border: 1px solid #3a3a3a;
            border-radius: 14px;
        }
        QLabel#ProjectTitle {
            font-size: 14pt;
            font-weight: 700;
        }
        QLabel#ProjectSubtitle {
            color: #999999;
            font-size: 10pt;
        }
        QLabel#SectionTitle {
            font-size: 12pt;
            font-weight: 600;
        }
        QLabel#StatTitle {
            color: #aaaaaa;
            font-size: 7pt;
        }
        QLabel#StatValue {
            font-size: 10pt;
            font-weight: 700;
        }
        QPushButton, QToolButton {
            background: #383838;
            border: 1px solid #505050;
            border-radius: 10px;
            padding: 6px 16px;
        }
        QPushButton:hover, QToolButton:hover {
            background: #454545;
        }
        QPushButton:pressed, QToolButton:pressed {
            background: #2a2a2a;
        }
        QLineEdit, QPlainTextEdit, QComboBox, QDateEdit {
            background: #121212;
            border: 1px solid #404040;
            border-radius: 10px;
            padding: 8px;
            color: #ececec;
            selection-background-color: #487cff;
        }
        QListWidget, QTableWidget {
            background: transparent;
            border: 1px solid #404040;
            border-radius: 10px;
            padding: 8px;
            color: #ececec;
            selection-background-color: #487cff;
        }
        QTableWidget::item {
            background: rgba(25, 25, 25, 140);
            border: none;
            padding: 4px 8px;
        }
        QTableWidget::item:alternate {
            background: rgba(35, 35, 35, 140);
        }
        QTableWidget::item:selected {
            background: rgba(72, 124, 255, 160);
            color: white;
        }
        QHeaderView::section {
            background: rgba(40, 40, 40, 180);
            color: #e0e0e0;
            padding: 8px;
            border: none;
            border-right: 1px solid #3a3a3a;
            border-bottom: 1px solid #3a3a3a;
            font-weight: 600;
        }
        QGroupBox {
            border: 1px solid #3a3a3a;
            border-radius: 12px;
            margin-top: 10px;
            padding-top: 12px;
            font-weight: 600;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 4px;
        }
        QLabel#MetaCaption {
            color: #888888;
            font-size: 9pt;
            font-weight: 600;
        }
        QLabel#MetaValue {
            color: #cccccc;
            font-size: 9pt;
        }
        QFrame#ResizeHandle {
            background: #3a3a3a;
            border: none;
        }
        QFrame#ResizeHandle:hover {
            background: #487cff;
        }
        QListWidget::item {
            border-radius: 10px;
            padding: 10px;
            margin: 2px 0;
        }
        QListWidget::item:selected {
            background: #2d4c8f;
            color: white;
        }
        #UpdateBanner {
            background: rgba(30, 60, 40, 220);
            border-top: 1px solid #2d6a3f;
        }
        #UpdateBanner QLabel#UpdateMsg {
            color: #6ee7a0;
            font-weight: 600;
        }
        #InstallBtn {
            background: #1e5c32;
            border: 1px solid #2d8a4a;
            color: white;
            font-weight: 700;
        }
        #InstallBtn:hover {
            background: #2d8a4a;
        }
    """)


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("YOUR APP NAME")   # change this
    app.setOrganizationName("ATS Inc")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
