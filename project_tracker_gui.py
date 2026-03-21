from __future__ import annotations

import sys
import threading
from pathlib import Path
from typing import Any, Optional

from PySide6.QtCore import QDate, Qt, QRectF, Signal
from PySide6.QtGui import QAction, QColor, QCursor, QIcon, QKeySequence, QPainter, QPainterPath, QPalette, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QAbstractItemView,
    QProgressDialog,
    QSizePolicy,
)

from project_tracker_backend import DEFAULT_TASKS, ProjectRecord, ProjectTrackerBackend, TaskRecord
from updater import UpdateInfo, check_for_update, download_and_apply

PHASES = sorted({item["phase"] for item in DEFAULT_TASKS} | {"General"})

PHASE_COLORS: dict[str, str] = {
    "Pre-Project": "#487cff",
    "Planning": "#a78bfa",
    "Materials": "#34d399",
    "Shipping": "#22d3ee",
    "Engineering": "#fb923c",
    "Installation": "#f59e0b",
    "Controls": "#e879f9",
    "Turnover": "#94a3b8",
    "Commissioning": "#f43f5e",
    "Closeout": "#4ade80",
    "Archive": "#818cf8",
    "Warranty": "#67e8f9",
    "Financial": "#fbbf24",
    "General": "#64748b",
}


class ProjectDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, project: Optional[ProjectRecord] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Project Details")
        self.setModal(True)
        self.resize(540, 380)

        self.job_name_edit = QLineEdit(project.job_name if project else "")
        self.job_number_edit = QLineEdit(project.job_number if project else "")
        self.pm_edit = QLineEdit(project.project_manager if project else "")
        self.sales_edit = QLineEdit(project.sales_engineer if project else "")

        self.completion_check = QCheckBox("Set target completion date")
        self.completion_edit = QDateEdit()
        self.completion_edit.setCalendarPopup(True)
        self.completion_edit.setDisplayFormat("yyyy-MM-dd")
        self.completion_edit.setMinimumDate(QDate(2000, 1, 1))

        if project and project.target_completion:
            parsed_date = QDate.fromString(project.target_completion, "yyyy-MM-dd")
            if parsed_date.isValid():
                self.completion_edit.setDate(parsed_date)
                self.completion_check.setChecked(True)
            else:
                self.completion_edit.setDate(QDate.currentDate())
                self.completion_check.setChecked(False)
        else:
            self.completion_edit.setDate(QDate.currentDate())
            self.completion_check.setChecked(False)

        self.completion_edit.setEnabled(self.completion_check.isChecked())
        self.completion_check.toggled.connect(self.completion_edit.setEnabled)

        completion_row = QHBoxLayout()
        completion_row.addWidget(self.completion_check)
        completion_row.addWidget(self.completion_edit)
        completion_row.setStretch(1, 1)

        self.liquid_damages_edit = QLineEdit(project.liquid_damages if project else "")
        self.warranty_edit = QLineEdit(project.warranty_period if project else "")
        self.notes_edit = QPlainTextEdit(project.notes if project else "")

        form_layout = QFormLayout()
        form_layout.addRow("Job name *", self.job_name_edit)
        form_layout.addRow("Job number *", self.job_number_edit)
        form_layout.addRow("Project manager", self.pm_edit)
        form_layout.addRow("Sales engineer", self.sales_edit)
        form_layout.addRow("Target completion", completion_row)
        form_layout.addRow("Liquid damages", self.liquid_damages_edit)
        form_layout.addRow("Warranty period", self.warranty_edit)
        form_layout.addRow("Notes", self.notes_edit)

        button_box = QDialogButtonBox()
        button_box.addButton(QDialogButtonBox.StandardButton.Ok)
        button_box.addButton(QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        main_layout = QVBoxLayout(self)
        main_layout.addLayout(form_layout)
        main_layout.addWidget(button_box)

    def get_data(self) -> ProjectRecord:
        return ProjectRecord(
            job_name=self.job_name_edit.text().strip(),
            job_number=self.job_number_edit.text().strip(),
            project_manager=self.pm_edit.text().strip(),
            sales_engineer=self.sales_edit.text().strip(),
            target_completion=(
                self.completion_edit.date().toString("yyyy-MM-dd")
                if self.completion_check.isChecked() and self.completion_edit.date().isValid()
                else None
            ),
            liquid_damages=self.liquid_damages_edit.text().strip(),
            warranty_period=self.warranty_edit.text().strip(),
            notes=self.notes_edit.toPlainText().strip(),
        )

    def accept(self) -> None:
        if not self.job_name_edit.text().strip():
            QMessageBox.warning(self, "Missing job name", "Enter a job name before saving.")
            return
        if not self.job_number_edit.text().strip():
            QMessageBox.warning(self, "Missing job number", "Enter a job number before saving.")
            return
        super().accept()


class TaskDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, task: Optional[TaskRecord] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Task Details")
        self.setModal(True)
        self.resize(440, 260)

        self.task_name_edit = QLineEdit(task.task_name if task else "")
        self.phase_combo = QComboBox()
        self.phase_combo.addItems(PHASES)
        if task and task.phase:
            phase_index = self.phase_combo.findText(task.phase)
            if phase_index >= 0:
                self.phase_combo.setCurrentIndex(phase_index)

        self.completed_check = QCheckBox("Completed")
        self.completed_check.setChecked(task.is_complete if task else False)

        self.completed_date_edit = QDateEdit()
        self.completed_date_edit.setCalendarPopup(True)
        self.completed_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.completed_date_edit.setDate(QDate.currentDate())
        if task and task.completed_date:
            parsed_date = QDate.fromString(task.completed_date, "yyyy-MM-dd")
            if parsed_date.isValid():
                self.completed_date_edit.setDate(parsed_date)

        self.notes_edit = QPlainTextEdit(task.notes if task else "")
        self.completed_check.toggled.connect(self.completed_date_edit.setEnabled)
        self.completed_date_edit.setEnabled(self.completed_check.isChecked())

        form_layout = QFormLayout()
        form_layout.addRow("Task", self.task_name_edit)
        form_layout.addRow("Phase", self.phase_combo)
        form_layout.addRow("Status", self.completed_check)
        form_layout.addRow("Completed date", self.completed_date_edit)
        form_layout.addRow("Notes", self.notes_edit)

        button_box = QDialogButtonBox()
        button_box.addButton(QDialogButtonBox.StandardButton.Ok)
        button_box.addButton(QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        main_layout = QVBoxLayout(self)
        main_layout.addLayout(form_layout)
        main_layout.addWidget(button_box)

    def get_data(self) -> dict[str, Any]:
        completed = self.completed_check.isChecked()
        return {
            "task_name": self.task_name_edit.text().strip(),
            "phase": self.phase_combo.currentText(),
            "is_complete": bool(completed),
            "completed_date": self.completed_date_edit.date().toString("yyyy-MM-dd") if completed else None,
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def accept(self) -> None:
        if not self.task_name_edit.text().strip():
            QMessageBox.warning(self, "Missing task name", "Enter a task name before saving.")
            return
        super().accept()


class StatCard(QFrame):
    def __init__(self, title: str, value: str = "0") -> None:
        super().__init__()
        self.setObjectName("StatCard")
        self.title_label = QLabel(title)
        self.title_label.setObjectName("StatTitle")
        self.value_label = QLabel(value)
        self.value_label.setObjectName("StatValue")

        card_layout = QVBoxLayout(self)
        card_layout.setContentsMargins(14, 12, 14, 12)
        card_layout.addWidget(self.title_label)
        card_layout.addWidget(self.value_label)

    def set_value(self, value: str) -> None:
        self.value_label.setText(value)


class SegmentedProgressBar(QWidget):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setFixedHeight(12)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self._segments: list[dict] = []

    def set_segments(self, segments: list[dict]) -> None:
        self._segments = [
            {**s, "color": QColor(PHASE_COLORS.get(s["phase"], "#487cff"))}
            for s in segments if s["total"] > 0
        ]
        self.update()

    def clear(self) -> None:
        self._segments = []
        self.update()

    def paintEvent(self, event) -> None:
        if not self._segments:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setBrush(QColor("#11151b"))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(self.rect(), 6, 6)
            return

        total_tasks = sum(s["total"] for s in self._segments)
        w = self.width()
        h = self.height()
        r = h // 2
        gap = 2

        rects: list[tuple[int, int, dict]] = []
        x = 0
        for seg in self._segments:
            seg_w = max(4, int(round(w * seg["total"] / total_tasks)))
            rects.append((x, seg_w, seg))
            x += seg_w + gap

        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)

        painter.setBrush(QColor("#11151b"))
        painter.drawRoundedRect(self.rect(), r, r)

        last_i = len(rects) - 1

        for i, (sx, seg_w, seg) in enumerate(rects):
            is_first = i == 0
            is_last = i == last_i

            dim = QColor(seg["color"])
            dim.setAlpha(60)
            painter.setBrush(dim)
            self._draw_segment(painter, sx, 0, seg_w, h, r, is_first, is_last)

            if seg["done"] > 0:
                fill_w = max(0, int(round(seg_w * seg["done"] / seg["total"])))
                if fill_w > 0:
                    painter.setBrush(QColor(seg["color"]))
                    fill_is_full = fill_w >= seg_w
                    self._draw_segment(
                        painter, sx, 0, fill_w, h, r,
                        left_round=is_first,
                        right_round=is_last and fill_is_full,
                    )

        painter.end()

    @staticmethod
    def _draw_segment(
            painter: QPainter,
            x: int, y: int, w: int, h: int, r: int,
            left_round: bool = False,
            right_round: bool = False,
    ) -> None:
        fx, fy, fw, fh = float(x), float(y), float(w), float(h)
        fr = min(float(r), fw / 2, fh / 2)

        path = QPainterPath()
        if left_round:
            path.moveTo(fx + fr, fy)
        else:
            path.moveTo(fx, fy)

        if right_round:
            path.lineTo(fx + fw - fr, fy)
            path.arcTo(QRectF(fx + fw - 2 * fr, fy, 2 * fr, 2 * fr), 90, -90)
        else:
            path.lineTo(fx + fw, fy)
            path.lineTo(fx + fw, fy + fh)

        if right_round:
            path.lineTo(fx + fw, fy + fh - fr)
            path.arcTo(QRectF(fx + fw - 2 * fr, fy + fh - 2 * fr, 2 * fr, 2 * fr), 0, -90)

        if left_round:
            path.lineTo(fx + fr, fy + fh)
            path.arcTo(QRectF(fx, fy + fh - 2 * fr, 2 * fr, 2 * fr), 270, -90)
        else:
            path.lineTo(fx, fy + fh)
            path.lineTo(fx, fy)

        if left_round:
            path.lineTo(fx, fy + fr)
            path.arcTo(QRectF(fx, fy, 2 * fr, 2 * fr), 180, -90)

        path.closeSubpath()
        painter.drawPath(path)


class ElidingLabel(QLabel):
    def __init__(self, text: str = "", parent: Optional[QWidget] = None) -> None:
        super().__init__(text, parent)
        self._full_text = text
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setMinimumWidth(40)

    def setText(self, text: str) -> None:  # type: ignore[override]
        self._full_text = text
        super().setText(text)
        self.update()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        metrics = self.fontMetrics()
        elided = metrics.elidedText(
            self._full_text, Qt.TextElideMode.ElideRight, self.width()
        )
        painter.setPen(self.palette().color(self.foregroundRole()))
        painter.drawText(self.rect(), int(self.alignment()), elided)


def _load_pixmap(image_path: Path) -> Optional[QPixmap]:
    """Load a QPixmap from path, returning None if missing or invalid."""
    if image_path.exists():
        px = QPixmap(str(image_path))
        if not px.isNull():
            return px
    return None


def _paint_watermark(
    painter: QPainter,
    pixmap: QPixmap,
    widget_w: int,
    widget_h: int,
    opacity: float,
    scale: float,
) -> None:
    """Draw a centred, scaled watermark onto an already-active painter."""
    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
    painter.setOpacity(opacity)
    max_dim = int(min(widget_w, widget_h) * scale)
    scaled = pixmap.scaled(
        max_dim, max_dim,
        Qt.AspectRatioMode.KeepAspectRatio,
        Qt.TransformationMode.SmoothTransformation,
    )
    x = (widget_w - scaled.width())  // 2
    y = (widget_h - scaled.height()) // 2
    painter.drawPixmap(x, y, scaled)


class _BackgroundWidget(QWidget):
    """Full-window central widget — paints watermark at 25% opacity."""

    _OPACITY = 0.25
    _SCALE   = 0.60

    def __init__(self, image_path: Path, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._pixmap = _load_pixmap(image_path)

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        if self._pixmap is None:
            return
        painter = QPainter(self)
        _paint_watermark(painter, self._pixmap, self.width(), self.height(), self._OPACITY, self._SCALE)
        painter.end()


class _WatermarkViewport(QWidget):
    """Custom table viewport — paints watermark at 25% opacity with no scroll flicker."""

    _OPACITY = 0.25
    _SCALE   = 0.65
    _BG      = QColor(18, 21, 27)

    def __init__(self, image_path: Path, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._pixmap = _load_pixmap(image_path)
        self.setAutoFillBackground(False)

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), self._BG)
        if self._pixmap is not None:
            _paint_watermark(painter, self._pixmap, self.width(), self.height(), self._OPACITY, self._SCALE)
        painter.end()


class UpdateBanner(QFrame):
    """Slim banner shown at the bottom of the window when an update is available."""

    install_clicked = Signal()

    def __init__(self, info: UpdateInfo, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setObjectName("UpdateBanner")
        self.setFixedHeight(44)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)

        msg = QLabel(
            f"🆕  Update available — v{info.latest_version} is ready. "
            f"You're on v{info.current_version}."
        )
        msg.setObjectName("UpdateMsg")
        layout.addWidget(msg, 1)

        if info.release_notes:
            notes_btn = QPushButton("What's new?")
            notes_btn.setFixedWidth(100)
            notes_btn.clicked.connect(lambda: QMessageBox.information(
                self, f"What's new in v{info.latest_version}",
                info.release_notes,
            ))
            layout.addWidget(notes_btn)

        install_btn = QPushButton("Install & Restart")
        install_btn.setFixedWidth(140)
        install_btn.setObjectName("InstallBtn")
        install_btn.clicked.connect(self.install_clicked)
        layout.addWidget(install_btn)

        dismiss_btn = QPushButton("✕")
        dismiss_btn.setFixedWidth(32)
        dismiss_btn.setToolTip("Dismiss")
        dismiss_btn.clicked.connect(self.hide)
        layout.addWidget(dismiss_btn)


class ResizeHandle(QFrame):
    MIN_SIDEBAR = 160
    MAX_SIDEBAR = 520

    def __init__(self, sidebar: QWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._sidebar = sidebar
        self._drag_start_x: Optional[float] = None
        self._drag_start_w: Optional[int] = None
        self.setFixedWidth(6)
        self.setObjectName("ResizeHandle")
        self.setCursor(QCursor(Qt.CursorShape.SplitHCursor))

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start_x = event.globalPosition().x()
            self._drag_start_w = self._sidebar.width()

    def mouseMoveEvent(self, event) -> None:
        if self._drag_start_x is None:
            return
        delta = event.globalPosition().x() - self._drag_start_x
        new_w = int(self._drag_start_w + delta)
        new_w = max(self.MIN_SIDEBAR, min(self.MAX_SIDEBAR, new_w))
        self._sidebar.setFixedWidth(new_w)

    def mouseReleaseEvent(self, event) -> None:
        self._drag_start_x = None
        self._drag_start_w = None


class MainWindow(QMainWindow):
    _update_ready = Signal()  # emitted from bg thread when a new version is found

    def __init__(self) -> None:
        super().__init__()
        self._update_ready.connect(self._show_update_banner)
        self.backend = ProjectTrackerBackend(Path(__file__).with_name("project_tracker_data.json"))
        self.current_project_id: Optional[int] = None
        self.current_tasks: list[TaskRecord] = []

        self._populating = False
        self._sort_column: Optional[int] = None
        self._sort_ascending: bool = True

        from version import __version__
        self.setWindowTitle(f"Project Tracking Tool v{__version__}")
        self.resize(1460, 860)
        self.setMinimumSize(1180, 700)

        _icon_path = Path(__file__).with_name("PTT_Normal.ico")
        if _icon_path.exists():
            _icon = QIcon(str(_icon_path))
            self.setWindowIcon(_icon)
            app = QApplication.instance()
            if isinstance(app, QApplication):
                app.setWindowIcon(_icon)

        self._build_ui()
        self._build_menu()
        self._build_shortcuts()
        self.refresh_project_list()

        # Warn if running from a cloud-synced folder
        self._check_sync_folder()

        # Check for updates in the background so startup is never delayed.
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    def _build_ui(self) -> None:
        central_widget = _BackgroundWidget(Path(__file__).with_name("PTT_Transparent.png"))
        self.setCentralWidget(central_widget)

        root_layout = QHBoxLayout(central_widget)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(0)

        sidebar = self._build_sidebar()
        sidebar.setFixedWidth(220)

        handle = ResizeHandle(sidebar, central_widget)

        root_layout.addWidget(sidebar)
        root_layout.addWidget(handle)
        root_layout.addWidget(self._build_main_panel(), 1)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

        # Update banner — hidden until a new version is detected
        self._update_banner: Optional[UpdateBanner] = None

    def _build_sidebar(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("Panel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(14, 14, 14, 14)
        panel_layout.setSpacing(10)

        title_label = QLabel("Projects")
        title_label.setObjectName("SectionTitle")
        panel_layout.addWidget(title_label)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search jobs, PM, sales engineer...")
        self.search_edit.textChanged.connect(self.refresh_project_list)
        panel_layout.addWidget(self.search_edit)

        button_row = QHBoxLayout()
        self.new_project_btn = QPushButton("New")
        self.new_project_btn.setMinimumWidth(72)
        self.new_project_btn.clicked.connect(self.create_project)
        self.import_btn = QPushButton("Import")
        self.import_btn.setMinimumWidth(100)
        self.import_btn.clicked.connect(self.import_workbook)
        button_row.addWidget(self.new_project_btn)
        button_row.addWidget(self.import_btn)
        panel_layout.addLayout(button_row)

        self.project_list = QListWidget()
        self.project_list.currentItemChanged.connect(self.on_project_selected)
        panel_layout.addWidget(self.project_list, 1)

        secondary_row = QHBoxLayout()
        self.edit_project_btn = QPushButton("Edit")
        self.edit_project_btn.setMinimumWidth(72)
        self.edit_project_btn.clicked.connect(self.edit_current_project)
        self.delete_project_btn = QPushButton("Delete")
        self.delete_project_btn.setMinimumWidth(72)
        self.delete_project_btn.clicked.connect(self.delete_current_project)
        secondary_row.addWidget(self.edit_project_btn)
        secondary_row.addWidget(self.delete_project_btn)
        panel_layout.addLayout(secondary_row)

        return panel

    def _build_main_panel(self) -> QWidget:
        panel = QWidget()
        panel.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        main_layout = QVBoxLayout(panel)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(12)

        main_layout.addWidget(self._build_project_header(), 0)
        main_layout.addWidget(self._build_task_table(), 10)
        return panel

    def _build_project_header(self) -> QWidget:
        wrapper = QFrame()
        wrapper.setObjectName("Panel")
        wrapper_layout = QVBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(12, 8, 12, 8)
        wrapper_layout.setSpacing(4)

        row = QHBoxLayout()
        row.setSpacing(8)
        row.setContentsMargins(0, 0, 0, 0)

        left_widget = QWidget()
        left_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        left_layout = QHBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        self.project_title = ElidingLabel("No project selected")
        self.project_title.setObjectName("ProjectTitle")
        self.project_title.setMinimumWidth(60)
        self.project_title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        left_layout.addWidget(self.project_title, 3)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setObjectName("HeaderSep")
        left_layout.addWidget(sep)

        self.job_number_value = ElidingLabel("—")
        self.pm_value = ElidingLabel("—")
        self.sales_value = ElidingLabel("—")
        self.completion_value = ElidingLabel("—")
        self.liquid_value = ElidingLabel("—")
        self.warranty_value = ElidingLabel("—")
        meta_pairs: list[tuple[str, ElidingLabel]] = [
            ("Job #", self.job_number_value),
            ("PM", self.pm_value),
            ("SE", self.sales_value),
            ("Due", self.completion_value),
            ("LD", self.liquid_value),
            ("Warranty", self.warranty_value),
        ]

        for meta_caption, val_label in meta_pairs:
            col = QVBoxLayout()
            col.setSpacing(0)
            col.setContentsMargins(0, 0, 0, 0)
            caption_lbl = QLabel(meta_caption)
            caption_lbl.setObjectName("MetaCaption")
            val_label.setObjectName("MetaValue")
            val_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            col.addWidget(caption_lbl)
            col.addWidget(val_label)
            left_layout.addLayout(col, 1)

        row.addWidget(left_widget, 1)

        right_widget = QWidget()
        right_widget.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Preferred)
        right_layout = QHBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(6)

        self.total_tasks_card = StatCard("Tasks")
        self.completed_card = StatCard("Done")
        self.pending_card = StatCard("Pending")
        self.progress_card = StatCard("Progress")
        for card in [self.total_tasks_card, self.completed_card, self.pending_card, self.progress_card]:
            card.setFixedWidth(76)
            right_layout.addWidget(card)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setObjectName("HeaderSep")
        right_layout.addWidget(sep2)

        self.add_task_btn = QPushButton("Add Task")
        self.add_task_btn.setFixedWidth(110)
        self.add_task_btn.clicked.connect(self.add_task)
        self.export_btn = QPushButton("Export")
        self.export_btn.setFixedWidth(76)
        self.export_btn.clicked.connect(self.export_snapshot)
        right_layout.addWidget(self.add_task_btn)
        right_layout.addWidget(self.export_btn)

        row.addWidget(right_widget, 0)
        wrapper_layout.addLayout(row)

        self.progress_bar = SegmentedProgressBar()
        wrapper_layout.addWidget(self.progress_bar)

        self.project_subtitle = QLabel()
        self.project_subtitle.hide()
        self.project_notes = QPlainTextEdit()
        self.project_notes.hide()

        return wrapper

    def _build_task_table(self) -> QWidget:
        wrapper = QFrame()
        wrapper.setObjectName("Panel")
        wrapper_layout = QVBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(16, 16, 16, 16)
        wrapper_layout.setSpacing(12)

        top_row = QHBoxLayout()
        title_label = QLabel("Tasks")
        title_label.setObjectName("SectionTitle")
        top_row.addWidget(title_label)
        top_row.addStretch(1)

        self.phase_filter = QComboBox()
        self.phase_filter.addItem("All phases")
        self.phase_filter.addItems(PHASES)
        self.phase_filter.currentTextChanged.connect(self.populate_tasks)
        top_row.addWidget(self.phase_filter)

        self.task_search_edit = QLineEdit()
        self.task_search_edit.setPlaceholderText("Filter tasks...")
        self.task_search_edit.textChanged.connect(self.populate_tasks)
        top_row.addWidget(self.task_search_edit)
        wrapper_layout.addLayout(top_row)

        self.task_table = QTableWidget(0, 6)
        self.task_table.setHorizontalHeaderLabels(
            ["Done", "Task", "Phase", "Completed Date", "Notes", "Actions"]
        )
        self.task_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.task_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.task_table.verticalHeader().setVisible(False)
        self.task_table.verticalHeader().setDefaultSectionSize(36)
        self.task_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.task_table.setAlternatingRowColors(False)
        self.task_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        _vp = _WatermarkViewport(Path(__file__).with_name("PTT_Transparent.png"))
        self.task_table.setViewport(_vp)

        header = self.task_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self._on_header_clicked)

        self.task_table.doubleClicked.connect(self._on_task_double_clicked)

        wrapper_layout.addWidget(self.task_table, 1)
        return wrapper

    def _check_sync_folder(self) -> None:
        """Warn if the app is running from a cloud-synced folder."""
        exe_path = str(Path(sys.executable).resolve()).lower()
        sync_indicators = ["onedrive", "dropbox", "google drive", "box sync", "icloud"]
        for indicator in sync_indicators:
            if indicator in exe_path:
                QMessageBox.warning(
                    self,
                    "Cloud Sync Folder Detected",
                    f"The app is running from a cloud-synced folder:\n\n"
                    f"{Path(sys.executable).resolve()}\n\n"
                    f"This can cause auto-updates to fail because {indicator.title()} "
                    f"locks files during sync.\n\n"
                    f"Please move the app to a local folder such as:\n"
                    f"C:\\Tools\\ProjectTrackingTool\\",
                )
                break

    # ── Auto-update ────────────────────────────────────────────────────────────

    def _check_update_bg(self) -> None:
        """Runs in a daemon thread — checks GitHub, posts result to UI thread."""
        info = check_for_update()
        if info:
            self._pending_update_info = info
            # Use the signal to safely cross from the background thread to the UI thread
            self._update_ready.emit()

    def _show_update_banner(self) -> None:
        info: UpdateInfo = getattr(self, "_pending_update_info", None)
        if info is None:
            return
        banner = UpdateBanner(info, self)
        banner.install_clicked.connect(lambda: self._do_install(info))
        self._update_banner = banner
        self.statusBar().addPermanentWidget(banner, 1)
        banner.show()
        self.status_bar.showMessage(f"Update available: v{info.latest_version}", 0)

    def _do_install(self, info: UpdateInfo) -> None:
        progress = QProgressDialog("Downloading update…", "Cancel", 0, 100, self)
        progress.setWindowTitle("Installing Update")
        progress.setModal(True)
        progress.setValue(0)
        progress.show()

        def on_progress(done: int, total: int) -> None:
            if total > 0:
                progress.setValue(int(done / total * 100))
            QApplication.processEvents()

        try:
            download_and_apply(info, progress_callback=on_progress)
        except RuntimeError as exc:
            progress.close()
            QMessageBox.critical(self, "Update failed", str(exc))

    # ── Menu ───────────────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("File")

        new_action = QAction("New Project", self)
        new_action.triggered.connect(self.create_project)

        import_action = QAction("Import Workbook", self)
        import_action.triggered.connect(self.import_workbook)

        self.export_menu_action = QAction("Export Snapshot", self)
        self.export_menu_action.triggered.connect(self.export_snapshot)
        self.export_menu_action.setEnabled(False)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)

        file_menu.addAction(new_action)
        file_menu.addAction(import_action)
        file_menu.addAction(self.export_menu_action)
        file_menu.addSeparator()
        file_menu.addAction(quit_action)

        # ── Help menu ──────────────────────────────────────────────────────────
        help_menu = self.menuBar().addMenu("Help")

        version_history_action = QAction("Version History && Recent Updates", self)
        version_history_action.triggered.connect(self._show_version_history)
        help_menu.addAction(version_history_action)

        help_menu.addSeparator()

        about_action = QAction("About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _show_version_history(self) -> None:
        """Fetch all releases from GitHub and display them in a scrollable dialog."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Version History & Recent Updates")
        dialog.setModal(True)
        dialog.resize(560, 480)

        layout = QVBoxLayout(dialog)
        layout.setSpacing(12)

        header = QLabel("Fetching release history from GitHub…")
        header.setObjectName("SectionTitle")
        layout.addWidget(header)

        text_area = QPlainTextEdit()
        text_area.setReadOnly(True)
        text_area.setObjectName("ReadOnlyNotes")
        layout.addWidget(text_area, 1)

        close_btn = QPushButton("Close")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(dialog.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        dialog.show()
        QApplication.processEvents()

        # Fetch releases in-place (dialog is already visible)
        from updater import GITHUB_OWNER, GITHUB_REPO
        try:
            import urllib.request, json as _json
            url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
            req = urllib.request.Request(
                url,
                headers={"Accept": "application/vnd.github+json",
                         "User-Agent": "ProjectTrackingTool"},
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                releases = _json.loads(resp.read().decode())

            if not releases:
                text_area.setPlainText("No releases found on GitHub.")
                header.setText("Version History")
                return

            lines = []
            for rel in releases:
                tag   = rel.get("tag_name", "").lstrip("vV")
                name  = rel.get("name", tag)
                date  = rel.get("published_at", "")[:10]
                notes = rel.get("body", "").strip() or "No release notes."
                lines.append(f"v{tag} — {name}  ({date})")
                lines.append("─" * 48)
                lines.append(notes)
                lines.append("")

            text_area.setPlainText("\n".join(lines))
            header.setText(f"Version History  ({len(releases)} release{'s' if len(releases) != 1 else ''})")

        except Exception as exc:
            text_area.setPlainText(
                f"Could not fetch release history.\n\nError: {exc}\n\n"
                "You can view the full history at:\n"
                f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/releases"
            )
            header.setText("Version History")

        dialog.exec()

    def _show_about(self) -> None:
        from version import __version__
        QMessageBox.information(
            self,
            "About Project Tracking Tool",
            f"Project Tracking Tool\n"
            f"Version {__version__}\n\n"
            f"Built for the ATS team.\n"
            f"© 2026 Justin Glave",
        )

    def _build_shortcuts(self) -> None:
        delete_shortcut = QAction("Delete task", self)
        delete_shortcut.setShortcut(QKeySequence(Qt.Key.Key_Delete))
        delete_shortcut.triggered.connect(self._delete_selected_task)
        self.addAction(delete_shortcut)

        enter_shortcut = QAction("Edit task", self)
        enter_shortcut.setShortcut(QKeySequence(Qt.Key.Key_Return))
        enter_shortcut.triggered.connect(self._edit_selected_task)
        self.addAction(enter_shortcut)

    def _selected_task_id(self) -> Optional[int]:
        row = self.task_table.currentRow()
        if row < 0:
            return None
        item = self.task_table.item(row, 1)
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _delete_selected_task(self) -> None:
        task_id = self._selected_task_id()
        if task_id is not None:
            self.delete_task(task_id)

    def _edit_selected_task(self) -> None:
        task_id = self._selected_task_id()
        if task_id is not None:
            self.edit_task(task_id)

    def _on_task_double_clicked(self) -> None:
        self._edit_selected_task()

    def refresh_project_list(self) -> None:
        search_text = self.search_edit.text().strip() if hasattr(self, "search_edit") else ""
        projects = self.backend.list_projects(search_text)
        selected_project_id = self.current_project_id

        self.project_list.blockSignals(True)
        self.project_list.clear()
        for project in projects:
            item_text = f"{project.job_name}\n{project.job_number}   •   {project.project_manager or 'No PM'}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, project.id)
            self.project_list.addItem(item)
        self.project_list.blockSignals(False)

        if projects:
            target_row = 0
            if selected_project_id is not None:
                for row_index in range(self.project_list.count()):
                    row_item = self.project_list.item(row_index)
                    if row_item.data(Qt.ItemDataRole.UserRole) == selected_project_id:
                        target_row = row_index
                        break
            self.project_list.setCurrentRow(target_row)
        else:
            self.current_project_id = None
            self.clear_project_display()

    def on_project_selected(
            self,
            current_item: Optional[QListWidgetItem],
            _previous_item: Optional[QListWidgetItem],
    ) -> None:
        self.current_project_id = (
            current_item.data(Qt.ItemDataRole.UserRole) if current_item else None
        )
        self.export_menu_action.setEnabled(self.current_project_id is not None)
        self.load_current_project()

    def load_current_project(self) -> None:
        if self.current_project_id is None:
            self.clear_project_display()
            return

        project = self.backend.get_project(self.current_project_id)
        if not project:
            self.clear_project_display()
            return

        self.project_title.setText(project.job_name)
        subtitle = f"{project.job_number}"
        if project.updated_at:
            subtitle += f"   •   Updated {project.updated_at}"
        self.project_subtitle.setText(subtitle)
        self.job_number_value.setText(project.job_number or "—")
        self.pm_value.setText(project.project_manager or "—")
        self.sales_value.setText(project.sales_engineer or "—")
        self.completion_value.setText(project.target_completion or "—")
        self.liquid_value.setText(project.liquid_damages or "—")
        self.warranty_value.setText(project.warranty_period or "—")
        self.project_notes.setPlainText(project.notes or "")

        self.current_tasks = self.backend.list_tasks(self.current_project_id)
        self._refresh_stats_only()
        self.populate_tasks()

    def clear_project_display(self) -> None:
        self.project_title.setText("No project selected")
        self.project_subtitle.setText("Create a project or import the Phoenix workbook to begin.")
        for value_widget in [
            self.job_number_value,
            self.pm_value,
            self.sales_value,
            self.completion_value,
            self.liquid_value,
            self.warranty_value,
        ]:
            value_widget.setText("—")
        self.project_notes.clear()
        self.total_tasks_card.set_value("0")
        self.completed_card.set_value("0")
        self.pending_card.set_value("0")
        self.progress_card.set_value("0%")
        self.progress_bar.clear()
        self.current_tasks = []
        self.task_table.setRowCount(0)
        self.export_menu_action.setEnabled(False)

    def _on_header_clicked(self, col: int) -> None:
        if col == 5:
            return
        if self._sort_column == col:
            self._sort_ascending = not self._sort_ascending
        else:
            self._sort_column = col
            self._sort_ascending = True

        header = self.task_table.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSortIndicator(
            col,
            Qt.SortOrder.AscendingOrder if self._sort_ascending else Qt.SortOrder.DescendingOrder,
        )
        self.populate_tasks()

    def _refresh_stats_only(self) -> None:
        """Refresh stat cards and progress bar without rebuilding the task table.
        Called by toggle_task so a checkbox click doesn't repopulate all rows."""
        if self.current_project_id is None:
            return
        summary = self.backend.get_project_summary(self.current_project_id)
        totals = summary["totals"]
        self.total_tasks_card.set_value(str(totals["tasks"]))
        self.completed_card.set_value(str(totals["completed"]))
        self.pending_card.set_value(str(totals["pending"]))
        self.progress_card.set_value(f"{totals['progress_percent']}%")

        phase_breakdown = summary["phase_breakdown"]
        segments = [
            {"phase": phase, "total": info["total"], "done": info["completed"]}
            for phase, info in phase_breakdown.items()
        ]
        self.progress_bar.set_segments(segments)

        # Sync the in-memory task list so sort/filter stays accurate
        self.current_tasks = self.backend.list_tasks(self.current_project_id)

    def populate_tasks(self) -> None:
        filtered_tasks = list(self.current_tasks)
        selected_phase = self.phase_filter.currentText() if hasattr(self, "phase_filter") else "All phases"
        search_text = self.task_search_edit.text().strip().casefold() if hasattr(self, "task_search_edit") else ""

        if selected_phase and selected_phase != "All phases":
            filtered_tasks = [task for task in filtered_tasks if task.phase == selected_phase]
        if search_text:
            filtered_tasks = [
                task
                for task in filtered_tasks
                if search_text in task.task_name.casefold()
                   or search_text in task.phase.casefold()
                   or search_text in (task.notes or "").casefold()
            ]

        col_keys = {
            0: lambda t: (0 if t.is_complete else 1),
            1: lambda t: t.task_name.casefold(),
            2: lambda t: t.phase.casefold(),
            3: lambda t: t.completed_date or "",
            4: lambda t: (t.notes or "").casefold(),
        }
        if self._sort_column in col_keys:
            filtered_tasks = sorted(
                filtered_tasks,
                key=col_keys[self._sort_column],
                reverse=not self._sort_ascending,
            )

        self._populating = True
        try:
            self.task_table.setRowCount(len(filtered_tasks))
            for row_index, task in enumerate(filtered_tasks):
                checkbox = QCheckBox()
                checkbox.setChecked(task.is_complete)
                checkbox.toggled.connect(
                    lambda checked, task_id=task.id: self.toggle_task(int(task_id), bool(checked))
                )
                self.task_table.setCellWidget(row_index, 0, self._centered_widget(checkbox))
                self.task_table.setItem(row_index, 1, QTableWidgetItem(task.task_name))

                phase_item = QTableWidgetItem(task.phase)
                phase_item.setForeground(QColor(PHASE_COLORS.get(task.phase, "#64748b")))
                self.task_table.setItem(row_index, 2, phase_item)
                self.task_table.setItem(row_index, 3, QTableWidgetItem(task.completed_date or ""))
                self.task_table.setItem(row_index, 4, QTableWidgetItem(task.notes or ""))
                self.task_table.setCellWidget(row_index, 5, self._task_actions_widget(task))

                for column_index in range(1, 5):
                    item = self.task_table.item(row_index, column_index)
                    if item is not None:
                        item.setData(Qt.ItemDataRole.UserRole, task.id)

        finally:
            self._populating = False

    def _task_actions_widget(self, task: TaskRecord) -> QWidget:
        container = QWidget()
        button_layout = QHBoxLayout(container)
        button_layout.setContentsMargins(0, 0, 0, 0)
        button_layout.setSpacing(4)

        edit_button = QToolButton()
        edit_button.setText("Edit")
        edit_button.clicked.connect(lambda: self.edit_task(int(task.id)))

        delete_button = QToolButton()
        delete_button.setText("Del")
        delete_button.clicked.connect(lambda: self.delete_task(int(task.id)))

        button_layout.addWidget(edit_button)
        button_layout.addWidget(delete_button)
        return container

    @staticmethod
    def _centered_widget(widget: QWidget) -> QWidget:
        container = QWidget()
        row_layout = QHBoxLayout(container)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.addStretch(1)
        row_layout.addWidget(widget)
        row_layout.addStretch(1)
        return container

    def create_project(self) -> None:
        dialog = ProjectDialog(self)
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        record = dialog.get_data()
        try:
            new_project_id = self.backend.create_project(record, include_default_tasks=True)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to create project", str(exc))
            return
        self.current_project_id = new_project_id
        self.refresh_project_list()
        self.status_bar.showMessage(f"Created project {record.job_name}", 4000)

    def edit_current_project(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        if not project:
            return
        dialog = ProjectDialog(self, project)
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        record = dialog.get_data()
        try:
            self.backend.update_project(
                self.current_project_id,
                job_name=record.job_name,
                job_number=record.job_number,
                project_manager=record.project_manager,
                sales_engineer=record.sales_engineer,
                target_completion=record.target_completion,
                liquid_damages=record.liquid_damages,
                warranty_period=record.warranty_period,
                notes=record.notes,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Unable to update project", str(exc))
            return
        self.refresh_project_list()
        self.load_current_project()
        self.status_bar.showMessage("Project updated", 4000)

    def delete_current_project(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        if not project:
            return
        answer = QMessageBox.question(
            self,
            "Delete project",
            f"Delete '{project.job_name}' and all of its tasks?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.backend.delete_project(self.current_project_id)
        self.current_project_id = None
        self.refresh_project_list()
        self.status_bar.showMessage("Project deleted", 4000)

    def import_workbook(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import workbook", "",
            "Excel Files (*.xlsx *.xlsm *.xltx *.xltm)",
        )
        if not file_path:
            return
        try:
            imported_project_id = self.backend.import_project_from_workbook(file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Import failed", str(exc))
            return
        imported_project = self.backend.get_project(imported_project_id)
        if imported_project and not imported_project.job_number:
            QMessageBox.warning(
                self, "No job number found",
                "The workbook did not contain a job number (cell H3 was empty).\n\n"
                "Duplicate detection was skipped. Importing this file again will "
                "create a second entry — edit the project to add a job number.",
            )
        self.current_project_id = imported_project_id
        self.refresh_project_list()
        self.status_bar.showMessage(f"Imported workbook: {Path(file_path).name}", 5000)

    def export_snapshot(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        default_name = f"{project.job_number or 'project'}_snapshot.json" if project else "project_snapshot.json"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export snapshot", default_name, "JSON Files (*.json)",
        )
        if not file_path:
            return
        try:
            output_path = self.backend.export_project_snapshot(self.current_project_id, file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))
            return
        self.status_bar.showMessage(f"Exported snapshot to {output_path}", 5000)

    def add_task(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        dialog = TaskDialog(self)
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        data = dialog.get_data()
        try:
            self.backend.add_task(
                project_id=self.current_project_id,
                task_name=data["task_name"],
                phase=data["phase"],
                completed_date=data["completed_date"],
                notes=data["notes"],
            )
        except Exception as exc:
            QMessageBox.critical(self, "Unable to add task", str(exc))
            return
        self.load_current_project()
        self.status_bar.showMessage("Task added", 4000)

    def edit_task(self, task_id: int) -> None:
        task = next((item for item in self.current_tasks if item.id == task_id), None)
        if not task:
            return
        dialog = TaskDialog(self, task)
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        data = dialog.get_data()
        try:
            self.backend.update_task(task_id, **data)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to update task", str(exc))
            return
        self.load_current_project()
        self.status_bar.showMessage("Task updated", 4000)

    def delete_task(self, task_id: int) -> None:
        task = next((item for item in self.current_tasks if item.id == task_id), None)
        if not task:
            return
        answer = QMessageBox.question(
            self, "Delete task", f"Delete task '{task.task_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.backend.delete_task(task_id)
        self.load_current_project()
        self.status_bar.showMessage("Task deleted", 4000)

    def toggle_task(self, task_id: int, checked: bool) -> None:
        if self._populating:
            return
        try:
            self.backend.set_task_completed(task_id, checked)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to update task", str(exc))
            return
        self._refresh_stats_only()


# ── Theme ──────────────────────────────────────────────────────────────────────

def apply_dark_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()

    color_roles = [
        (QPalette.ColorRole.Window, QColor(25, 28, 34)),
        (QPalette.ColorRole.WindowText, QColor(230, 230, 230)),
        (QPalette.ColorRole.Base, QColor(18, 21, 27)),
        (QPalette.ColorRole.AlternateBase, QColor(31, 35, 43)),
        (QPalette.ColorRole.ToolTipBase, QColor(240, 240, 240)),
        (QPalette.ColorRole.ToolTipText, QColor(20, 20, 20)),
        (QPalette.ColorRole.Text, QColor(230, 230, 230)),
        (QPalette.ColorRole.Button, QColor(40, 45, 55)),
        (QPalette.ColorRole.ButtonText, QColor(235, 235, 235)),
        (QPalette.ColorRole.BrightText, QColor(255, 90, 90)),
        (QPalette.ColorRole.Highlight, QColor(72, 124, 255)),
        (QPalette.ColorRole.HighlightedText, QColor(255, 255, 255)),
        (QPalette.ColorRole.Link, QColor(102, 169, 255)),
    ]
    for role, color in color_roles:
        palette.setColor(role, color)

    app.setPalette(palette)

    app.setStyleSheet(
        """
        QWidget {
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 11pt;
        }
        QMainWindow, QMenuBar, QMenu, QStatusBar {
            background: #191c22;
            color: #e8e8e8;
        }
        #Panel, #StatCard {
            background: rgba(32, 36, 44, 160);
            border: 1px solid #2b313d;
            border-radius: 14px;
        }
        QLabel#ProjectTitle {
            font-size: 14pt;
            font-weight: 700;
        }
        QLabel#ProjectSubtitle {
            color: #98a4b8;
            font-size: 10pt;
        }
        QLabel#SectionTitle {
            font-size: 12pt;
            font-weight: 600;
        }
        QLabel#StatTitle {
            color: #9eadc5;
            font-size: 8pt;
        }
        QLabel#StatValue {
            font-size: 13pt;
            font-weight: 700;
        }
        QPushButton, QToolButton {
            background: #2c3442;
            border: 1px solid #3a4557;
            border-radius: 10px;
            padding: 6px 16px;
        }
        QPushButton:hover, QToolButton:hover {
            background: #354055;
        }
        QPushButton:pressed, QToolButton:pressed {
            background: #253247;
        }
        QLineEdit, QPlainTextEdit, QComboBox, QDateEdit {
            background: #12151b;
            border: 1px solid #313746;
            border-radius: 10px;
            padding: 8px;
            color: #ececec;
            selection-background-color: #487cff;
        }
        QListWidget, QTableWidget {
            background: transparent;
            border: 1px solid #313746;
            border-radius: 10px;
            padding: 8px;
            color: #ececec;
            selection-background-color: #487cff;
        }
        QTableWidget::item {
            background: rgba(22, 26, 34, 140);
            color: #ececec;
            border: none;
            padding: 4px 8px;
        }
        QTableWidget::item:alternate {
            background: rgba(30, 35, 46, 140);
        }
        QTableWidget::item:selected {
            background: rgba(72, 124, 255, 160);
            color: white;
        }
        QHeaderView::section {
            background: rgba(30, 35, 46, 180);
            color: #dfe5f2;
            padding: 8px;
            border: none;
            border-right: 1px solid #2f3644;
            border-bottom: 1px solid #2f3644;
            font-weight: 600;
        }
        QPlainTextEdit#ReadOnlyNotes {
            background: #0e1016;
            color: #7a8599;
            border: 1px solid #252b36;
        }
        QGroupBox {
            border: 1px solid #2f3644;
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
            color: #6b7a95;
            font-size: 9pt;
            font-weight: 600;
        }
        QLabel#MetaValue {
            color: #d0d8eb;
            font-size: 9pt;
        }
        QFrame#ResizeHandle {
            background: #2b313d;
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
            background: #2d6a3f;
            border: 1px solid #3d8a52;
            color: white;
            font-weight: 700;
        }
        #InstallBtn:hover {
            background: #3d8a52;
        }
        """
    )


def main() -> int:
    app = QApplication(sys.argv)
    apply_dark_theme(app)
    window = MainWindow()
    window.show()
    return int(app.exec())


if __name__ == "__main__":
    raise SystemExit(main())
