from __future__ import annotations

import json
import logging
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


def _resource_path(filename: str) -> Path:
    """Locate a bundled asset whether running from source or a PyInstaller exe."""
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / filename
    return Path(__file__).with_name(filename)


def _app_data_path() -> Path:
    """Return the writable user data file path, migrating from legacy location if needed."""
    data_dir = Path(os.environ.get("APPDATA", Path.home())) / "ATS Inc" / "Project Tracking Tool"
    data_dir.mkdir(parents=True, exist_ok=True)
    new_path = data_dir / "project_tracker_data.json"

    # One-time migration: copy data from old location (next to exe) if new file doesn't exist
    if not new_path.exists():
        legacy = Path(sys.executable).with_name("project_tracker_data.json")
        if legacy.exists():
            import shutil
            shutil.copy2(legacy, new_path)

    return new_path


def _backup_data_file(data_path: Path) -> None:
    """Copy the data file to a timestamped backup, keeping the most recent 10 backups."""
    if not data_path.exists():
        return
    backup_dir = data_path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"project_tracker_data_{timestamp}.json"
    try:
        shutil.copy2(data_path, dest)
    except OSError:
        logger.exception("Failed to write backup to %s", dest)
        return
    # Prune oldest backups, keep last 10
    backups = sorted(backup_dir.glob("project_tracker_data_*.json"))
    for old in backups[:-10]:
        try:
            old.unlink()
        except OSError:
            pass
from typing import Any, Optional

from PySide6.QtCore import QDate, QFileSystemWatcher, QSize, Qt, QRectF, QTimer, Signal, QSettings, QUrl
from PySide6.QtGui import QAction, QColor, QCursor, QDesktopServices, QFont, QIcon, QKeySequence, QPainter, QPainterPath, QPixmap
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
    QMenu,
    QProgressDialog,
    QScrollArea,
    QSizePolicy,
    QToolTip,
    QTreeWidget,
    QTreeWidgetItem,
    QStackedWidget,
    QInputDialog,
)

import shutil

from project_tracker_backend import DEFAULT_TASKS, ChangeOrderRecord, NoteRecord, ProjectRecord, ProjectTrackerBackend, TaskRecord
from updater import UpdateInfo, check_for_update, download_and_apply
from financials_dialog import FinancialsDialog
from financials_excel import ExcelFinancialsProvider, SnapshotFinancialsProvider
from user_auth import UserManager, UserRecord

PHASES = sorted({item["phase"] for item in DEFAULT_TASKS} | {"General"})

# Fields compared when detecting changes between sessions
_CHANGE_FIELDS: dict[str, str] = {
    "job_name":            "Job Name",
    "job_number":          "Job Number",
    "project_manager":     "Project Manager",
    "sales_engineer":      "Sales Engineer",
    "target_completion":   "Target Completion",
    "booked_date":         "Booked Date",
    "contract_value":      "Contract Value",
    "owner":               "Owner",
    "contracted_with":     "Contracted With",
    "general_contractor":  "General Contractor",
    "liquid_damages":      "Liquid Damages",
    "warranty_period":     "Warranty Period",
    "job_subtype":         "Job Sub-Type",
    "group_ops_manager":   "Group Ops Manager",
    "group_ops_supervisor":"Group Ops Supervisor",
}

_LOCAL_SNAPSHOT_PATH = (
    Path(os.environ.get("APPDATA", Path.home()))
    / "ATS Inc" / "Project Tracking Tool" / "session_snapshot.json"
)

_SESSION_PATH = (
    Path(os.environ.get("APPDATA", Path.home()))
    / "ATS Inc" / "Project Tracking Tool" / "session.json"
)

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
        self.resize(580, 620)

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

        # Extended fields from Odin email
        self.booked_date_edit       = QLineEdit(project.booked_date if project else "")
        self.group_ops_mgr_edit     = QLineEdit(project.group_ops_manager if project else "")
        self.group_ops_sup_edit     = QLineEdit(project.group_ops_supervisor if project else "")
        self.job_subtype_edit       = QLineEdit(project.job_subtype if project else "")
        self.owner_edit             = QLineEdit(project.owner if project else "")
        self.contracted_with_edit   = QLineEdit(project.contracted_with if project else "")
        self.general_contractor_edit= QLineEdit(project.general_contractor if project else "")
        self.contract_value_edit    = QLineEdit(project.contract_value if project else "")
        self.job_docs_edit          = QLineEdit(project.job_docs if project else "")
        self.div25_url_edit         = QLineEdit(project.div25_url if project else "")

        self.template_combo = QComboBox()
        self.template_combo.addItem("Standard", "standard")
        self.template_combo.addItem("Phoenix", "phoenix")
        # Only shown when creating a new project (no existing project passed in)
        self.template_combo.setVisible(project is None)

        form_layout = QFormLayout()
        if project is None:
            form_layout.addRow("Task Template",       self.template_combo)
        form_layout.addRow("Job name *",          self.job_name_edit)
        form_layout.addRow("Job number *",        self.job_number_edit)
        form_layout.addRow("Project manager",     self.pm_edit)
        form_layout.addRow("Sales engineer",      self.sales_edit)
        form_layout.addRow("Target completion",   completion_row)
        form_layout.addRow("Liquid damages",      self.liquid_damages_edit)
        form_layout.addRow("Warranty period",     self.warranty_edit)
        form_layout.addRow("Booked date",         self.booked_date_edit)
        form_layout.addRow("Contract value",      self.contract_value_edit)
        form_layout.addRow("Div25 URL",           self.div25_url_edit)
        form_layout.addRow("Group ops manager",   self.group_ops_mgr_edit)
        form_layout.addRow("Group ops supervisor",self.group_ops_sup_edit)
        form_layout.addRow("Job sub-type",        self.job_subtype_edit)
        form_layout.addRow("Owner",               self.owner_edit)
        form_layout.addRow("Contracted with",     self.contracted_with_edit)
        form_layout.addRow("General contractor",  self.general_contractor_edit)
        form_layout.addRow("Job docs path",       self.job_docs_edit)
        form_layout.addRow("Notes",               self.notes_edit)

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
            booked_date=self.booked_date_edit.text().strip(),
            group_ops_manager=self.group_ops_mgr_edit.text().strip(),
            group_ops_supervisor=self.group_ops_sup_edit.text().strip(),
            job_subtype=self.job_subtype_edit.text().strip(),
            owner=self.owner_edit.text().strip(),
            contracted_with=self.contracted_with_edit.text().strip(),
            general_contractor=self.general_contractor_edit.text().strip(),
            contract_value=self.contract_value_edit.text().strip(),
            job_docs=self.job_docs_edit.text().strip(),
            div25_url=self.div25_url_edit.text().strip(),
        )

    def get_template(self) -> str:
        return self.template_combo.currentData()

    def accept(self) -> None:
        if not self.job_name_edit.text().strip():
            QMessageBox.warning(self, "Missing job name", "Enter a job name before saving.")
            return
        if not self.job_number_edit.text().strip():
            QMessageBox.warning(self, "Missing job number", "Enter a job number before saving.")
            return
        super().accept()


# ── Phoenix component helpers ──────────────────────────────────────────────────

class PrimaryButton(QPushButton):
    """Red primary-action button per Phoenix design system."""
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(36)
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class SecondaryButton(QPushButton):
    """Blue secondary-action button per Phoenix design system."""
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("secondaryButton")
        self.setMinimumHeight(36)
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class TertiaryButton(QPushButton):
    """Outline low-emphasis button per Phoenix design system."""
    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setObjectName("tertiaryButton")
        self.setMinimumHeight(36)
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class PhoenixTable(QTableWidget):
    """Read-only display table with Phoenix styling defaults."""
    def __init__(self, rows: int, cols: int, parent=None):
        super().__init__(rows, cols, parent)
        self.verticalHeader().setVisible(False)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setAlternatingRowColors(True)


# ── Drag-reorder table ────────────────────────────────────────────────────────

class ReorderableTaskTable(QTableWidget):
    """QTableWidget with internal row drag-drop that emits rowsReordered after a drop."""

    rowsReordered = Signal()

    def __init__(self, rows: int, cols: int, parent: Optional[QWidget] = None) -> None:
        super().__init__(rows, cols, parent)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setDragDropOverwriteMode(False)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)

    def dropEvent(self, event: Any) -> None:  # type: ignore[override]
        super().dropEvent(event)
        self.rowsReordered.emit()


class TaskDialog(QDialog):
    def __init__(self, parent: Optional[QWidget] = None, task: Optional[TaskRecord] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Task Details")
        self.setModal(True)
        self.resize(440, 300)

        self.task_name_edit = QLineEdit(task.task_name if task else "")
        self.phase_combo = QComboBox()
        self.phase_combo.addItems(PHASES)
        if task and task.phase:
            phase_index = self.phase_combo.findText(task.phase)
            if phase_index >= 0:
                self.phase_combo.setCurrentIndex(phase_index)

        self.due_date_check = QCheckBox("Has due date")
        self.due_date_edit = QDateEdit()
        self.due_date_edit.setCalendarPopup(True)
        self.due_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.due_date_edit.setDate(QDate.currentDate())
        if task and task.due_date:
            self.due_date_check.setChecked(True)
            parsed = QDate.fromString(task.due_date, "yyyy-MM-dd")
            if parsed.isValid():
                self.due_date_edit.setDate(parsed)
        else:
            self.due_date_edit.setEnabled(False)
        self.due_date_check.toggled.connect(self.due_date_edit.setEnabled)

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
        form_layout.addRow("Due date", self.due_date_check)
        form_layout.addRow("", self.due_date_edit)
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
        has_due = self.due_date_check.isChecked()
        return {
            "task_name": self.task_name_edit.text().strip(),
            "phase": self.phase_combo.currentText(),
            "due_date": self.due_date_edit.date().toString("yyyy-MM-dd") if has_due else None,
            "is_complete": bool(completed),
            "completed_date": self.completed_date_edit.date().toString("yyyy-MM-dd") if completed else None,
            "notes": self.notes_edit.toPlainText().strip(),
        }

    def accept(self) -> None:
        if not self.task_name_edit.text().strip():
            QMessageBox.warning(self, "Missing task name", "Enter a task name before saving.")
            return
        super().accept()


class NoteDialog(QDialog):
    """Add or edit a single note."""

    def __init__(self, parent: Optional[QWidget] = None,
                 note: Optional[NoteRecord] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Edit Note" if note else "Add Note")
        self.setModal(True)
        self.resize(520, 280)

        from PySide6.QtCore import QDate as _QDate
        today = _QDate.currentDate().toString("yyyy-MM-dd")

        self.date_edit     = QLineEdit(note.date if note else today)
        self.content_edit  = QPlainTextEdit(note.content if note else "")
        self.status_combo  = QComboBox()
        self.status_combo.addItems(["Open", "Closed"])
        if note and note.status:
            self.status_combo.setCurrentText(note.status)
        self.closeout_edit = QPlainTextEdit(note.closeout_comment if note else "")

        form = QFormLayout()
        form.addRow("Date",             self.date_edit)
        form.addRow("Note / Comment",   self.content_edit)
        form.addRow("Status",           self.status_combo)
        form.addRow("Closeout Comment", self.closeout_edit)

        btns = QDialogButtonBox()
        btns.addButton(QDialogButtonBox.StandardButton.Ok)
        btns.addButton(QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        lay = QVBoxLayout(self)
        lay.addLayout(form)
        lay.addWidget(btns)

    def get_data(self) -> dict:
        return {
            "date":              self.date_edit.text().strip(),
            "content":           self.content_edit.toPlainText().strip(),
            "status":            self.status_combo.currentText(),
            "closeout_comment":  self.closeout_edit.toPlainText().strip(),
        }

    def accept(self) -> None:
        if not self.content_edit.toPlainText().strip():
            QMessageBox.warning(self, "Empty note", "Enter a note before saving.")
            return
        super().accept()


class NotesWindow(QDialog):
    """Floating notes log for a project — matches the Excel Job Progress Notes layout."""

    def __init__(self, project_id: int, project_name: str,
                 backend: Any, view_only: bool = False,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.project_id   = project_id
        self.backend      = backend
        self._view_only   = view_only
        self.setWindowTitle(f"Job Progress Notes — {project_name}")
        self.resize(1000, 560)
        self.setMinimumSize(700, 400)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # Toolbar
        toolbar = QHBoxLayout()
        title_lbl = QLabel("Job Progress Notes")
        title_lbl.setObjectName("SectionTitle")
        toolbar.addWidget(title_lbl)
        toolbar.addStretch()
        if not view_only:
            add_btn = PrimaryButton("+ Add Note")
            add_btn.setFixedWidth(110)
            add_btn.clicked.connect(self._add_note)
            toolbar.addWidget(add_btn)
        layout.addLayout(toolbar)

        # Table
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["#", "Date", "Note or Comment", "Status", "Closeout Comment"]
        )
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.table.setAlternatingRowColors(False)
        if not view_only:
            self.table.doubleClicked.connect(self._edit_selected)

        hdr = self.table.horizontalHeader()
        hdr.setSectionsMovable(False)
        for c in range(5):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        hdr.resizeSection(0, 40)
        hdr.resizeSection(1, 100)
        hdr.resizeSection(2, 320)
        hdr.resizeSection(3, 100)
        hdr.resizeSection(4, 280)
        hdr.setStretchLastSection(False)
        layout.addWidget(self.table, 1)

        # Bottom buttons
        btn_row = QHBoxLayout()
        if not view_only:
            edit_btn = PrimaryButton("Edit")
            edit_btn.setFixedWidth(90)
            edit_btn.clicked.connect(self._edit_selected)
            del_btn = PrimaryButton("Delete")
            del_btn.setFixedWidth(90)
            del_btn.clicked.connect(self._delete_selected)
            btn_row.addWidget(edit_btn)
            btn_row.addWidget(del_btn)
        close_btn = TertiaryButton("Close")
        close_btn.setFixedWidth(90)
        close_btn.clicked.connect(self.accept)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._refresh()

    def _refresh(self) -> None:
        notes = self.backend.list_notes(self.project_id)
        self.table.setRowCount(len(notes))
        for r, note in enumerate(notes):
            is_open = note.status == "Open"
            bg = QColor("#FFCCCC") if is_open else QColor("#CCFFCC")
            status_color = QColor("#CC0000") if is_open else QColor("#006400")

            items = [
                QTableWidgetItem(str(note.note_number)),
                QTableWidgetItem(note.date),
                QTableWidgetItem(note.content),
                QTableWidgetItem(note.status),
                QTableWidgetItem(note.closeout_comment),
            ]
            for c, item in enumerate(items):
                item.setBackground(bg)
                if c == 3:
                    item.setForeground(status_color)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                # Show full text as tooltip for content and closeout columns
                if c == 2 and note.content:
                    item.setToolTip(note.content)
                if c == 4 and note.closeout_comment:
                    item.setToolTip(note.closeout_comment)
                item.setData(Qt.ItemDataRole.UserRole, note.id)
                self.table.setItem(r, c, item)

    def _selected_id(self) -> Optional[int]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _add_note(self) -> None:
        if self._view_only:
            return
        dlg = NoteDialog(self)
        if dlg.exec() != int(QDialog.DialogCode.Accepted):
            return
        d = dlg.get_data()
        self.backend.add_note(self.project_id, d["content"], d["date"],
                              d["status"], d["closeout_comment"])
        self._refresh()

    def _edit_selected(self) -> None:
        if self._view_only:
            return
        note_id = self._selected_id()
        if note_id is None:
            return
        note = next((n for n in self.backend.list_notes(self.project_id)
                     if n.id == note_id), None)
        if note is None:
            return
        dlg = NoteDialog(self, note)
        if dlg.exec() != int(QDialog.DialogCode.Accepted):
            return
        self.backend.update_note(note_id, **dlg.get_data())
        self._refresh()

    def _delete_selected(self) -> None:
        if self._view_only:
            return
        note_id = self._selected_id()
        if note_id is None:
            return
        ans = QMessageBox.question(self, "Delete note", "Delete this note?",
                                   QMessageBox.StandardButton.Yes |
                                   QMessageBox.StandardButton.No,
                                   QMessageBox.StandardButton.No)
        if ans == QMessageBox.StandardButton.Yes:
            self.backend.delete_note(note_id)
            self._refresh()


class ChangeOrderDialog(QDialog):
    """Add or edit a single Change Order."""

    def __init__(self, parent: Optional[QWidget] = None,
                 co: Optional[ChangeOrderRecord] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Edit Change Order" if co else "Add Change Order")
        self.setModal(True)
        self.resize(600, 520)

        def _le(val: str = "") -> QLineEdit:
            return QLineEdit(val)

        def _combo(options: list[str], current: str = "") -> QComboBox:
            cb = QComboBox()
            cb.addItems(options)
            if current:
                cb.setCurrentText(current)
            return cb

        status_opts = ["Pending", "Accepted", "Rejected"]

        self.cop_number      = _le(co.cop_number      if co else "")
        self.reference       = _le(co.reference       if co else "")
        self.description     = _le(co.description     if co else "")
        self.creation_date   = _le(co.creation_date   if co else "")
        self.ats_price       = _le(co.ats_price       if co else "")
        self.ats_direct_cost = _le(co.ats_direct_cost if co else "")
        self.ats_status      = _combo(status_opts, co.ats_status if co else "Pending")
        self.booked_in_portal= _le(co.booked_in_portal if co else "")
        self.ats_booked_co   = _le(co.ats_booked_co  if co else "")
        self.mech_co         = _le(co.mech_co         if co else "")
        self.sub_quoted_price= _le(co.sub_quoted_price if co else "")
        self.sub_plug_number = _le(co.sub_plug_number if co else "")
        self.sub_status      = _combo(status_opts, co.sub_status if co else "Pending")
        self.sub_co_sent     = _le(co.sub_co_sent     if co else "")
        self.sub_co_number   = _le(co.sub_co_number   if co else "")
        self.notes_edit      = QPlainTextEdit(co.notes if co else "")
        self.notes_edit.setFixedHeight(60)

        form = QFormLayout()
        form.addRow("COP #",             self.cop_number)
        form.addRow("Reference",          self.reference)
        form.addRow("Description",        self.description)
        form.addRow("Creation Date",      self.creation_date)
        form.addRow("ATS Price",          self.ats_price)
        form.addRow("ATS Direct Cost",    self.ats_direct_cost)
        form.addRow("ATS Status",         self.ats_status)
        form.addRow("Booked in Portal",   self.booked_in_portal)
        form.addRow("ATS Booked CO#",     self.ats_booked_co)
        form.addRow("Mech CO#",           self.mech_co)
        form.addRow("Sub Quoted Price",   self.sub_quoted_price)
        form.addRow("Sub Plug #",         self.sub_plug_number)
        form.addRow("Sub Status",         self.sub_status)
        form.addRow("Sub CO Sent",        self.sub_co_sent)
        form.addRow("Sub CO#",            self.sub_co_number)
        form.addRow("Notes",              self.notes_edit)

        btns = QDialogButtonBox()
        btns.addButton(QDialogButtonBox.StandardButton.Ok)
        btns.addButton(QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        inner.setLayout(form)
        scroll.setWidget(inner)

        lay = QVBoxLayout(self)
        lay.addWidget(scroll, 1)
        lay.addWidget(btns)

    def get_data(self) -> ChangeOrderRecord:
        return ChangeOrderRecord(
            cop_number      = self.cop_number.text().strip(),
            reference       = self.reference.text().strip(),
            description     = self.description.text().strip(),
            creation_date   = self.creation_date.text().strip(),
            ats_price       = self.ats_price.text().strip(),
            ats_direct_cost = self.ats_direct_cost.text().strip(),
            ats_status      = self.ats_status.currentText(),
            booked_in_portal= self.booked_in_portal.text().strip(),
            ats_booked_co   = self.ats_booked_co.text().strip(),
            mech_co         = self.mech_co.text().strip(),
            sub_quoted_price= self.sub_quoted_price.text().strip(),
            sub_plug_number = self.sub_plug_number.text().strip(),
            sub_status      = self.sub_status.currentText(),
            sub_co_sent     = self.sub_co_sent.text().strip(),
            sub_co_number   = self.sub_co_number.text().strip(),
            notes           = self.notes_edit.toPlainText().strip(),
        )


class ChangeOrderWindow(QDialog):
    """Floating Change Order tracker window."""

    # Status background colors
    _STATUS_BG = {
        "Accepted": QColor("#CCFFCC"),
        "Rejected": QColor("#FFCCCC"),
        "Pending":  QColor("#FFFACD"),
    }

    def __init__(self, project_id: int, project_name: str,
                 backend: Any, view_only: bool = False,
                 parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.project_id = project_id
        self.backend    = backend
        self._view_only = view_only
        self.setWindowTitle(f"Change Order Log — {project_name}")
        self.resize(1200, 600)
        self.setMinimumSize(900, 400)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        # ── Summary bar ───────────────────────────────────────────────────────
        self._summary_bar = QWidget()
        sb_layout = QHBoxLayout(self._summary_bar)
        sb_layout.setContentsMargins(0, 0, 0, 0)
        sb_layout.setSpacing(4)

        _tooltips = {
            "ats_base":    "Base Price — sum of all ATS Price entries",
            "ats_current": "Base Price + all Accepted ATS change orders",
        }
        self._sum_labels: dict[str, QLabel] = {}
        for key, caption in [
            ("ats_base",     "ATS Base"),
            ("ats_accepted", "ATS Accepted"),
            ("ats_pending",  "ATS Pending"),
            ("ats_current",  "ATS Current"),
            ("sub_accepted", "Sub Accepted"),
            ("sub_pending",  "Sub Pending"),
            ("sub_current",  "Sub Current"),
        ]:
            box = QFrame()
            box.setObjectName("StatCard")
            bl = QVBoxLayout(box)
            bl.setContentsMargins(8, 4, 8, 4)
            bl.setSpacing(1)
            cap = QLabel(caption)
            cap.setObjectName("MetaCaption")
            if key in _tooltips:
                box.setToolTip(_tooltips[key])
                cap.setToolTip(_tooltips[key])
            val = QLabel("$0")
            val.setObjectName("StatValue")
            bl.addWidget(cap)
            bl.addWidget(val)
            self._sum_labels[key] = val
            sb_layout.addWidget(box)

        layout.addWidget(self._summary_bar)

        # ── Toolbar ───────────────────────────────────────────────────────────
        toolbar = QHBoxLayout()
        toolbar.addStretch()
        if not view_only:
            add_btn = PrimaryButton("+ Add CO")
            add_btn.setFixedWidth(100)
            add_btn.clicked.connect(self._add_co)
            toolbar.addWidget(add_btn)
        layout.addLayout(toolbar)

        # ── Table ─────────────────────────────────────────────────────────────
        headers = [
            "COP#", "Reference", "Description", "Date",
            "ATS Price", "Direct Cost", "ATS Status", "Portal",
            "Booked CO#", "Mech CO#",
            "Sub Quoted", "Plug #", "Sub Price",
            "Sub Status", "CO Sent", "Sub CO#", "Notes",
        ]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        if not view_only:
            self.table.doubleClicked.connect(self._edit_selected)

        hdr = self.table.horizontalHeader()
        hdr.setSectionsMovable(False)
        for c in range(17):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Interactive)
        hdr.resizeSection(0, 70)   # COP#
        hdr.resizeSection(1, 90)   # Reference
        hdr.resizeSection(2, 260)  # Description
        hdr.resizeSection(3, 90)   # Date
        hdr.resizeSection(4, 90)   # ATS Price
        hdr.resizeSection(5, 90)   # Direct Cost
        hdr.resizeSection(6, 90)   # ATS Status
        hdr.resizeSection(7, 70)   # Portal
        hdr.resizeSection(8, 90)   # Booked CO#
        hdr.resizeSection(9, 80)   # Mech CO#
        hdr.resizeSection(10, 90)  # Sub Quoted
        hdr.resizeSection(11, 70)  # Plug #
        hdr.resizeSection(12, 90)  # Sub Price
        hdr.resizeSection(13, 90)  # Sub Status
        hdr.resizeSection(14, 75)  # CO Sent
        hdr.resizeSection(15, 75)  # Sub CO#
        hdr.resizeSection(16, 200) # Notes
        hdr.setStretchLastSection(False)
        layout.addWidget(self.table, 1)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        if not view_only:
            edit_btn = PrimaryButton("Edit");   edit_btn.setFixedWidth(80);  edit_btn.clicked.connect(self._edit_selected)
            del_btn  = PrimaryButton("Delete"); del_btn.setFixedWidth(80);   del_btn.clicked.connect(self._delete_selected)
            btn_row.addWidget(edit_btn)
            btn_row.addWidget(del_btn)
        close_btn = TertiaryButton("Close"); close_btn.setFixedWidth(80); close_btn.clicked.connect(self.accept)
        btn_row.addStretch()
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self._refresh()

    def _refresh(self) -> None:
        cos = self.backend.list_change_orders(self.project_id)
        summary = self.backend.get_co_summary(self.project_id)

        # Update summary bar
        for key, lbl in self._sum_labels.items():
            lbl.setText(f"${summary[key]:,.0f}")

        self.table.setRowCount(len(cos))
        for r, co in enumerate(cos):
            sub_price = co.sub_quoted_price if co.sub_quoted_price else co.sub_plug_number
            ats_bg  = self._STATUS_BG.get(co.ats_status, QColor("#FFFFFF"))
            sub_bg  = self._STATUS_BG.get(co.sub_status, QColor("#FFFFFF"))
            def _fmt_price(v: str) -> str:
                if not v:
                    return ""
                try:
                    return f"${float(v):,.2f}"
                except ValueError:
                    return v

            row_vals = [
                co.cop_number, co.reference, co.description, co.creation_date,
                _fmt_price(co.ats_price), _fmt_price(co.ats_direct_cost),
                co.ats_status, co.booked_in_portal,
                co.ats_booked_co, co.mech_co,
                co.sub_quoted_price, co.sub_plug_number, _fmt_price(sub_price),
                co.sub_status, co.sub_co_sent, co.sub_co_number, co.notes,
            ]
            for c, val in enumerate(row_vals):
                item = QTableWidgetItem(str(val))
                bg = ats_bg if c <= 9 else (sub_bg if c <= 15 else QColor("#FFFFFF"))
                item.setBackground(bg)
                if val:
                    item.setToolTip(str(val))
                item.setData(Qt.ItemDataRole.UserRole, co.id)
                self.table.setItem(r, c, item)

    def _selected_id(self) -> Optional[int]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _add_co(self) -> None:
        if self._view_only:
            return
        dlg = ChangeOrderDialog(self)
        if dlg.exec() != int(QDialog.DialogCode.Accepted):
            return
        self.backend.add_change_order(self.project_id, dlg.get_data())
        self._refresh()

    def _edit_selected(self) -> None:
        if self._view_only:
            return
        co_id = self._selected_id()
        if co_id is None:
            return
        co = next((c for c in self.backend.list_change_orders(self.project_id)
                   if c.id == co_id), None)
        if co is None:
            return
        dlg = ChangeOrderDialog(self, co)
        if dlg.exec() != int(QDialog.DialogCode.Accepted):
            return
        self.backend.update_change_order(co_id, dlg.get_data())
        self._refresh()

    def _delete_selected(self) -> None:
        if self._view_only:
            return
        co_id = self._selected_id()
        if co_id is None:
            return
        ans = QMessageBox.question(self, "Delete CO", "Delete this change order?",
                                   QMessageBox.StandardButton.Yes |
                                   QMessageBox.StandardButton.No,
                                   QMessageBox.StandardButton.No)
        if ans == QMessageBox.StandardButton.Yes:
            self.backend.delete_change_order(co_id)
            self._refresh()


class StatCard(QFrame):
    def __init__(self, title: str, value: str = "0") -> None:
        super().__init__()
        self.setObjectName("StatCard")
        self.title_label = QLabel(title)
        self.title_label.setObjectName("StatTitle")
        self.value_label = QLabel(value)
        self.value_label.setObjectName("StatValue")

        card_layout = QVBoxLayout(self)
        card_layout.setContentsMargins(8, 6, 8, 6)
        card_layout.setSpacing(2)
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
        self._rects: list[tuple[int, int, dict]] = []  # (x, width, seg) for hit-testing
        self.setMouseTracking(True)

    def set_segments(self, segments: list[dict]) -> None:
        self._segments = [
            {**s, "color": QColor(PHASE_COLORS.get(s["phase"], "#487cff"))}
            for s in segments if s["total"] > 0
        ]
        self.update()

    def clear(self) -> None:
        self._segments = []
        self._rects = []
        self.update()

    def mouseMoveEvent(self, event) -> None:
        mx = event.position().x()
        for x, w, seg in self._rects:
            if x <= mx <= x + w:
                pct = int(round(seg["done"] / seg["total"] * 100)) if seg["total"] else 0
                QToolTip.showText(
                    event.globalPosition().toPoint(),
                    f"{seg['phase']}\n{seg['done']} / {seg['total']} complete ({pct}%)",
                    self,
                )
                return
        QToolTip.hideText()

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

        self._rects = rects  # save for mouseMoveEvent hit-testing

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
            notes_btn = TertiaryButton("What's new?")
            notes_btn.setFixedWidth(120)
            notes_btn.clicked.connect(lambda: QMessageBox.information(
                self, f"What's new in v{info.latest_version}",
                info.release_notes,
            ))
            layout.addWidget(notes_btn)

        install_btn = QPushButton("Install && Restart")
        install_btn.setFixedWidth(150)
        install_btn.setObjectName("InstallBtn")
        install_btn.clicked.connect(self.install_clicked)
        layout.addWidget(install_btn)

        dismiss_btn = TertiaryButton("✕")
        dismiss_btn.setFixedWidth(44)
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


class _HeaderResizeHandle(QFrame):
    """Drag handle between project title and meta fields in the header row."""
    _MIN_TITLE_W = 80
    _MAX_TITLE_W = 700

    def __init__(self, title_widget: QWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._title_widget = title_widget
        self._drag_start_x: Optional[float] = None
        self._drag_start_w: Optional[int] = None
        self.setFixedWidth(6)
        self.setObjectName("ResizeHandle")
        self.setCursor(QCursor(Qt.CursorShape.SplitHCursor))

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start_x = event.globalPosition().x()
            self._drag_start_w = self._title_widget.width()

    def mouseMoveEvent(self, event) -> None:
        if self._drag_start_x is None:
            return
        delta = event.globalPosition().x() - self._drag_start_x
        new_w = int(self._drag_start_w + delta)
        new_w = max(self._MIN_TITLE_W, min(self._MAX_TITLE_W, new_w))
        self._title_widget.setFixedWidth(new_w)

    def mouseReleaseEvent(self, event) -> None:
        self._drag_start_x = None
        self._drag_start_w = None


class _VResizeHandle(QFrame):
    """Horizontal drag strip between the project header and task table.
    Drag up to grow the header, drag down to shrink it."""

    _MIN_H = 80
    _MAX_H = 300

    def __init__(self, header: QWidget, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._header = header
        self._drag_start_y: Optional[float] = None
        self._drag_start_h: Optional[int] = None
        self.setFixedHeight(6)
        self.setObjectName("VResizeHandle")
        self.setCursor(QCursor(Qt.CursorShape.SizeVerCursor))

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start_y = event.globalPosition().y()
            self._drag_start_h = self._header.height()

    def mouseMoveEvent(self, event) -> None:
        if self._drag_start_y is None:
            return
        delta = event.globalPosition().y() - self._drag_start_y
        new_h = int(self._drag_start_h + delta)
        new_h = max(self._MIN_H, min(self._MAX_H, new_h))
        self._header.setFixedHeight(new_h)

    def mouseReleaseEvent(self, event) -> None:
        self._drag_start_y = None
        self._drag_start_h = None


class DataLocationDialog(QDialog):
    """Dialog for configuring where the shared data file lives."""

    def __init__(self, current_folder: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Data File Location")
        self.setModal(True)
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        info = QLabel(
            "Set the folder where the shared project database is stored.\n"
            "Point all users to the same synced folder (e.g. SharePoint / OneDrive)\n"
            "so everyone works from the same data."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        path_row = QHBoxLayout()
        self.path_edit = QLineEdit(current_folder)
        self.path_edit.setPlaceholderText("e.g. C:\\Users\\you\\ATS Inc\\Phoenix - ATS Job Tracker")
        browse_btn = SecondaryButton("Browse...")
        browse_btn.clicked.connect(self._browse)
        path_row.addWidget(self.path_edit, 1)
        path_row.addWidget(browse_btn)
        layout.addLayout(path_row)

        self.status_label = QLabel("")
        self.status_label.setObjectName("errorLabel")
        layout.addWidget(self.status_label)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)  # type: ignore[arg-type]
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Shared Data Folder", self.path_edit.text())
        if folder:
            self.path_edit.setText(folder)

    def _accept(self) -> None:
        folder = self.path_edit.text().strip()
        if folder and not Path(folder).exists():
            self.status_label.setText(f"Folder does not exist: {folder}")
            return
        self.accept()

    def selected_folder(self) -> str:
        return self.path_edit.text().strip()


class LoginDialog(QDialog):
    """Prompt the user to log in with username + password."""

    def __init__(self, user_manager: UserManager, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._user_manager = user_manager
        self._logged_in_user: Optional[UserRecord] = None
        self.setWindowTitle("Log In")
        self.setModal(True)
        self.setFixedWidth(340)

        form = QFormLayout()
        self._username_edit = QLineEdit()
        self._username_edit.setPlaceholderText("Username")
        self._password_edit = QLineEdit()
        self._password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self._password_edit.setPlaceholderText("Password")
        self._remember_check = QCheckBox("Remember me on this computer")
        form.addRow("Username:", self._username_edit)
        form.addRow("Password:", self._password_edit)

        self._error_lbl = QLabel("")
        self._error_lbl.setObjectName("errorLabel")
        self._error_lbl.setWordWrap(True)
        self._error_lbl.setVisible(False)

        login_btn = PrimaryButton("Log In")
        login_btn.setDefault(True)
        login_btn.clicked.connect(self._attempt_login)
        self._password_edit.returnPressed.connect(self._attempt_login)

        forgot_lbl = QLabel("Forgot your password? Contact your administrator.")
        forgot_lbl.setObjectName("hint")
        forgot_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(self._remember_check)
        layout.addWidget(self._error_lbl)
        layout.addWidget(login_btn)
        layout.addWidget(forgot_lbl)

    def _attempt_login(self) -> None:
        username = self._username_edit.text().strip()
        password = self._password_edit.text()
        user = self._user_manager.authenticate(username, password)
        if user is None:
            self._error_lbl.setText("Incorrect username or password.")
            self._error_lbl.setVisible(True)
            self._password_edit.clear()
            return
        self._logged_in_user = user
        if self._remember_check.isChecked():
            _SESSION_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(_SESSION_PATH, "w", encoding="utf-8") as f:
                json.dump({"username": user.username}, f)
        else:
            _SESSION_PATH.unlink(missing_ok=True)
        self.accept()

    def logged_in_user(self) -> Optional[UserRecord]:
        return self._logged_in_user


class ChangePasswordDialog(QDialog):
    """Force a user to pick a new password (shown when must_change_password is True)."""

    def __init__(self, user_manager: UserManager, username: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._user_manager = user_manager
        self._username = username
        self.setWindowTitle("Change Password")
        self.setModal(True)
        self.setFixedWidth(360)

        self._new_edit = QLineEdit()
        self._new_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self._new_edit.setPlaceholderText("New password (8+ characters)")
        self._confirm_edit = QLineEdit()
        self._confirm_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self._confirm_edit.setPlaceholderText("Confirm new password")

        self._error_lbl = QLabel("")
        self._error_lbl.setObjectName("errorLabel")
        self._error_lbl.setWordWrap(True)
        self._error_lbl.setVisible(False)

        save_btn = PrimaryButton("Save Password")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save)

        form = QFormLayout()
        form.addRow("New password:", self._new_edit)
        form.addRow("Confirm:", self._confirm_edit)

        layout = QVBoxLayout(self)
        intro = QLabel(f"Hi {username}! You must set a new password before continuing.")
        intro.setWordWrap(True)
        layout.addWidget(intro)
        layout.addLayout(form)
        layout.addWidget(self._error_lbl)
        layout.addWidget(save_btn)

    def _save(self) -> None:
        pw = self._new_edit.text()
        confirm = self._confirm_edit.text()
        if pw != confirm:
            self._error_lbl.setText("Passwords do not match.")
            self._error_lbl.setVisible(True)
            return
        if len(pw) < UserManager.MIN_PASSWORD_LENGTH:
            self._error_lbl.setText(f"Password must be at least {UserManager.MIN_PASSWORD_LENGTH} characters.")
            self._error_lbl.setVisible(True)
            return
        try:
            self._user_manager.change_password(self._username, pw)
        except Exception as exc:
            self._error_lbl.setText(str(exc))
            self._error_lbl.setVisible(True)
            return
        self.accept()


class SelfChangePasswordDialog(QDialog):
    """Let any logged-in user change their own password (requires current password)."""

    def __init__(self, user_manager: UserManager, username: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._user_manager = user_manager
        self._username = username
        self.setWindowTitle("Change My Password")
        self.setModal(True)
        self.setFixedWidth(380)

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"Changing password for: <b>{username}</b>"))

        form = QFormLayout()
        self._current_pw = QLineEdit()
        self._current_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self._new_pw = QLineEdit()
        self._new_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self._confirm_pw = QLineEdit()
        self._confirm_pw.setEchoMode(QLineEdit.EchoMode.Password)
        form.addRow("Current password:", self._current_pw)
        form.addRow("New password:", self._new_pw)
        form.addRow("Confirm new password:", self._confirm_pw)
        layout.addLayout(form)

        self._error_lbl = QLabel("")
        self._error_lbl.setObjectName("errorLabel")
        self._error_lbl.setWordWrap(True)
        self._error_lbl.setVisible(False)
        layout.addWidget(self._error_lbl)

        save_btn = PrimaryButton("Save Password")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self._save)
        layout.addWidget(save_btn)

    def _save(self) -> None:
        current = self._current_pw.text()
        new_pw = self._new_pw.text()
        confirm = self._confirm_pw.text()

        if self._user_manager.authenticate(self._username, current) is None:
            self._error_lbl.setText("Current password is incorrect.")
            self._error_lbl.setVisible(True)
            self._current_pw.clear()
            return
        if new_pw != confirm:
            self._error_lbl.setText("New passwords do not match.")
            self._error_lbl.setVisible(True)
            return
        try:
            self._user_manager.change_password(self._username, new_pw)
            QMessageBox.information(self, "Done", "Password changed successfully.")
            self.accept()
        except ValueError as exc:
            self._error_lbl.setText(str(exc))
            self._error_lbl.setVisible(True)


class ManageUsersDialog(QDialog):
    """Admin UI for creating, listing, and managing user accounts."""

    _ROLE_LABELS = ["User", "View Only", "Admin"]
    _ROLE_KEYS = ["user", "view_only", "admin"]

    def __init__(self, user_manager: UserManager, current_user: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._user_manager = user_manager
        self._current_user = current_user
        self.setWindowTitle("Manage Users")
        self.setModal(True)
        self.setMinimumWidth(460)

        layout = QVBoxLayout(self)

        # User list (shows username + role)
        layout.addWidget(QLabel("Users:"))
        self._user_list = QListWidget()
        layout.addWidget(self._user_list)

        # Add user row
        add_row = QHBoxLayout()
        self._new_username = QLineEdit()
        self._new_username.setPlaceholderText("New username")
        self._new_password = QLineEdit()
        self._new_password.setEchoMode(QLineEdit.EchoMode.Password)
        self._new_password.setPlaceholderText("Temp password (8+ chars)")
        self._new_role_combo = QComboBox()
        for label, key in zip(self._ROLE_LABELS, self._ROLE_KEYS):
            self._new_role_combo.addItem(label, key)
        add_btn = PrimaryButton("Add User")
        add_btn.clicked.connect(self._add_user)
        add_row.addWidget(self._new_username, 2)
        add_row.addWidget(self._new_password, 2)
        add_row.addWidget(self._new_role_combo, 1)
        add_row.addWidget(add_btn)
        layout.addLayout(add_row)

        self._error_lbl = QLabel("")
        self._error_lbl.setObjectName("errorLabel")
        self._error_lbl.setWordWrap(True)
        self._error_lbl.setVisible(False)
        layout.addWidget(self._error_lbl)

        # Action buttons
        btn_row = QHBoxLayout()
        del_btn = PrimaryButton("Delete Selected User")
        del_btn.clicked.connect(self._delete_user)
        reset_btn = PrimaryButton("Reset Password")
        reset_btn.setToolTip("Set a temporary password — user must change it on next login")
        reset_btn.clicked.connect(self._reset_password)
        role_btn = SecondaryButton("Change Role...")
        role_btn.clicked.connect(self._change_role)
        btn_row.addWidget(del_btn)
        btn_row.addWidget(reset_btn)
        btn_row.addWidget(role_btn)
        layout.addLayout(btn_row)

        close_btn = TertiaryButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        self._refresh_list()

    def _selected_username(self) -> Optional[str]:
        item = self._user_list.currentItem()
        if item is None:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _refresh_list(self) -> None:
        from user_auth import ROLE_LABELS
        self._user_list.clear()
        for username in self._user_manager.list_users():
            user = self._user_manager.get_user(username)
            role_label = ROLE_LABELS.get(user.role if user else "user", "User")
            item = QListWidgetItem(f"{username}  [{role_label}]")
            item.setData(Qt.ItemDataRole.UserRole, username)
            self._user_list.addItem(item)

    def _add_user(self) -> None:
        username = self._new_username.text().strip()
        password = self._new_password.text()
        role = self._new_role_combo.currentData() or "user"
        try:
            self._user_manager.create_user(username, password, must_change_password=True, role=role)
            self._new_username.clear()
            self._new_password.clear()
            self._new_role_combo.setCurrentIndex(0)
            self._error_lbl.setVisible(False)
            self._refresh_list()
        except ValueError as exc:
            self._error_lbl.setText(str(exc))
            self._error_lbl.setVisible(True)

    def _delete_user(self) -> None:
        username = self._selected_username()
        if username is None:
            return
        if username.casefold() == self._current_user.casefold():
            QMessageBox.warning(self, "Cannot Delete", "You cannot delete your own account.")
            return
        reply = QMessageBox.question(
            self, "Delete User",
            f"Delete user '{username}'? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            try:
                self._user_manager.delete_user(username)
            except ValueError as exc:
                QMessageBox.warning(self, "Cannot Delete", str(exc))
                return
            self._refresh_list()

    def _reset_password(self) -> None:
        username = self._selected_username()
        if username is None:
            return
        new_pw, ok = QInputDialog.getText(
            self, "Reset Password",
            f"Enter a temporary password for '{username}':",
            QLineEdit.EchoMode.Password,
        )
        if not ok or not new_pw:
            return
        try:
            self._user_manager.reset_password(username, new_pw)
            QMessageBox.information(self, "Done", f"Password for '{username}' has been reset.\nThey will be prompted to set a new one on next login.")
        except ValueError as exc:
            QMessageBox.warning(self, "Error", str(exc))

    def _change_role(self) -> None:
        username = self._selected_username()
        if username is None:
            return
        user = self._user_manager.get_user(username)
        if user is None:
            return
        # Prevent demoting the last admin
        if user.role == "admin":
            all_users = self._user_manager.list_users()
            admin_count = sum(
                1 for u in all_users
                if (rec := self._user_manager.get_user(u)) and rec.role == "admin"
            )
            if admin_count <= 1:
                QMessageBox.warning(
                    self, "Cannot Change Role",
                    "There must be at least one administrator. Assign another admin first."
                )
                return
        current_idx = self._ROLE_KEYS.index(user.role) if user.role in self._ROLE_KEYS else 0
        chosen, ok = QInputDialog.getItem(
            self, "Change Role",
            f"Select a new role for '{username}':",
            self._ROLE_LABELS, current_idx, False,
        )
        if not ok:
            return
        new_role = self._ROLE_KEYS[self._ROLE_LABELS.index(chosen)]
        self._user_manager.set_role(username, new_role)
        self._refresh_list()


class RecentChangesDialog(QDialog):
    """Collapsible tree showing what changed in the shared data since last open."""

    def __init__(self, changes: list[dict], since: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Recent Changes")
        self.setModal(True)
        self.setMinimumWidth(520)
        self.setMinimumHeight(320)
        self.resize(560, 400)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        since_lbl = QLabel(f"Changes since last opened:  {since}")
        since_lbl.setObjectName("hint")
        layout.addWidget(since_lbl)

        tree = QTreeWidget()
        tree.setHeaderHidden(True)
        tree.setRootIsDecorated(True)
        tree.setIndentation(20)
        tree.setAnimated(True)
        layout.addWidget(tree)

        for change in changes:
            is_new = change["type"] == "new"
            badge = "NEW" if is_new else "UPDATED"
            badge_color = "#4caf50" if is_new else "#fb923c"
            by = change.get("updated_by", "")
            by_suffix = f"  (by {by})" if by else ""
            label = f"[{badge}]  {change['job_number']}  —  {change['job_name']}{by_suffix}"
            top = QTreeWidgetItem(tree, [label])
            top.setForeground(0, QColor(badge_color))
            font = top.font(0)
            font.setBold(True)
            top.setFont(0, font)

            for field_label, old_val, new_val in change["fields"]:
                if is_new:
                    text = f"  {field_label}:  {new_val}"
                else:
                    old_disp = old_val if old_val else "(empty)"
                    new_disp = new_val if new_val else "(empty)"
                    text = f"  {field_label}:  {old_disp}  →  {new_disp}"
                child = QTreeWidgetItem(top, [text])
                child.setForeground(0, QColor("#cccccc"))

        tree.collapseAll()

        close_btn = TertiaryButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)


class ActivityLogDialog(QDialog):
    """Log of all activity for a project — newest entries first.
    Admin users see a Remove button on each row."""

    _ACTION_COLORS = {
        "created":     "#4caf50",
        "completed":   "#2196f3",
        "updated":     "#fb923c",
        "deleted":     "#ef5350",
        "uncompleted": "#9e9e9e",
    }

    def __init__(
        self,
        activities: list,
        project_name: str,
        backend: Any = None,
        is_admin: bool = False,
        parent: Optional[QWidget] = None,
    ) -> None:
        super().__init__(parent)
        self._backend = backend
        self._is_admin = is_admin
        self._activities = list(activities)
        self._project_name = project_name

        self.setWindowTitle(f"Activity Log — {project_name}")
        self.setMinimumWidth(680)
        self.setMinimumHeight(420)
        self.resize(800 if is_admin else 760, 500)

        self._layout = QVBoxLayout(self)
        self._layout.setSpacing(8)

        self._count_lbl = QLabel()
        self._count_lbl.setObjectName("hint")
        self._layout.addWidget(self._count_lbl)

        col_count = 6 if is_admin else 5
        self._table = PhoenixTable(0, col_count)
        headers = ["Timestamp", "User", "Action", "Item", "Details"]
        if is_admin:
            headers.append("")
        self._table.setHorizontalHeaderLabels(headers)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self._table.setColumnWidth(0, 155)
        self._table.setColumnWidth(1, 100)
        self._table.setColumnWidth(2, 90)
        if is_admin:
            self._table.setColumnWidth(5, 90)
            self._table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
            self._table.verticalHeader().setDefaultSectionSize(32)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.verticalHeader().setVisible(False)
        self._table.setAlternatingRowColors(True)
        self._layout.addWidget(self._table)

        close_btn = TertiaryButton("Close")
        close_btn.clicked.connect(self.accept)
        self._layout.addWidget(close_btn)

        self._populate()

    def _populate(self) -> None:
        self._table.setRowCount(0)
        self._count_lbl.setText(
            f"{len(self._activities)} event(s)" if self._activities else "No activity recorded yet."
        )
        for row, act in enumerate(self._activities):
            self._table.insertRow(row)
            color = QColor(self._ACTION_COLORS.get(act.action, "#888888"))
            entity_label = f"[{act.entity_type.title()}] {act.entity_name}"
            for col, val in enumerate([act.timestamp, act.user, act.action, entity_label, act.details]):
                item = QTableWidgetItem(str(val))
                if col == 2:
                    item.setForeground(color)
                self._table.setItem(row, col, item)

            if self._is_admin:
                remove_btn = SecondaryButton("Remove")
                remove_btn.setFixedSize(80, 26)
                remove_btn.setProperty("activity_id", act.id)
                remove_btn.clicked.connect(self._on_remove)
                self._table.setCellWidget(row, 5, remove_btn)

    def _on_remove(self) -> None:
        btn = self.sender()
        activity_id = btn.property("activity_id")
        reply = QMessageBox.question(
            self, "Remove Entry",
            "Remove this activity log entry? This cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        if self._backend is not None:
            self._backend.delete_activity(activity_id)
        self._activities = [a for a in self._activities if a.id != activity_id]
        self._populate()


class BulkExportDialog(QDialog):
    """Let the user pick multiple projects to export into one Excel workbook."""

    def __init__(self, projects: list, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Bulk Export to Excel")
        self.setMinimumWidth(460)
        self.setMinimumHeight(380)

        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        layout.addWidget(QLabel("Select projects to export:"))

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        for proj in projects:
            display = f"{proj.job_number}  —  {proj.job_name}" if proj.job_number else proj.job_name
            item = QListWidgetItem(display)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            item.setData(Qt.ItemDataRole.UserRole, proj.id)
            self._list.addItem(item)
        layout.addWidget(self._list)

        btn_row = QHBoxLayout()
        sel_all = SecondaryButton("Select All")
        sel_none = SecondaryButton("Select None")
        sel_all.clicked.connect(self._select_all)
        sel_none.clicked.connect(self._select_none)
        btn_row.addWidget(sel_all)
        btn_row.addWidget(sel_none)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel  # type: ignore[arg-type]
        )
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _select_all(self) -> None:
        for i in range(self._list.count()):
            self._list.item(i).setCheckState(Qt.CheckState.Checked)

    def _select_none(self) -> None:
        for i in range(self._list.count()):
            self._list.item(i).setCheckState(Qt.CheckState.Unchecked)

    def selected_ids(self) -> list[int]:
        ids = []
        for i in range(self._list.count()):
            item = self._list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                ids.append(item.data(Qt.ItemDataRole.UserRole))
        return ids


class MainWindow(QMainWindow):
    _update_ready = Signal()  # emitted from bg thread when a new version is found

    def __init__(self, current_user: str = "") -> None:
        super().__init__()
        self._update_ready.connect(self._show_update_banner)
        self._current_user = current_user
        self.backend = ProjectTrackerBackend(self._resolve_data_path())
        self.backend.current_user = current_user
        self.current_project_id: Optional[int] = None
        self._financials_provider: Optional[ExcelFinancialsProvider] = self._build_financials_provider()
        self.current_tasks: list[TaskRecord] = []

        self._populating = False
        self._sort_column: Optional[int] = None
        self._sort_ascending: bool = True
        self._div25_url: str = ""
        self._show_test_jobs: bool = False
        self._compact_mode: bool = False

        # File watcher — auto-refresh when the shared JSON changes
        self._file_watcher = QFileSystemWatcher(self)
        self._file_watcher_debounce = QTimer(self)
        self._file_watcher_debounce.setSingleShot(True)
        self._file_watcher_debounce.setInterval(1500)  # wait 1.5s after last change
        self._file_watcher_debounce.timeout.connect(self._on_data_file_changed)
        self._start_file_watcher()

        from version import __version__
        self.setWindowTitle(f"Project Tracking Tool v{__version__}")
        self.resize(1460, 860)
        self.setMinimumSize(1180, 700)
        self.setAcceptDrops(True)

        _icon_path = _resource_path("PTT_Normal.ico")
        if _icon_path.exists():
            _icon = QIcon(str(_icon_path))
            self.setWindowIcon(_icon)
            app = QApplication.instance()
            if isinstance(app, QApplication):
                app.setWindowIcon(_icon)

        self._build_ui()
        self._build_menu()
        self._build_shortcuts()
        if self._current_user_view_only():
            self._apply_view_only_restrictions()
        self.refresh_project_list()
        QTimer.singleShot(0, self.refresh_project_list)
        QTimer.singleShot(200, self._check_recent_changes)

        # Warn if running from a cloud-synced folder
        self._check_sync_folder()

        # Check for updates on startup, then every 30 minutes while the app is open.
        threading.Thread(target=self._check_update_bg, daemon=True).start()
        self._update_poll_timer = QTimer(self)
        self._update_poll_timer.setInterval(30 * 60 * 1000)  # 30 minutes
        self._update_poll_timer.timeout.connect(
            lambda: threading.Thread(target=self._check_update_bg, daemon=True).start()
        )
        self._update_poll_timer.start()

    def _build_ui(self) -> None:
        self._stack = QStackedWidget()
        self.setCentralWidget(self._stack)

        # Page 0: logged-out screen
        self._stack.addWidget(self._build_logged_out_panel())

        # Page 1: full app content
        app_widget = _BackgroundWidget(_resource_path("PTT_Transparent.png"))
        root_layout = QHBoxLayout(app_widget)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(0)

        sidebar = self._build_sidebar()
        sidebar.setFixedWidth(242)

        root_layout.addWidget(sidebar)
        root_layout.addWidget(self._build_main_panel(), 1)
        self._stack.addWidget(app_widget)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self._user_status_lbl = QLabel()
        self._user_status_lbl.setObjectName("hint")
        self.status_bar.addPermanentWidget(self._user_status_lbl)
        self.status_bar.showMessage("Ready")
        self._update_banner: Optional[UpdateBanner] = None

        # Show correct initial page
        um = UserManager(self._users_path())
        if self._current_user or not um.has_any_users():
            self._stack.setCurrentIndex(1)
            if self._current_user:
                self._update_status_bar_user()
        else:
            self._stack.setCurrentIndex(0)

    def _build_logged_out_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)

        title = QLabel("Project Tracking Tool")
        title.setObjectName("dialogTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        msg = QLabel("You are logged out.")
        msg.setObjectName("dialogSubtitle")
        msg.setAlignment(Qt.AlignmentFlag.AlignCenter)

        login_btn = PrimaryButton("Log In")
        login_btn.setFixedWidth(140)
        login_btn.clicked.connect(self._do_login)

        layout.addStretch(1)
        layout.addWidget(title)
        layout.addSpacing(12)
        layout.addWidget(msg)
        layout.addSpacing(20)
        layout.addWidget(login_btn, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch(1)

        return panel

    def _build_sidebar(self) -> QWidget:
        panel = QFrame()
        panel.setObjectName("Panel")
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 16, 16, 16)
        panel_layout.setSpacing(10)

        title_label = QLabel("Projects")
        title_label.setObjectName("SectionTitle")
        panel_layout.addWidget(title_label)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search jobs, PM, sales engineer...")
        self.search_edit.textChanged.connect(self.refresh_project_list)
        panel_layout.addWidget(self.search_edit)

        sort_row = QHBoxLayout()
        sort_row.setSpacing(4)
        self.sort_combo = QComboBox()
        self.sort_combo.addItem("Last Updated", "updated")
        self.sort_combo.addItem("Name", "name")
        self.sort_combo.addItem("Job Number", "job_number")
        self.sort_combo.setCurrentIndex(2)
        self.sort_combo.currentIndexChanged.connect(self.refresh_project_list)
        self.sort_dir_btn = QToolButton()
        self.sort_dir_btn.setText("↑ A–Z")
        self.sort_dir_btn.setCheckable(True)
        self.sort_dir_btn.setChecked(True)
        self.sort_dir_btn.setToolTip("Toggle sort direction")
        self.sort_dir_btn.clicked.connect(self._toggle_sort_direction)
        sort_row.addWidget(self.sort_combo, 1)
        sort_row.addWidget(self.sort_dir_btn)
        panel_layout.addLayout(sort_row)

        button_row = QHBoxLayout()
        self.new_project_btn = PrimaryButton("New")
        self.new_project_btn.setMinimumWidth(72)
        self.new_project_btn.clicked.connect(self.create_project)
        self.import_btn = SecondaryButton("Import")
        self.import_btn.setMinimumWidth(100)
        self.import_btn.clicked.connect(self.import_workbook)
        button_row.addWidget(self.new_project_btn)
        button_row.addWidget(self.import_btn)
        panel_layout.addLayout(button_row)

        self.import_email_btn = SecondaryButton("📧 Import Email")
        self.import_email_btn.setMinimumWidth(160)
        self.import_email_btn.setToolTip("Import project from Odin assignment email (.eml)")
        self.import_email_btn.clicked.connect(self.import_email)
        panel_layout.addWidget(self.import_email_btn)

        self.project_list = QListWidget()
        self.project_list.currentItemChanged.connect(self.on_project_selected)
        self.project_list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.project_list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.project_list.setSpacing(1)
        panel_layout.addWidget(self.project_list, 1)

        self._fin_data_label = QLabel("")
        self._fin_data_label.setObjectName("FinDataMeta")
        self._fin_data_label.setVisible(False)
        panel_layout.addWidget(self._fin_data_label)

        secondary_row = QHBoxLayout()
        self.edit_project_btn = SecondaryButton("Edit")
        self.edit_project_btn.setMinimumWidth(72)
        self.edit_project_btn.clicked.connect(self.edit_current_project)
        self.delete_project_btn = PrimaryButton("Delete")
        self.delete_project_btn.setMinimumWidth(72)
        self.delete_project_btn.clicked.connect(self.delete_current_project)
        self.pin_project_btn = SecondaryButton("📌 Pin")
        self.pin_project_btn.setMinimumWidth(72)
        self.pin_project_btn.setToolTip("Pin this project to the top of the list")
        self.pin_project_btn.clicked.connect(self._toggle_pin)
        secondary_row.addWidget(self.edit_project_btn)
        secondary_row.addWidget(self.delete_project_btn)
        secondary_row.addWidget(self.pin_project_btn)
        panel_layout.addLayout(secondary_row)

        return panel

    def _build_main_panel(self) -> QWidget:
        panel = QWidget()
        panel.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        main_layout = QVBoxLayout(panel)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        self._header_widget = self._build_project_header()
        main_layout.addWidget(self._header_widget, 0)

        v_handle = _VResizeHandle(self._header_widget, panel)
        main_layout.addWidget(v_handle)

        main_layout.addWidget(self._build_task_table(), 1)
        return panel

    def _build_project_header(self) -> QWidget:
        wrapper = QFrame()
        wrapper.setObjectName("Panel")
        wrapper_layout = QVBoxLayout(wrapper)
        wrapper_layout.setContentsMargins(12, 8, 12, 8)
        wrapper_layout.setSpacing(4)

        # Project title centered above the info row
        self.project_title = ElidingLabel("No project selected")
        self.project_title.setObjectName("ProjectTitle")
        self.project_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.project_title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        wrapper_layout.addWidget(self.project_title)

        row = QHBoxLayout()
        row.setSpacing(8)
        row.setContentsMargins(0, 0, 0, 0)

        left_widget = QWidget()
        left_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        left_layout = QHBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        self.job_number_value = ElidingLabel("—")
        self.pm_value = ElidingLabel("—")
        self.sales_value = ElidingLabel("—")
        self.completion_value = ElidingLabel("—")
        self.liquid_value = ElidingLabel("—")
        self.warranty_value = ElidingLabel("—")
        self.booked_value = ElidingLabel("—")
        self.contract_value_value = ElidingLabel("—")
        # Div25 is a clickable button not a label
        self.div25_btn = QPushButton("Div25 →")
        self.div25_btn.setObjectName("Div25Btn")
        self.div25_btn.setToolTip("Open Div25 project page")
        self.div25_btn.setFixedWidth(80)
        self.div25_btn.setMinimumHeight(42)
        self.div25_btn.setEnabled(False)
        self.div25_btn.clicked.connect(self._open_div25)

        meta_pairs: list[tuple[str, ElidingLabel]] = [
            ("Job #",     self.job_number_value),
            ("PM",        self.pm_value),
            ("SE",        self.sales_value),
            ("Due",       self.completion_value),
            ("Booked",    self.booked_value),
            ("Contract",  self.contract_value_value),
            ("Warranty",  self.warranty_value),
        ]

        for meta_caption, val_label in meta_pairs:
            col = QVBoxLayout()
            col.setSpacing(0)
            col.setContentsMargins(0, 0, 0, 0)
            caption_lbl = QLabel(meta_caption)
            caption_lbl.setObjectName("MetaCaption")
            val_label.setObjectName("MetaValue")
            val_label.setMinimumWidth(90)
            val_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
            col.addWidget(caption_lbl)
            col.addWidget(val_label)
            left_layout.addLayout(col, 2)

        # Add Div25 button — no caption, just the button aligned to bottom
        div25_col = QVBoxLayout()
        div25_col.setSpacing(0)
        div25_col.setContentsMargins(0, 0, 0, 0)
        div25_col.addWidget(self.div25_btn, 1)
        left_layout.addLayout(div25_col, 0)

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
        for card, width in [
            (self.total_tasks_card, 72),
            (self.completed_card, 66),
            (self.pending_card, 78),
            (self.progress_card, 88),
        ]:
            card.setFixedWidth(width)
            right_layout.addWidget(card)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.VLine)
        sep2.setObjectName("HeaderSep")
        right_layout.addWidget(sep2)

        # Export button with dropdown menu
        self.export_btn = SecondaryButton("Export")
        self.export_btn.setFixedWidth(92)
        export_menu = QMenu(self.export_btn)
        export_menu.addAction("Export to Excel (.xlsx)", self.export_excel)
        export_menu.addAction("Export Snapshot (.json)", self.export_snapshot)
        self.export_btn.setMenu(export_menu)

        right_layout.addWidget(self.export_btn)

        row.addWidget(right_widget, 0)
        wrapper_layout.addLayout(row)

        self.progress_bar = SegmentedProgressBar()
        wrapper_layout.addWidget(self.progress_bar)

        self.project_subtitle = QLabel()
        self.project_subtitle.setObjectName("ProjectSubtitle")
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
        top_row.setSpacing(4)
        title_label = QLabel("Tasks")
        title_label.setObjectName("SectionTitle")
        top_row.addWidget(title_label)

        _tf = QFont()
        _tf.setPointSize(9)
        _sp = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

        def _toolbar_btn(label: str, max_w: int, tip: str = "", cls: type[QPushButton] = SecondaryButton) -> QPushButton:
            btn = cls(label)
            btn.setMaximumWidth(max_w)
            btn.setMinimumWidth(36)
            btn.setSizePolicy(_sp)
            if tip:
                btn.setToolTip(tip)
            return btn

        self.notes_btn = _toolbar_btn("📝 Notes", 100, "Open job progress notes")
        self.notes_btn.clicked.connect(self._open_notes)
        top_row.addWidget(self.notes_btn)

        self.co_btn = _toolbar_btn("🚀 CO Log", 114, "Open change order log")
        self.co_btn.clicked.connect(self._open_change_orders)
        top_row.addWidget(self.co_btn)

        self.project_info_btn = _toolbar_btn("ℹ️ Info", 92, "View all project details")
        self.project_info_btn.clicked.connect(self._show_project_info)
        top_row.addWidget(self.project_info_btn)

        self.financials_btn = _toolbar_btn("💰 Financials", 130, "View financial data from ODIN")
        self.financials_btn.clicked.connect(self._open_financials)
        top_row.addWidget(self.financials_btn)

        self.activity_log_btn = _toolbar_btn("📜 Activity", 116, "View activity log for this project")
        self.activity_log_btn.clicked.connect(self._open_activity_log)
        top_row.addWidget(self.activity_log_btn)

        top_row.addStretch(1)

        self.phase_filter = QComboBox()
        self.phase_filter.setMaximumWidth(118)
        self.phase_filter.setMinimumWidth(36)
        self.phase_filter.setSizePolicy(_sp)
        self.phase_filter.addItem("All phases")
        self.phase_filter.addItems(PHASES)
        self.phase_filter.currentTextChanged.connect(self.populate_tasks)
        top_row.addWidget(self.phase_filter)

        self.add_task_btn = _toolbar_btn("Add Task", 92, cls=PrimaryButton)
        self.add_task_btn.clicked.connect(self.add_task)
        top_row.addWidget(self.add_task_btn)

        self.task_search_edit = QLineEdit()
        self.task_search_edit.setMaximumWidth(110)
        self.task_search_edit.setMinimumWidth(36)
        self.task_search_edit.setSizePolicy(_sp)
        self.task_search_edit.setPlaceholderText("Filter tasks...")
        self.task_search_edit.textChanged.connect(self.populate_tasks)
        top_row.addWidget(self.task_search_edit)

        self.template_apply_combo = QComboBox()
        self.template_apply_combo.addItem("Templates")
        self.template_apply_combo.addItem("Standard", "standard")
        self.template_apply_combo.addItem("Phoenix", "phoenix")
        self.template_apply_combo.setMaximumWidth(108)
        self.template_apply_combo.setMinimumWidth(36)
        self.template_apply_combo.setSizePolicy(_sp)
        self.template_apply_combo.activated.connect(self._apply_template_from_combo)
        top_row.addWidget(self.template_apply_combo)

        self.bulk_complete_btn = _toolbar_btn("✓ All", 68, "Mark all visible tasks complete")
        self.bulk_complete_btn.clicked.connect(self._bulk_complete_tasks)
        top_row.addWidget(self.bulk_complete_btn)

        self.bulk_uncomplete_btn = _toolbar_btn("✗ All", 68, "Mark all visible tasks incomplete")
        self.bulk_uncomplete_btn.clicked.connect(self._bulk_uncomplete_tasks)
        top_row.addWidget(self.bulk_uncomplete_btn)

        self.compact_btn = _toolbar_btn("⊟ Compact", 100, "Toggle compact row view", cls=TertiaryButton)
        self.compact_btn.setCheckable(True)
        self.compact_btn.clicked.connect(self._toggle_compact_view)
        top_row.addWidget(self.compact_btn)

        wrapper_layout.addLayout(top_row)

        self.task_table = ReorderableTaskTable(0, 6)
        self.task_table.setHorizontalHeaderLabels(
            ["Done", "Task", "Phase", "Due Date", "Completed Date", "Notes"]
        )
        self.task_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.task_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.task_table.verticalHeader().setVisible(False)
        self.task_table.verticalHeader().setDefaultSectionSize(36)
        self.task_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.task_table.setAlternatingRowColors(False)
        self.task_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        _vp = _WatermarkViewport(_resource_path("PTT_Transparent.png"))
        self.task_table.setViewport(_vp)

        header = self.task_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(False)
        # Default column widths
        header.resizeSection(1, 280)
        header.resizeSection(2, 110)
        header.resizeSection(3, 110)
        header.resizeSection(4, 130)
        header.setSectionsClickable(True)
        header.sectionClicked.connect(self._on_header_clicked)

        self.task_table.doubleClicked.connect(self._on_task_double_clicked)
        self.task_table.rowsReordered.connect(self._on_tasks_reordered)
        self.task_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.task_table.customContextMenuRequested.connect(self._on_task_context_menu)

        wrapper_layout.addWidget(self.task_table, 1)
        return wrapper

    def _open_div25(self) -> None:
        if self._div25_url:
            QDesktopServices.openUrl(QUrl(self._div25_url))

    def _open_notes(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        name = project.job_name if project else "Project"
        dlg = NotesWindow(self.current_project_id, name, self.backend,
                          view_only=self._current_user_view_only(), parent=self)
        dlg.exec()

    def _open_change_orders(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        name = project.job_name if project else "Project"
        dlg = ChangeOrderWindow(self.current_project_id, name, self.backend,
                                view_only=self._current_user_view_only(), parent=self)
        dlg.exec()

    def _open_activity_log(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        name = project.job_name if project else "Project"
        activities = self.backend.list_activity(self.current_project_id)
        dlg = ActivityLogDialog(
            activities, name,
            backend=self.backend,
            is_admin=self._current_user_is_admin(),
            parent=self,
        )
        dlg.exec()

    def _apply_template_from_combo(self, index: int) -> None:
        if self._current_user_view_only():
            self.template_apply_combo.setCurrentIndex(0)
            return
        if index == 0:
            return  # "Templates" header selected — do nothing
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            self.template_apply_combo.setCurrentIndex(0)
            return
        template = self.template_apply_combo.itemData(index)
        template_name = self.template_apply_combo.itemText(index)
        confirm = QMessageBox.question(
            self,
            "Replace tasks?",
            f"This will delete ALL current tasks and replace them with the {template_name} template.\n\nContinue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        self.template_apply_combo.setCurrentIndex(0)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.backend.replace_project_tasks(self.current_project_id, template)
        self.load_current_project()
        self.status_bar.showMessage(f"Applied {template_name} template", 4000)

    def _show_project_info(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        if not project:
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"Project Info — {project.job_name}")
        dlg.setModal(True)
        dlg.resize(620, 600)

        inner = QWidget()
        form = QFormLayout(inner)
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(8)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        def _row(label: str, value: str) -> None:
            lbl = QLabel(value or "—")
            lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            lbl.setWordWrap(True)
            lbl.setMinimumWidth(300)
            form.addRow(f"<b>{label}</b>", lbl)

        _row("Job Number",           project.job_number)
        _row("Job Name",             project.job_name)
        _row("Project Manager",      project.project_manager)
        _row("Sales Engineer",       project.sales_engineer)
        _row("Target Completion",    project.target_completion or "")
        _row("Booked Date",          project.booked_date)
        _row("Contract Value",       project.contract_value)
        _row("Liquid Damages",       project.liquid_damages)
        _row("Warranty Period",      project.warranty_period)
        _row("Job Sub-Type",         project.job_subtype)
        _row("Owner",                project.owner)
        _row("Contracted With",      project.contracted_with)
        _row("General Contractor",   project.general_contractor)
        _row("Group Ops Manager",    project.group_ops_manager)
        _row("Group Ops Supervisor", project.group_ops_supervisor)
        _row("Div25 URL",            project.div25_url)
        _row("Job Docs Path",        project.job_docs)

        if project.notes:
            notes_edit = QPlainTextEdit(project.notes)
            notes_edit.setReadOnly(True)
            notes_edit.setFixedHeight(120)
            notes_edit.setMinimumWidth(300)
            form.addRow("<b>Notes</b>", notes_edit)

        # ── ODIN Financial Summary ──────────────────────────────────────── #
        if self._financials_provider and project.job_number:
            snap = self._financials_provider.get_financials(project.job_number)
            if snap.contract_value:
                sep = QFrame()
                sep.setFrameShape(QFrame.Shape.HLine)
                sep.setFrameShadow(QFrame.Shadow.Sunken)
                form.addRow(sep)

                fin_title = QLabel("<b>ODIN Financial Data</b>")
                form.addRow(fin_title)

                def _fin_row(label: str, value: str) -> None:
                    lbl = QLabel(value)
                    lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
                    lbl.setMinimumWidth(300)
                    form.addRow(f"<b>{label}</b>", lbl)

                diff = snap.differential_margin
                arrow = "▲" if diff >= 0 else "▼"
                diff_color = "#4caf50" if diff >= 0.02 else ("#f44336" if diff <= -0.02 else "#ff9800")
                diff_lbl = QLabel(f"<span style='color:{diff_color}; font-weight:bold'>{arrow} {abs(diff)*100:.1f}%</span>")
                diff_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

                _fin_row("Contract Value",  f"${snap.contract_value:,.2f}")
                _fin_row("Billed to Date",  f"${snap.billed_to_date:,.2f}")
                _fin_row("Actual Cost",     f"${snap.actual_cost:,.2f}")
                _fin_row("Booked Margin",   f"{snap.booked_margin*100:.1f}%")
                _fin_row("Actual Margin",   f"{snap.actual_margin*100:.1f}%")
                form.addRow("<b>Differential</b>", diff_lbl)
                _fin_row("Status (ODIN)",   snap.status)
                if snap.last_refreshed:
                    _fin_row("Data as of",  snap.last_refreshed)

        scroll = QScrollArea()
        scroll.setWidget(inner)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        close_btn = TertiaryButton("Close")
        close_btn.clicked.connect(dlg.accept)

        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.addWidget(scroll, 1)
        layout.addWidget(close_btn)

        dlg.exec()

    def _start_file_watcher(self) -> None:
        """Watch the shared data JSON for changes made by other users."""
        data_path = str(self._resolve_data_path())
        if data_path not in self._file_watcher.files():
            self._file_watcher.addPath(data_path)
        self._file_watcher.fileChanged.connect(self._on_file_watcher_event)

    def _on_file_watcher_event(self, path: str) -> None:
        # Re-add the path in case OneDrive replaced the file (common with cloud sync)
        if path not in self._file_watcher.files():
            self._file_watcher.addPath(path)
        # Debounce: restart the timer so rapid writes only trigger one refresh
        self._file_watcher_debounce.start()

    def _on_data_file_changed(self) -> None:
        """Called after debounce — reload backend and refresh UI."""
        self.backend = ProjectTrackerBackend(self._resolve_data_path())
        self.refresh_project_list()

    @staticmethod
    def _resolve_data_path() -> Path:
        """Return the data file path — custom shared folder if configured, else default."""
        settings = QSettings("ATSInc", "ProjectTrackingTool")
        custom_folder = str(settings.value("dataFolder", "")).strip()
        if custom_folder:
            folder = Path(custom_folder)
            folder.mkdir(parents=True, exist_ok=True)
            return folder / "project_tracker_data.json"
        return _app_data_path()

    def _open_data_location_settings(self) -> None:
        settings = QSettings("ATSInc", "ProjectTrackingTool")
        current_folder = str(settings.value("dataFolder", "")).strip()

        dlg = DataLocationDialog(current_folder, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        new_folder = dlg.selected_folder()

        # If clearing back to default
        if not new_folder:
            settings.remove("dataFolder")
            self._reload_backend()
            return

        new_path = Path(new_folder) / "project_tracker_data.json"

        # Offer to copy existing data if target doesn't have a file yet
        if not new_path.exists():
            old_path = _app_data_path()
            if old_path.exists() and old_path != new_path:
                reply = QMessageBox.question(
                    self,
                    "Copy Existing Data?",
                    f"No data file found in the selected folder.\n\n"
                    f"Copy your existing data to:\n{new_path}\n\n"
                    f"(Recommended if you're setting this up for the first time)",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                )
                if reply == QMessageBox.StandardButton.Yes:
                    shutil.copy2(old_path, new_path)

        settings.setValue("dataFolder", new_folder)
        self._reload_backend()
        QMessageBox.information(
            self,
            "Data Location Updated",
            f"Data file location set to:\n{new_path}\n\n"
            f"The app is now using this location.",
        )

    def _reload_backend(self) -> None:
        """Reinitialize the backend from the current resolved path and refresh the UI."""
        self.current_project_id = None
        self.backend = ProjectTrackerBackend(self._resolve_data_path())
        self.backend.current_user = self._current_user
        self._start_file_watcher()
        self.refresh_project_list()

    # ── Financials ─────────────────────────────────────────────────────── #

    # ── Recent-changes snapshot ─────────────────────────────────────────── #

    def _save_session_snapshot(self) -> None:
        """Save a local snapshot of all projects for next-open diff."""
        try:
            projects = self.backend.list_projects("", include_test=True)
            data = {
                "saved_at": datetime.now().replace(microsecond=0).isoformat(sep=" "),
                "projects": {
                    str(p.id): {field: getattr(p, field, "") or "" for field in _CHANGE_FIELDS}
                    for p in projects
                },
            }
            _LOCAL_SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(_LOCAL_SNAPSHOT_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except (OSError, ValueError):
            pass

    def _check_recent_changes(self) -> None:
        """Compare current data against local snapshot; show dialog if anything changed."""
        if not _LOCAL_SNAPSHOT_PATH.exists():
            self._save_session_snapshot()
            return

        try:
            with open(_LOCAL_SNAPSHOT_PATH, encoding="utf-8") as f:
                snapshot = json.load(f)
        except (OSError, ValueError):
            self._save_session_snapshot()
            return

        since = snapshot.get("saved_at", "last session")
        old_projects: dict = snapshot.get("projects", {})
        current_projects = self.backend.list_projects("", include_test=True)

        changes: list[dict] = []
        for project in current_projects:
            pid = str(project.id)
            if pid not in old_projects:
                fields = [
                    (label, "", str(getattr(project, field, "") or ""))
                    for field, label in _CHANGE_FIELDS.items()
                    if getattr(project, field, "")
                ]
                changes.append({
                    "type": "new",
                    "job_number": project.job_number or "—",
                    "job_name": project.job_name or "—",
                    "updated_by": project.created_by or "",
                    "fields": fields,
                })
            else:
                old = old_projects[pid]
                diffs = [
                    (label, str(old.get(field, "") or ""), str(getattr(project, field, "") or ""))
                    for field, label in _CHANGE_FIELDS.items()
                    if str(old.get(field, "") or "") != str(getattr(project, field, "") or "")
                ]
                if diffs:
                    changes.append({
                        "type": "updated",
                        "job_number": project.job_number or "—",
                        "job_name": project.job_name or "—",
                        "updated_by": project.updated_by or "",
                        "fields": diffs,
                    })

        self._save_session_snapshot()

        if changes:
            dlg = RecentChangesDialog(changes, since, parent=self)
            dlg.exec()

    @staticmethod
    def _build_financials_provider() -> Optional[ExcelFinancialsProvider | SnapshotFinancialsProvider]:
        settings = QSettings("ATSInc", "ProjectTrackingTool")
        file_path = str(settings.value("financialsFile", ""))
        snapshot_path = MainWindow._resolve_data_path().parent / "financial_snapshot.json"
        if file_path and Path(file_path).exists():
            return ExcelFinancialsProvider(file_path, sheet_name="Phoenix Projects", snapshot_path=snapshot_path)
        if snapshot_path.exists():
            return SnapshotFinancialsProvider(snapshot_path)
        return None

    def _open_financials_file_settings(self) -> None:
        settings = QSettings("ATSInc", "ProjectTrackingTool")
        current_file = str(settings.value("financialsFile", ""))

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Financial Data File",
            str(Path(current_file).parent) if current_file else str(Path.home()),
            "Excel Files (*.xlsb *.xlsx *.xlsm);;All Files (*)",
        )
        if not file_path:
            return

        settings.setValue("financialsFile", file_path)
        snapshot_path = self._resolve_data_path().parent / "financial_snapshot.json"
        self._financials_provider = ExcelFinancialsProvider(file_path, sheet_name="Phoenix Projects", snapshot_path=snapshot_path)
        QMessageBox.information(
            self,
            "Financial Data File Set",
            f"Financial data will now be read from:\n{file_path}",
        )
        self.refresh_project_list()

    def _open_financials(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return

        if self._financials_provider is None:
            reply = QMessageBox.question(
                self,
                "No Financial Data File",
                "No financial data file has been configured.\n\n"
                "Would you like to select your ODIN tracking Excel file now?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if reply == QMessageBox.StandardButton.Yes:
                self._open_financials_file_settings()
            if self._financials_provider is None:
                return

        project = self.backend.get_project(self.current_project_id)
        if not project:
            return

        if not project.job_number:
            QMessageBox.information(
                self, "Missing Job Number", "This project does not have a job number yet."
            )
            return

        dlg = FinancialsDialog(
            job_number=project.job_number,
            provider=self._financials_provider,
            parent=self,
        )
        dlg.exec()
        self.refresh_project_list()

    def _check_sync_folder(self) -> None:
        """Warn if the app *executable* is running from a cloud-synced folder."""
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
        info: Optional[UpdateInfo] = getattr(self, "_pending_update_info", None)
        if info is None:
            return
        # Don't show a second banner if one is already visible
        if self._update_banner and self._update_banner.isVisible():
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

    def _users_path(self) -> Path:
        """Return path to users.json, co-located with the data file."""
        return self.backend.db_path.parent / "users.json"

    def _update_status_bar_user(self) -> None:
        from user_auth import ROLE_LABELS
        if self._current_user:
            role = self._current_user_role()
            role_label = ROLE_LABELS.get(role, "")
            suffix = f"  [{role_label}]" if role_label else ""
            self._user_status_lbl.setText(f"  Logged in as: {self._current_user}{suffix}")
        else:
            self._user_status_lbl.setText("")

    def _current_user_role(self) -> str:
        """Return the role of the logged-in user.

        Returns 'admin' when no users exist yet (first-time setup),
        'none' when users exist but nobody is logged in, or the user's
        actual role string otherwise.
        """
        if not self._current_user:
            um = UserManager(self._users_path())
            return "admin" if not um.has_any_users() else "none"
        um = UserManager(self._users_path())
        user = um.get_user(self._current_user)
        return user.role if user is not None else "user"

    def _current_user_is_admin(self) -> bool:
        return self._current_user_role() == "admin"

    def _current_user_view_only(self) -> bool:
        return self._current_user_role() == "view_only"

    def _do_login(self) -> None:
        um = UserManager(self._users_path())
        if not um.has_any_users():
            return
        dlg = LoginDialog(um, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        user = dlg.logged_in_user()
        if user is None:
            return
        if user.must_change_password:
            ChangePasswordDialog(um, user.username, parent=self).exec()
        self._current_user = user.username
        self.backend.current_user = user.username
        self._reset_controls()
        if self._current_user_view_only():
            self._apply_view_only_restrictions()
        self._update_status_bar_user()
        self._update_auth_menu()
        self._stack.setCurrentIndex(1)
        self.refresh_project_list()

    def _do_logout(self) -> None:
        reply = QMessageBox.question(
            self, "Log Out",
            f"Log out of '{self._current_user}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self._current_user = ""
        self.backend.current_user = ""
        _SESSION_PATH.unlink(missing_ok=True)
        self.current_project_id = None
        self._update_status_bar_user()
        self._update_auth_menu()
        self._stack.setCurrentIndex(0)

    def _reset_controls(self) -> None:
        """Re-enable all modification controls before applying role restrictions."""
        self.new_project_btn.setEnabled(True)
        self.edit_project_btn.setEnabled(True)
        self.delete_project_btn.setEnabled(True)
        self.pin_project_btn.setEnabled(True)
        self.pin_project_btn.setToolTip("Pin this project to the top of the list")
        self.task_table.setDragEnabled(True)
        self.bulk_complete_btn.setEnabled(True)
        self.bulk_complete_btn.setToolTip("Mark all visible tasks complete")
        self.bulk_uncomplete_btn.setEnabled(True)
        self.bulk_uncomplete_btn.setToolTip("Mark all visible tasks incomplete")
        self.import_btn.setEnabled(True)
        self.import_email_btn.setEnabled(True)
        self.import_email_btn.setToolTip("Import project from Odin assignment email (.eml)")
        self.add_task_btn.setEnabled(True)
        self.notes_btn.setEnabled(True)
        self.notes_btn.setToolTip("Open job progress notes")
        self.co_btn.setEnabled(True)
        self.co_btn.setToolTip("Open change order log")
        self.activity_log_btn.setEnabled(True)
        self.activity_log_btn.setToolTip("View activity log for this project")
        self.template_apply_combo.setEnabled(True)
        self._new_project_action.setEnabled(True)
        self._import_action.setEnabled(True)

    def _update_auth_menu(self) -> None:
        """Toggle menu item visibility based on logged-in state and role."""
        logged_in = bool(self._current_user)
        view_only = self._current_user_view_only()
        self._login_action.setVisible(not logged_in)
        self._logout_action.setVisible(logged_in)
        self._change_pw_action.setVisible(logged_in)
        self._manage_users_action_ref.setVisible(logged_in and self._current_user_is_admin())
        for action in self._logged_in_only_actions:
            action.setVisible(logged_in)
        # Hide write actions from view-only users
        self._new_project_action.setVisible(logged_in and not view_only)
        self._import_action.setVisible(logged_in and not view_only)

    def _open_manage_users(self) -> None:
        if not self._current_user_is_admin():
            QMessageBox.warning(self, "Access Denied", "Only administrators can manage users.")
            return
        um = UserManager(self._users_path())
        dlg = ManageUsersDialog(um, self._current_user, parent=self)
        dlg.exec()

    def _open_change_my_password(self) -> None:
        if not self._current_user:
            return
        um = UserManager(self._users_path())
        dlg = SelfChangePasswordDialog(um, self._current_user, parent=self)
        dlg.exec()

    def _apply_view_only_restrictions(self) -> None:
        """Disable all data-modification controls for view-only users."""
        tip = "Your account is view-only"
        for btn in (
            self.new_project_btn, self.edit_project_btn, self.delete_project_btn,
            self.pin_project_btn, self.import_btn, self.import_email_btn, self.add_task_btn,
            self.notes_btn, self.co_btn, self.template_apply_combo,
        ):
            btn.setEnabled(False)
            btn.setToolTip(tip)
        self._new_project_action.setEnabled(False)
        self._import_action.setEnabled(False)
        self.task_table.setDragEnabled(False)
        for btn in (self.bulk_complete_btn, self.bulk_uncomplete_btn):
            btn.setEnabled(False)
            btn.setToolTip(tip)

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("File")

        # ── Auth actions (mutually exclusive) ─────────────────────────────
        self._login_action = QAction("Log In...", self)
        self._login_action.triggered.connect(self._do_login)

        self._logout_action = QAction("Log Out", self)
        self._logout_action.triggered.connect(self._do_logout)

        # ── Logged-in actions ─────────────────────────────────────────────
        self._new_project_action = QAction("New Project", self)
        self._new_project_action.triggered.connect(self.create_project)

        self._import_action = QAction("Import Workbook", self)
        self._import_action.triggered.connect(self.import_workbook)

        self.export_excel_action = QAction("Export to Excel (.xlsx)", self)
        self.export_excel_action.triggered.connect(self.export_excel)
        self.export_excel_action.setEnabled(False)

        self.export_menu_action = QAction("Export Snapshot (.json)", self)
        self.export_menu_action.triggered.connect(self.export_snapshot)
        self.export_menu_action.setEnabled(False)

        self.bulk_export_excel_action = QAction("Bulk Export to Excel...", self)
        self.bulk_export_excel_action.triggered.connect(self._bulk_export_excel)

        self._data_location_action = QAction("Data Location...", self)
        self._data_location_action.triggered.connect(self._open_data_location_settings)

        self._financials_action = QAction("Financial Data File...", self)
        self._financials_action.triggered.connect(self._open_financials_file_settings)

        self._manage_users_action_ref = QAction("Manage Users...", self)
        self._manage_users_action_ref.triggered.connect(self._open_manage_users)

        self._change_pw_action = QAction("Change My Password...", self)
        self._change_pw_action.triggered.connect(self._open_change_my_password)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.close)

        # All actions that should only appear when logged in
        self._logged_in_only_actions = [
            self._new_project_action,
            self._import_action,
            self.export_excel_action,
            self.export_menu_action,
            self.bulk_export_excel_action,
            self._data_location_action,
            self._financials_action,
        ]

        file_menu.addAction(self._login_action)
        file_menu.addAction(self._new_project_action)
        file_menu.addAction(self._import_action)
        file_menu.addSeparator()
        file_menu.addAction(self.export_excel_action)
        file_menu.addAction(self.export_menu_action)
        file_menu.addAction(self.bulk_export_excel_action)
        file_menu.addSeparator()
        file_menu.addAction(self._data_location_action)
        file_menu.addAction(self._financials_action)
        file_menu.addAction(self._manage_users_action_ref)
        file_menu.addAction(self._change_pw_action)
        file_menu.addAction(self._logout_action)
        file_menu.addSeparator()
        file_menu.addAction(quit_action)

        self._update_auth_menu()

        # ── View menu ──────────────────────────────────────────────────────────
        view_menu = self.menuBar().addMenu("View")

        refresh_action = QAction("Refresh", self)
        refresh_action.setShortcut(QKeySequence("F5"))
        refresh_action.triggered.connect(self.refresh_project_list)
        view_menu.addAction(refresh_action)

        # ── Help menu ──────────────────────────────────────────────────────────
        help_menu = self.menuBar().addMenu("Help")

        version_history_action = QAction("Version History && Recent Updates", self)
        version_history_action.triggered.connect(self._show_version_history)
        help_menu.addAction(version_history_action)

        help_menu.addSeparator()

        email_support_action = QAction("Email Support", self)
        email_support_action.triggered.connect(self._email_support)
        help_menu.addAction(email_support_action)

        help_menu.addSeparator()

        self.test_jobs_action = QAction("Show Test Jobs", self)
        self.test_jobs_action.triggered.connect(self._toggle_test_jobs)
        help_menu.addAction(self.test_jobs_action)

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

        close_btn = TertiaryButton("Close")
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

    def _toggle_test_jobs(self) -> None:
        if not self._show_test_jobs:
            # Check if test jobs already exist; if not, create them
            existing = self.backend.list_projects(include_test=True)
            has_test = any(p.is_test for p in existing)
            if not has_test:
                self.backend.create_test_jobs()
            self._show_test_jobs = True
            self.test_jobs_action.setText("Hide Test Jobs")
            self.status_bar.showMessage("Test jobs visible", 4000)
        else:
            self._show_test_jobs = False
            self.test_jobs_action.setText("Show Test Jobs")
            self.status_bar.showMessage("Test jobs hidden", 4000)
            # If a test job is currently selected, deselect it
            if self.current_project_id is not None:
                proj = self.backend.get_project(self.current_project_id)
                if proj and proj.is_test:
                    self.current_project_id = None
        self.refresh_project_list()

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

    @staticmethod
    def _email_support() -> None:
        QDesktopServices.openUrl(QUrl("mailto:Justing@atsinc.org"))

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

    def _toggle_sort_direction(self) -> None:
        asc = self.sort_dir_btn.isChecked()
        self.sort_dir_btn.setText("↑ A–Z" if asc else "↓ Z–A")
        self.refresh_project_list()

    def refresh_project_list(self) -> None:
        search_text = self.search_edit.text().strip() if hasattr(self, "search_edit") else ""
        sort_by = self.sort_combo.currentData() if hasattr(self, "sort_combo") else "updated"
        sort_asc = self.sort_dir_btn.isChecked() if hasattr(self, "sort_dir_btn") else False
        projects = self.backend.list_projects(
            search_text,
            include_test=getattr(self, "_show_test_jobs", False),
            sort_by=sort_by,
            sort_asc=sort_asc,
        )
        selected_project_id = self.current_project_id

        self.project_list.blockSignals(True)
        self.project_list.clear()
        for project in projects:
            fin_html = ""
            if self._financials_provider and project.job_number:
                try:
                    snap = self._financials_provider.get_financials(project.job_number)
                    if snap.contract_value:
                        diff = snap.differential_margin
                        arrow = "▲" if diff >= 0 else "▼"
                        pct = abs(diff) * 100
                        margin_color = "#4caf50" if pct <= 1.0 else ("#ff9800" if pct <= 5.0 else "#f44336")
                        fin_html = (
                            f'<span style="color:#4caf50">${snap.contract_value:,.0f}</span>'
                            f'<span style="color:{margin_color}">&nbsp;&nbsp;{arrow}{diff*100:.1f}%</span>'
                        )
                    else:
                        fin_html = '<span style="color:#888888">ODIN: No data</span>'
                except (TypeError, ValueError, AttributeError, KeyError):
                    logger.exception(
                        "Failed to build financial line for job %s", project.job_number
                    )

            job_name = project.job_name or ""
            if len(job_name) > 36:
                job_name = job_name[:34] + "…"
            if project.pinned:
                job_name = "📌 " + job_name

            item = QListWidgetItem()
            item.setData(Qt.ItemDataRole.UserRole, project.id)

            widget = QWidget()
            widget.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
            layout = QVBoxLayout(widget)
            layout.setContentsMargins(4, 6, 4, 16)
            layout.setSpacing(3)

            num_lbl = QLabel(project.job_number or "")
            num_lbl.setObjectName("hint")
            layout.addWidget(num_lbl)

            name_lbl = QLabel(job_name)
            layout.addWidget(name_lbl)

            if fin_html:
                fin_lbl = QLabel(fin_html)
                fin_lbl.setTextFormat(Qt.TextFormat.RichText)
                layout.addWidget(fin_lbl)

            widget.adjustSize()
            item_w = self.project_list.viewport().width() or 196
            item.setSizeHint(QSize(item_w, widget.sizeHint().height() + 12))
            self.project_list.addItem(item)
            self.project_list.setItemWidget(item, widget)
        self.project_list.blockSignals(False)

        # Update the "data as of" label (item 8)
        if self._financials_provider:
            ts = self._financials_provider.data_as_of
            if ts:
                self._fin_data_label.setText(f"ODIN data as of {ts}")
                self._fin_data_label.setVisible(True)
            else:
                self._fin_data_label.setVisible(False)
        else:
            self._fin_data_label.setVisible(False)

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
        self.export_excel_action.setEnabled(self.current_project_id is not None)
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
        booked_raw = project.booked_date or ""
        self.booked_value.setText(booked_raw.split("T")[0].split(" ")[0] or "—")
        self.contract_value_value.setText(
            f"${float(project.contract_value):,.0f}" if project.contract_value else "—"
        )
        self._div25_url = project.div25_url or ""
        self.div25_btn.setEnabled(bool(self._div25_url))
        self.div25_btn.setToolTip(self._div25_url or "No Div25 URL")
        self.project_notes.setPlainText(project.notes or "")
        if not self._current_user_view_only():
            self._update_pin_button(project.pinned)

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
            self.booked_value,
            self.contract_value_value,
        ]:
            value_widget.setText("—")
        self._div25_url = ""
        self.div25_btn.setEnabled(False)
        self.div25_btn.setToolTip("No Div25 URL")
        self._update_pin_button(False)
        self.project_notes.clear()
        self.total_tasks_card.set_value("0")
        self.completed_card.set_value("0")
        self.pending_card.set_value("0")
        self.progress_card.set_value("0%")
        self.progress_bar.clear()
        self.current_tasks = []
        self.task_table.setRowCount(0)
        self.export_menu_action.setEnabled(False)
        self.export_excel_action.setEnabled(False)

    def _on_header_clicked(self, col: int) -> None:
        if col == 6:
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
            3: lambda t: t.due_date or "",
            4: lambda t: t.completed_date or "",
            5: lambda t: (t.notes or "").casefold(),
        }
        if self._sort_column in col_keys:
            filtered_tasks = sorted(
                filtered_tasks,
                key=col_keys[self._sort_column],
                reverse=not self._sort_ascending,
            )

        today_str = QDate.currentDate().toString("yyyy-MM-dd")

        self._populating = True
        try:
            self.task_table.setRowCount(len(filtered_tasks))
            for row_index, task in enumerate(filtered_tasks):
                checkbox = QCheckBox()
                checkbox.setChecked(task.is_complete)
                if self._current_user_view_only():
                    checkbox.setEnabled(False)
                else:
                    checkbox.toggled.connect(
                        lambda checked, task_id=task.id: self.toggle_task(int(task_id), bool(checked))
                    )
                self.task_table.setCellWidget(row_index, 0, self._centered_widget(checkbox))
                self.task_table.setItem(row_index, 1, QTableWidgetItem(task.task_name))

                phase_item = QTableWidgetItem(task.phase)
                phase_item.setForeground(QColor(PHASE_COLORS.get(task.phase, "#64748b")))
                self.task_table.setItem(row_index, 2, phase_item)

                # Due date — highlight overdue (past due, not completed) in red
                due_item = QTableWidgetItem(task.due_date or "")
                if task.due_date and not task.is_complete and task.due_date < today_str:
                    due_item.setForeground(QColor("#f44336"))
                self.task_table.setItem(row_index, 3, due_item)

                self.task_table.setItem(row_index, 4, QTableWidgetItem(task.completed_date or ""))
                self.task_table.setItem(row_index, 5, QTableWidgetItem(task.notes or ""))

                for column_index in range(1, 6):
                    item = self.task_table.item(row_index, column_index)
                    if item is not None:
                        item.setData(Qt.ItemDataRole.UserRole, task.id)

        finally:
            self._populating = False

    def _on_task_context_menu(self, pos) -> None:
        menu = QMenu(self)

        add_action = QAction("Add Task", self)
        add_action.triggered.connect(self.add_task)
        menu.addAction(add_action)

        if not self._current_user_view_only():
            item = self.task_table.itemAt(pos)
            row = self.task_table.rowAt(pos.y()) if item is None else item.row()
            task_id = None
            if row >= 0:
                id_item = self.task_table.item(row, 1)
                if id_item:
                    task_id = id_item.data(Qt.ItemDataRole.UserRole)

            if task_id is not None:
                menu.addSeparator()
                edit_action = QAction("Edit Task", self)
                edit_action.triggered.connect(lambda: self.edit_task(task_id))
                menu.addAction(edit_action)

                delete_action = QAction("Delete Task", self)
                delete_action.triggered.connect(lambda: self.delete_task(task_id))
                menu.addAction(delete_action)

        menu.exec(self.task_table.viewport().mapToGlobal(pos))

    @staticmethod
    def _centered_widget(widget: QWidget) -> QWidget:
        container = QWidget()
        row_layout = QHBoxLayout(container)
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.addStretch(1)
        row_layout.addWidget(widget)
        row_layout.addStretch(1)
        return container

    def _save_new_project(self, record: ProjectRecord, status_msg: str = "", task_template: str = "standard") -> bool:
        """Create a new project from record, refresh list, update status. Returns True on success."""
        try:
            new_id = self.backend.create_project(record, include_default_tasks=True, task_template=task_template)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to create project", str(exc))
            return False
        self.current_project_id = new_id
        self.refresh_project_list()
        self.status_bar.showMessage(status_msg or f"Created project {record.job_name}", 4000)
        return True

    def create_project(self) -> None:
        if self._current_user_view_only():
            return
        dialog = ProjectDialog(self)
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        self._save_new_project(dialog.get_data(), task_template=dialog.get_template())

    def edit_current_project(self) -> None:
        if self._current_user_view_only():
            return
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
                booked_date=record.booked_date,
                group_ops_manager=record.group_ops_manager,
                group_ops_supervisor=record.group_ops_supervisor,
                job_subtype=record.job_subtype,
                owner=record.owner,
                contracted_with=record.contracted_with,
                general_contractor=record.general_contractor,
                contract_value=record.contract_value,
                job_docs=record.job_docs,
                div25_url=record.div25_url,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Unable to update project", str(exc))
            return
        self.refresh_project_list()
        self.load_current_project()
        self.status_bar.showMessage("Project updated", 4000)

    def delete_current_project(self) -> None:
        if self._current_user_view_only():
            return
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

    def _toggle_pin(self) -> None:
        if self._current_user_view_only():
            return
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        if not project:
            return
        new_pinned = not project.pinned
        self.backend.update_project(self.current_project_id, pinned=new_pinned)
        self._update_pin_button(new_pinned)
        self.refresh_project_list()
        msg = "Project pinned to top" if new_pinned else "Project unpinned"
        self.status_bar.showMessage(msg, 3000)

    def _update_pin_button(self, pinned: bool) -> None:
        if pinned:
            self.pin_project_btn.setText("📌 Unpin")
            self.pin_project_btn.setToolTip("Unpin this project")
        else:
            self.pin_project_btn.setText("📌 Pin")
            self.pin_project_btn.setToolTip("Pin this project to the top of the list")

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(u.toLocalFile().lower().endswith(".eml") for u in urls):
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event) -> None:
        if self._current_user_view_only():
            return
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".eml"):
                self._process_email_import(path)
                return

    def import_email(self) -> None:
        if self._current_user_view_only():
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Odin Assignment Email",
            "",
            "Email Files (*.eml)",
        )
        if file_path:
            self._process_email_import(file_path)

    def _process_email_import(self, eml_path: str) -> None:
        try:
            record, is_duplicate = self.backend.import_project_from_email(eml_path)
        except Exception as exc:
            QMessageBox.critical(self, "Email import failed", str(exc))
            return

        if not record.job_name and not record.job_number:
            QMessageBox.warning(
                self, "Import failed",
                "Could not extract project data from this email.\n\n"
                "Make sure it is an Odin assignment email."
            )
            return

        if is_duplicate:
            # Find the existing project id
            existing = next(
                (p for p in self.backend.list_projects(include_test=True)
                 if p.job_number.strip() == record.job_number.strip()),
                None,
            )
            answer = QMessageBox.question(
                self,
                "Project already exists",
                f"Job #{record.job_number} — '{record.job_name}' already exists.\n\n"
                f"Do you want to update the existing project with data from this email?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer == QMessageBox.StandardButton.Yes and existing and existing.id is not None:
                try:
                    self.backend.update_project_from_email(existing.id, record)
                except Exception as exc:
                    QMessageBox.critical(self, "Update failed", str(exc))
                    return
                self.current_project_id = existing.id
                self.refresh_project_list()
                self.status_bar.showMessage(
                    f"Updated project {record.job_number} from email", 5000
                )
            return

        # New project — open dialog pre-filled for review before saving
        dialog = ProjectDialog(self, record)
        dialog.setWindowTitle(f"Review Import — {record.job_name}")
        if dialog.exec() != int(QDialog.DialogCode.Accepted):
            return
        self._save_new_project(
            dialog.get_data(),
            f"Imported project {record.job_number} from email",
        )

    def import_workbook(self) -> None:
        if self._current_user_view_only():
            return
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

    def export_excel(self) -> None:
        if self.current_project_id is None:
            QMessageBox.information(self, "No project selected", "Select a project first.")
            return
        project = self.backend.get_project(self.current_project_id)
        default_name = f"{project.job_number or 'project'}_report.xlsx" if project else "project_report.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export to Excel", default_name, "Excel Files (*.xlsx)",
        )
        if not file_path:
            return
        try:
            output_path = self.backend.export_project_to_excel(self.current_project_id, file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))
            return
        self.status_bar.showMessage(f"Exported Excel report to {output_path}", 5000)

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

    def _bulk_export_excel(self) -> None:
        projects = self.backend.list_projects(include_test=self._show_test_jobs)
        if not projects:
            QMessageBox.information(self, "No projects", "There are no projects to export.")
            return
        dlg = BulkExportDialog(projects, parent=self)
        if dlg.exec() != int(QDialog.DialogCode.Accepted):
            return
        ids = dlg.selected_ids()
        if not ids:
            QMessageBox.information(self, "Nothing selected", "Select at least one project to export.")
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Bulk Export to Excel", "bulk_export.xlsx", "Excel Files (*.xlsx)",
        )
        if not file_path:
            return
        try:
            output_path = self.backend.export_projects_to_excel(ids, file_path)
        except Exception as exc:
            QMessageBox.critical(self, "Export failed", str(exc))
            return
        self.status_bar.showMessage(f"Exported {len(ids)} project(s) to {output_path}", 6000)

    def add_task(self) -> None:
        if self._current_user_view_only():
            return
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
                due_date=data["due_date"],
                completed_date=data["completed_date"],
                notes=data["notes"],
            )
        except Exception as exc:
            QMessageBox.critical(self, "Unable to add task", str(exc))
            return
        self.load_current_project()
        self.status_bar.showMessage("Task added", 4000)

    def edit_task(self, task_id: int) -> None:
        if self._current_user_view_only():
            return
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
        if self._current_user_view_only():
            return
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
        if self._populating or self._current_user_view_only():
            return
        try:
            self.backend.set_task_completed(task_id, checked)
        except Exception as exc:
            QMessageBox.critical(self, "Unable to update task", str(exc))
            return
        self._refresh_stats_only()

    def _on_tasks_reordered(self) -> None:
        """Called after a drag-drop reorder in the task table. Persists new sort_order values."""
        if self._current_user_view_only():
            return
        for row_index in range(self.task_table.rowCount()):
            item = self.task_table.item(row_index, 1)
            if item is None:
                continue
            task_id = item.data(Qt.ItemDataRole.UserRole)
            if task_id is not None:
                self.backend.update_task(int(task_id), sort_order=row_index)
        # Reload to sync current_tasks list with new order
        if self.current_project_id is not None:
            self.current_tasks = self.backend.list_tasks(self.current_project_id)

    def _bulk_complete_tasks(self) -> None:
        """Mark all currently visible tasks as complete."""
        if self._current_user_view_only():
            return
        task_ids = self._visible_task_ids()
        if not task_ids:
            return
        reply = QMessageBox.question(
            self, "Mark All Complete",
            f"Mark {len(task_ids)} visible task(s) as complete?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        for task_id in task_ids:
            self.backend.set_task_completed(task_id, True)
        self.load_current_project()
        self.status_bar.showMessage(f"Marked {len(task_ids)} task(s) complete", 4000)

    def _bulk_uncomplete_tasks(self) -> None:
        """Mark all currently visible tasks as incomplete."""
        if self._current_user_view_only():
            return
        task_ids = self._visible_task_ids()
        if not task_ids:
            return
        reply = QMessageBox.question(
            self, "Mark All Incomplete",
            f"Mark {len(task_ids)} visible task(s) as incomplete?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        for task_id in task_ids:
            self.backend.set_task_completed(task_id, False)
        self.load_current_project()
        self.status_bar.showMessage(f"Marked {len(task_ids)} task(s) incomplete", 4000)

    def _visible_task_ids(self) -> list[int]:
        """Return task IDs for all rows currently displayed in the task table."""
        ids = []
        for row_index in range(self.task_table.rowCount()):
            item = self.task_table.item(row_index, 1)
            if item is not None:
                task_id = item.data(Qt.ItemDataRole.UserRole)
                if task_id is not None:
                    ids.append(int(task_id))
        return ids

    def _toggle_compact_view(self) -> None:
        self._compact_mode = self.compact_btn.isChecked()
        row_height = 24 if self._compact_mode else 36
        self.task_table.verticalHeader().setDefaultSectionSize(row_height)
        # Resize all existing rows
        for row_index in range(self.task_table.rowCount()):
            self.task_table.setRowHeight(row_index, row_height)
        # Hide Notes column (col 5) in compact mode
        self.task_table.setColumnHidden(5, self._compact_mode)
        self.compact_btn.setText("⊞ Normal" if self._compact_mode else "⊟ Compact")


# ── Theme ──────────────────────────────────────────────────────────────────────

_EMBEDDED_QSS = """
QMainWindow { background-color: #0a0e27; color: #ffffff; }
QWidget { color: #ffffff; font-family: "Segoe UI", Arial, sans-serif; font-size: 11pt; }
QMenuBar { background-color: #0a0e27; color: #ffffff; border-bottom: 1px solid #2d3748; padding: 4px 0px; spacing: 16px; }
QMenuBar::item:selected { background-color: #1f2937; color: #3b82f6; }
QMenuBar::item:pressed { background-color: #1e3a8a; }
QMenu { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 4px; padding: 4px 0px; }
QMenu::item { padding: 8px 16px; }
QMenu::item:selected { background-color: #1f2937; color: #3b82f6; }
QMenu::item:pressed { background-color: #1e3a8a; }
QMenu::separator { background-color: #2d3748; height: 1px; margin: 4px 0px; }
QPushButton, QToolButton { background-color: #dc2626; color: #ffffff; border: none; border-radius: 6px; padding: 6px 14px; font-weight: 600; font-size: 11pt; }
QPushButton:hover, QToolButton:hover { background-color: #b91c1c; }
QPushButton:pressed, QToolButton:pressed { background-color: #991b1b; }
QPushButton:focus { outline: none; border: 2px solid #3b82f6; }
QPushButton:disabled, QToolButton:disabled { background-color: #4b5563; color: #6b7280; }
QPushButton#secondaryButton { background-color: #1e3a8a; }
QPushButton#secondaryButton:hover { background-color: #1e40af; }
QPushButton#tertiaryButton { background-color: transparent; border: 1px solid #4b5563; color: #3b82f6; }
QPushButton#tertiaryButton:hover { background-color: #1f2937; border: 1px solid #3b82f6; }
QLineEdit { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 6px; padding: 6px 8px; selection-background-color: #3b82f6; }
QLineEdit:focus { border: 2px solid #3b82f6; }
QLineEdit:disabled { background-color: #050810; color: #6b7280; }
QTextEdit, QPlainTextEdit { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 6px; padding: 6px 8px; selection-background-color: #3b82f6; }
QTextEdit:focus, QPlainTextEdit:focus { border: 2px solid #3b82f6; }
QComboBox { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 6px; padding: 6px 8px; }
QComboBox:focus { border: 2px solid #3b82f6; }
QComboBox:disabled { background-color: #050810; color: #6b7280; }
QComboBox::drop-down { border: none; padding-right: 8px; }
QComboBox QAbstractItemView { background-color: #141829; color: #ffffff; selection-background-color: #3b82f6; border: 1px solid #2d3748; outline: none; }
QDateEdit { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 6px; padding: 6px 8px; }
QDateEdit:focus { border: 2px solid #3b82f6; }
QSpinBox, QDoubleSpinBox { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; border-radius: 6px; padding: 6px 8px; }
QSpinBox:focus, QDoubleSpinBox:focus { border: 2px solid #3b82f6; }
QSpinBox::up-button, QDoubleSpinBox::up-button, QSpinBox::down-button, QDoubleSpinBox::down-button { background-color: #050810; border: none; width: 20px; }
QCheckBox { color: #ffffff; spacing: 8px; }
QCheckBox::indicator { width: 18px; height: 18px; border-radius: 4px; border: 1px solid #4b5563; background-color: #141829; }
QCheckBox::indicator:hover { border: 1px solid #3b82f6; background-color: #1f2937; }
QCheckBox::indicator:checked { background-color: #10b981; border: 1px solid #10b981; }
QLabel { color: #ffffff; font-family: "Segoe UI", sans-serif; }
QLabel#ProjectTitle { font-size: 14pt; font-weight: 700; color: #ffffff; }
QLabel#ProjectSubtitle { color: #9ca3af; font-size: 10pt; }
QLabel#SectionTitle { font-size: 12pt; font-weight: 600; color: #ffffff; }
QLabel#StatTitle { color: #9ca3af; font-size: 7pt; }
QLabel#StatValue { font-size: 10pt; font-weight: 700; }
QLabel#FinDataMeta { color: #9ca3af; font-size: 8pt; }
QLabel#MetaCaption { color: #9ca3af; font-size: 9pt; font-weight: 600; }
QLabel#MetaValue { color: #d1d5db; font-size: 9pt; }
QTabWidget::pane { border: 1px solid #2d3748; background-color: #141829; }
QTabBar::tab { background-color: #050810; color: #9ca3af; padding: 6px 18px; border: 1px solid #2d3748; border-bottom: none; border-radius: 6px 6px 0 0; font-weight: 500; }
QTabBar::tab:selected { background-color: #141829; color: #ffffff; font-weight: 600; border-bottom: 3px solid #dc2626; }
QTabBar::tab:hover:!selected { background-color: #1f2937; color: #d1d5db; }
QTableWidget, QTableView { background-color: transparent; alternate-background-color: rgba(10, 14, 39, 140); gridline-color: #2d3748; border: 1px solid #2d3748; border-radius: 6px; color: #ffffff; }
QTableWidget::item, QTableView::item { background-color: rgba(20, 24, 41, 140); padding: 3px 6px; border: none; color: #ffffff; }
QTableWidget::item:alternate, QTableView::item:alternate { background-color: rgba(10, 14, 39, 140); }
QTableWidget::item:selected, QTableView::item:selected { background-color: #1e40af; color: #ffffff; }
QTableWidget::item:hover, QTableView::item:hover { background-color: #1f2937; }
QHeaderView::section { background-color: rgba(5, 8, 16, 180); color: #e5e7eb; padding: 6px 8px; border: none; border-right: 1px solid #2d3748; border-bottom: 1px solid #2d3748; font-weight: 600; }
QHeaderView::section:hover { background-color: #1f2937; }
QListWidget { background: transparent; border: 1px solid #2d3748; border-radius: 10px; padding: 8px; color: #ffffff; }
QListWidget::item { background: transparent; border-radius: 6px; padding: 8px; margin: 2px 0; color: #ffffff; }
QListWidget::item:selected { background: #1e3a8a; color: white; }
QListWidget::item:hover:!selected { background: #1f2937; }
QTreeWidget { background: transparent; border: 1px solid #2d3748; border-radius: 10px; padding: 4px; color: #ececec; outline: none; }
QTreeWidget::item { border-radius: 6px; padding: 5px 8px; margin: 1px 0; }
QTreeWidget::item:selected { background: #1e3a8a; color: white; }
QTreeWidget::item:hover:!selected { background: #1f2937; }
QScrollBar:vertical { background-color: #0a0e27; width: 8px; border: none; }
QScrollBar::handle:vertical { background-color: #4b5563; border-radius: 4px; min-height: 20px; }
QScrollBar::handle:vertical:hover { background-color: #6b7280; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; border: none; background: none; }
QScrollBar:horizontal { background-color: #0a0e27; height: 8px; border: none; }
QScrollBar::handle:horizontal { background-color: #4b5563; border-radius: 4px; min-width: 20px; }
QScrollBar::handle:horizontal:hover { background-color: #6b7280; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; border: none; background: none; }
QGroupBox { color: #ffffff; border: 1px solid #2d3748; border-radius: 8px; margin-top: 12px; padding-top: 12px; font-weight: 600; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0px 4px; }
QDialog { background-color: #0a0e27; }
QMessageBox QLabel { color: #ffffff; }
QMessageBox QPushButton { min-width: 80px; }
QSplitter::handle { background-color: #2d3748; }
QSplitter::handle:hover { background-color: #3b82f6; }
QFrame[frameShape="4"], QFrame[frameShape="5"] { border: 1px solid #2d3748; background-color: transparent; }
QToolTip { background-color: #141829; color: #ffffff; border: 1px solid #2d3748; padding: 6px 10px; border-radius: 4px; }
QStatusBar { background-color: #050810; color: #d1d5db; border-top: 1px solid #2d3748; padding: 2px 12px; }
QProgressBar { border: 1px solid #2d3748; border-radius: 6px; background-color: #050810; text-align: center; color: #ffffff; }
QProgressBar::chunk { background-color: #dc2626; border-radius: 4px; }
#Panel, #StatCard { background: rgba(20, 24, 41, 180); border: 1px solid #2d3748; border-radius: 14px; }
#UpdateBanner { background: rgba(30, 58, 138, 220); border-top: 1px solid #3b82f6; }
QLabel#UpdateMsg { color: #93c5fd; font-weight: 600; }
#InstallBtn { background: #dc2626; border: 1px solid #ef4444; color: white; font-weight: 700; }
#InstallBtn:hover { background: #b91c1c; }
#RestoreBtn { background: #92400e; border: 1px solid #f59e0b; color: #f59e0b; font-weight: 700; }
#RestoreBtn:hover { background: #b45309; }
QLabel#PassBadge { background: #10b981; color: white; border-radius: 8px; font-weight: 700; font-size: 10pt; padding: 2px 6px; }
QLabel#FailBadge { background: #ef4444; color: white; border-radius: 8px; font-weight: 700; font-size: 10pt; padding: 2px 6px; }
QLabel#ArchivedBadge { background: #92400e; color: #f59e0b; border-radius: 8px; font-weight: 700; font-size: 10pt; padding: 0px 10px; }
QFrame#ResizeHandle, QFrame#VResizeHandle { background: #2d3748; border: none; }
QFrame#ResizeHandle:hover, QFrame#VResizeHandle:hover { background: #3b82f6; }
QPushButton#Div25Btn { background: #1e3a5f; border: 1px solid #2d5a8e; border-radius: 6px; color: #5ba3f5; font-weight: 600; padding: 2px 6px; }
QPushButton#Div25Btn:hover { background: #2d5a8e; color: #87c3ff; }
QPushButton#Div25Btn:disabled { background: #1a1a1a; border: 1px solid #333333; color: #555555; }
QPlainTextEdit#ReadOnlyNotes { background: #050810; color: #9ca3af; border: 1px solid #2d3748; }
QLabel#errorLabel { color: #ef4444; }
QLabel#dialogTitle { font-size: 18pt; font-weight: 700; color: #ffffff; }
QLabel#dialogSubtitle { font-size: 11pt; color: #9ca3af; }
"""


def apply_phoenix_theme(app: QApplication) -> None:
    app.setStyle("Fusion")
    qss_path = str(_resource_path("phoenix_style.qss"))
    if os.path.exists(qss_path):
        with open(qss_path, "r") as fh:
            app.setStyleSheet(fh.read())
    else:
        app.setStyleSheet(_EMBEDDED_QSS)



def main() -> int:
    app = QApplication(sys.argv)
    apply_phoenix_theme(app)
    settings = QSettings("ATSInc", "ProjectTrackingTool")

    # ── User login ──────────────────────────────────────────────────────── #
    # Resolve the users.json path from the configured data folder.
    custom_folder = str(  settings.value("dataFolder", "")).strip()
    if custom_folder:
        users_path = Path(custom_folder) / "users.json"
    else:
        users_path = _app_data_path().parent / "users.json"

    user_manager = UserManager(users_path)
    current_user = ""

    if user_manager.has_any_users():
        # Try remember-me session first
        auto_logged_in = False
        if _SESSION_PATH.exists():
            try:
                session_data = json.loads(_SESSION_PATH.read_text(encoding="utf-8"))
                remembered = session_data.get("username", "")
                if remembered and user_manager.get_user(remembered):
                    current_user = remembered
                    auto_logged_in = True
            except (OSError, ValueError, KeyError):
                pass

        if not auto_logged_in:
            login_dlg = LoginDialog(user_manager)
            if login_dlg.exec() != QDialog.DialogCode.Accepted:
                return 0
            user = login_dlg.logged_in_user()
            if user is None:
                return 0
            current_user = user.username
            # If first login, force password change
            if user.must_change_password:
                chpw_dlg = ChangePasswordDialog(user_manager, current_user)
                chpw_dlg.exec()  # They must complete this; closing is treated as done
    # If no users exist yet, the app opens without requiring login so the admin
    # can set up accounts via File > Manage Users.

    # Auto-backup on open (keep last 10 backups)
    _backup_data_file(_app_data_path())

    window = MainWindow(current_user=current_user)
    window.showMaximized()
    return int(app.exec())


if __name__ == "__main__":
    raise SystemExit(main())