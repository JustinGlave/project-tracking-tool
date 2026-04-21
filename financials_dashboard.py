from __future__ import annotations

from typing import Protocol

from PySide6.QtCore import Qt, QSortFilterProxyModel, QAbstractTableModel, QModelIndex, QPersistentModelIndex
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTabWidget,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from financials_dialog import FinancialsDialog
from financials_models import FinancialSnapshot


class DashboardProvider(Protocol):
    def get_all_financials(self) -> list[FinancialSnapshot]: ...
    def get_financials(self, job_number: str) -> FinancialSnapshot: ...


def _valid(s: FinancialSnapshot) -> bool:
    name = s.job_name.strip()
    return bool(name) and name != "0.0"


# ── Financial Overview table ────────────────────────────────────────────────

_FIN_COLUMNS = [
    ("Job #",          "job_number"),
    ("Job Name",       "job_name"),
    ("PM",             "project_manager"),
    ("Status",         "status"),
    ("Contract Value", "contract_value"),
    ("Billed to Date", "billed_to_date"),
    ("Actual Cost",    "actual_cost"),
    ("Booked Margin",  "booked_margin"),
    ("Actual Margin",  "actual_margin"),
    ("Diff Margin",    "differential_margin"),
]

_FIN_MONEY_COLS = {4, 5, 6}
_FIN_PCT_COLS   = {7, 8, 9}
_FIN_DIFF_COL   = 9


class _FinModel(QAbstractTableModel):
    def __init__(self, rows: list[FinancialSnapshot], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows = rows
        self._totals = self._compute_totals()

    def _compute_totals(self) -> list:
        n = len(self._rows)
        cv = sum(s.contract_value for s in self._rows)
        bd = sum(s.billed_to_date for s in self._rows)
        ac = sum(s.actual_cost for s in self._rows)
        bm = (sum(s.booked_margin for s in self._rows) / n) if n else 0.0
        am = (sum(s.actual_margin for s in self._rows) / n) if n else 0.0
        dm = (sum(s.differential_margin for s in self._rows) / n) if n else 0.0
        return ["", f"{n} projects", "", "TOTALS / AVG", cv, bd, ac, bm, am, dm]

    def rowCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        return len(self._rows) + 1

    def columnCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        return len(_FIN_COLUMNS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return _FIN_COLUMNS[section][0]
        return None

    def data(self, index: QModelIndex | QPersistentModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        is_totals = row == len(self._rows)
        val = self._totals[col] if is_totals else getattr(self._rows[row], _FIN_COLUMNS[col][1])

        if role == Qt.ItemDataRole.DisplayRole:
            return self._fmt(col, val, is_totals)
        if role == Qt.ItemDataRole.UserRole:
            return val
        if role == Qt.ItemDataRole.ForegroundRole:
            if col == _FIN_DIFF_COL:
                v = self._totals[col] if is_totals else self._rows[row].differential_margin
                if v >= 0.02:   return QColor("#4caf50")
                if v <= -0.02:  return QColor("#f44336")
                return QColor("#ff9800")
            if is_totals:
                return QColor("#aaaaaa")
        if role == Qt.ItemDataRole.FontRole and is_totals:
            f = QFont(); f.setBold(True); return f
        if role == Qt.ItemDataRole.BackgroundRole and is_totals:
            return QColor("#1e1e1e")
        return None

    @staticmethod
    def _fmt(col: int, val, is_totals: bool) -> str:
        if col in _FIN_MONEY_COLS:
            return f"${val:,.0f}" if isinstance(val, float) else str(val)
        if col in _FIN_PCT_COLS:
            prefix = "avg " if is_totals else ""
            return f"{prefix}{val * 100:.1f}%" if isinstance(val, float) else str(val)
        return str(val) if val else "—"

    def snapshot_at(self, row: int) -> FinancialSnapshot | None:
        return self._rows[row] if row < len(self._rows) else None


# ── Labor table ─────────────────────────────────────────────────────────────

_LAB_COLUMNS = [
    ("Job #",              "job_number"),
    ("Job Name",           "job_name"),
    ("PM",                 "project_manager"),
    ("Status",             "status"),
    ("PM Hours",           "pm_hours_actual"),
    ("Tech Hours",         "tech_hours_actual"),
    ("Total Hours",        "total_hours_actual"),
    ("PM Cost",            "pm_cost_actual"),
    ("Tech Cost",          "tech_cost_actual"),
    ("Total Labor Cost",   "total_labor_actual"),
    ("Labor Budget Rem",   None),   # derived: rem_pct / rem_usd
    ("Total Labor Budget", None),   # derived: total_labor_actual + labor_rem_usd
]

_LAB_HOUR_COLS  = {4, 5, 6}
_LAB_MONEY_COLS = {7, 8, 9, 11}
_LAB_REM_COL    = 10
_LAB_BUDGET_COL = 11


def _labor_rem_usd(s: FinancialSnapshot) -> float:
    return s.labor_rem_usd

def _total_labor_budget(s: FinancialSnapshot) -> float:
    return s.total_labor_actual + s.labor_rem_usd


class _LaborModel(QAbstractTableModel):
    def __init__(self, rows: list[FinancialSnapshot], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._rows = rows
        self._totals = self._compute_totals()

    def _compute_totals(self) -> list:
        n = len(self._rows)
        pmh  = sum(s.pm_hours_actual for s in self._rows)
        tech = sum(s.tech_hours_actual for s in self._rows)
        tot  = sum(s.total_hours_actual for s in self._rows)
        pmc  = sum(s.pm_cost_actual for s in self._rows)
        tec  = sum(s.tech_cost_actual for s in self._rows)
        tlc  = sum(s.total_labor_actual for s in self._rows)
        rem  = sum(_labor_rem_usd(s) for s in self._rows)
        bud  = sum(_total_labor_budget(s) for s in self._rows)
        return ["", f"{n} projects", "", "TOTALS", pmh, tech, tot, pmc, tec, tlc, rem, bud]

    def rowCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        return len(self._rows) + 1

    def columnCount(self, parent: QModelIndex | QPersistentModelIndex = QModelIndex()) -> int:
        return len(_LAB_COLUMNS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return _LAB_COLUMNS[section][0]
        return None

    def _raw(self, row: int, col: int):
        if row == len(self._rows):
            return self._totals[col]
        s = self._rows[row]
        if col == _LAB_REM_COL:
            return _labor_rem_usd(s)
        if col == _LAB_BUDGET_COL:
            return _total_labor_budget(s)
        return getattr(s, _LAB_COLUMNS[col][1])

    def data(self, index: QModelIndex | QPersistentModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        is_totals = row == len(self._rows)

        if role == Qt.ItemDataRole.DisplayRole:
            return self._fmt(row, col)
        if role == Qt.ItemDataRole.UserRole:
            return self._raw(row, col)
        if role == Qt.ItemDataRole.ForegroundRole and is_totals:
            return QColor("#aaaaaa")
        if role == Qt.ItemDataRole.FontRole and is_totals:
            f = QFont(); f.setBold(True); return f
        if role == Qt.ItemDataRole.BackgroundRole and is_totals:
            return QColor("#1e1e1e")
        return None

    def _fmt(self, row: int, col: int) -> str:
        is_totals = row == len(self._rows)
        if col == _LAB_REM_COL:
            if is_totals:
                return f"${self._totals[col]:,.0f}"
            s = self._rows[row]
            return f"{s.labor_rem_pct * 100:.1f}%  /  ${s.labor_rem_usd:,.0f}"
        val = self._raw(row, col)
        if col in _LAB_HOUR_COLS:
            return f"{val:,.1f} hrs" if isinstance(val, float) else str(val)
        if col in _LAB_MONEY_COLS:
            return f"${val:,.0f}" if isinstance(val, float) else str(val)
        return str(val) if val else "—"

    def snapshot_at(self, row: int) -> FinancialSnapshot | None:
        return self._rows[row] if row < len(self._rows) else None


# ── Shared table factory ─────────────────────────────────────────────────────

def _make_table(model: QAbstractTableModel, on_double_click) -> tuple[QTableView, QSortFilterProxyModel]:
    proxy = QSortFilterProxyModel()
    proxy.setSourceModel(model)
    proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    proxy.setFilterKeyColumn(-1)
    proxy.setSortRole(Qt.ItemDataRole.UserRole)

    table = QTableView()
    table.setModel(proxy)
    table.setSortingEnabled(True)
    table.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
    table.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
    table.setAlternatingRowColors(True)
    table.verticalHeader().setVisible(False)
    table.horizontalHeader().setStretchLastSection(True)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    table.doubleClicked.connect(on_double_click)
    return table, proxy


# ── Dialog ───────────────────────────────────────────────────────────────────

class FinancialsDashboardDialog(QDialog):
    def __init__(self, provider: DashboardProvider, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._provider = provider

        self.setWindowTitle("Financials Dashboard — All Projects")
        self.resize(1200, 700)
        self.setMinimumSize(900, 500)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        # Toolbar
        toolbar = QHBoxLayout()
        title = QLabel("All Projects — Financial Dashboard")
        title.setObjectName("FinancialDialogTitle")
        toolbar.addWidget(title)
        toolbar.addStretch(1)

        self._search = QLineEdit()
        self._search.setPlaceholderText("Filter by job #, name, PM, status…")
        self._search.setFixedWidth(260)
        self._search.textChanged.connect(self._on_filter)
        toolbar.addWidget(self._search)

        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self._refresh)
        toolbar.addWidget(refresh_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        toolbar.addWidget(close_btn)
        root.addLayout(toolbar)

        self._status_label = QLabel("")
        self._status_label.setObjectName("FinancialMeta")
        root.addWidget(self._status_label)

        # Tabs
        self._tabs = QTabWidget()
        root.addWidget(self._tabs, 1)

        # Financial Overview tab
        self._fin_model = _FinModel([])
        self._fin_table, self._fin_proxy = _make_table(self._fin_model, self._open_from_fin)
        fin_tab = QWidget()
        fin_layout = QVBoxLayout(fin_tab)
        fin_layout.setContentsMargins(0, 6, 0, 0)
        fin_layout.addWidget(self._fin_table)
        self._tabs.addTab(fin_tab, "Financial Overview")

        # Labor tab
        self._lab_model = _LaborModel([])
        self._lab_table, self._lab_proxy = _make_table(self._lab_model, self._open_from_labor)
        lab_tab = QWidget()
        lab_layout = QVBoxLayout(lab_tab)
        lab_layout.setContentsMargins(0, 6, 0, 0)
        lab_layout.addWidget(self._lab_table)
        self._tabs.addTab(lab_tab, "Labor Hours & Cost")

        # Warranty & Archived tab
        self._war_model = _FinModel([])
        self._war_table, self._war_proxy = _make_table(self._war_model, self._open_from_war)
        war_tab = QWidget()
        war_layout = QVBoxLayout(war_tab)
        war_layout.setContentsMargins(0, 6, 0, 0)
        war_layout.addWidget(self._war_table)
        self._tabs.addTab(war_tab, "Warranty & Archived")

        self._search.textChanged.connect(self._on_filter)

        hint = QLabel("Double-click a row to view full financial details for that project.")
        hint.setObjectName("FinancialMeta")
        root.addWidget(hint)

        self._apply_styles()
        self._refresh()

    def _refresh(self) -> None:
        if hasattr(self._provider, "force_refresh"):
            self._provider.force_refresh()
        all_snaps = [s for s in self._provider.get_all_financials() if _valid(s)]
        all_snaps.sort(key=lambda s: s.job_number)

        inactive = [s for s in all_snaps if "warranty" in s.status.lower() or "archiv" in s.status.lower()]
        active   = [s for s in all_snaps if s not in inactive]

        def _load(model_attr, proxy_attr, table_attr, cls, rows):
            m = cls(rows)
            setattr(self, model_attr, m)
            getattr(self, proxy_attr).setSourceModel(m)
            getattr(self, table_attr).resizeColumnsToContents()

        _load("_fin_model", "_fin_proxy", "_fin_table", _FinModel,   active)
        _load("_lab_model", "_lab_proxy", "_lab_table", _LaborModel, active)
        _load("_war_model", "_war_proxy", "_war_table", _FinModel,   inactive)

        ts = f"  |  Data as of {all_snaps[0].last_refreshed}" if all_snaps and all_snaps[0].last_refreshed else ""
        self._status_label.setText(
            f"{len(active)} active  |  {len(inactive)} warranty/archived{ts}"
        )
        self._tabs.setTabText(0, f"Financial Overview ({len(active)})")
        self._tabs.setTabText(1, f"Labor Hours & Cost ({len(active)})")
        self._tabs.setTabText(2, f"Warranty & Archived ({len(inactive)})")

    def _on_filter(self, text: str) -> None:
        for proxy in (self._fin_proxy, self._lab_proxy, self._war_proxy):
            proxy.setFilterFixedString(text)

    def _open_from_fin(self, index) -> None:
        snap = self._fin_model.snapshot_at(self._fin_proxy.mapToSource(index).row())
        if snap:
            FinancialsDialog(job_number=snap.job_number, provider=self._provider, parent=self).exec()

    def _open_from_labor(self, index) -> None:
        snap = self._lab_model.snapshot_at(self._lab_proxy.mapToSource(index).row())
        if snap:
            FinancialsDialog(job_number=snap.job_number, provider=self._provider, parent=self).exec()

    def _open_from_war(self, index) -> None:
        snap = self._war_model.snapshot_at(self._war_proxy.mapToSource(index).row())
        if snap:
            FinancialsDialog(job_number=snap.job_number, provider=self._provider, parent=self).exec()

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QLabel#FinancialDialogTitle { font-size: 14pt; font-weight: 700; }
            QLabel#FinancialMeta { color: #888888; font-size: 9pt; }
            QTabWidget::pane { border: 1px solid #3a3a3a; border-radius: 8px; }
            QTabBar::tab {
                padding: 6px 18px;
                background: #1e1e1e;
                border: 1px solid #3a3a3a;
                border-bottom: none;
                border-radius: 6px 6px 0 0;
            }
            QTabBar::tab:selected { background: #2a2a2a; font-weight: 600; }
            QTableView {
                border: none;
                gridline-color: #2a2a2a;
            }
            QTableView::item:selected { background: #1a3a5c; }
            QHeaderView::section {
                background: #1e1e1e;
                border: none;
                border-bottom: 1px solid #3a3a3a;
                padding: 4px 8px;
                font-weight: 600;
            }
            """
        )
