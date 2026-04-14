from __future__ import annotations

from typing import Protocol

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from financials_models import FinancialSnapshot


class FinancialsProvider(Protocol):
    def get_financials(self, job_number: str) -> FinancialSnapshot: ...


class _ValueLabel(QLabel):
    def __init__(self, text: str = "—", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.setObjectName("FinancialValue")


class _SectionCard(QFrame):
    def __init__(self, title: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("FinancialCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setObjectName("FinancialSectionTitle")
        layout.addWidget(title_label)

        self.body = QWidget()
        self.body_layout = QFormLayout(self.body)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(6)
        layout.addWidget(self.body)

    def add_row(self, label: str, widget: QWidget) -> None:
        self.body_layout.addRow(label, widget)


class FinancialsDialog(QDialog):
    def __init__(
        self,
        job_number: str,
        provider: FinancialsProvider,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._provider = provider
        self._job_number = job_number.strip()
        self._snapshot = FinancialSnapshot.empty(self._job_number)

        self.setWindowTitle(f"Financials — {self._job_number or 'No Job Number'}")
        self.resize(1000, 800)
        self.setMinimumSize(800, 600)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # ── Toolbar ──────────────────────────────────────────────────── #
        toolbar = QHBoxLayout()
        self.title_label = QLabel(f"Financials for Job {self._job_number or '—'}")
        self.title_label.setObjectName("FinancialDialogTitle")
        toolbar.addWidget(self.title_label)
        toolbar.addStretch(1)

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.clicked.connect(self.refresh_data)
        toolbar.addWidget(self.refresh_btn)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        toolbar.addWidget(close_btn)
        root.addLayout(toolbar)

        self.last_refreshed_label = QLabel("Last refreshed: —")
        self.last_refreshed_label.setObjectName("FinancialMeta")
        root.addWidget(self.last_refreshed_label)

        # ── Scrollable grid ──────────────────────────────────────────── #
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root.addWidget(scroll, 1)

        container = QWidget()
        scroll.setWidget(container)
        grid = QGridLayout(container)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setSpacing(10)

        self.summary_card  = _SectionCard("Project Summary")
        self.billing_card  = _SectionCard("Billing & Costs")
        self.margins_card  = _SectionCard("Margins")
        self.labor_card    = _SectionCard("Labor")
        self.budgets_card  = _SectionCard("Budget Remaining")
        self.notes_card    = _SectionCard("Notes")

        grid.addWidget(self.summary_card,  0, 0)
        grid.addWidget(self.billing_card,  0, 1)
        grid.addWidget(self.margins_card,  1, 0)
        grid.addWidget(self.labor_card,    1, 1)
        grid.addWidget(self.budgets_card,  2, 0, 1, 2)
        grid.addWidget(self.notes_card,    3, 0, 1, 2)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)

        self._build_fields()
        self._apply_styles()
        self.refresh_data()

    # ------------------------------------------------------------------ #
    # Field construction                                                   #
    # ------------------------------------------------------------------ #

    def _build_fields(self) -> None:
        # Summary
        self.job_name = _ValueLabel()
        self.pm       = _ValueLabel()
        self.sales    = _ValueLabel()
        self.status   = _ValueLabel()

        self.summary_card.add_row("Job Name",        self.job_name)
        self.summary_card.add_row("Project Manager", self.pm)
        self.summary_card.add_row("Sales",           self.sales)
        self.summary_card.add_row("Status",          self.status)

        # Billing
        self.contract_value  = _ValueLabel()
        self.billed_to_date  = _ValueLabel()
        self.amount_paid     = _ValueLabel()
        self.estimated_cost  = _ValueLabel()
        self.actual_cost     = _ValueLabel()

        self.billing_card.add_row("Contract Value",       self.contract_value)
        self.billing_card.add_row("Billed to Date",       self.billed_to_date)
        self.billing_card.add_row("Amount Paid to Date",  self.amount_paid)
        self.billing_card.add_row("Estimated At Completion Cost", self.estimated_cost)
        self.billing_card.add_row("Actual Cost",          self.actual_cost)

        # Margins
        self.booked_margin  = _ValueLabel()
        self.actual_margin  = _ValueLabel()
        self.diff_margin    = _ValueLabel()

        self.margins_card.add_row("Booked Margin",       self.booked_margin)
        self.margins_card.add_row("Actual Margin",       self.actual_margin)
        self.margins_card.add_row("Differential Margin", self.diff_margin)

        # Labor
        self.pm_hours         = _ValueLabel()
        self.tech_hours       = _ValueLabel()
        self.total_hours      = _ValueLabel()
        self.pm_cost          = _ValueLabel()
        self.tech_cost        = _ValueLabel()
        self.total_labor      = _ValueLabel()

        self.labor_card.add_row("PM Hours",             self.pm_hours)
        self.labor_card.add_row("Technician Hours",     self.tech_hours)
        self.labor_card.add_row("Total Hours",          self.total_hours)
        self.labor_card.add_row("PM Cost",              self.pm_cost)
        self.labor_card.add_row("Technician Cost",      self.tech_cost)
        self.labor_card.add_row("Total Labor Cost",     self.total_labor)

        # Budget remaining (% remaining / $ remaining)
        self.labor_rem    = _ValueLabel()
        self.material_rem = _ValueLabel()
        self.warranty_rem = _ValueLabel()
        self.travel_rem   = _ValueLabel()
        self.sub_rem      = _ValueLabel()
        self.odc_rem      = _ValueLabel()

        self.budgets_card.add_row("Labor Budget",      self.labor_rem)
        self.budgets_card.add_row("Material",          self.material_rem)
        self.budgets_card.add_row("Warranty",          self.warranty_rem)
        self.budgets_card.add_row("Travel",            self.travel_rem)
        self.budgets_card.add_row("Subcontract",       self.sub_rem)
        self.budgets_card.add_row("Other Direct Cost", self.odc_rem)

        hint = QLabel("  % remaining  /  $ remaining")
        hint.setObjectName("FinancialMeta")
        self.budgets_card.body_layout.addRow("", hint)

        # Notes
        self.notes_text = QTextEdit()
        self.notes_text.setReadOnly(True)
        self.notes_text.setMinimumHeight(80)
        self.notes_card.add_row("", self.notes_text)

    # ------------------------------------------------------------------ #
    # Styles                                                               #
    # ------------------------------------------------------------------ #

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QLabel#FinancialDialogTitle {
                font-size: 15pt;
                font-weight: 700;
            }
            QLabel#FinancialMeta {
                color: #888888;
                font-size: 9pt;
            }
            QFrame#FinancialCard {
                border: 1px solid #3a3a3a;
                border-radius: 12px;
                background: rgba(38, 38, 38, 140);
            }
            QLabel#FinancialSectionTitle {
                font-size: 11pt;
                font-weight: 700;
            }
            QLabel#FinancialValue {
                font-size: 10pt;
                color: #e8e8e8;
            }
            QTextEdit {
                border: 1px solid #404040;
                border-radius: 8px;
                background: #151515;
            }
            """
        )

    # ------------------------------------------------------------------ #
    # Data helpers                                                         #
    # ------------------------------------------------------------------ #

    @staticmethod
    def _money(value: float) -> str:
        return f"${value:,.2f}"

    @staticmethod
    def _pct(value: float) -> str:
        return f"{value * 100:.1f}%"

    @staticmethod
    def _rem_pair(pct: float, usd: float) -> str:
        return f"{pct * 100:.1f}%  /  ${usd:,.2f}"

    # ------------------------------------------------------------------ #
    # Refresh                                                              #
    # ------------------------------------------------------------------ #

    def refresh_data(self) -> None:
        if not self._job_number:
            QMessageBox.information(
                self, "Missing job number", "This project does not have a job number yet."
            )
            return

        self.refresh_btn.setEnabled(False)
        try:
            if hasattr(self._provider, "force_refresh"):
                self._provider.force_refresh()
            self._snapshot = self._provider.get_financials(self._job_number)
            self._render_snapshot()
        except Exception as exc:
            QMessageBox.critical(self, "Financial refresh failed", str(exc))
        finally:
            self.refresh_btn.setEnabled(True)

    def _render_snapshot(self) -> None:
        s = self._snapshot
        self.title_label.setText(f"Financials for Job {s.job_number}")
        self.last_refreshed_label.setText(f"Last refreshed: {s.last_refreshed or '—'}")

        self.job_name.setText(s.job_name or "—")
        self.pm.setText(s.project_manager or "—")
        self.sales.setText(s.sales_person or "—")
        self.status.setText(s.status or "—")

        self.contract_value.setText(self._money(s.contract_value))
        self.billed_to_date.setText(self._money(s.billed_to_date))
        self.amount_paid.setText(self._money(s.amount_paid_to_date))
        self.estimated_cost.setText(self._money(s.estimated_cost))
        self.actual_cost.setText(self._money(s.actual_cost))

        self.booked_margin.setText(self._pct(s.booked_margin))
        self.actual_margin.setText(self._pct(s.actual_margin))
        diff = s.differential_margin
        self.diff_margin.setText(self._pct(diff))
        if diff >= 0.02:
            self.diff_margin.setStyleSheet("color: #4caf50; font-weight: bold;")  # green
        elif diff <= -0.02:
            self.diff_margin.setStyleSheet("color: #f44336; font-weight: bold;")  # red
        else:
            self.diff_margin.setStyleSheet("color: #ff9800; font-weight: bold;")  # yellow

        self.pm_hours.setText(f"{s.pm_hours_actual:,.1f} hrs")
        self.tech_hours.setText(f"{s.tech_hours_actual:,.1f} hrs")
        self.total_hours.setText(f"{s.total_hours_actual:,.1f} hrs")
        self.pm_cost.setText(self._money(s.pm_cost_actual))
        self.tech_cost.setText(self._money(s.tech_cost_actual))
        self.total_labor.setText(self._money(s.total_labor_actual))

        self.labor_rem.setText(self._rem_pair(s.labor_rem_pct, s.labor_rem_usd))
        self.material_rem.setText(self._rem_pair(s.material_rem_pct, s.material_rem_usd))
        self.warranty_rem.setText(self._rem_pair(s.warranty_rem_pct, s.warranty_rem_usd))
        self.travel_rem.setText(self._rem_pair(s.travel_rem_pct, s.travel_rem_usd))
        self.sub_rem.setText(self._rem_pair(s.subcontract_rem_pct, s.subcontract_rem_usd))
        self.odc_rem.setText(self._rem_pair(s.odc_rem_pct, s.odc_rem_usd))

        self.notes_text.setPlainText("\n".join(s.notes) if s.notes else "")
