from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass(slots=True)
class FinancialSnapshot:
    job_number: str
    job_id: Optional[int] = None
    job_name: str = ""
    gom: str = ""
    gos: str = ""
    project_manager: str = ""
    sales_person: str = ""
    status: str = ""

    contract_value: float = 0.0
    billed_to_date: float = 0.0
    amount_paid_to_date: float = 0.0
    actual_cost: float = 0.0
    estimated_cost: float = 0.0

    # Margin percentages (0.0–1.0)
    booked_margin: float = 0.0
    actual_margin: float = 0.0
    differential_margin: float = 0.0

    pm_hours_actual: float = 0.0
    tech_hours_actual: float = 0.0
    pm_cost_actual: float = 0.0
    tech_cost_actual: float = 0.0

    # Cost buckets — rem_pct = remaining budget as fraction (0.0–1.0)
    #                rem_usd = remaining budget in dollars
    labor_rem_pct: float = 0.0
    labor_rem_usd: float = 0.0
    material_rem_pct: float = 0.0
    material_rem_usd: float = 0.0
    warranty_rem_pct: float = 0.0
    warranty_rem_usd: float = 0.0
    travel_rem_pct: float = 0.0
    travel_rem_usd: float = 0.0
    subcontract_rem_pct: float = 0.0
    subcontract_rem_usd: float = 0.0
    odc_rem_pct: float = 0.0
    odc_rem_usd: float = 0.0

    last_refreshed: Optional[str] = None
    notes: list[str] = field(default_factory=list)

    @property
    def total_labor_actual(self) -> float:
        return round(self.pm_cost_actual + self.tech_cost_actual, 2)

    @property
    def total_hours_actual(self) -> float:
        return round(self.pm_hours_actual + self.tech_hours_actual, 2)

    @classmethod
    def empty(cls, job_number: str) -> "FinancialSnapshot":
        return cls(
            job_number=job_number,
            notes=["No financial data loaded yet."],
        )

    def touch(self) -> None:
        self.last_refreshed = datetime.now().replace(microsecond=0).isoformat(sep=" ")
