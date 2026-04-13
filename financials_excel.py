from __future__ import annotations

import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

from financials_models import FinancialSnapshot

logger = logging.getLogger(__name__)

# How long (seconds) before re-reading the file
_CACHE_TTL = 300  # 5 minutes

# Row index (0-based) where the column headers live in the PM sheet
_HEADER_ROW = 8

# Column indices in the PM-named sheet (e.g. "Justin Glave")
_COL_JOB_NUMBER    = 1
_COL_JOB_NAME      = 2
_COL_CONTRACT      = 3
_COL_PM            = 4
_COL_BILLED        = 5
_COL_PAID          = 6
_COL_SALES         = 7
_COL_EAC           = 8   # Estimated At Completion Cost
_COL_BOOKED_MARGIN = 9
_COL_ACTUAL_COST   = 10
_COL_ACTUAL_MARGIN = 11
_COL_DIFF_MARGIN   = 12
_COL_STATUS        = 13
_COL_PM_HOURS      = 14
_COL_TECH_HOURS    = 15
_COL_LABOR_REM_PCT = 16
_COL_MAT_REM_PCT   = 17
_COL_WARR_REM_PCT  = 18
_COL_TRAV_REM_PCT  = 19
_COL_SUB_REM_PCT   = 20
_COL_ODC_REM_PCT   = 21
_COL_PM_COST       = 22
_COL_TECH_COST     = 23
_COL_LABOR_REM_USD = 24
_COL_MAT_REM_USD   = 25
_COL_WARR_REM_USD  = 26
_COL_TRAV_REM_USD  = 27
_COL_SUB_REM_USD   = 28
_COL_ODC_REM_USD   = 29


def _flt(val) -> float:
    """Safely coerce a cell value to float."""
    try:
        return float(val) if val is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _str(val) -> str:
    """Safely coerce a cell value to str."""
    if val is None:
        return ""
    return str(val).strip()


def _job_key(raw) -> Optional[str]:
    """Normalise a job-number cell to a string key, or None if not a valid number."""
    if raw is None:
        return None
    try:
        n = int(float(raw))
        if n <= 0:
            return None
        return str(n)
    except (TypeError, ValueError):
        return None


class ExcelFinancialsProvider:
    """
    Reads financial data from the ODIN-connected Excel tracking workbook (.xlsb).

    The workbook is expected to have a sheet whose name matches ``sheet_name``
    (defaults to the first non-empty sheet whose row 8, col 1 == "Job Number").
    Data rows start at row 9 (0-based index 9).

    Results are cached for ``_CACHE_TTL`` seconds so repeated lookups within
    the same session don't re-read the file.
    """

    def __init__(self, file_path: str, sheet_name: str = "") -> None:
        self._file_path = file_path
        self._sheet_name = sheet_name
        self._cache: dict[str, FinancialSnapshot] = {}
        self._cache_mtime: float = 0.0
        self._cache_time: float = 0.0

    # ------------------------------------------------------------------ #
    # Public API                                                           #
    # ------------------------------------------------------------------ #

    @property
    def data_as_of(self) -> str:
        """Return the timestamp of the last successful file load, or empty string."""
        for snap in self._cache.values():
            return snap.last_refreshed or ""
        return ""

    def get_financials(self, job_number: str) -> FinancialSnapshot:
        """Return a FinancialSnapshot for *job_number*, reading the file if needed."""
        key = _job_key(job_number)
        if key is None:
            return FinancialSnapshot.empty(job_number)

        self._refresh_if_needed()

        snap = self._cache.get(key)
        if snap is None:
            snap = FinancialSnapshot.empty(job_number)
            snap.notes = [f"Job {job_number} not found in financial data file."]
            snap.touch()
        return snap

    # ------------------------------------------------------------------ #
    # Internal helpers                                                     #
    # ------------------------------------------------------------------ #

    def _refresh_if_needed(self) -> None:
        path = Path(self._file_path)
        if not path.exists():
            logger.warning("Financial data file not found: %s", self._file_path)
            return

        now = time.monotonic()
        try:
            mtime = path.stat().st_mtime
        except OSError:
            mtime = 0.0

        # Skip re-read if cache is fresh and file hasn't changed
        if (
            self._cache
            and (now - self._cache_time) < _CACHE_TTL
            and mtime == self._cache_mtime
        ):
            return

        try:
            self._load(path)
            self._cache_time = now
            self._cache_mtime = mtime
        except Exception:
            logger.exception("Failed to read financial data file: %s", self._file_path)

    def _load(self, path: Path) -> None:
        import pyxlsb  # imported lazily so the rest of the app works without it

        refreshed = datetime.now().replace(microsecond=0).isoformat(sep=" ")
        new_cache: dict[str, FinancialSnapshot] = {}

        with pyxlsb.open_workbook(str(path)) as wb:
            sheet_name = self._sheet_name or self._detect_sheet(wb)
            if not sheet_name:
                logger.warning("Could not find a valid data sheet in %s", path)
                return

            with wb.get_sheet(sheet_name) as sheet:
                for row_idx, row in enumerate(sheet.rows()):
                    if row_idx <= _HEADER_ROW:
                        continue  # skip title / header rows

                    vals = [c.v for c in row]
                    if len(vals) <= _COL_JOB_NUMBER:
                        continue

                    key = _job_key(vals[_COL_JOB_NUMBER])
                    if key is None:
                        continue

                    snap = FinancialSnapshot(
                        job_number=key,
                        job_name=_str(vals[_COL_JOB_NAME] if len(vals) > _COL_JOB_NAME else None),
                        project_manager=_str(vals[_COL_PM] if len(vals) > _COL_PM else None),
                        sales_person=_str(vals[_COL_SALES] if len(vals) > _COL_SALES else None),
                        status=_str(vals[_COL_STATUS] if len(vals) > _COL_STATUS else None),

                        contract_value=_flt(vals[_COL_CONTRACT] if len(vals) > _COL_CONTRACT else None),
                        billed_to_date=_flt(vals[_COL_BILLED] if len(vals) > _COL_BILLED else None),
                        amount_paid_to_date=_flt(vals[_COL_PAID] if len(vals) > _COL_PAID else None),
                        estimated_cost=_flt(vals[_COL_EAC] if len(vals) > _COL_EAC else None),
                        actual_cost=_flt(vals[_COL_ACTUAL_COST] if len(vals) > _COL_ACTUAL_COST else None),

                        booked_margin=_flt(vals[_COL_BOOKED_MARGIN] if len(vals) > _COL_BOOKED_MARGIN else None),
                        actual_margin=_flt(vals[_COL_ACTUAL_MARGIN] if len(vals) > _COL_ACTUAL_MARGIN else None),
                        differential_margin=_flt(vals[_COL_DIFF_MARGIN] if len(vals) > _COL_DIFF_MARGIN else None),

                        pm_hours_actual=_flt(vals[_COL_PM_HOURS] if len(vals) > _COL_PM_HOURS else None),
                        tech_hours_actual=_flt(vals[_COL_TECH_HOURS] if len(vals) > _COL_TECH_HOURS else None),
                        pm_cost_actual=_flt(vals[_COL_PM_COST] if len(vals) > _COL_PM_COST else None),
                        tech_cost_actual=_flt(vals[_COL_TECH_COST] if len(vals) > _COL_TECH_COST else None),

                        labor_rem_pct=_flt(vals[_COL_LABOR_REM_PCT] if len(vals) > _COL_LABOR_REM_PCT else None),
                        material_rem_pct=_flt(vals[_COL_MAT_REM_PCT] if len(vals) > _COL_MAT_REM_PCT else None),
                        warranty_rem_pct=_flt(vals[_COL_WARR_REM_PCT] if len(vals) > _COL_WARR_REM_PCT else None),
                        travel_rem_pct=_flt(vals[_COL_TRAV_REM_PCT] if len(vals) > _COL_TRAV_REM_PCT else None),
                        subcontract_rem_pct=_flt(vals[_COL_SUB_REM_PCT] if len(vals) > _COL_SUB_REM_PCT else None),
                        odc_rem_pct=_flt(vals[_COL_ODC_REM_PCT] if len(vals) > _COL_ODC_REM_PCT else None),

                        labor_rem_usd=_flt(vals[_COL_LABOR_REM_USD] if len(vals) > _COL_LABOR_REM_USD else None),
                        material_rem_usd=_flt(vals[_COL_MAT_REM_USD] if len(vals) > _COL_MAT_REM_USD else None),
                        warranty_rem_usd=_flt(vals[_COL_WARR_REM_USD] if len(vals) > _COL_WARR_REM_USD else None),
                        travel_rem_usd=_flt(vals[_COL_TRAV_REM_USD] if len(vals) > _COL_TRAV_REM_USD else None),
                        subcontract_rem_usd=_flt(vals[_COL_SUB_REM_USD] if len(vals) > _COL_SUB_REM_USD else None),
                        odc_rem_usd=_flt(vals[_COL_ODC_REM_USD] if len(vals) > _COL_ODC_REM_USD else None),

                        last_refreshed=refreshed,
                        notes=[],
                    )
                    new_cache[key] = snap

        self._cache = new_cache
        logger.info("Loaded %d financial records from %s", len(new_cache), path)

    @staticmethod
    def _detect_sheet(wb) -> str:
        """Find the first sheet whose row-8 col-1 cell looks like a job-number header."""
        import pyxlsb
        for name in wb.sheets:
            try:
                with wb.get_sheet(name) as sheet:
                    for row_idx, row in enumerate(sheet.rows()):
                        if row_idx == _HEADER_ROW:
                            vals = [c.v for c in row]
                            if len(vals) > _COL_JOB_NUMBER:
                                v = vals[_COL_JOB_NUMBER]
                                if isinstance(v, str) and "job number" in v.lower():
                                    return name
                            break
            except Exception:
                continue
        return ""
