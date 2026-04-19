from __future__ import annotations

import dataclasses
import json
import shutil
import tempfile
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

from financials_models import FinancialSnapshot

try:
    import pyxlsb as _pyxlsb  # top-level so PyInstaller bundles it
    _PYXLSB_AVAILABLE = True
except ImportError:
    _PYXLSB_AVAILABLE = False

logger = logging.getLogger(__name__)

# How long (seconds) before re-reading the file
_CACHE_TTL = 300  # 5 minutes

# Row index (0-based) where the column headers live in the PM sheet
_HEADER_ROW = 8

# Column indices (0-based) in the PM-named sheet (e.g. "Justin Glave")
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

    def __init__(self, file_path: str, sheet_name: str = "", snapshot_path: Optional[Path] = None) -> None:
        self._file_path = file_path
        self._sheet_name = sheet_name
        self._snapshot_path = snapshot_path
        self._cache: dict[str, FinancialSnapshot] = {}
        self._cache_mtime: float = 0.0
        self._cache_time: float = 0.0
        self._load_error: str = ""

    # ------------------------------------------------------------------ #
    # Public API                                                           #
    # ------------------------------------------------------------------ #

    def force_refresh(self) -> None:
        """Reset cache so the file is re-read on the next get_financials call."""
        self._cache_time = 0.0
        self._cache_mtime = 0.0
        self._cache = {}
        self._load_error = ""

    @property
    def data_as_of(self) -> str:
        """Return the timestamp of the last successful file load, or empty string."""
        for snap in self._cache.values():
            return snap.last_refreshed or ""
        return ""

    def get_all_financials(self) -> list[FinancialSnapshot]:
        """Return all cached FinancialSnapshots, reading the file if needed."""
        self._refresh_if_needed()
        return list(self._cache.values())

    def get_financials(self, job_number: str) -> FinancialSnapshot:
        """Return a FinancialSnapshot for *job_number*, reading the file if needed."""
        key = _job_key(job_number)
        if key is None:
            return FinancialSnapshot.empty(job_number)

        self._refresh_if_needed()

        if self._load_error:
            snap = FinancialSnapshot.empty(job_number)
            snap.notes = [self._load_error]
            snap.touch()
            return snap

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

        if not _PYXLSB_AVAILABLE:
            self._load_error = (
                "Financial data is unavailable — please reinstall the app "
                "using ProjectTrackingToolSetup.exe to enable this feature."
            )
            return

        try:
            self._load(path)
            self._cache_time = now
            self._cache_mtime = mtime
        except Exception:
            logger.exception("Failed to read financial data file: %s", self._file_path)

    def _load(self, path: Path) -> None:

        refreshed = datetime.now().replace(microsecond=0).isoformat(sep=" ")
        new_cache: dict[str, FinancialSnapshot] = {}

        # Copy to a temp file so we can read it even while Excel has it open
        with tempfile.NamedTemporaryFile(suffix=path.suffix, delete=False) as tmp:
            tmp_path = tmp.name
        try:
            shutil.copy2(str(path), tmp_path)
            self._load_rows(Path(tmp_path), new_cache, refreshed)
        finally:
            Path(tmp_path).unlink(missing_ok=True)

        self._cache = new_cache
        logger.info("Loaded %d financial records from %s", len(new_cache), path)
        if self._snapshot_path:
            self._save_snapshot()

    def _load_rows(self, path: Path, new_cache: dict, refreshed: str) -> None:

        def _c(vals: list, idx: int):
            return vals[idx] if len(vals) > idx else None

        with _pyxlsb.open_workbook(str(path)) as wb:
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
                        job_name=_str(_c(vals, _COL_JOB_NAME)),
                        project_manager=_str(_c(vals, _COL_PM)),
                        sales_person=_str(_c(vals, _COL_SALES)),
                        status=_str(_c(vals, _COL_STATUS)),

                        contract_value=_flt(_c(vals, _COL_CONTRACT)),
                        billed_to_date=_flt(_c(vals, _COL_BILLED)),
                        amount_paid_to_date=_flt(_c(vals, _COL_PAID)),
                        estimated_cost=_flt(_c(vals, _COL_EAC)),
                        actual_cost=_flt(_c(vals, _COL_ACTUAL_COST)),

                        booked_margin=_flt(_c(vals, _COL_BOOKED_MARGIN)),
                        actual_margin=_flt(_c(vals, _COL_ACTUAL_MARGIN)),
                        differential_margin=_flt(_c(vals, _COL_DIFF_MARGIN)),

                        pm_hours_actual=_flt(_c(vals, _COL_PM_HOURS)),
                        tech_hours_actual=_flt(_c(vals, _COL_TECH_HOURS)),
                        pm_cost_actual=_flt(_c(vals, _COL_PM_COST)),
                        tech_cost_actual=_flt(_c(vals, _COL_TECH_COST)),

                        labor_rem_pct=_flt(_c(vals, _COL_LABOR_REM_PCT)),
                        material_rem_pct=_flt(_c(vals, _COL_MAT_REM_PCT)),
                        warranty_rem_pct=_flt(_c(vals, _COL_WARR_REM_PCT)),
                        travel_rem_pct=_flt(_c(vals, _COL_TRAV_REM_PCT)),
                        subcontract_rem_pct=_flt(_c(vals, _COL_SUB_REM_PCT)),
                        odc_rem_pct=_flt(_c(vals, _COL_ODC_REM_PCT)),

                        labor_rem_usd=_flt(_c(vals, _COL_LABOR_REM_USD)),
                        material_rem_usd=_flt(_c(vals, _COL_MAT_REM_USD)),
                        warranty_rem_usd=_flt(_c(vals, _COL_WARR_REM_USD)),
                        travel_rem_usd=_flt(_c(vals, _COL_TRAV_REM_USD)),
                        subcontract_rem_usd=_flt(_c(vals, _COL_SUB_REM_USD)),
                        odc_rem_usd=_flt(_c(vals, _COL_ODC_REM_USD)),

                        last_refreshed=refreshed,
                        notes=[],
                    )
                    new_cache[key] = snap

    @staticmethod
    def _detect_sheet(wb) -> str:
        """Find the first sheet whose header row contains a 'Job Number' cell."""
        for name in wb.sheets:
            try:
                with wb.get_sheet(name) as sheet:
                    for row_idx, row in enumerate(sheet.rows()):
                        if row_idx == _HEADER_ROW:
                            for cell in row:
                                if isinstance(cell.v, str) and "job number" in cell.v.lower():
                                    return name
                            break
            except Exception:
                continue
        return ""

    def _save_snapshot(self) -> None:
        """Write the current cache to a JSON snapshot file for other machines to read."""
        try:
            data = {
                "saved_at": datetime.now().replace(microsecond=0).isoformat(sep=" "),
                "records": {k: dataclasses.asdict(v) for k, v in self._cache.items()},
            }
            self._snapshot_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self._snapshot_path, "w", encoding="utf-8") as f:
                json.dump(data, f)
            logger.info("Financial snapshot saved to %s", self._snapshot_path)
        except Exception:
            logger.exception("Failed to save financial snapshot to %s", self._snapshot_path)


class SnapshotFinancialsProvider:
    """
    Read-only provider that loads financial data from a JSON snapshot file.
    Used on machines that don't have the live .xlsb file.
    """

    def __init__(self, snapshot_path: Path) -> None:
        self._snapshot_path = snapshot_path
        self._cache: dict[str, FinancialSnapshot] = {}
        self._loaded = False
        self._load()

    def _load(self) -> None:
        try:
            with open(self._snapshot_path, encoding="utf-8") as f:
                data = json.load(f)
            for key, record in data.get("records", {}).items():
                record.pop("job_id", None)  # drop fields not in current dataclass
                self._cache[key] = FinancialSnapshot(**record)
            logger.info("Loaded %d records from financial snapshot %s", len(self._cache), self._snapshot_path)
        except Exception:
            logger.exception("Failed to load financial snapshot from %s", self._snapshot_path)

    @property
    def data_as_of(self) -> str:
        for snap in self._cache.values():
            return f"{snap.last_refreshed} (snapshot)" if snap.last_refreshed else ""
        return ""

    def get_all_financials(self) -> list[FinancialSnapshot]:
        """Return all cached FinancialSnapshots."""
        return list(self._cache.values())

    def get_financials(self, job_number: str) -> FinancialSnapshot:
        key = _job_key(job_number)
        if key is None:
            return FinancialSnapshot.empty(job_number)
        snap = self._cache.get(key)
        if snap is None:
            snap = FinancialSnapshot.empty(job_number)
            snap.notes = [f"Job {job_number} not found in financial snapshot."]
        return snap

    def force_refresh(self) -> None:
        self._cache = {}
        self._load()
