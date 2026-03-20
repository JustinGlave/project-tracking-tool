from __future__ import annotations

import json
import logging
import tempfile
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

DEFAULT_TEMPLATE_NAME = "Phoenix Job Tracking"


@dataclass(slots=True)
class ProjectRecord:
    id: Optional[int] = None
    job_name: str = ""
    job_number: str = ""
    project_manager: str = ""
    sales_engineer: str = ""
    target_completion: Optional[str] = None
    liquid_damages: str = ""
    warranty_period: str = ""
    notes: str = ""
    created_at: Optional[str] = None
    updated_at: Optional[str] = None


@dataclass(slots=True)
class TaskRecord:
    id: Optional[int] = None
    project_id: Optional[int] = None
    task_name: str = ""
    phase: str = "General"
    sort_order: int = 0
    completed_date: Optional[str] = None
    is_complete: bool = False
    notes: str = ""


# Tuple of dicts — immutable at the container level, preventing accidental
# runtime mutation of the default task list.
DEFAULT_TASKS: tuple[dict[str, Any], ...] = (
    {"phase": "Pre-Project", "task_name": "Sales-Ops Turnover"},
    {"phase": "Pre-Project", "task_name": "Review of Specs/Add."},
    {"phase": "Pre-Project", "task_name": "Contract Review"},
    {"phase": "Planning", "task_name": "Job Plan Developed"},
    {"phase": "Materials", "task_name": "Phoenix Material Submittal"},
    {"phase": "Materials", "task_name": "Phoenix Material Submittal Approved"},
    {"phase": "Materials", "task_name": "Phoenix Material Delivery Confirmations"},
    {"phase": "Materials", "task_name": "Valves Ordered"},
    {"phase": "Shipping", "task_name": "Flow Curves Archived in Job Folder"},
    {"phase": "Shipping", "task_name": "Shipping Tracking to Customer"},
    {"phase": "Shipping", "task_name": "Shipping confirmed to Customer/ATS"},
    {"phase": "Engineering", "task_name": "Phoenix Drawing Package Submittal"},
    {"phase": "Engineering", "task_name": "PM Book, and Checkout Sheet Book Developed"},
    {"phase": "Engineering", "task_name": "Engineering Re-estimate"},
    {"phase": "Installation", "task_name": "Hire Elec. Sub Request Form"},
    {"phase": "Installation", "task_name": "Elec Install Standards reviewed with installer"},
    {"phase": "Controls", "task_name": "DDC Developed"},
    {"phase": "Controls", "task_name": "Graphics Developed"},
    {"phase": "Turnover", "task_name": "Service Turnover"},
    {"phase": "Commissioning", "task_name": "Submitted Commissioning Documents"},
    {"phase": "Commissioning", "task_name": "Checkout/PTP w/ Documentation"},
    {"phase": "Commissioning", "task_name": "Start-Up Complete All docs to CX and Ops Teams"},
    {"phase": "Commissioning", "task_name": "Commissioning Complete"},
    {"phase": "Closeout", "task_name": "All Punch Lists Complete"},
    {"phase": "Closeout", "task_name": "Close-out Documents Complete"},
    {"phase": "Closeout", "task_name": "Owner Training"},
    {"phase": "Archive", "task_name": "Job Back-up Archived"},
    {"phase": "Archive", "task_name": "As-Built Drawings Completed"},
    {"phase": "Archive", "task_name": "Archive Drawings (PDF)"},
    {"phase": "Archive", "task_name": "Scan Check-out Sheets to Job File"},
    {"phase": "Warranty", "task_name": "Warranty Letter"},
    {"phase": "Materials", "task_name": "Return Excess Materials"},
    {"phase": "Financial", "task_name": "Back-up Database Set-up Dial-up Database Tracking Form"},
    {"phase": "Financial", "task_name": "Resolve Trailing Costs"},
    {"phase": "Financial", "task_name": "Resolve All Change-orders"},
    {"phase": "Archive", "task_name": "Send record drawings to service Account Manager"},
    {"phase": "Financial", "task_name": "Final Billing and Payment"},
)


class ProjectTrackerBackend:
    """JSON-backed project tracker service.

    Storage notes
    -------------
    All data lives in a single JSON file. This is intentionally simple and
    portable for a single-user desktop tool. There is no file-level locking,
    so running two instances against the same file simultaneously can produce
    lost updates. If multi-process access ever becomes a requirement, replace
    the load/save calls with a proper database (SQLite with WAL mode is the
    natural next step).

    Workbook import notes
    ---------------------
    ``import_project_from_workbook`` reads header fields from hard-coded cell
    addresses (C3, H3, C4, …) that match the Phoenix Job Tracking template
    layout. If the template is ever restructured those addresses will need
    updating here.
    """

    def __init__(self, db_path: str | Path = "project_tracker_data.json") -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._initialize_storage()

    def _initialize_storage(self) -> None:
        if self.db_path.exists():
            data = self._load_data()
            changed = False
            if "projects" not in data:
                data["projects"] = []
                changed = True
            if "tasks" not in data:
                data["tasks"] = []
                changed = True
            if "next_project_id" not in data:
                data["next_project_id"] = self._next_id(data["projects"])
                changed = True
            if "next_task_id" not in data:
                data["next_task_id"] = self._next_id(data["tasks"])
                changed = True
            if changed:
                self._save_data(data)
            return

        self._save_data(
            {
                "projects": [],
                "tasks": [],
                "next_project_id": 1,
                "next_task_id": 1,
            }
        )

    def _load_data(self) -> dict[str, Any]:
        if not self.db_path.exists():
            return {
                "projects": [],
                "tasks": [],
                "next_project_id": 1,
                "next_task_id": 1,
            }
        return json.loads(self.db_path.read_text(encoding="utf-8"))

    def _save_data(self, data: dict[str, Any]) -> None:
        # Write to a temp file alongside the target, then rename — this ensures
        # atomic saves on all major platforms and prevents a half-written
        # JSON file if the process is interrupted mid-save.
        tmp_fd, tmp_path_str = tempfile.mkstemp(
            dir=self.db_path.parent, suffix=".tmp"
        )
        tmp_path = Path(tmp_path_str)
        try:
            with open(tmp_fd, "w", encoding="utf-8") as fh:
                json.dump(data, fh, indent=2)
            tmp_path.replace(self.db_path)
        except Exception:
            tmp_path.unlink(missing_ok=True)
            raise

    @staticmethod
    def _next_id(items: list[dict[str, Any]]) -> int:
        if not items:
            return 1
        return max(int(item.get("id", 0)) for item in items) + 1

    # ---------- project methods ----------

    def create_project(
        self,
        project: ProjectRecord,
        include_default_tasks: bool = True,
    ) -> int:
        data = self._load_data()
        now = self._now_iso()
        new_project_id = int(data["next_project_id"])

        for existing_project in data["projects"]:
            if (
                str(existing_project.get("job_number", "")).strip() == project.job_number.strip()
                and project.job_number.strip()
            ):
                raise ValueError(
                    f"Project with job number '{project.job_number}' already exists."
                )

        project_record = {
            "id": new_project_id,
            "job_name": project.job_name.strip(),
            "job_number": project.job_number.strip(),
            "project_manager": project.project_manager.strip(),
            "sales_engineer": project.sales_engineer.strip(),
            "target_completion": self._normalize_date(project.target_completion),
            "liquid_damages": project.liquid_damages.strip(),
            "warranty_period": project.warranty_period.strip(),
            "notes": project.notes.strip(),
            "created_at": now,
            "updated_at": now,
        }
        data["projects"].append(project_record)
        data["next_project_id"] = new_project_id + 1

        if include_default_tasks:
            self._insert_default_tasks(data, new_project_id)

        self._save_data(data)
        return new_project_id

    def update_project(self, project_id: int, **changes: Any) -> None:
        allowed_fields = {
            "job_name",
            "job_number",
            "project_manager",
            "sales_engineer",
            "target_completion",
            "liquid_damages",
            "warranty_period",
            "notes",
        }
        unknown_keys = set(changes) - allowed_fields
        if unknown_keys:
            logger.warning("update_project received unknown fields (ignored): %s", unknown_keys)

        updates = {key: value for key, value in changes.items() if key in allowed_fields}
        if not updates:
            return

        if "target_completion" in updates:
            updates["target_completion"] = self._normalize_date(updates["target_completion"])

        data = self._load_data()
        target_project = self._find_project_dict(data, project_id)
        if target_project is None:
            return

        new_job_number = str(updates.get("job_number", target_project["job_number"])).strip()
        if new_job_number:
            for existing_project in data["projects"]:
                if (
                    int(existing_project["id"]) != project_id
                    and str(existing_project.get("job_number", "")).strip() == new_job_number
                ):
                    raise ValueError(
                        f"Project with job number '{new_job_number}' already exists."
                    )

        for field_name, field_value in updates.items():
            target_project[field_name] = field_value.strip() if isinstance(field_value, str) else field_value

        target_project["updated_at"] = self._now_iso()
        self._save_data(data)

    def delete_project(self, project_id: int) -> None:
        data = self._load_data()
        data["projects"] = [item for item in data["projects"] if int(item["id"]) != project_id]
        data["tasks"] = [item for item in data["tasks"] if int(item["project_id"]) != project_id]
        self._save_data(data)

    def get_project(self, project_id: int) -> Optional[ProjectRecord]:
        data = self._load_data()
        target_project = self._find_project_dict(data, project_id)
        return self._project_from_dict(target_project) if target_project else None

    def list_projects(self, search_text: str = "") -> list[ProjectRecord]:
        data = self._load_data()
        search_value = search_text.strip().casefold()

        project_dicts = data["projects"]
        if search_value:
            project_dicts = [
                item
                for item in project_dicts
                if search_value in str(item.get("job_name", "")).casefold()
                or search_value in str(item.get("job_number", "")).casefold()
                or search_value in str(item.get("project_manager", "")).casefold()
                or search_value in str(item.get("sales_engineer", "")).casefold()
            ]

        project_dicts = sorted(
            project_dicts,
            key=lambda item: (
                str(item.get("updated_at", "")),
                str(item.get("job_name", "")).casefold(),
            ),
            reverse=True,
        )
        return [self._project_from_dict(item) for item in project_dicts]

    # ---------- task methods ----------

    def add_task(
        self,
        project_id: int,
        task_name: str,
        phase: str = "General",
        sort_order: Optional[int] = None,
        completed_date: Optional[str] = None,
        notes: str = "",
    ) -> int:
        cleaned_task_name = self._clean_text(task_name)
        cleaned_phase = self._clean_text(phase) or "General"

        data = self._load_data()
        target_project = self._find_project_dict(data, project_id)
        if target_project is None:
            raise ValueError(f"Project {project_id} was not found.")

        for existing_task in data["tasks"]:
            if (
                int(existing_task["project_id"]) == project_id
                and str(existing_task["task_name"]).casefold() == cleaned_task_name.casefold()
            ):
                raise ValueError(
                    f"Task '{cleaned_task_name}' already exists for this project."
                )

        new_sort_order = (
            sort_order if sort_order is not None
            else self._next_sort_order_from_data(data, project_id)
        )
        normalized_completed_date = self._normalize_date(completed_date)
        new_task_id = int(data["next_task_id"])

        task_record = {
            "id": new_task_id,
            "project_id": project_id,
            "task_name": cleaned_task_name,
            "phase": cleaned_phase,
            "sort_order": int(new_sort_order),
            "completed_date": normalized_completed_date,
            "is_complete": bool(normalized_completed_date),
            "notes": notes.strip(),
        }
        data["tasks"].append(task_record)
        data["next_task_id"] = new_task_id + 1
        target_project["updated_at"] = self._now_iso()
        self._save_data(data)
        return new_task_id

    def update_task(self, task_id: int, **changes: Any) -> None:
        allowed_fields = {
            "task_name",
            "phase",
            "sort_order",
            "completed_date",
            "is_complete",
            "notes",
        }
        unknown_keys = set(changes) - allowed_fields
        if unknown_keys:
            logger.warning("update_task received unknown fields (ignored): %s", unknown_keys)

        updates = {key: value for key, value in changes.items() if key in allowed_fields}
        if not updates:
            return

        if "completed_date" in updates:
            updates["completed_date"] = self._normalize_date(updates["completed_date"])
        if "completed_date" in updates and "is_complete" not in updates:
            updates["is_complete"] = bool(updates["completed_date"])
        if "task_name" in updates:
            updates["task_name"] = self._clean_text(str(updates["task_name"]))
        if "phase" in updates:
            updates["phase"] = self._clean_text(str(updates["phase"])) or "General"

        data = self._load_data()
        target_task = self._find_task_dict(data, task_id)
        if target_task is None:
            return

        updated_task_name = str(updates.get("task_name", target_task["task_name"])).strip()
        for existing_task in data["tasks"]:
            if (
                int(existing_task["id"]) != task_id
                and int(existing_task["project_id"]) == int(target_task["project_id"])
                and str(existing_task["task_name"]).casefold() == updated_task_name.casefold()
            ):
                raise ValueError(
                    f"Task '{updated_task_name}' already exists for this project."
                )

        for field_name, field_value in updates.items():
            target_task[field_name] = field_value

        owning_project = self._find_project_dict(data, int(target_task["project_id"]))
        if owning_project is not None:
            owning_project["updated_at"] = self._now_iso()

        self._save_data(data)

    def delete_task(self, task_id: int) -> None:
        data = self._load_data()
        target_task = self._find_task_dict(data, task_id)
        if target_task is None:
            return

        owning_project_id = int(target_task["project_id"])
        data["tasks"] = [item for item in data["tasks"] if int(item["id"]) != task_id]

        owning_project = self._find_project_dict(data, owning_project_id)
        if owning_project is not None:
            owning_project["updated_at"] = self._now_iso()

        self._save_data(data)

    def list_tasks(self, project_id: int, phase: Optional[str] = None) -> list[TaskRecord]:
        data = self._load_data()
        task_dicts = [item for item in data["tasks"] if int(item["project_id"]) == project_id]

        if phase:
            task_dicts = [item for item in task_dicts if str(item["phase"]) == phase]

        task_dicts = sorted(
            task_dicts,
            key=lambda item: (
                int(item.get("sort_order", 0)),
                str(item.get("task_name", "")).casefold(),
            ),
        )
        return [self._task_from_dict(item) for item in task_dicts]

    def set_task_completed(
        self,
        task_id: int,
        completed: bool,
        completed_date: Optional[str] = None,
    ) -> None:
        if completed and not completed_date:
            completed_date = date.today().isoformat()
        self.update_task(
            task_id,
            is_complete=bool(completed),
            completed_date=completed_date if completed else None,
        )

    def get_project_summary(self, project_id: int) -> dict[str, Any]:
        project_record = self.get_project(project_id)
        task_records = self.list_tasks(project_id)
        total_tasks = len(task_records)
        completed_tasks = sum(1 for item in task_records if item.is_complete)
        pending_tasks = total_tasks - completed_tasks
        progress_percent = (
            round((completed_tasks / total_tasks) * 100, 1) if total_tasks else 0.0
        )

        phase_breakdown: dict[str, dict[str, int]] = {}
        for task_record in task_records:
            bucket = phase_breakdown.setdefault(
                task_record.phase, {"total": 0, "completed": 0, "pending": 0}
            )
            bucket["total"] += 1
            if task_record.is_complete:
                bucket["completed"] += 1
            else:
                bucket["pending"] += 1

        return {
            "project": asdict(project_record) if project_record else None,
            "totals": {
                "tasks": total_tasks,
                "completed": completed_tasks,
                "pending": pending_tasks,
                "progress_percent": progress_percent,
            },
            "phase_breakdown": phase_breakdown,
        }

    # ---------- workbook import ----------

    def import_project_from_workbook(
        self,
        workbook_path: str | Path,
        sheet_name: Optional[str] = None,
        create_missing_tasks: bool = True,
    ) -> int:
        workbook_file = Path(workbook_path)
        workbook = load_workbook(workbook_file, data_only=True)
        sheet = workbook[sheet_name] if sheet_name else workbook[workbook.sheetnames[0]]

        job_number = self._value(sheet, "H3")
        if not job_number:
            logger.warning(
                "Imported workbook '%s' has no job number in cell H3. "
                "Duplicate detection will be skipped for this import.",
                workbook_file.name,
            )

        project = ProjectRecord(
            job_name=self._value(sheet, "C3"),
            job_number=job_number,
            project_manager=self._value(sheet, "C4"),
            sales_engineer=self._value(sheet, "H4"),
            target_completion=self._value(sheet, "E5"),
            liquid_damages=self._value(sheet, "E6"),
            warranty_period=self._value(sheet, "E7"),
            notes=f"Imported from workbook: {workbook_file.name}",
        )

        imported_project_id = self.create_project(project, include_default_tasks=True)
        existing_tasks = {
            task.task_name.casefold(): task
            for task in self.list_tasks(imported_project_id)
        }

        imported_task_items = self._extract_tasks_from_sheet(sheet)
        for imported_item in imported_task_items:
            task_key = imported_item["task_name"].casefold()
            existing_task = existing_tasks.get(task_key)
            if existing_task:
                self.update_task(
                    int(existing_task.id),
                    phase=imported_item["phase"],
                    sort_order=imported_item["sort_order"],
                    completed_date=imported_item["completed_date"],
                    is_complete=bool(imported_item["completed_date"]),
                )
            elif create_missing_tasks:
                self.add_task(
                    project_id=imported_project_id,
                    task_name=imported_item["task_name"],
                    phase=imported_item["phase"],
                    sort_order=imported_item["sort_order"],
                    completed_date=imported_item["completed_date"],
                )

        return imported_project_id

    def export_project_snapshot(self, project_id: int, export_path: str | Path) -> Path:
        output_file = Path(export_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "summary": self.get_project_summary(project_id),
            "tasks": [asdict(task) for task in self.list_tasks(project_id)],
        }
        output_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        return output_file

    # ---------- internal helpers ----------

    @staticmethod
    def _insert_default_tasks(data: dict[str, Any], project_id: int) -> None:
        for index, task_template in enumerate(DEFAULT_TASKS, start=1):
            new_task_id = int(data["next_task_id"])
            data["tasks"].append(
                {
                    "id": new_task_id,
                    "project_id": project_id,
                    "task_name": task_template["task_name"],
                    "phase": task_template["phase"],
                    "sort_order": index,
                    "completed_date": None,
                    "is_complete": False,
                    "notes": "",
                }
            )
            data["next_task_id"] = new_task_id + 1

    @staticmethod
    def _next_sort_order_from_data(data: dict[str, Any], project_id: int) -> int:
        matching_orders = [
            int(item.get("sort_order", 0))
            for item in data["tasks"]
            if int(item["project_id"]) == project_id
        ]
        return (max(matching_orders) + 1) if matching_orders else 1

    @staticmethod
    def _find_project_dict(
        data: dict[str, Any], project_id: int
    ) -> Optional[dict[str, Any]]:
        for project_dict in data["projects"]:
            if int(project_dict["id"]) == project_id:
                return project_dict
        return None

    @staticmethod
    def _find_task_dict(
        data: dict[str, Any], task_id: int
    ) -> Optional[dict[str, Any]]:
        for task_dict in data["tasks"]:
            if int(task_dict["id"]) == task_id:
                return task_dict
        return None

    def _extract_tasks_from_sheet(self, sheet: Any) -> list[dict[str, Any]]:
        left_rows = [10, 12, 14, 17, 20, 22, 24, 27, 29, 31, 33, 36, 38, 40, 42, 44, 51, 54]
        right_rows = [10, 12, 14, 17, 20, 22, 24, 27, 31, 33, 36, 38, 40, 42, 44, 47, 49, 51, 54]

        extracted_tasks: list[dict[str, Any]] = []
        for order_value, row_idx in enumerate(left_rows, start=1):
            task_name = self._clean_text(self._value(sheet, f"B{row_idx}"))
            if task_name:
                extracted_tasks.append(
                    {
                        "task_name": task_name,
                        "phase": self._infer_phase(task_name),
                        "sort_order": order_value,
                        "completed_date": self._value(sheet, f"D{row_idx}"),
                    }
                )

        start_order = len(extracted_tasks) + 1
        for offset, row_idx in enumerate(right_rows, start=0):
            task_name = self._clean_text(self._value(sheet, f"F{row_idx}"))
            if task_name:
                extracted_tasks.append(
                    {
                        "task_name": task_name,
                        "phase": self._infer_phase(task_name),
                        "sort_order": start_order + offset,
                        "completed_date": self._value(sheet, f"H{row_idx}"),
                    }
                )
        return extracted_tasks

    @staticmethod
    def _infer_phase(task_name: str) -> str:
        lowered = task_name.casefold()
        rules = {
            "materials": ["material", "valves", "shipping"],
            "engineering": ["drawing", "ddc", "graphics", "re-estimate", "flow curves"],
            "installation": ["elec", "install", "sub request"],
            "commissioning": ["commission", "checkout", "start-up", "service turnover"],
            "closeout": ["punch", "close-out", "owner training", "warranty"],
            "archive": ["archive", "as-built", "scan", "record drawings", "job back-up"],
            "financial": ["billing", "payment", "change-orders", "trailing costs", "database"],
            "planning": ["turnover", "review", "contract", "job plan"],
        }
        for phase_name, keywords in rules.items():
            if any(keyword in lowered for keyword in keywords):
                return phase_name.title()
        return "General"

    @staticmethod
    def _project_from_dict(project_dict: dict[str, Any]) -> ProjectRecord:
        return ProjectRecord(
            id=project_dict["id"],
            job_name=project_dict["job_name"],
            job_number=project_dict["job_number"],
            project_manager=project_dict["project_manager"],
            sales_engineer=project_dict["sales_engineer"],
            target_completion=project_dict["target_completion"],
            liquid_damages=project_dict["liquid_damages"],
            warranty_period=project_dict["warranty_period"],
            notes=project_dict["notes"],
            created_at=project_dict["created_at"],
            updated_at=project_dict["updated_at"],
        )

    @staticmethod
    def _task_from_dict(task_dict: dict[str, Any]) -> TaskRecord:
        return TaskRecord(
            id=task_dict["id"],
            project_id=task_dict["project_id"],
            task_name=task_dict["task_name"],
            phase=task_dict["phase"],
            sort_order=task_dict["sort_order"],
            completed_date=task_dict["completed_date"],
            is_complete=bool(task_dict["is_complete"]),
            notes=task_dict["notes"],
        )

    @staticmethod
    def _clean_text(value: Any) -> str:
        if value is None:
            return ""
        return " ".join(str(value).replace("\n", " ").split()).strip()

    @staticmethod
    def _value(sheet: Any, cell_ref: str) -> str:
        value = sheet[cell_ref].value
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()
        return ProjectTrackerBackend._clean_text(value)

    @staticmethod
    def _normalize_date(value: Any) -> Optional[str]:
        if value is None:
            return None
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()

        text = str(value).strip()
        if not text:
            return None

        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d"):
            try:
                return datetime.strptime(text, fmt).date().isoformat()
            except ValueError:
                continue

        # No known format matched — raise rather than silently passing through
        # a raw string that could corrupt date comparisons downstream.
        raise ValueError(
            f"Unrecognized date format: '{text}'. "
            "Expected one of: YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY, YYYY/MM/DD."
        )

    @staticmethod
    def _now_iso() -> str:
        return datetime.now().replace(microsecond=0).isoformat(sep=" ")


if __name__ == "__main__":
    # Demo: creates a temporary file so it never pollutes the working directory.
    # json and tempfile are already imported at module level — no re-import needed.
    with tempfile.NamedTemporaryFile(
        suffix=".json", delete=False, mode="w"
    ) as _tmp:
        demo_path = _tmp.name

    backend = ProjectTrackerBackend(demo_path)
    demo_project_id = backend.create_project(
        ProjectRecord(
            job_name="Phoenix Demo Job",
            job_number="PHX-001",
            project_manager="Justin",
            sales_engineer="ATS",
            target_completion="2026-06-30",
        )
    )
    print(json.dumps(backend.get_project_summary(demo_project_id), indent=2))
    print(f"\nTemp data written to: {demo_path}")
