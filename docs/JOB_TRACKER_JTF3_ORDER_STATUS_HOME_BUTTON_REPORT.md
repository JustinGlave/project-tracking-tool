# JTF-3 — Home Order-Status Button (Implementation Report)

Implements the single-button addition to the existing Home dashboard, per the operator-corrected JTF-3 scope (no new Home page, no Home redesign).

- **Branch:** `feature/job-tracker-rss-webpro-orders` (JTF-1 + JTF-2 pushed; JTF-3 commits are local)
- **Base:** `main @ 8a85aae` (v1.8.6)
- **Commit added by JTF-3:**
  - `b9716d3` — feat(jtf-3): Home order-status button + OrderStatusDialog modal
  - (this report will be the next commit)
- **Test result:** 54 / 54 pass (45 prior + 9 new JTF-3)

---

## 1. Files changed

| File | Change |
|---|---|
| `project_tracker_backend.py` | New module constant `ORDER_SIGNAL_TASKS = frozenset({"Valves Ordered"})`; new helper `_project_has_order_signal(project_id, tasks)`; new `order_status: Optional[str] = None` parameter on `list_projects` (`"ordered"` / `"missing"` / `None`); new `get_order_status_rollup() -> dict`. |
| `project_tracker_gui.py` | New `OrderStatusDialog` class (modal, tabs + tables, double-click drill-down); one added button on the existing Home dashboard (`self._dash_order_status_btn`) with text refreshed in `_refresh_dashboard`; new `_open_order_status_dialog` + `_select_project_by_id` handlers; added `QTabWidget` import. |
| `tests/test_regressions.py` | New `JTF3OrderStatusTests` class with 9 tests. |
| `docs/JOB_TRACKER_JTF3_ORDER_STATUS_HOME_BUTTON_REPORT.md` | This report. |

No other files modified. JTF-1 / JTF-2 changes are untouched.

---

## 2. Order-status rule

```python
# project_tracker_backend.py
ORDER_SIGNAL_TASKS: frozenset[str] = frozenset({"Valves Ordered"})

def _project_has_order_signal(project_id: int, tasks: list[dict]) -> bool:
    targets = {name.casefold() for name in ORDER_SIGNAL_TASKS}
    for task in tasks:
        if int(task.get("project_id", 0)) != project_id:
            continue
        if str(task.get("task_name", "")).casefold() not in targets:
            continue
        if bool(task.get("is_complete", False)):
            return True
    return False
```

- **A project is "ordered" iff** at least one of its tasks named in `ORDER_SIGNAL_TASKS` is marked `is_complete = True`.
- **Otherwise "missing"** — covers both incomplete signal task and *no* signal task at all (e.g., a project that had its "Valves Ordered" task deleted).
- Case-insensitive task-name match so a re-cased copy (`"valves ordered"`) still counts.
- No new schema field; no financial inference. Reads only `is_complete` on existing `TaskRecord` rows.

Widening the rule later is one-line: add another task name to the `frozenset`.

---

## 3. Backend filter behavior

`ProjectTrackerBackend.list_projects` gained one optional parameter:

```python
def list_projects(
    self,
    search_text: str = "",
    include_test: bool = True,
    sort_by: str = "updated",
    sort_asc: bool = False,
    has_rss: Optional[bool] = None,
    order_status: Optional[str] = None,   # ← JTF-3
) -> list[ProjectRecord]: ...
```

| `order_status` | Result |
|---|---|
| `None` | No order filter applied (legacy default — preserves callers that don't opt in). |
| `"ordered"` | Returns only projects where `_project_has_order_signal` is True. |
| `"missing"` | Returns only projects where `_project_has_order_signal` is False. |
| anything else | Raises `ValueError("order_status must be 'ordered', 'missing', or None …")`. |

The filter runs **after** `include_test`, `search_text`, and `has_rss`, and **before** sort — composing cleanly with all existing filters (verified by `test_order_status_composes_with_rss_filter` and `test_order_status_composes_with_text_search`).

### Rollup

`ProjectTrackerBackend.get_order_status_rollup() -> dict` does one `_load_data` pass and returns:

```python
{
    "ordered_count": int,
    "missing_count": int,
    "ordered": [{"id", "job_name", "job_number", "project_manager", "updated_at"}, ...],
    "missing": [{"id", "job_name", "job_number", "project_manager", "updated_at"}, ...],
}
```

Rows in each list are sorted by `job_number.casefold()` so the modal renders deterministically. Test jobs (`is_test=True`) are excluded.

---

## 4. Home button behavior

The Home dashboard layout is **unchanged** apart from one new row added between the stat cards and the lists row:

```
┌─ Home ───────────────────────────────────────────────────────────────┐
│  Welcome to ATS Job Tracker                                          │
│  Select a project from the sidebar…                                  │
├──────────────────────────────────────────────────────────────────────┤
│  [ Projects ]   [ Incomplete Tasks ]   [ Total Tasks ]               │  ← unchanged
│  ┌────────────────────────────────────────────────────────────────┐  │
│  │ Valve/Parts Order Status — 12 missing / 3 ordered              │  │  ← NEW (one button)
│  └────────────────────────────────────────────────────────────────┘  │
│  Top 5 by Contract Value         5 Most Recently Added               │  ← unchanged
│  (table)                          (table)                            │
│  Recent Activity                                                      │  ← unchanged
│  (table)                                                              │
└──────────────────────────────────────────────────────────────────────┘
```

- Single `QPushButton`, left-aligned with a stretch on the right so it doesn't span the full width.
- Static initial label `"Valve/Parts Order Status"`; live label set by `_refresh_dashboard` after backend rollup: `"Valve/Parts Order Status — {missing} missing / {ordered} ordered"`.
- Tooltip: *"View jobs grouped by whether their "Valves Ordered" task is complete."*
- Click handler: `_open_order_status_dialog`.
- If the backend rollup raises, the button label falls back to the static text and a `QMessageBox.critical` surfaces the error on click.

No StatCards were modified. No tables were modified. No removed widgets.

---

## 5. Modal behavior — `OrderStatusDialog`

Modal opened by the Home button:

```
┌─ Valve / Parts Order Status ─────────────────────────────────────────┐
│  Jobs by parts/valves order status                                   │
│  Based on the existing "Valves Ordered" task.                        │
│  Double-click a row to open that project.                            │
│                                                                       │
│  ┌─ Missing Orders (12) ─┬─ Ordered (3) ──────────────────────────┐  │
│  │ Job #   Job Name                 PM           Updated          │  │
│  │ A-1     Alpha Project            Justin G.    2026-05-30 14:22 │  │
│  │ B-2     Beta Project             Lisa Park    2026-05-30 09:14 │  │
│  │ …                                                              │  │
│  └────────────────────────────────────────────────────────────────┘  │
│                                                       [   Close   ]  │
└──────────────────────────────────────────────────────────────────────┘
```

Behavior:
- `QTabWidget` with two tabs: **Missing Orders (N)** and **Ordered (N)**. Counts in the tab labels.
- Each tab is a `QTableWidget` with columns `Job #` | `Job Name` | `PM` | `Updated` (4 cols).
  - Job # column is `ResizeToContents`; Job Name stretches; PM and Updated are `ResizeToContents`.
  - Updated column displays the first 16 chars of `updated_at` with `T` replaced by space (e.g. `2026-05-30 14:22`).
  - Alternating row colors. Rows non-editable, row-selection mode.
- **Double-click a row** → `OrderStatusDialog.project_selected.emit(project_id)` + `self.accept()` (closes modal).
- **MainWindow** wires `project_selected → _select_project_by_id`, which sets the corresponding sidebar `QListWidget` row as current (firing the normal `on_project_selected` chain). Same path as a user clicking the project in the sidebar.
- **Close** button at the bottom-right; `setAutoDefault(False)` so Enter doesn't fire it accidentally.

Per spec, the optional project-row indicator was **not** added. Rows already carry `📌` (pinned) and `📎` (RSS, JTF-1); adding a third glyph would clutter them.

---

## 6. Tests & validation

### New regression tests (9, all in `JTF3OrderStatusTests`)

| Test | Asserts |
|---|---|
| `test_order_status_ordered_returns_only_projects_with_completed_valves_task` | Three-project fixture; `order_status="ordered"` → `["Alpha"]`. |
| `test_order_status_missing_returns_incomplete_or_absent_signal` | Same fixture; `order_status="missing"` → `["Beta", "Gamma"]` (B has incomplete task; G has no task). |
| `test_order_status_none_returns_all_projects` | Same fixture; `order_status=None` → all three. |
| `test_order_status_invalid_value_raises` | `order_status="maybe"` → `ValueError`. |
| `test_order_status_composes_with_rss_filter` | Beta also gets RSS; `has_rss=True, order_status="missing"` → `["Beta"]`. |
| `test_order_status_composes_with_text_search` | `search="Alpha", order_status="ordered"` → `["Alpha"]`; `"Alpha"+"missing"` → `[]`. |
| `test_rollup_counts_match_filter_results` | Rollup `ordered_count=1`, `missing_count=2`; lists match filter output. |
| `test_rollup_excludes_test_jobs` | Adding `is_test=True` project doesn't change rollup counts; name absent from rows. |
| `test_rollup_rows_carry_minimum_columns_needed_by_modal` | Every row has `id`, `job_name`, `job_number`, `project_manager`, `updated_at`. |

### Validation run

| Step | Result |
|---|---|
| `py_compile project_tracker_backend.py project_tracker_gui.py tests/test_regressions.py` | clean |
| Full suite from main repo CWD | **54 / 54 pass** (45 prior + 9 new JTF-3) |
| Import smoke | `ORDER_SIGNAL_TASKS` is `frozenset({"Valves Ordered"})`; `order_status` confirmed in `list_projects.__code__.co_varnames`; `get_order_status_rollup` returns the expected key set. |

Interactive GUI smoke not run by this report (avoids popping a window). To verify visually: launch `python project_tracker_gui.py` from the main repo, return to the Home view (Home button in sidebar or close any open project), confirm the new button appears with live counts, click it, switch tabs, double-click a row, verify the project opens in the sidebar.

---

## 7. Next step recommendation

The feature branch now carries all three operator-approved JTF items (JTF-1 RSS filter, JTF-2 multi-WebPro, JTF-3 order-status button) on top of v1.8.6 main. Suggested next steps, in order:

1. **Operator smoke test**: launch source-mode and verify the three new affordances end-to-end (RSS filter dropdown, WebPro multi-ID editor, order-status Home button + modal).
2. **Push `feature/job-tracker-rss-webpro-orders`** to refresh the remote with the JTF-2 + JTF-3 commits (currently only JTF-1 was pushed when this branch was set up).
3. **Release prep** as a single v1.8.7 (PATCH per Job Tracker semver convention since these are additive UX rather than a redesign): bump `version.py`, add a "What's New in v1.8.7" README section, run build, release. None of these have been touched per the JTF-3 spec.

No follow-up coding required for JTF unless the operator wants to:
- Widen `ORDER_SIGNAL_TASKS` (e.g., add `"Phoenix Material Submittal Approved"`).
- Surface the JTF-3 modal from elsewhere (a sidebar shortcut, a menu entry).
- Add a sidebar filter dropdown for order status (mirrors the JTF-1 RSS pattern).

---

## 8. Constraints confirmation

- ✅ **No new Home page was created.** The Home dashboard is the same `QStackedWidget` index 0 view built by `_build_dashboard`; one button row was inserted between existing rows.
- ✅ **No Home redesign occurred.** Stat cards row, lists row, and activity table are byte-identical to prior code. One new button row; no widget removed; no font / spacing / color changes outside the new button.
- ✅ **No schema migration occurred.** No new field on `ProjectRecord` or `TaskRecord`. `ORDER_SIGNAL_TASKS` is a module-level Python constant.
- ✅ **No financials / auth changes.** `financials_*.py` and `user_auth.py` untouched.
- ✅ **No updater / build / installer changes.** `updater.py`, `build.bat`, `installer.iss` untouched.
- ✅ **No `version.py` change.** Still `1.8.6` on the branch tip.
- ✅ **No release / tag / publish occurred.** JTF-3 commits live on the local feature branch only; nothing pushed for JTF-3; no tag created; no GitHub release.
- ✅ **JTF-1 RSS behavior preserved.** Four JTF-1 tests still pass; the RSS filter dropdown and 📎 indicator are untouched.
- ✅ **JTF-2 WebPro behavior preserved.** Twelve JTF-2 tests still pass; `WebProIdsDialog` and ProjectDialog integration are untouched.
