# Job Tracker — RSS / WebPro / Order-Status Feature Audit

Planning document only. No source code, schema, version, build, or release files have been modified. No release published.

Auditor pass executed against `main @ 8a85aae` (v1.8.6, commons retrofit). Worktree `claude/busy-kowalevski-154a16` is 13 commits behind origin/main; main repo was used as the source of truth for this audit.

---

## 1. Repo state

| Item | Value |
|---|---|
| Main-repo branch | `main` |
| Main-repo HEAD | `8a85aae` — "gitignore: ignore .venv/ + .venv*/ + venv variants (post-release housekeeping)" |
| Main-repo origin sync | 0 / 0 (clean, up to date) |
| Main-repo working tree | clean |
| `version.py` | `__version__ = "1.8.6"` |
| Latest release tag | `v1.8.6` (also `v1.8.6-rc1`, `job-tracker-retrofit-v1.8.5-pre`) |
| v1.8.6 nature | Wave 8b commons retrofit + release hardening + starter_package removal — **no functional changes** per CHANGELOG |
| Project data location | `%APPDATA%\ATS Inc\Project Tracking Tool\project_tracker_data.json` (default), or a shared-folder override via **File → Data Location…** |
| Worktree branch (this session) | `claude/busy-kowalevski-154a16` — 13 commits behind origin/main; **stale relative to v1.8.6**; not used for audit findings below |

Because v1.8.6 introduced no functional changes, every finding below applies equally to v1.8.5 and v1.8.6 code. No feature branch has been created (per spec).

---

## 2. RSS — current state

### Storage shape

- `ProjectRecord.rss_files: list` ([project_tracker_backend.py:48](project_tracker_backend.py))
- Stored as `"rss_files"` in the per-project JSON record ([backend:434](project_tracker_backend.py)).
- Legacy migration from a single-string `csv_file_path` is already in place (`_migrate_rss_files`, [backend:205-213](project_tracker_backend.py)).
- A project may already hold **zero, one, or many** RSS feeds. The data model is list-native.
- Each entry is a dict; observed keys include `name`, `path`, and the parsed rows. Specific dict shape is owned by `RSSViewDialog` / `_attach_rss` in `project_tracker_gui.py`.

### Attach flow

- Notes window → **Attach RSS** button (`attach_btn` at [gui:904](project_tracker_gui.py)) → `_attach_rss` ([gui:1025](project_tracker_gui.py)).
- File dialog accepts `*.csv *.xlsx *.xlsm`.
- Preview shown, table name prompted, confirmed.
- If existing feeds: **Replace All** vs **Add New** choice ([gui:1040-1064](project_tracker_gui.py)).

### View flow

- Notes window → **RSS** button → `_open_rss` (around [gui:1069](project_tracker_gui.py)).
- 0 feeds: no-op (button effectively idle for empty projects).
- 1 feed: opens `RSSViewDialog` directly.
- N feeds: opens `RSSSelectDialog` picker ([gui:791-825](project_tracker_gui.py)) → user picks → `RSSViewDialog` opens that entry.

### Visibility on project list / cards

- **None.** `refresh_project_list` ([gui:4759-4826](project_tracker_gui.py)) renders only: job number, pinned-prefixed name, and an optional ODIN-financials line.
- No RSS-presence indicator (icon, color, or text) on any project row today.

### Search / filter integration

- **None.** `ProjectTrackerBackend.list_projects` ([backend:543-565](project_tracker_backend.py)) only matches text against `job_name`, `job_number`, `project_manager`, `sales_engineer`.
- No `has_rss` parameter, no filter dropdown in the sidebar, no Home-page surface for RSS presence/absence.

### Minimal path for filter WITH-RSS / WITHOUT-RSS

**Recommendation: filter dropdown in the sidebar + a small row indicator. Not a Home-page surface for v1.**

Reasoning: the filter is sidebar-shaped (it modifies the project list, just like Sort By already does). A Home-page button would imply a separate view; the sidebar dropdown matches existing patterns and is one click cheaper.

Minimum v1 = three changes:

1. New optional parameter on `list_projects(has_rss: Optional[bool] = None)`. None = no filter; True = `len(rss_files) > 0`; False = empty list.
2. New small `QComboBox` in the sidebar next to the existing Sort By: **All / Has RSS / No RSS**. Wire `currentIndexChanged → refresh_project_list`.
3. Tiny visual indicator on each project row — e.g. append a `📎` glyph after the job name when `len(rss_files) > 0`. Makes the filter discoverable; costs ~3 lines in `refresh_project_list`.

Search-box enhancement is **not recommended** — search is a freetext field; presence is a boolean facet. Mixing them confuses users.

---

## 3. WebPro ID — current state

### Storage shape

- `ProjectRecord.webpro_id: str = ""` ([backend:41](project_tracker_backend.py))
- Stored as `"webpro_id"` in JSON ([backend:431](project_tracker_backend.py)).
- Single string. No list, no normalization beyond `.strip()`.
- `update_project` only accepts `"webpro_id"` as an allowed key ([backend:471](project_tracker_backend.py)).
- `_project_from_dict` reads `project_dict.get("webpro_id", "")` ([backend:2073](project_tracker_backend.py)) — tolerant of missing field.

### UI

- Single button in the project header — `self.webpro_id_btn` (object name `WebProIdBtn`) at [gui:3087-3102](project_tracker_gui.py).
- Fixed width 110px, min height 42px. Shows current value or "—".
- Edit flow: `_edit_webpro_id` ([gui:3321-3333](project_tracker_gui.py)) uses a single-line `QInputDialog.getText()` for one value; saves via `update_project(..., webpro_id=text.strip())`.
- Header repopulation sets text in `load_current_project` ([gui:4895](project_tracker_gui.py)) and clears it in `clear_project_display` ([gui:4936](project_tracker_gui.py)).
- Disabled for view-only users.

### Search / filter / exports

- **Not in search.** `list_projects` does not match against WebPro ID.
- **Not in Excel export.** The `info_fields` list in `_write_project_sheets` ([backend:1426-1446](project_tracker_backend.py)) covers job name/number/PM/SE/contract/owner/contractor/div25 etc. but **does not include WebPro ID**. Other exports (`export_project_snapshot`) write `asdict()` of the record, which would carry whatever field shape we keep.
- **Not in financials.** Grepped `financials_*.py` — zero hits for `webpro`.

This is great news for migration: a single-string → list-of-strings change touches only `project_tracker_gui.py` + `project_tracker_backend.py`. Financials, exports, and other tools are unaffected.

### Plan for multiple WebPro IDs (backwards-compatible)

See §6 for full migration mechanics. Short version:

- Keep `"webpro_id": str` as the canonical legacy field. Add `"webpro_ids": list[str]` as the new field.
- `_project_from_dict` reads BOTH and reconciles: if `webpro_ids` present, use it; else if `webpro_id` present, wrap in a single-item list.
- Backend write path always writes both keys — `webpro_id` as the first ID (or `""` if empty) + `webpro_ids` as the full list. This way any older app reading the file still sees a sensible single value.
- New editor: a small dialog with add/remove/edit rows instead of single-line input.
- Header button display: when 1 ID, show as today; when N IDs, show e.g. `"12345 +2"` with full list in tooltip; clicking the button opens the editor either way.

---

## 4. Valves / Parts ordered — current state

### What already exists

The task templates carry **explicit** order-tracking tasks:

```
DEFAULT_TASKS (and PHOENIX_TASKS, same entries) at project_tracker_backend.py:147-185
  ...
  {"phase": "Materials", "task_name": "Phoenix Material Submittal"},
  {"phase": "Materials", "task_name": "Phoenix Material Submittal Approved"},
  {"phase": "Materials", "task_name": "Phoenix Material Delivery Confirmations"},
  {"phase": "Materials", "task_name": "Valves Ordered"},
  ...
  {"phase": "Materials", "task_name": "Return Excess Materials"},
```

Each task is a `TaskRecord` with `is_complete: bool` and `completed_date: Optional[str]`. The "Valves Ordered" task being marked complete IS the canonical "valves have been ordered" signal in today's data.

### What does NOT exist

- No explicit `valves_ordered: bool` field on `ProjectRecord`.
- No "parts ordered" concept distinct from the Materials-phase tasks.
- No purchase-order entity; `ChangeOrderRecord` is for COs (change orders / pricing) not POs (purchase orders).
- The financials provider tracks material-remaining percentages (`material_rem_pct`, `material_rem_usd`) but those are dollar-budget indicators, not ordered/not-ordered booleans.

### Per-spec classification

| Class | Verdict | Reason |
|---|---|---|
| A. Already structured and searchable | Partial | "Valves Ordered" task exists per project; not currently queryable by name |
| **B. Derivable from existing tasks/statuses** | **Yes** | `list_tasks(project_id)` + look up `task_name == "Valves Ordered"` + `is_complete` |
| C. Requires new explicit project fields | No (not for v1) | Adding a redundant project-level flag would drift from the task source of truth |
| D. Ambiguous / operator decision required | **Yes (one decision)** | The phrase "valves OR parts ordered" — see §9 |

### How "has parts ordered" can be derived

The most surgical definition for v1:

```
def project_has_valves_ordered(backend, project_id) -> bool:
    return any(
        t.task_name == "Valves Ordered" and t.is_complete
        for t in backend.list_tasks(project_id)
    )
```

If the operator wants a broader "materials in motion" signal, it could be widened:

```
ORDER_SIGNAL_TASKS = {"Valves Ordered", "Phoenix Material Submittal Approved"}
def project_has_orders_placed(...) -> bool:
    return any(
        t.task_name in ORDER_SIGNAL_TASKS and t.is_complete
        for t in backend.list_tasks(project_id)
    )
```

Both are pure read-side derivations — no schema change needed. The operator decision is just **which task names** count.

A small performance note: today's `list_projects` does not pull tasks. A Home-screen rollup "N jobs missing orders" needs a single backend pass that joins projects to their tasks. Easiest is one new method `get_order_status_rollup() -> dict` that loads data once and tallies — same shape as `get_dashboard_stats`.

---

## 5. Home dashboard — current composition

`_build_dashboard` at [project_tracker_gui.py:2888-2978](project_tracker_gui.py) is the Home view. Today's layout:

```
┌─────────────────────────────────────────────────────┐
│  Welcome to ATS Job Tracker                         │
│  Select a project from the sidebar…                 │
├─────────────────────────────────────────────────────┤
│  [ Projects ]   [ Incomplete Tasks ]  [ Total ]    │  ← StatCards row, NOT clickable
├──────────────────────────┬──────────────────────────┤
│  Top 5 by Contract Value │  5 Most Recently Added   │
├──────────────────────────┴──────────────────────────┤
│  Recent Activity                                    │
│  (scrolling table, last 20)                         │
└─────────────────────────────────────────────────────┘
```

- StatCards are display-only; no `clicked` signal currently.
- `_refresh_dashboard` ([gui:2980+](project_tracker_gui.py)) repopulates everything from `backend.get_dashboard_stats()`.

### Recommended order-status surface

**A 4th StatCard, made clickable, showing the count of jobs missing orders.**

Mockup:

```
[ Projects ]  [ Incomplete Tasks ]  [ Total Tasks ]  [ ⚠️ Missing Parts Orders ]
                                                       ────────clickable────────
```

- Label: **"Missing Parts Orders"** (or **"Valves Not Ordered"** if operator chooses the narrow definition — see §9).
- Value: integer count of jobs where the chosen ORDER_SIGNAL task(s) are incomplete.
- Click handler: open a small modal with a sortable table of `Job # | Job Name | PM | Days Since Created` for the missing-orders subset. Double-click a row to jump to that project (same selection signal as the sidebar).

Why this and not Option B/C:

- Same `StatCard` widget already exists and matches the visual rhythm.
- Single row, no dashboard re-layout, no new sections.
- Click-to-drill-down preserves "Home is a summary, project view is the detail" mental model.
- If the count is zero the card just shows `0` — no special empty state needed.

Alternative ("ordered" view) is largely useless on its own — completed work doesn't need surfacing. Stick with the "missing" framing.

---

## 6. Data migration plan

### RSS

No data-shape change. Already list-native. **No migration required.** Filter/indicator additions are read-side only.

### WebPro IDs

**Old shape (current):**
```json
{
  "id": 42,
  "job_name": "...",
  "webpro_id": "12345"
}
```

**New shape (proposed):**
```json
{
  "id": 42,
  "job_name": "...",
  "webpro_id": "12345",          // ← retained as legacy/compat field = first item
  "webpro_ids": ["12345", "67890"]  // ← new canonical field
}
```

**Migration behavior:**

- On read, in `_project_from_dict`:
  - If `webpro_ids` exists and is a non-empty list → use it.
  - Else if `webpro_id` exists and is non-empty → return `[webpro_id]`.
  - Else return `[]`.
- On write, in `create_project`/`update_project`:
  - Normalize incoming list: `.strip()` each entry, drop empties, dedupe case-insensitively while preserving the first-seen casing, preserve order.
  - Store `webpro_ids = normalized_list` AND `webpro_id = normalized_list[0] if normalized_list else ""`.
  - Dual-write keeps older readers (the next-to-ship-old-version case) functional indefinitely.

**Backup behavior:** Existing auto-backup-on-open (`_backup_data_file`, gui:43-78) runs unchanged. First launch after the upgrade creates a backup of the pre-migration file before the first save lands the new field. No new backup logic required.

**Test cases** (see §8 for full list):
- Existing single-ID record loads as a 1-element list.
- Missing both fields loads as empty list (no crash — H3 from earlier audit already hardens this path).
- Dual-write round-trips: write `["abc-1", "ABC-1"]` → read returns `["abc-1"]` (deduped); JSON contains both `webpro_id="abc-1"` and `webpro_ids=["abc-1"]`.
- Old app version reading the new file sees `webpro_id="abc-1"` and continues to work.
- Whitespace-only entries dropped.

**Rollback safety:** Because `webpro_id` is still written, an emergency rollback to v1.8.6 leaves every project showing the first ID. No data loss; users only lose access to the additional IDs until they re-upgrade. The additional IDs remain in JSON untouched (just unused) by the older app.

**Preservation of existing values:** Guaranteed by the read-path fallback and the dual-write. Every existing single `webpro_id` will load as `[webpro_id]` and be re-written as both fields on the next save.

### Order status

No data-shape change. Derived from existing tasks. **No migration required.**

---

## 7. Implementation sequence

Each phase is a separate review-able commit set. Phases are sequenced by independence, not dependency.

### JTF-1 — RSS filter + indicator (sidebar)

**Touches:** `project_tracker_backend.py`, `project_tracker_gui.py`, `tests/test_regressions.py`.

- Backend: add `has_rss: Optional[bool] = None` to `list_projects()`. New filter applied after text-search filter, before sort.
- GUI: add `QComboBox` to the sidebar between search and sort. Three options: All / Has RSS / No RSS. Default All. Wire to `refresh_project_list`. Append `📎` to project name in `refresh_project_list` when `len(project.rss_files) > 0`.
- Tests: see §8.

**Effort:** S. **Risk:** very low. No schema change.

### JTF-2 — Multiple WebPro IDs (dual-write migration)

**Touches:** `project_tracker_backend.py`, `project_tracker_gui.py`, `tests/test_regressions.py`.

- Backend: add `webpro_ids: list[str]` to `ProjectRecord` dataclass with `field(default_factory=list)`. Add to allowed-fields in `update_project`. Add to `create_project` write path. Implement `_normalize_webpro_ids(values: list[str]) -> list[str]` (strip, drop empties, casefold-dedupe preserving order). Read path in `_project_from_dict` reconciles old/new fields per §6. Write path dual-writes.
- GUI: replace `_edit_webpro_id` with a small `WebProIdsDialog` (table or vertical list of inputs, add/remove buttons). Update header button: display `id[0]` when 1, `f"{id[0]} +{N-1}"` when N>1, "—" when empty; tooltip shows full list.
- Tests: see §8.

**Effort:** M. **Risk:** low. Read path is defensive; write path is additive.

### JTF-3 — Order-status derivation (backend pure-read addition)

**Touches:** `project_tracker_backend.py`, `tests/test_regressions.py`.

- Add `ORDER_SIGNAL_TASKS: frozenset[str]` module constant — initially `{"Valves Ordered"}` (the narrow definition; operator can widen per §9).
- Add `get_order_status_rollup() -> dict` returning `{"missing_count": int, "missing_projects": [{"id", "job_name", "job_number", "project_manager", "days_since_created"}]}`. Implementation: one `_load_data` call, iterate projects + tasks once. Exclude `is_test`. Tag projects as "missing" if they have NO task with name in `ORDER_SIGNAL_TASKS` marked complete.
- Tests: see §8.

**Effort:** S. **Risk:** very low. Pure-read, no UI yet. **Blocks JTF-4.**

### JTF-4 — Home dashboard "Missing Parts Orders" card

**Touches:** `project_tracker_gui.py`, `tests/test_regressions.py` (UI test optional).

- Add `self._dash_missing_orders_card = StatCard("Missing Parts Orders", "—")` to the cards row in `_build_dashboard`.
- Make `StatCard` clickable if not already (check object — `StatCard` at gui:1423 already supports `clicked` signal per existing task-card usage at gui:3136-3138). Wire `clicked → self._open_missing_orders_modal`.
- `_refresh_dashboard`: read `backend.get_order_status_rollup()`, set the card value.
- `_open_missing_orders_modal`: new dialog showing the missing-projects table; double-click selects the project and closes (use existing project-selection signal).
- Tests: see §8.

**Effort:** S-M. **Risk:** low. Depends on JTF-3.

### JTF-5 — Source-mode validation + test sweep

**Touches:** `tests/test_regressions.py`, possibly README "What's New" entry only.

- Full test suite run, source-mode launch via `python project_tracker_gui.py` to smoke-check the new dropdown / WebPro editor / dashboard card.
- README "What's New in v1.8.7" entry covering the new capabilities.
- No version.py bump in this audit — that lands as part of the actual implementation phase, **not** here.

**Effort:** S. **Risk:** none.

### Sequencing dependencies

```
JTF-1 (RSS filter)      ──┐
JTF-2 (WebPro multi)    ──┼──→ JTF-5 (validate + docs)
JTF-3 (order derive)  → JTF-4 (Home card) ─┘
```

JTF-1, JTF-2, and JTF-3 can land in any order or in parallel. JTF-4 requires JTF-3.

---

## 8. Tests required (before / alongside implementation)

Add to `tests/test_regressions.py` as a new class — match the existing `V185RegressionTests` pattern.

### RSS (JTF-1)

- `test_list_projects_has_rss_filter_includes_only_projects_with_attachments`
- `test_list_projects_has_rss_filter_excludes_projects_with_attachments_when_false`
- `test_list_projects_has_rss_none_returns_all_projects`
- `test_list_projects_rss_filter_composes_with_text_search`

### WebPro IDs (JTF-2)

- `test_legacy_single_webpro_id_loads_as_single_item_list`
- `test_new_webpro_ids_list_loads_directly`
- `test_missing_both_webpro_fields_loads_as_empty_list`
- `test_create_project_writes_both_webpro_id_and_webpro_ids`
- `test_update_project_dedupes_case_insensitive_webpro_ids`
- `test_update_project_strips_whitespace_and_drops_empty_webpro_ids`
- `test_old_app_reading_new_file_still_sees_first_webpro_id`  (simulated by reading the JSON directly and asserting `webpro_id` key)
- `test_webpro_ids_preserve_insertion_order`

### Order-status derivation (JTF-3)

- `test_order_status_rollup_counts_jobs_missing_valves_ordered_task`
- `test_order_status_rollup_excludes_test_jobs`
- `test_order_status_rollup_marks_completed_when_order_signal_task_complete`
- `test_order_status_rollup_handles_project_with_no_tasks` (e.g. all deleted)
- `test_order_status_rollup_uses_only_configured_task_names`

### Home dashboard card (JTF-4)

- Unit-test the rollup separately; UI test optional. If added: `test_missing_orders_card_text_matches_rollup_count`.

### Regression sweep

- All existing v1.8.5 tests still pass (no regression).
- WebPro IDs migration does not alter any other field — round-trip diff test for a fully-populated project record (excluding the WebPro fields).

---

## 9. Operator decisions needed

These block implementation; they don't block this audit.

### D1. Definition of "valves or parts ordered"

| Option | Tasks counted as "ordered" | Reasoning |
|---|---|---|
| **D1a (recommended)** | `{"Valves Ordered"}` | Narrowest. Matches the literal phrase. Cleanest signal. |
| D1b | `{"Valves Ordered", "Phoenix Material Submittal Approved"}` | "Materials in motion" — broader; useful if submittal-approved is when ATS actually places the PO. |
| D1c | All "Materials" phase tasks complete | Whole-phase rollup. Likely too strict for an early-warning surface. |
| D1d | New explicit `parts_ordered: bool` field on `ProjectRecord` | Avoids any task-name coupling. Costs schema change + migration + a new UI control. **Not recommended for v1.** |

Pick one. D1a is the v1 default if no decision is made.

### D2. WebPro IDs editor UX

| Option | Description |
|---|---|
| **D2a (recommended)** | Small modal with table rows (`ID` + `Remove` button), `+ Add` button at top, OK/Cancel. Same dialog used for create and edit. |
| D2b | Inline editor on the project header — series of small chips with `×` to remove and a `+` to add. More modern; bigger header footprint. |
| D2c | Keep the existing single-line input but accept comma-separated values. Cheapest, but operator must remember the convention. |

### D3. Project-row indicator for RSS

| Option | Description |
|---|---|
| **D3a (recommended)** | Suffix the job name with `📎` when `rss_files` non-empty. Cheap, immediately visible. |
| D3b | A small badge widget next to the job name. Nicer-looking; more layout work. |
| D3c | No indicator; rely solely on the dropdown. Discoverability suffers. |

### D4. Home-card click target

| Option | Description |
|---|---|
| **D4a (recommended)** | Click opens a modal with the missing-jobs list; double-click a row to select that project and close the modal. |
| D4b | Click filters the sidebar project list to those jobs (i.e. wires through a new "Missing orders" entry in the new RSS-style filter dropdown). |
| D4c | Click toggles an indicator color but no drill-down. Lowest value. |

---

## 10. Recommendation

**Status: ready after decisions.**

The audit is complete. Three of four decisions are recommendations (defaults if no input); only **D1** (definition of "valves or parts ordered") materially shapes the feature, and only because the literal phrasing is ambiguous between the narrow "Valves Ordered" task and the broader "Materials in motion" reading.

If you accept D1a / D2a / D3a / D4a as written, the implementation is unblocked and can proceed in this order:

1. JTF-1 (RSS filter + 📎 indicator) — surgical, ~50 lines + tests.
2. JTF-2 (multiple WebPro IDs + dual-write) — moderate, ~100 lines + tests.
3. JTF-3 (order-status backend method) — small.
4. JTF-4 (Home Missing-Orders card + drill-down modal) — small-medium, depends on JTF-3.
5. JTF-5 (smoke + docs) — small.

Total estimate: 1 release worth of work, candidate version **v1.8.7** (PATCH if framed as additive UX; MINOR if framed as new features — operator's call).

---

## Confirmation of constraints honored

- ✅ No source code changed.
- ✅ No data schema changed.
- ✅ `version.py` not touched (still v1.8.6 in main repo).
- ✅ `build.bat`, `installer.iss`, `updater.py`, `commons/` not touched.
- ✅ Financials and auth code not touched.
- ✅ No release published.
- ✅ No feature branch created.
- ✅ Only one new file written: this audit document at the main repo root.
