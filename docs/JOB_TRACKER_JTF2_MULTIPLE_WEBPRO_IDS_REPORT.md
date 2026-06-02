# JTF-2 — Multiple WebPro IDs (Implementation Report)

Implements the multi-WebPro support approved by the operator on 2026-06-02 (decision D2a) on top of JTF-1.

- **Branch:** `feature/job-tracker-rss-webpro-orders` (pushed; tracks `origin/feature/job-tracker-rss-webpro-orders`)
- **Base:** `main @ 8a85aae` (v1.8.6)
- **Commits added by JTF-2:**
  - `39318ce` — feat(jtf-2): multiple WebPro IDs with backward-compatible dual-write
  - (this report will be committed next)
- **Test result:** 45 / 45 pass (33 prior + 12 new JTF-2)

---

## 1. Files changed

| File | Change |
|---|---|
| `project_tracker_backend.py` | Added `webpro_ids: list[str]` field to `ProjectRecord`; new `_normalize_webpro_ids()` + `_migrate_webpro_ids()` helpers; `_project_from_dict` reads new field with legacy fallback; `create_project` and `update_project` dual-write both keys; `list_projects` text search now includes WebPro IDs. |
| `project_tracker_gui.py` | New `WebProIdsDialog` class; `_edit_webpro_id` rewired to use it; header button display switched to "—" / single ID / "N WebPro IDs"; `ProjectDialog` swapped its single QLineEdit for a button that opens the same dialog and now round-trips the full list; `edit_current_project` caller passes `webpro_ids=…` (preserves multi-ID projects on whole-project edit). |
| `tests/test_regressions.py` | New `JTF2WebProIDsTests` class with 12 tests. |
| `docs/JOB_TRACKER_JTF2_MULTIPLE_WEBPRO_IDS_REPORT.md` | This report. |

No other files modified. Audit + JTF-1 commits are still on the branch and unaffected.

---

## 2. Data compatibility behavior

### Storage shape

**Before JTF-2** (legacy, v1.6.0 – v1.8.6):
```json
{ "id": 1, "job_name": "...", "webpro_id": "12345" }
```

**After JTF-2** (v1.8.7+, dual-write):
```json
{
  "id": 1,
  "job_name": "...",
  "webpro_id": "12345",                 // ← legacy mirror = webpro_ids[0] (or "")
  "webpro_ids": ["12345", "67890"]      // ← new canonical field
}
```

### Read path (`_project_from_dict` → `_migrate_webpro_ids`)

Resolution precedence:
1. Stored `webpro_ids` (list) → normalized and returned.
2. Stored `webpro_id` (legacy single string) → wrapped in a 1-item list.
3. Neither present → empty list.

`_project_from_dict` then sets `ProjectRecord.webpro_ids` to the resolved list AND `ProjectRecord.webpro_id` to `webpro_ids[0] if webpro_ids else ""`. Any caller that still reads `.webpro_id` (display code, exports if any are added later) keeps seeing the first ID without change.

### Write path (`create_project` / `update_project`)

Callers may supply either field:
- `update_project(webpro_ids=["a", "b"])` → canonical.
- `update_project(webpro_id="a")` → legacy single-string call.
- `update_project(webpro_ids=…, webpro_id=…)` → `webpro_ids` wins.

Backend reconciliation always:
1. Normalizes via `_normalize_webpro_ids`: strip whitespace, drop empties, dedupe **case-insensitive preserving first-seen casing and insertion order**.
2. Dual-writes `"webpro_id": normalized[0] if normalized else ""` and `"webpro_ids": normalized`.

### Rollback safety

If a user downgrades to v1.8.6 (or earlier) after a v1.8.7+ save:
- The older app reads `webpro_id` (the legacy key, still populated with the first ID).
- The older app does not understand `webpro_ids` and ignores it.
- Any additional IDs become invisible until the user re-upgrades. **No data is removed from the JSON file by the downgrade itself.**

If the older app *writes* the project after a downgrade, the new write only includes `webpro_id` (since the v1.8.6 code path only writes that key). At that point the `webpro_ids` list in JSON may go stale relative to `webpro_id`. On the next v1.8.7+ load, `_migrate_webpro_ids` would prefer the existing list — meaning the older-app's edit to the single field could be lost. This is an **inherent risk of forward-incompatible writes during a downgrade**, not a JTF-2 regression; the same risk exists for any new field. Documented here for transparency.

### Preservation of existing values

- A project with only the legacy `"webpro_id": "12345"` loads as `webpro_ids=["12345"]` and re-saves as both keys, preserving `"12345"`.
- A project with no WebPro at all loads as `[]` / `""` — also no change.
- A project being edited through the main Edit Project dialog **no longer loses additional IDs** because `ProjectDialog` round-trips the full list via internal state (the bug that would have arisen from naively keeping the single QLineEdit was caught during JTF-2 implementation; see §3 below).

---

## 3. GUI editor behavior

### New `WebProIdsDialog` (small modal)

```
┌─────────────────────────────────┐
│ WebPro IDs                      │
├─────────────────────────────────┤
│ Add one or more WebPro IDs…     │
│ ┌─────────────────────────────┐ │
│ │ 12345                       │ │
│ │ 67890                       │ │  ← QListWidget (selectable)
│ │                             │ │
│ └─────────────────────────────┘ │
│ ┌─────────────────┐ ┌────────┐  │
│ │ Enter WebPro ID │ │ + Add  │  │
│ └─────────────────┘ └────────┘  │
│ ┌─────────────────────────────┐ │
│ │ Remove Selected             │ │
│ └─────────────────────────────┘ │
│                  [ Cancel ] [ OK ] │
└─────────────────────────────────┘
```

Behavior:
- **Add**: typing + Enter or `+ Add` button. Live case-insensitive dedupe (visible empty/dupe input clears silently rather than spamming a warning).
- **Remove**: select one or more rows, click **Remove Selected**. Removes from highest index first so earlier rows don't shift mid-loop.
- **OK** returns the visible list; backend re-normalizes on save (final source of truth for whitespace / dedupe / empty drop).
- **Cancel** discards changes.
- `Add` / `Remove Selected` buttons have `setAutoDefault(False)` so Enter in the input field doesn't accidentally trigger them — it goes to `_add_from_input` via `returnPressed`.

### Header button display

| State | Text | Tooltip |
|---|---|---|
| No IDs | `—` | "Click to add WebPro IDs" |
| 1 ID | The ID string | `WebPro ID: <id>` + "Click to edit." |
| N > 1 IDs | `"N WebPro IDs"` | bulleted list of all IDs + "Click to edit." |

Same fixed 110px width / 42px height as before — no header layout change. View-only users have the button disabled (unchanged from prior behavior).

### `ProjectDialog` integration

Previously: a single-line `QLineEdit` labeled "WebPro ID" that overwrote whatever was stored on save. This would have silently dropped additional IDs after JTF-2.

Now: a labeled "WebPro IDs" `QPushButton` whose text reflects the current count (`"Add WebPro IDs…"` / `"WebPro ID: 12345"` / `"3 WebPro IDs"`). Clicking opens the same `WebProIdsDialog`. The dialog mutates an internal `self._webpro_ids: list[str]` and `get_data()` returns a `ProjectRecord` with `webpro_ids=…`. The `edit_current_project` caller passes that list through.

Result: editing a multi-ID project through either the header button OR the main Edit Project dialog preserves all IDs.

### Open behavior

The audit noted there is no "open" semantics for WebPro IDs — they are stored display values, not URLs. Clicking the header button always opens the editor. Nothing changed here.

---

## 4. Search behavior

WebPro IDs are now included in the existing project text search.

`ProjectTrackerBackend.list_projects(search_text=…)` builds a small per-record haystack via `_migrate_webpro_ids(item)` (so legacy `webpro_id`-only records also match) and adds one OR clause to the existing match expression. No new public parameter; no UI changes; the placeholder text on the sidebar search box still reads "Search jobs, PM, sales engineer..." — extending that string to mention WebPro is left for a follow-up polish pass since the change here is additive (no regression for users who don't search by WebPro).

This was the audit's "defer unless trivial" item. A single OR clause counts as trivial; both new and legacy fields are covered.

---

## 5. Tests & validation

### New regression tests (12)

All in `JTF2WebProIDsTests` (`tests/test_regressions.py`):

| Area | Test |
|---|---|
| Read path | `test_legacy_single_webpro_id_loads_as_single_item_list` |
| Read path | `test_new_webpro_ids_list_loads_directly` |
| Read path | `test_missing_both_webpro_fields_loads_as_empty_list` |
| Write path | `test_create_project_writes_both_webpro_id_and_webpro_ids` |
| Write path | `test_create_project_with_legacy_single_field_still_works` |
| Normalization | `test_update_project_dedupes_case_insensitive_preserves_order` |
| Normalization | `test_update_project_strips_whitespace_and_drops_empty` |
| Reconciliation | `test_update_project_via_legacy_webpro_id_still_works` |
| Forward compat | `test_existing_app_reading_new_file_sees_first_id_as_webpro_id` |
| Cleanup | `test_clearing_webpro_ids_removes_legacy_value_too` |
| Search | `test_list_projects_search_matches_against_any_webpro_id` |
| Search | `test_list_projects_search_also_matches_legacy_webpro_id_only` |

### Validation run

| Step | Result |
|---|---|
| `py_compile project_tracker_backend.py project_tracker_gui.py tests/test_regressions.py` | clean |
| Full suite from main repo CWD | **45 / 45 pass** (33 prior + 12 new JTF-2) |
| Import smoke | `'webpro_ids' in ProjectRecord.__dataclass_fields__` → True; `_normalize_webpro_ids(['  a  ', '', 'A', 'b', 'a'])` → `['a', 'b']` |

Interactive GUI smoke for the new modal was not run by this report (avoids popping a window). Operator can verify by launching `python project_tracker_gui.py` from the main repo: click the WebPro ID header button on any project — the new modal should appear with Add/Remove/OK/Cancel.

---

## 6. Next step — JTF-3: Order-status derivation

JTF-3 is now unblocked. Per the audit (decision D1a):

- Add `ORDER_SIGNAL_TASKS: frozenset[str]` module constant — initially `{"Valves Ordered"}`.
- Add `ProjectTrackerBackend.get_order_status_rollup() -> dict` returning `{"missing_count": int, "missing_projects": [...]}`. Pure read-side; one `_load_data` pass; excludes `is_test` jobs.
- No schema change required (derived from existing tasks).
- Backend-only; UI lands in JTF-4 (Home dashboard "Missing Parts Orders" card).
- Tests: rollup count, test-job exclusion, configured-task-name behavior, project-with-no-tasks edge case.

---

## 7. Constraints confirmation

- ✅ **Existing WebPro IDs preserved.** Read fallback wraps legacy single string in a 1-item list; dual-write keeps the legacy key populated; `ProjectDialog` round-trips the full list end-to-end (no silent drop on whole-project edit).
- ✅ **No RSS behavior regression.** JTF-1 sidebar dropdown and 📎 row indicator are untouched; the four JTF-1 regression tests still pass.
- ✅ **No financials / auth changes.** `financials_*.py` and `user_auth.py` untouched.
- ✅ **No updater / build / installer changes.** `updater.py`, `build.bat`, `installer.iss` untouched.
- ✅ **No release / tag / publish occurred.** Commits live on `feature/job-tracker-rss-webpro-orders` (pushed to GitHub for review visibility but no PR opened, no tag created, no GitHub release).
- ✅ **No `version.py` change.** Still `1.8.6` on the branch tip.
- ✅ **UI scope discipline.** One new dialog (`WebProIdsDialog`), one rewired method (`_edit_webpro_id`), one button replacement in `ProjectDialog`, one header-button display switch. No broader header redesign, no menu changes, no Home-card work.
