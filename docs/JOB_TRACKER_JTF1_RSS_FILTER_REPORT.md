# JTF-1 — RSS Filter + Project-Row Indicator (Implementation Report)

Implements the RSS-presence filter and visual indicator approved by the operator on 2026-06-02 (decision D3a). Scope is intentionally narrow per the JTF-1 spec.

- **Branch:** `feature/job-tracker-rss-webpro-orders`
- **Base:** `main @ 8a85aae` (v1.8.6)
- **Status:** implemented + tests passing + not pushed
- **Commits on this branch:**
  1. `a64906e` — docs: add JTF feature audit + docs/ output convention
  2. `85d3771` — feat(jtf-1): RSS filter dropdown + project-row paperclip indicator

---

## 1. Files changed

| File | Change |
|---|---|
| `project_tracker_backend.py` | `list_projects()` gains `has_rss: Optional[bool] = None` parameter; filter applied between text-search and sort using `_migrate_rss_files()` for legacy-aware presence detection. |
| `project_tracker_gui.py` | New `QComboBox` `self.rss_filter_combo` in the sidebar (between search and sort). `refresh_project_list()` reads its `currentData()` and forwards `has_rss=…`. Project rows append `📎` after the job name when `project.rss_files` is non-empty. |
| `tests/test_regressions.py` | New `JTF1RSSFilterTests` class with four tests. |
| `CLAUDE.md` | (preceding docs commit) added "docs/" convention for audit/feature/design output. |
| `docs/JOB_TRACKER_RSS_WEBPRO_ORDER_FEATURE_AUDIT.md` | (preceding docs commit) the approved audit. |
| `docs/JOB_TRACKER_JTF1_RSS_FILTER_REPORT.md` | This report. |

No other files modified.

---

## 2. Backend filter behavior

Signature (full default-arg compatibility preserved):

```python
def list_projects(
    self,
    search_text: str = "",
    include_test: bool = True,
    sort_by: str = "updated",
    sort_asc: bool = False,
    has_rss: Optional[bool] = None,
) -> list[ProjectRecord]:
```

Filter semantics:

| `has_rss` | Result |
|---|---|
| `None` | No RSS filter applied (legacy default; behavior unchanged). |
| `True` | Returns only projects whose stored `rss_files` resolves to a non-empty list. |
| `False` | Returns only projects whose stored `rss_files` is empty. |

Implementation note: the filter calls `_migrate_rss_files(item)` (the same helper used by `_project_from_dict`), so a legacy project that still carries the pre-`rss_files` single-string `csv_file_path` field is correctly counted as "has RSS". This avoids drift between the row indicator (which uses the migrated `ProjectRecord.rss_files`) and the filter (which reads from the raw dict for performance).

Composition with other parameters is preserved:
- Text search runs first; RSS filter runs against the post-search candidate set.
- Sort runs last and is unaffected.
- `include_test` is unaffected.

---

## 3. GUI dropdown behavior

A new `QComboBox` sits in the sidebar between the search box and the existing sort row:

```
┌─ sidebar ───────────────────────┐
│  [ Home ]   Title               │
│  ┌─────────────────────────────┐│
│  │ Search jobs, PM, ...        ││  ← QLineEdit (unchanged)
│  └─────────────────────────────┘│
│  ┌─────────────────────────────┐│
│  │ All projects        ▾       ││  ← NEW: rss_filter_combo
│  └─────────────────────────────┘│
│  ┌─────────────────┬──[ ↑ A–Z ]┐│
│  │ Last Updated ▾  │           ││  ← sort row (unchanged)
│  └─────────────────┴───────────┘│
│  ...                            │
└─────────────────────────────────┘
```

Items:
- `"All projects"` → data `None` (default, current behavior)
- `"📎 Has RSS"` → data `True`
- `"No RSS"` → data `False`

Wiring: `currentIndexChanged → refresh_project_list`. The same handler text-search and sort already use, so all three filters compose without special-case logic.

Tooltip on the combo box: *"Filter projects by whether they have an RSS attachment."*

Layout impact: one additional `QComboBox` row in the sidebar `panel_layout`. No widget removed; no font, color, or spacing changes; no overall sidebar width change.

---

## 4. RSS indicator behavior

In `refresh_project_list()`, after the existing pinned-prefix handling:

```python
if project.pinned:
    job_name = "📌 " + job_name
if project.rss_files:
    job_name = f"{job_name} 📎"
```

Behavior:
- Indicator appears after the (possibly truncated) job name, before being placed in the row's `QLabel`.
- Coexists with the pinned `📌` prefix — a pinned project with RSS shows `"📌 Job Name 📎"`.
- Pure visual additon — no click handler, no tooltip change.
- Driven by `project.rss_files`, which `_project_from_dict` builds via `_migrate_rss_files`, so legacy `csv_file_path`-only projects show the indicator correctly.

---

## 5. Tests & validation

### New regression tests

In `tests/test_regressions.py`, class `JTF1RSSFilterTests`:

| Test | Asserts |
|---|---|
| `test_has_rss_true_returns_only_projects_with_rss` | Three projects: A (no RSS), B (RSS), C (RSS). `has_rss=True` → `["Beta", "Gamma"]`. |
| `test_has_rss_false_returns_only_projects_without_rss` | Same fixture. `has_rss=False` → `["Alpha"]`. |
| `test_has_rss_none_returns_all_projects` | Same fixture. `has_rss=None` → all three names. |
| `test_has_rss_filter_composes_with_text_search` | `search_text="B", has_rss=True` → `["Beta"]`. Sanity check that text-only `search_text="B"` also returns just Beta (no cross-contamination). |

Fixture sets up three projects directly via `backend.create_project` then attaches RSS to two via `backend.update_project(..., rss_files=[{...}])`, which is the same path `_attach_rss` uses in the GUI.

### Validation steps run

| Step | Command | Result |
|---|---|---|
| Syntax check | `python -m py_compile project_tracker_backend.py project_tracker_gui.py tests/test_regressions.py` | clean |
| Full test suite | `python -m unittest tests.test_regressions -v` | **33 / 33 pass** (29 existing + 4 new JTF-1) |
| Import smoke | `python -c "import project_tracker_gui; ..."` | imports OK; `has_rss` confirmed in `list_projects.__code__.co_varnames` |

GUI smoke at the dropdown / row-indicator level was not run interactively as part of this report (avoids popping a window during automated work). The operator may launch `python project_tracker_gui.py` from the main repo to verify the dropdown and indicator render as described.

---

## 6. Next step — JTF-2: Multiple WebPro IDs

JTF-2 is unblocked. Per the audit (D2a):
- Backend: add `webpro_ids: list[str]` to `ProjectRecord`; dual-write `webpro_id` (first item) + `webpro_ids` (full list); read path reconciles legacy `webpro_id` into a 1-item list.
- GUI: replace single-input `_edit_webpro_id` with a small modal (`WebProIdsDialog`): rows of `ID + Remove`, `+ Add`, OK/Cancel. Header button shows first ID + `"+N"` for additional, full list in tooltip.
- Tests: legacy load, list load, missing both, dual-write round-trip, dedupe, whitespace handling, insertion-order preservation, old-app-reading-new-file compatibility.

JTF-2 should be a separate branch commit set on this same feature branch.

---

## 7. Constraints confirmation

- ✅ **No schema migration occurred.** `rss_files` list shape unchanged; no field added or renamed; no migration code added.
- ✅ **No financials / auth changes.** `financials_*.py` and `user_auth.py` untouched.
- ✅ **No updater / build / installer changes.** `updater.py`, `build.bat`, `installer.iss` untouched.
- ✅ **No release / tag / publish occurred.** Commits live only on `feature/job-tracker-rss-webpro-orders` locally; nothing pushed; no tag created; no GitHub release.
- ✅ **No `version.py` change.** Still `1.8.6`.
- ✅ **Scope discipline.** UI change is two widgets (one `QComboBox`, one f-string suffix). No broader sidebar redesign, no search-box rewrite.
