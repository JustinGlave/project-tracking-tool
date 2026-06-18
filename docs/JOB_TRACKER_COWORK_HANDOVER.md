# Job Tracker — Cowork Handover

**Purpose:** Transfer the in-flight JTF feature work (RSS filter, multiple WebPro IDs, Home order-status button) to a fresh Claude Cowork session with zero loss of context.

**Written:** 2026-06-02 · **Author context:** end of the JTF-1/2/3 implementation + merge-gate session.

**Read this top to bottom once before touching anything.** The single most important item is §1 (the worktree trap). Everything else is reference.

---

## 0. TL;DR — where things stand

- Three operator-approved features (JTF-1, JTF-2, JTF-3) are **fully implemented, tested (54/54 pass), committed, and pushed** to `origin/feature/job-tracker-rss-webpro-orders`.
- The merge gate concluded **A. Merge-ready**.
- **NOT yet done:** (a) merge to `main`, (b) the v1.8.7 release (version bump, README, build, GitHub release).
- `version.py` is still `1.8.6`. That is correct — the bump belongs to the release step, not the feature work.
- Nothing has been released, tagged, or published. No `gh release`, no tag.

**The next decision is the operator's:** merge now, or do anything else first. Do not merge without explicit go-ahead.

---

## 1. ⚠️ CRITICAL: the worktree trap (read this first)

There are **two working directories** for this repo on this machine:

| Path | Role | Use it? |
|---|---|---|
| `C:\Users\justing\PycharmProjects\Job Tracker` | **The real repo.** All JTF work, commits, branch, tests, and the `.venv` live here. | **YES — work here.** |
| `C:\Users\justing\PycharmProjects\Job Tracker\.claude\worktrees\busy-kowalevski-154a16` | A stale git worktree from an earlier session, **13+ commits behind**, on branch `claude/busy-kowalevski-154a16`. | **NO — ignore entirely.** |

The previous session's *shell* defaulted to the worktree, so every git/python command was explicitly prefixed with `git -C "C:/Users/justing/PycharmProjects/Job Tracker"` or `cd "C:/Users/justing/PycharmProjects/Job Tracker" && …`. **Do the same.** If you run a bare `git status` or `python -m unittest` from the worktree you will see wrong/old results.

**Concrete failure mode already observed:** running `python -m unittest tests.test_regressions` from the worktree picked up the stale copy and reported 15 errors for code that was actually correct. Running it from the real repo reported `OK`. **Always `cd "C:/Users/justing/PycharmProjects/Job Tracker"` first.**

Verify you're in the right place:
```bash
cd "C:/Users/justing/PycharmProjects/Job Tracker"
git branch --show-current     # → feature/job-tracker-rss-webpro-orders
git log --oneline -1          # → ee9e9c7 docs: JTF feature branch merge gate report
```

---

## 2. Repo facts (verified live this session)

| Fact | Value |
|---|---|
| Real repo root | `C:\Users\justing\PycharmProjects\Job Tracker` |
| GitHub | `https://github.com/JustinGlave/project-tracking-tool` |
| Current branch (real repo) | `feature/job-tracker-rss-webpro-orders` |
| Feature tip (local == origin) | `ee9e9c7` |
| `main` tip (local == origin) | `8a85aae` (v1.8.6, commons retrofit) |
| Working tree | clean |
| `version.py` | `1.8.6` (unchanged — bump at release) |
| Python venv | `C:\Users\justing\PycharmProjects\Job Tracker\.venv\Scripts\python.exe` |
| Test result | `Ran 54 tests … OK` |
| App stack | Python 3 · PySide6 · `phoenix_commons` (commons submodule, editable-installed) · openpyxl · pyxlsb |
| Data file | `%APPDATA%\ATS Inc\Project Tracking Tool\project_tracker_data.json` (or a shared-folder override via File → Data Location) |

**v1.8.6 context:** the repo was retrofitted to a shared `phoenix_commons` package (theme, widgets, paths, updater facade through commons). `paths.py` at the root wraps `phoenix_commons.paths.resource_path`. Source mode requires `pip install -e ./commons` (already done in the existing `.venv`). None of the JTF work touched commons.

---

## 3. The 8 commits on the feature branch (oldest → newest)

```
a64906e docs: add JTF feature audit + docs/ output convention
85d3771 feat(jtf-1): RSS filter dropdown + project-row paperclip indicator
410f601 docs(jtf-1): RSS filter implementation report
39318ce feat(jtf-2): multiple WebPro IDs with backward-compatible dual-write
a3e1821 docs(jtf-2): multiple WebPro IDs implementation report
b9716d3 feat(jtf-3): Home order-status button + OrderStatusDialog modal
b29729e docs(jtf-3): Home order-status button implementation report
ee9e9c7 docs: JTF feature branch merge gate report
```

Linear history, no mid-feature merges from main. `feat`/`docs` split per JTF. All pushed.

**Files changed vs `origin/main`** (9 files, +2102 / −13):

| File | What changed |
|---|---|
| `project_tracker_backend.py` | RSS filter, WebPro multi-ID model + helpers, order-status helper + rollup + filter |
| `project_tracker_gui.py` | RSS dropdown + 📎 indicator, `WebProIdsDialog`, `OrderStatusDialog` + Home button |
| `tests/test_regressions.py` | 25 new tests across 3 JTF classes |
| `CLAUDE.md` | +4 lines: "docs go in `docs/`" convention |
| `docs/*.md` (5 files) | audit + 3 per-feature reports + merge-gate report |

---

## 4. What each feature does (precise behavior + code anchors)

All line numbers verified live against the feature tip `ee9e9c7`. They will drift if code is edited — re-grep the symbol name if a number looks off.

### JTF-1 — RSS filter + project-row indicator

**Behavior:** Sidebar dropdown **All projects / 📎 Has RSS / No RSS** filters the project list. Project rows get a `📎` suffix when `rss_files` is non-empty (coexists with the `📌` pinned prefix).

| Symbol | Location |
|---|---|
| `list_projects(..., has_rss: Optional[bool] = None)` | `project_tracker_backend.py:640` |
| RSS filter block (uses `_migrate_rss_files`) | `project_tracker_backend.py:~655` |
| `self.rss_filter_combo` (dropdown) | `project_tracker_gui.py:2984` |
| `refresh_project_list` reads `currentData()` | `project_tracker_gui.py:5025` |
| `📎` row suffix | in `refresh_project_list` (search `rss_files`) |
| Tests | `JTF1RSSFilterTests` — `tests/test_regressions.py:720` (4 tests) |

Rule: filter uses `_migrate_rss_files(item)` so a legacy single-string `csv_file_path` still counts as "has RSS."

### JTF-2 — Multiple WebPro IDs (backward-compatible)

**Behavior:** A job can have many WebPro IDs. Header button shows `—` / the single ID / `"N WebPro IDs"` (full list in tooltip). Clicking opens a modal (list + Add + Remove Selected + OK/Cancel). The Edit-Project dialog routes through the same modal.

**Data contract (the part you must not break):**
- `ProjectRecord` has BOTH `webpro_id: str` (legacy mirror) and `webpro_ids: list[str]` (canonical).
- **Read** precedence: `webpro_ids` (list) → else legacy `webpro_id` wrapped as `[id]` → else `[]`.
- **Write** is a **dual-write**: every save sets `webpro_ids = normalized_list` AND `webpro_id = normalized_list[0] if normalized_list else ""`. This keeps older app builds (which only read `webpro_id`) working — forward/backward compatible, **no migration script, no required field.**
- Normalize = strip, drop empties, case-insensitive dedupe preserving first-seen casing + insertion order.

| Symbol | Location |
|---|---|
| `webpro_ids: list[str]` field | `project_tracker_backend.py:42` |
| `_normalize_webpro_ids()` | `project_tracker_backend.py:242` |
| `_migrate_webpro_ids()` | `project_tracker_backend.py:272` |
| `create_project` dual-write | `project_tracker_backend.py:474` |
| `update_project` reconcile branch | `project_tracker_backend.py:~570` |
| `_project_from_dict` reads `_migrate_webpro_ids` | `project_tracker_backend.py:2222` |
| search includes WebPro haystack | in `list_projects` (search `_migrate_webpro_ids`) |
| `class WebProIdsDialog` | `project_tracker_gui.py:2624` |
| header button `_edit_webpro_id` | `project_tracker_gui.py:3557` |
| `ProjectDialog._edit_webpro_ids` | `project_tracker_gui.py:340` |
| Tests | `JTF2WebProIDsTests` — `tests/test_regressions.py:543` (12 tests) |

**Trap that was already fixed:** the naive "keep the single `QLineEdit` in ProjectDialog" path would silently drop extra IDs on whole-project edit. It now round-trips the full list and `edit_current_project` passes `webpro_ids=…` (not `webpro_id=…`). Do not regress this.

### JTF-3 — Home order-status button (ONE button, no redesign)

**Behavior:** ONE new button on the **existing** Home dashboard: `"Valve/Parts Order Status — X missing / Y ordered"`. Opens a modal with two tabs (**Missing Orders (N)** / **Ordered (N)**), each a table of Job # | Job Name | PM | Updated. Double-click a row selects that project in the sidebar and closes the modal.

**Rule (operator decision D1a):** a job is "ordered" iff it has a task named **"Valves Ordered"** (case-insensitive) with `is_complete == True`. Derived from existing task data — **no schema field added.** To widen later, edit the single frozenset literal.

| Symbol | Location |
|---|---|
| `ORDER_SIGNAL_TASKS = frozenset({"Valves Ordered"})` | `project_tracker_backend.py:210` |
| `_project_has_order_signal()` | `project_tracker_backend.py:213` |
| `list_projects(..., order_status=None)` (`"ordered"`/`"missing"`) | `project_tracker_backend.py:640` (filter `~689`) |
| `get_order_status_rollup()` | `project_tracker_backend.py:2000` |
| `class OrderStatusDialog` (+ `project_selected` Signal) | `project_tracker_gui.py:2707` |
| Home button `_dash_order_status_btn` | `project_tracker_gui.py:3133` |
| button text refresh in `_refresh_dashboard` | `project_tracker_gui.py:3220` |
| `_open_order_status_dialog` | `project_tracker_gui.py:4224` |
| `_select_project_by_id` (drill-down) | `project_tracker_gui.py:4238` |
| Tests | `JTF3OrderStatusTests` — `tests/test_regressions.py:440` (9 tests) |

---

## 5. What is NOT done (the actual remaining work)

### 5a. Merge to main (operator go-ahead required)

Exact fast-forward plan (no conflicts — feature is linear on top of main):
```bash
cd "C:/Users/justing/PycharmProjects/Job Tracker"
git checkout main
git fetch origin
git merge --ff-only feature/job-tracker-rss-webpro-orders
git push origin main
# optional cleanup:
# git branch -d feature/job-tracker-rss-webpro-orders
# git push origin --delete feature/job-tracker-rss-webpro-orders
```
Alternatively open a PR: `https://github.com/JustinGlave/project-tracking-tool/pull/new/feature/job-tracker-rss-webpro-orders`.

### 5b. v1.8.7 release (separate operation, after merge)

This is the next *feature-complete* deliverable. Steps, in order:

1. **Bump `version.py`** `1.8.6` → `1.8.7` (PATCH: additive UX on top of v1.8.6, not a redesign).
2. **README "What's New in v1.8.7"** — add a section covering the 3 features. Also bump the "Current Version" line.
3. *(Optional polish)* update the sidebar search placeholder to mention WebPro (it now searches WebPro IDs but still says "Search jobs, PM, sales engineer…").
4. **Build** (see §6). Validate the updater zip.
5. **Manual smoke** of the frozen exe + installer.
6. **Release:** `gh release create v1.8.7 …` + upload the 3 assets.

None of these have been started.

---

## 6. Build & release procedure (from CLAUDE.md — do not deviate)

**Never run `cmd /c build.bat` from Bash — it silently fails.** Run each step directly.

**Step 1 — PyInstaller (Bash, from real repo root):**
```bash
.venv/Scripts/pyinstaller --onedir --windowed --icon=PTT_Normal.ico --name=ProjectTrackingTool "--add-data=PTT_Transparent.png;." "--add-data=PTT_Normal.ico;." "--add-data=phoenix_style.qss;." "--add-data=pyxlsb;pyxlsb" --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --collect-submodules=openpyxl --collect-submodules=PySide6.QtCore --collect-submodules=PySide6.QtGui --collect-submodules=PySide6.QtWidgets --hidden-import=pyxlsb -y project_tracker_gui.py
```
> Note: this command is the historical one-folder spec. The repo also has `ProjectTrackingTool.spec` and `build.bat`; confirm with the operator whether commons (`phoenix_commons`) needs an explicit `--add-data`/`--collect-submodules` in the frozen build, since v1.8.6's commons retrofit may have updated the bundling. **Verify the frozen exe actually loads `phoenix_commons` before releasing.**

**Step 2 — Inno Setup installer (PowerShell):**
```powershell
& "C:\Users\justing\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.8.7 "C:\Users\justing\PycharmProjects\Job Tracker\installer.iss"
```
> **Hard rule (preserve):** do NOT add an AppId GUID to `installer.iss`. v1.6.0–v1.8.6 users rely on the AppName-hashed upgrade detection; adding a GUID breaks in-place upgrades.

**Step 3 — Zip archives (PowerShell). The auto-updater zip must contain the folder CONTENTS (exe + `_internal/`), use the `\*` suffix:**
```powershell
Compress-Archive -Path 'dist\ProjectTrackingTool\*' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force
Compress-Archive -Path 'dist\ProjectTrackingTool' -DestinationPath 'dist\ProjectTrackingTool_FullInstall.zip' -Force
```

**Validate the updater zip before releasing:**
```powershell
& "C:\Users\justing\PycharmProjects\Job Tracker\.venv\Scripts\python.exe" -c "from updater import _validate_update_zip; from pathlib import Path; _validate_update_zip(Path('dist/ProjectTrackingTool.zip')); print('OK')"
```

**Release assets to upload:** `ProjectTrackingToolSetup.exe` (installer), `ProjectTrackingTool.zip` (auto-updater payload — full folder contents), `ProjectTrackingTool_FullInstall.zip` (manual full install).

---

## 7. Standing constraints / guardrails (these held through all JTF work — keep holding them)

- **No `Co-Authored-By`** trailers in commits or PRs.
- **New docs go in `docs/`** at the repo root (not repo root loose, not `commons/docs/`).
- Address-book deletion approval is tied to the app's `admin` role (no per-username constant).
- Do not touch, during feature work: `financials_*.py`, `user_auth.py`, `updater.py`, `build.bat`, `installer.iss`, `commons/`, `paths.py`, `phoenix_style.qss`, `version.py`. (Release work touches `version.py` + README only.)
- No new schema fields were added (WebPro reused the dataclass with a non-required field + JSON fallback; order-status is derived). Keep it that way unless the operator approves a migration.

---

## 8. How to verify state in 30 seconds (run these first in any new session)

```bash
cd "C:/Users/justing/PycharmProjects/Job Tracker"
git branch --show-current                 # feature/job-tracker-rss-webpro-orders
git status -s                             # (empty = clean)
git log --oneline -1                      # ee9e9c7
git log --oneline origin/main..HEAD       # the 8 JTF commits
cat version.py | tail -1                  # __version__ = "1.8.6"
.venv/Scripts/python.exe -m unittest tests.test_regressions   # Ran 54 tests ... OK
```

Source-mode launch (opens the GUI — closes cleanly on exit):
```bash
cd "C:/Users/justing/PycharmProjects/Job Tracker"
.venv/Scripts/python.exe project_tracker_gui.py
```

---

## 9. Test inventory (`tests/test_regressions.py`)

| Class | Line | Count | Area |
|---|---|---|---|
| `AuthRegressionTests` | 32 | 5 | pre-existing auth |
| `BackendRegressionTests` | 114 | 7 | pre-existing backend |
| `UpdaterRegressionTests` | 238 | 3 | pre-existing updater |
| `V185RegressionTests` | 270 | 14 | v1.8.5 fix release |
| `JTF3OrderStatusTests` | 440 | 9 | order-status filter + rollup |
| `JTF2WebProIDsTests` | 543 | 12 | WebPro multi-ID model |
| `JTF1RSSFilterTests` | 720 | 4 | RSS filter |
| **Total** | | **54** | all pass |

(Class order in the file is non-sequential by JTF number — that's cosmetic, not a bug.)

---

## 10. Existing reference docs (read for depth, all in `docs/`)

| Doc | Use it for |
|---|---|
| `JOB_TRACKER_RSS_WEBPRO_ORDER_FEATURE_AUDIT.md` | Original audit; the 4 operator decisions (D1a/D2a/D3a/D4a) and why; the full feature rationale. |
| `JOB_TRACKER_JTF1_RSS_FILTER_REPORT.md` | JTF-1 implementation detail. |
| `JOB_TRACKER_JTF2_MULTIPLE_WEBPRO_IDS_REPORT.md` | JTF-2 detail incl. rollback/downgrade safety analysis. |
| `JOB_TRACKER_JTF3_ORDER_STATUS_HOME_BUTTON_REPORT.md` | JTF-3 detail. |
| `JOB_TRACKER_JTF_FEATURE_MERGE_GATE_REPORT.md` | The merge-readiness verdict (A. Merge-ready), full diff scope, exact merge plan. |

There is also `AUDIT_METHODOLOGY.md` (in `docs/` per the convention) — a reusable "verify-before-claiming" code-audit process from an earlier session. Not JTF-related but useful if Cowork is asked to audit.

---

## 11. Known gotchas / non-issues

- **CRLF warnings on commit** (`LF will be replaced by CRLF`) are benign Windows line-ending notices, not errors.
- **Tests must run from the real repo CWD** (see §1). The `.venv` is only in the real repo, not the worktree.
- **Operator decisions are locked:** D1a (`{"Valves Ordered"}`), D2a (modal editor), D3a (📎 suffix), D4a (modal + double-click drill-down). Don't re-litigate without the operator.
- **Order-status filter exists in the backend** (`order_status=` on `list_projects`) but is intentionally NOT surfaced as a sidebar dropdown — only the Home button consumes it. That was the approved scope; adding a sidebar order filter would be new scope.
- **No project-row indicator for ordered/missing** — deliberately skipped (rows already carry 📌 and 📎; spec said skip unless trivial).

---

## 12. Suggested first message to the Cowork session

> "Job Tracker is at `C:\Users\justing\PycharmProjects\Job Tracker` (NOT the `.claude/worktrees/...` worktree). Branch `feature/job-tracker-rss-webpro-orders` (tip `ee9e9c7`) holds 3 finished features, 54/54 tests green, merge-ready. Read `docs/JOB_TRACKER_COWORK_HANDOVER.md` first. I want to [merge to main / cut the v1.8.7 release / something else]."

---

## 13. Confirmation of state at handover

- ✅ Feature branch pushed; local == origin (`ee9e9c7`).
- ✅ 54/54 tests pass (verified live in this session).
- ✅ Working tree clean.
- ✅ `version.py` still `1.8.6` (release-step bump pending, by design).
- ✅ No merge to main performed.
- ✅ No tag, no GitHub release, no asset upload.
- ✅ No changes to financials / auth / updater / build / installer / commons / version (feature scope respected).
- ✅ Merge-gate verdict: **A. Merge-ready** (pending operator's visual smoke sign-off, which is the only gate not automatable).
