# Project Tracking Tool — Claude Instructions

## Retrofit state

**Wave 8b (commons retrofit) in progress.** B1 (commons submodule +
requirements-dev + CI minor edit) landed 2026-05-27 by explicit
operator-approved early-open override (the doctrinal cooldown floor
was 2026-06-09, computed from Wave 8a's 2026-05-26 merge per
MIGRATION_RULES § Frequency limits; floor breached intentionally with
no unresolved technical blockers). B2-B11 sequence in
`phoenix-commons/docs/ui-platform-baseline-v1/WAVE_8B_IMPLEMENTATION_BRIEF.md`.

The commons submodule lives at `commons/` (pinned to `phoenix-commons`
`main`). Source mode requires `pip install -e ./commons` per ADR-015.

Critical preservation rules for Wave 8b:

- **No AppId GUID added to installer.iss** — v1.6.0..v1.8.5 users have
  AppName-hashed default; adding one would break upgrade detection.
- Full-folder updater payload preserved (`expected_internal=True`).
- All Excel / financials / user-auth code paths preserved-local.
- `version.py` stays at `1.8.5` (Decision #1 tag-skip).

## Build Process

**Never run `cmd /c build.bat` from Bash.** It silently fails in the Git Bash environment. Instead, run each build step directly:

### Step 1 — PyInstaller (Bash)
```
.venv/Scripts/pyinstaller --onedir --windowed --icon=PTT_Normal.ico --name=ProjectTrackingTool "--add-data=PTT_Transparent.png;." "--add-data=PTT_Normal.ico;." "--add-data=phoenix_style.qss;." "--add-data=pyxlsb;pyxlsb" --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --collect-submodules=openpyxl --collect-submodules=PySide6.QtCore --collect-submodules=PySide6.QtGui --collect-submodules=PySide6.QtWidgets --hidden-import=pyxlsb -y project_tracker_gui.py
```

### Step 2 — Inno Setup installer (PowerShell)
Inno Setup is installed at: `C:\Users\justing\AppData\Local\Programs\Inno Setup 6\ISCC.exe`
```powershell
& "C:\Users\justing\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=<VERSION> "C:\Users\justing\PycharmProjects\Job Tracker\installer.iss"
```

### Step 3 — Zip archives (PowerShell)

The auto-updater zip must contain the **contents** of the PyInstaller folder (exe + `_internal/` at top level), not just the exe — `_validate_update_zip` rejects packages missing `_internal/`. Use the `\*` suffix on the source path:

```powershell
Compress-Archive -Path 'dist\ProjectTrackingTool\*' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force
Compress-Archive -Path 'dist\ProjectTrackingTool' -DestinationPath 'dist\ProjectTrackingTool_FullInstall.zip' -Force
```

Validate the auto-updater zip before releasing:

```powershell
& "C:\Users\justing\PycharmProjects\Job Tracker\.venv\Scripts\python.exe" -c "from updater import _validate_update_zip; from pathlib import Path; _validate_update_zip(Path('dist/ProjectTrackingTool.zip')); print('OK')"
```

### Release assets to upload
- `dist\ProjectTrackingToolSetup.exe` — installer for new users
- `dist\ProjectTrackingTool.zip` — auto-updater payload (full folder contents)
- `dist\ProjectTrackingTool_FullInstall.zip` — full folder, for manual installs

---

## Access control

- Address book deletion approval is tied to the app's `admin` role. There is no per-username approver constant to keep in sync.

## Commit style

- No `Co-Authored-By` trailers in commit messages or PR descriptions.

## Documentation output

- New audit reports, feature plans, and design docs go in `docs/` at the Job Tracker repo root — not at the repo root, not in `commons/docs/`.
