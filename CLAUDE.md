# Project Tracking Tool — Claude Instructions

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
```powershell
Compress-Archive -Path 'dist\ProjectTrackingTool\ProjectTrackingTool.exe' -DestinationPath 'dist\ProjectTrackingTool.zip' -Force
Compress-Archive -Path 'dist\ProjectTrackingTool' -DestinationPath 'dist\ProjectTrackingTool_FullInstall.zip' -Force
```

### Release assets to upload
- `dist\ProjectTrackingToolSetup.exe` — installer for new users
- `dist\ProjectTrackingTool.zip` — exe only, used by auto-updater
- `dist\ProjectTrackingTool_FullInstall.zip` — full folder, for manual installs

---

## Access control

- Address book deletion approval is tied to the app's `admin` role. There is no per-username approver constant to keep in sync.

## Commit style

- No `Co-Authored-By` trailers in commit messages or PR descriptions.
