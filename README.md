# Project Tracking Tool

A desktop application for tracking Phoenix project tasks, built for the ATS team.

---

## What It Does

- Create and manage projects imported from the Phoenix Job Tracking workbook
- Track tasks by phase with color-coded progress
- Visual segmented progress bar showing completion by phase
- Search and filter tasks by phase or keyword
- Export project snapshots to JSON
- Auto-updates — when a new version is released, the app notifies you and installs it with one click

---

## Getting Started

### For Users (Running the App)

1. Download the latest `ProjectTrackingTool.exe` from the [Releases](../../releases) page
2. Place it in a folder of your choice (e.g. `C:\Tools\ProjectTracker\`)
3. Copy your asset files into the same folder:
   - `PTT_Normal.ico`
   - `PTT_Transparent.png`
4. Double-click `ProjectTrackingTool.exe` to launch

No Python or other software required.

---

## File Structure

```
project_tracker_gui.py       — Main UI
project_tracker_backend.py   — Data and storage logic
updater.py                   — Auto-update system
version.py                   — Current version number
build.bat                    — Builds the .exe (developers only)
PTT_Normal.ico               — App icon
PTT_Transparent.png          — Watermark image
```

---

## For Developers — Releasing an Update

1. Make your code changes
2. Bump the version in `version.py`
3. Run `build.bat` to build a new `dist\ProjectTrackingTool.exe`
4. Test the exe
5. Push changes to GitHub:
   ```
   git add .
   git commit -m "v1.x.x - description of changes"
   git push
   ```
6. Go to **Releases → Draft a new release** on GitHub
7. Tag it `v1.x.x`, write release notes, upload the exe, publish

Users will see an update banner in the app automatically on next launch.

---

## Built With

- Python 3
- PySide6 (Qt for Python)
- openpyxl (Excel import)
- PyInstaller (exe packaging)
