# Project Tracking Tool

A desktop application for tracking ATS project tasks, built for the ATS team.

**Current Version: v1.0.15**

---

## What It Does

- Create and manage projects with full job details (PM, SE, contract value, owner, etc.)
- Two task templates — **Standard** and **Phoenix** — applied at job creation or swapped any time
- Track tasks by phase with color-coded progress matching the segmented progress bar
- Add notes and change orders to each job
- Visual segmented progress bar showing completion by phase
- Search and filter tasks by phase or keyword
- Export projects to Excel or JSON snapshot
- Dark mode / light mode toggle with preference saved across sessions
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

## How-To Guides

### Creating a New Job

1. Click **New** in the sidebar (or **File → New Project**)
2. Select a **Task Template**:
   - **Standard** — full default task list
   - **Phoenix** — streamlined list for Phoenix jobs (excludes tasks not applicable)
3. Fill in job details (Job Name and Job Number are required)
4. Click **OK** — the job is created with the selected task list pre-loaded

### Changing a Job's Task Template After Creation

If you selected the wrong template when creating a job, you can reset it:

1. Select the job in the sidebar
2. In the task bar, click the **Templates** dropdown (next to the Filter Tasks box)
3. Choose **Standard** or **Phoenix**
4. Confirm the prompt — all current tasks will be replaced with the selected template

> **Note:** This replaces all tasks. Any completed tasks or custom tasks will be lost.

### Adding and Managing Tasks

- Click **Add Task** (in the task bar between All Phases and Filter Tasks) to add a custom task
- Check the **Done** checkbox on any row to mark a task complete
- Use **Edit** / **Del** buttons on each row to modify or remove individual tasks
- Sort tasks by clicking any column header
- Filter by phase using the **All Phases** dropdown, or search by keyword in the **Filter Tasks** box

### Adding Notes

1. With a job selected, click **Notes** in the task bar
2. Click **+ Add Note** to create a new note with a date and content
3. Notes can be marked **Open** or **Closed** and include a closeout comment
4. Double-click any note row to edit it

### Adding Change Orders

1. With a job selected, click **Change Orders** in the task bar
2. Click **+ Add CO** to enter a new change order
3. Fields include COP#, description, ATS pricing, sub pricing, and status tracking
4. The summary bar at the top shows running totals:
   - **ATS Base** — sum of all ATS Price entries
   - **ATS Current** — Base Price plus all Accepted change orders

### Viewing All Project Details

Click **Project Info** (next to Change Orders in the task bar) to open a popup showing every field entered for the job — owner, contractor, contract value, warranty, Div25 URL, and more.

### Importing a Job from an Odin Email

1. Click **Import Email** in the sidebar
2. Paste the full text of the Odin assignment email
3. The tool extracts job name, number, PM, SE, contract value, and other fields automatically
4. Review the pre-filled dialog and click **OK**

### Exporting a Project

- **File → Export to Excel (.xlsx)** — generates a formatted Excel report for the selected job
- **File → Export Snapshot (.json)** — saves a full JSON backup of the selected job
- The **Export** button in the header also provides both options

### Using Dark Mode / Light Mode

Go to **View → Dark Mode** to toggle between dark and light themes. Your preference is saved and restored on next launch.

### Training Mode — Test Jobs

Use **Help → Show Test Jobs** to load 5 pre-built demo jobs covering a range of scenarios:

| Job | Status | Template | Highlights |
|-----|--------|----------|------------|
| PNNL - Building 3000 Controls | Early stage | Standard | 2 notes, 1 approved CO |
| Hanford Site - HVAC Controls | Mid-progress | Phoenix | 3 notes, 1 approved + 1 pending CO |
| Boeing Renton - Building 4-20 BAS | Nearly complete | Standard | 3 notes, 3 approved COs |
| Microsoft Campus - Lab 7 Automation | Just started | Standard | 1 note, 1 pending CO |
| Richland Schools - District HVAC | Fully closed out | Phoenix | 3 notes, 2 approved COs |

Click **Help → Hide Test Jobs** to remove them from the sidebar. Test jobs are never shown by default and do not interfere with real project data.

---

## File Structure

```
project_tracker_gui.py       — Main UI
project_tracker_backend.py   — Data and storage logic
updater.py                   — Auto-update system
version.py                   — Current version number
build.bat                    — Builds the .exe and install zip (developers only)
PTT_Normal.ico               — App icon
PTT_Transparent.png          — Watermark image
```

---

## For Developers — Releasing an Update

1. Make your code changes
2. Bump the version in `version.py`
3. Run `build.bat` to build `dist\ProjectTrackingTool.exe` and the install zip
4. Test the exe
5. Push changes to GitHub:
   ```
   git add .
   git commit -m "v1.x.x - description of changes"
   git push
   ```
6. Create a release: `gh release create v1.x.x --title "v1.x.x" --notes "..."`
7. Upload the exe and zip to the release

Users will see an update banner in the app automatically on next launch.

---

## Built With

- Python 3
- PySide6 (Qt for Python)
- openpyxl (Excel export)
- PyInstaller (exe packaging)
