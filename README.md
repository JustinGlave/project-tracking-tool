# Project Tracking Tool

A desktop application for tracking ATS project tasks, built for the ATS team.

**Current Version: v1.3.0**

---

## What It Does

- Create and manage projects with full job details (PM, SE, contract value, owner, contractor, Div25 URL, etc.)
- Two task templates — **Standard** and **Phoenix** — applied at job creation or swapped any time
- Track tasks by phase with color-coded progress matching the segmented progress bar
- Add notes and change orders to each job
- **Role-based access** — Admin, User, and View Only roles with per-role restrictions
- **Pin projects** to the top of the list with a 📌 indicator
- **Task due dates** — set due dates on tasks; overdue tasks highlight red
- **Drag to reorder tasks** within a project
- **Compact view** — toggle a condensed row layout to see more tasks at once
- **Bulk complete / uncomplete** — mark all visible tasks done or undone in one click (with confirmation)
- **Activity log** — every create, edit, complete, and delete action is logged per project with timestamp and user
- **Bulk Excel export** — select multiple projects and export them all into one formatted workbook
- **Financials Dashboard** — view financial data across all projects in one window, with separate tabs for active jobs, labor hours & cost, and warranty/archived jobs; sortable columns, live search, and totals row
- **Home Dashboard** — landing screen shows project count, overdue tasks, tasks due this week, total tasks, top-5 projects by contract value, 5 most recently added projects, and a live recent-activity feed
- **Task Notes History** — each task carries a timestamped note thread; view and add entries from the right-click context menu ("Notes History")
- **Keyboard shortcuts** — Ctrl+N (new project), Ctrl+T (new task), Ctrl+F (project search), Ctrl+Shift+F (task search), Ctrl+E (export), Ctrl+Shift+E (bulk export), Escape (clear active search)
- Visual segmented progress bar showing completion by phase
- Search and filter tasks by phase or keyword
- Sort the project list by Last Updated, Name, or Job Number — ascending or descending
- **Shared database** — point all users to a shared folder (SharePoint / OneDrive) so everyone works from the same data in real time
- Export a single project to Excel or JSON snapshot
- **Phoenix Controls dark navy UI** — consistent design system across all ATS tools
- Auto-backup on open — keeps the last 10 backups in a `backups/` subfolder
- Auto-updates — when a new version is released, the app notifies you and installs it with one click

---

## Getting Started

### For Users (Installing the App)

1. Download `ProjectTrackingToolSetup.exe` from the [Releases](../../releases) page
2. Run the installer — no admin rights required
3. A desktop shortcut is created automatically
4. Double-click the shortcut to launch

Your project data is stored in `%APPDATA%\ATS Inc\Project Tracking Tool\` and is never touched by updates or reinstalls.

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
2. In the task bar, click the **Templates** dropdown
3. Choose **Standard** or **Phoenix**
4. Confirm the prompt — all current tasks will be replaced with the selected template

> **Note:** This replaces all tasks. Any completed tasks or custom tasks will be lost.

### Home Dashboard

When no project is selected the main panel shows the Home Dashboard:

- **Stat cards** — Projects, Due This Week, Overdue, Total Tasks
- **Top 5 by Contract Value** and **5 Most Recently Added** project tables
- **Recent Activity** feed showing the last 20 actions across all projects

Click any project in the sidebar to open it; click an empty area of the sidebar or press **Escape** to return to the dashboard.

### Adding and Managing Tasks

- Click **Add Task** to add a custom task
- Check the **Done** checkbox on any row to mark a task complete
- **Right-click** any task row for a context menu — **Add Task**, **Edit Task**, **Delete Task**, **Notes History**
- Set a **Due Date** on any task — overdue incomplete tasks highlight red automatically
- **Drag rows** to reorder tasks within the list
- Sort tasks by clicking any column header
- Filter by phase using the **All phases** dropdown, or search by keyword in the **Filter tasks** box
- Use **✓ All** / **✗ All** to bulk complete or uncomplete all currently visible tasks (confirmation required)
- Toggle **Compact** to shrink row height and hide the Notes column for a denser view

### Task Notes History

Each task maintains a timestamped note thread separate from the single inline Notes field:

1. **Right-click** any task row and choose **Notes History**
2. The dialog shows all previous notes in chronological order
3. Type in the input box and click **Add Note** to append a new entry — every entry is stamped with the date, time, and your username
4. When you edit a task and change the Notes field, the new text is automatically appended to the history thread as well

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| **Ctrl+N** | New project |
| **Ctrl+T** | New task |
| **Ctrl+F** | Focus project search |
| **Ctrl+Shift+F** | Focus task search |
| **Ctrl+E** | Export to Excel (current project) |
| **Ctrl+Shift+E** | Bulk export to Excel |
| **Escape** | Clear the active search box |
| **Delete** | Delete selected task |
| **Enter / Return** | Edit selected task |

### Pinning a Project

Click **📌 Pin** (bottom of the sidebar) to pin the selected project to the top of the list. Pinned projects always float above unpinned ones regardless of sort order. Click **📌 Unpin** to release it.

### Adding Notes

1. With a job selected, click **📝 Notes** in the task bar
2. Click **+ Add Note** to create a new note with a date and content
3. Notes can be marked **Open** or **Closed** and include a closeout comment
4. Double-click any note row to edit it

### Adding Change Orders

1. With a job selected, click **🚀 CO Log** in the task bar
2. Click **+ Add CO** to enter a new change order
3. Fields include COP#, description, ATS pricing, sub pricing, and status tracking
4. The summary bar at the top shows running totals for ATS and Sub contracts

### Viewing All Project Details

Click **ℹ️ Info** in the task bar to open a popup showing every field entered for the job — owner, contractor, contract value, warranty, Div25 URL, and more.

### Viewing the Financials Dashboard

Click **📊 All Financials** in the sidebar (visible once a financial data file is configured) to open the dashboard.

The dashboard shows financial data for all projects from your ODIN tracking workbook across four tabs:

| Tab | Contents |
|-----|----------|
| **Financial Overview** | Active jobs — contract value, billed to date, actual cost, booked/actual/differential margins |
| **Labor Hours & Cost** | Active jobs — PM hours, tech hours, total hours, PM cost, tech cost, total labor cost, labor budget remaining, total labor budget |
| **Warranty & Archived** | Jobs with a warranty or archived status — same financial columns as the overview |

- **Search bar** filters all tabs simultaneously by job #, name, PM, or status
- **Click any column header** to sort by that column
- **Double-click any row** to open the full per-project financial detail view
- **Refresh** re-reads the ODIN workbook and updates all tabs
- The totals row at the bottom of each tab sums money columns and averages margin percentages

> Jobs with no name (blank or "0.0") are excluded from all tabs.

### Viewing the Activity Log

Click **📜 Activity** in the task bar to open the activity log for the selected project. Every task creation, edit, completion, and deletion is recorded with a timestamp and the user who made the change.

- **Admin users** see a **Remove** button on each row to delete individual log entries (with confirmation)

### Importing a Job from an Odin Email

1. Click **Import Email** in the sidebar
2. Paste the full text of the Odin assignment email
3. The tool extracts job name, number, PM, SE, contract value, and other fields automatically
4. Review the pre-filled dialog and click **OK**

### Exporting a Project

- **File → Export to Excel (.xlsx)** — generates a formatted Excel report for the selected job
- **File → Export Snapshot (.json)** — saves a full JSON backup of the selected job
- **File → Bulk Export to Excel...** — pick multiple projects and export them all into a single workbook, one set of sheets per project
- The **Export** button in the header also provides single-project options

### Sorting the Project List

Use the sort controls below the search box in the sidebar:

- **Dropdown** — choose to sort by **Last Updated**, **Name**, or **Job Number**
- **↑ A–Z / ↓ Z–A button** — toggle between ascending and descending order

### Setting Up a Shared Database (Multi-User)

All users can share a single database by pointing the app at a synced folder (SharePoint, OneDrive, etc.):

1. The database owner sets up a shared SharePoint/OneDrive folder and invites all users
2. Each user installs OneDrive and syncs the shared folder to their local machine
3. In the app, go to **File → Data Location...**
4. Click **Browse...** and select the local synced folder
5. Click **OK** — if no data file exists there yet, the app will offer to copy your existing data over
6. Repeat steps 3–5 on every user's machine, pointing to their local copy of the same synced folder

> **Note:** The app automatically retries saves if OneDrive briefly locks the data file during sync. Conflicts are rare on a small team but can occur if two users save simultaneously.

### User Accounts and Roles

Accounts are managed by an admin via **File → Manage Users...**

| Role | Can do |
|------|--------|
| **Admin** | Everything — create/edit/delete projects and tasks, manage users, remove activity log entries |
| **User** | Create/edit/delete projects and tasks; cannot manage user accounts |
| **View Only** | Read-only access — can view all data but cannot make any changes |

Use **File → Change My Password...** to update your own password at any time.

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
financials_models.py         — FinancialSnapshot dataclass
financials_excel.py          — ODIN Excel / JSON snapshot provider
financials_dialog.py         — Per-project financial detail dialog
financials_dashboard.py      — All-projects financials dashboard dialog
user_auth.py                 — User account and authentication system
updater.py                   — Auto-update system
version.py                   — Current version number
phoenix_style.qss            — Phoenix Controls unified QSS stylesheet
build.bat                    — Builds the exe, installer, and zips (developers only)
installer.iss                — Inno Setup installer script (developers only)
PTT_Normal.ico               — App icon
PTT_Transparent.png          — Watermark image
```

---

## For Developers — Releasing an Update

1. Make your code changes
2. Bump the version in `version.py`
3. Run `build.bat` — this produces:
   - `dist\ProjectTrackingTool\ProjectTrackingTool.exe` — test this first
   - `dist\ProjectTrackingToolSetup.exe` — installer for new users
   - `dist\ProjectTrackingTool.zip` — exe only, used by the auto-updater
   - `dist\ProjectTrackingTool_FullInstall.zip` — full folder, for manual installs
4. Test the exe and the installer
5. Push changes to GitHub:
   ```
   git add .
   git commit -m "v1.x.x - description of changes"
   git push
   ```
6. Create a GitHub release:
   ```
   gh release create v1.x.x --title "v1.x.x" --notes "Description of what changed"
   ```
7. Upload the release assets:
   ```
   gh release upload v1.x.x dist/ProjectTrackingToolSetup.exe dist/ProjectTrackingTool.zip
   ```

Users will see an update banner in the app automatically on next launch. The banner downloads `ProjectTrackingTool.zip` and replaces the exe in-place.

---

## Built With

- Python 3
- PySide6 (Qt for Python)
- Phoenix Controls Unified Design System (dark navy QSS theme)
- openpyxl (Excel export)
- PyInstaller (exe packaging)
- Inno Setup 6 (installer)
