# CLAUDE.md — ATS Desktop App Starter

This project uses the same design system, build pipeline, and auto-updater as the
**Project Tracking Tool** (github.com/JustinGlave/project-tracking-tool).

---

## Tech Stack

- **Python 3** with **PySide6** (Qt for Python) for the UI
- **openpyxl** for any Excel export features
- **PyInstaller** (`--onedir --windowed`) to package as a Windows exe
- **Inno Setup 6** to build the installer (`installer.iss`)
- **GitHub Releases** for distribution and auto-updates (`updater.py`)

---

## File Structure

```
app_gui.py              — Main UI (rename to match app name)
app_backend.py          — Data and storage logic
updater.py              — Auto-update system (edit GITHUB_OWNER/GITHUB_REPO)
version.py              — Version number (bump before every release)
build.bat               — Builds exe + installer + zips
installer.iss           — Inno Setup installer script
PTT_Normal.ico          — App icon (replace with your own)
PTT_Transparent.png     — Watermark/logo (replace with your own)
```

---

## Design System

### Accent Color
`#487cff` — used for selection highlights, hover states on resize handles,
menu item hover, and the install button.

### Font
`Segoe UI, Arial, sans-serif` at `11pt` base. All widgets inherit this.

### Themes
Two themes are built in and toggled via **View → Dark Mode**. The preference
is saved in `QSettings` and restored on next launch.

Both themes use the **Fusion** Qt style as a base, then override the palette
and apply a full QSS stylesheet.

#### Dark Theme palette (key values)
| Role | Value |
|------|-------|
| Window | `#1c1c1c` |
| Base (inputs/lists) | `#121212` |
| Button | `#2d2d2d` |
| Text | `#e6e6e6` |
| Highlight | `#487cff` |

#### Light Theme palette (key values)
| Role | Value |
|------|-------|
| Window | `rgb(210, 212, 218)` — mid-grey, not pure white |
| Base (inputs/lists) | `rgb(225, 227, 232)` |
| Button | `rgb(195, 198, 206)` |
| Text | `rgb(25, 25, 25)` |
| Highlight | `#487cff` |

### Widget Styling Conventions
| Widget | Style |
|--------|-------|
| Panels / cards | `border-radius: 14px`, semi-transparent background, 1px border |
| Buttons | `border-radius: 10px`, `padding: 6px 16px` |
| Inputs (QLineEdit, QComboBox, etc.) | `border-radius: 10px`, `padding: 8px` |
| Tables / lists | `border-radius: 10px`, transparent background |
| Table headers | Bold, 8px padding, no outer border |
| List items | `border-radius: 10px`, `padding: 10px`, `margin: 2px 0` |

### Named Widget IDs (setObjectName)
These are used by the QSS stylesheet — keep names consistent:
| Name | Used for |
|------|----------|
| `Panel` | Section containers / cards |
| `StatCard` | Small stat display cards |
| `ProjectTitle` | 14pt bold heading |
| `ProjectSubtitle` | 10pt muted subtitle |
| `SectionTitle` | 12pt section heading |
| `StatTitle` | 7pt muted label above a stat value |
| `StatValue` | 10pt bold stat value |
| `MetaCaption` | 9pt bold field label |
| `MetaValue` | 9pt field value |
| `ResizeHandle` | Horizontal drag handle between panels |
| `VResizeHandle` | Vertical drag handle |
| `UpdateBanner` | Auto-update banner at bottom of window |
| `UpdateMsg` | Label inside the update banner |
| `InstallBtn` | Green install button inside update banner |

---

## Layout Pattern

```
QMainWindow
└── Central QWidget (QHBoxLayout)
    ├── Left sidebar (QListWidget, fixed ~220px width)
    │   ├── New / action buttons at top
    │   └── Item list
    └── Right main area (QWidget, stretch)
        ├── Header area (title, meta fields)
        ├── Content area (table or stacked views)
        └── Update banner (hidden until update available)
```

---

## Path Helpers

Always use these two helpers so the app works both from source and as a
PyInstaller bundle:

```python
import sys, os
from pathlib import Path

def _resource_path(filename: str) -> str:
    """Path to a bundled asset (icon, image). Works from source and exe."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", ""))
    else:
        base = Path(__file__).parent
    return str(base / filename)

def _app_data_path(filename: str) -> str:
    """Path to user data in %APPDATA%\\<Publisher>\\<AppName>\\."""
    base = Path(os.environ.get("APPDATA", Path.home())) / "ATS Inc" / "YOUR APP NAME"
    base.mkdir(parents=True, exist_ok=True)
    return str(base / filename)
```

---

## Auto-Updater

`updater.py` is a drop-in module. Edit two lines at the top:

```python
GITHUB_OWNER = "JustinGlave"
GITHUB_REPO  = "your-repo-name"
```

The updater looks for a release asset named `<AppName>.zip` containing the exe.
`build.bat` produces this zip automatically.

In the GUI, call `check_for_update()` in a background `QThread` on startup,
then show the `UpdateBanner` widget if an update is found.

---

## Build & Release Workflow

1. Edit code
2. Bump version in `version.py`
3. Run `build.bat` — produces:
   - `dist\<AppName>\<AppName>.exe` — test this first
   - `dist\<AppName>Setup.exe` — installer for new users
   - `dist\<AppName>.zip` — auto-updater asset
4. Test the exe and installer
5. `git add . && git commit -m "v1.x.x - description" && git push`
6. `gh release create v1.x.x --title "v1.x.x" --notes "..."`
7. `gh release upload v1.x.x dist/<AppName>Setup.exe dist/<AppName>.zip`

---

## Installer Notes

- `PrivilegesRequired=lowest` — no admin required
- Installs to `{localappdata}\ATS Inc\<AppName>\`
- User data goes to `{userappdata}\ATS Inc\<AppName>\` (separate from app files)
- Uses `{userdesktop}` for the desktop shortcut (not `{commondesktop}`)
- Uninstaller asks whether to delete user data before removing it
- Version is passed from `build.bat` via `/DMyAppVersion=%VERSION%`
