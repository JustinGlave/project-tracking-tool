# Project Tracking Tool — Setup & Update Guide

## What You'll End Up With

- A `ProjectTrackingTool.exe` your team can run with no Python installed
- Every time you push a new version to GitHub, users see an in-app
  "Update Available" banner and can install it with one click
- You never have to manually send `.exe` files around again

---

## Part 1 — One-Time Setup (do this once)

### Step 1 — Create a GitHub Account

1. Go to **https://github.com** and click **Sign up**
2. Choose a username (e.g. `justinglave`), enter your email, create a password
3. Verify your email address

### Step 2 — Create a Repository

1. After logging in, click the **+** icon (top right) → **New repository**
2. Fill in:
   - **Repository name:** `project-tracking-tool`
   - **Description:** Project Tracking Tool desktop app
   - **Visibility:** Private ✓ (only you can see it)
   - Check **Add a README file**
3. Click **Create repository**

### Step 3 — Update the Code with Your GitHub Details

Open `updater.py` and change these two lines at the top:

```python
GITHUB_OWNER = "YOUR_GITHUB_USERNAME"   # e.g. "justinglave"
GITHUB_REPO  = "project-tracking-tool"
```

Replace `YOUR_GITHUB_USERNAME` with the username you just created.

### Step 4 — Install Required Tools

Open a Command Prompt in your project folder and run:

```
pip install pyinstaller
```

### Step 5 — Install Git (to upload code to GitHub)

1. Download from **https://git-scm.com/download/win**
2. Install with all defaults
3. Open a new Command Prompt and verify: `git --version`

### Step 6 — Upload Your Code to GitHub

In Command Prompt, navigate to your project folder:

```
cd C:\Users\YourName\PycharmProjects\Job Tracker
```

Then run these commands one by one:

```
git init
git add .
git commit -m "Initial release v1.0.0"
git branch -M main
git remote add origin https://github.com/JustinGlave/project-tracking-tool.git
git push -u origin main
```

When prompted, sign in with your GitHub username and password.

---

## Part 2 — Building Your First Executable

### Step 1 — Check Your File List

Make sure all these files are in your project folder:

```
project_tracker_gui.py
project_tracker_backend.py
updater.py
version.py
build.bat
PTT_Normal.ico
PTT_Transparent.png
```

### Step 2 — Run the Build Script

Double-click `build.bat` (or run it from Command Prompt).

It will create:
```
dist\
  ProjectTrackingTool.exe   ← this is what you distribute
```

Test it by double-clicking the exe — make sure everything works.

### Step 3 — Create Your First GitHub Release

1. Go to your repository on GitHub
2. Click **Releases** (right side of the page) → **Draft a new release**
3. Click **Choose a tag** → type `v1.0.0` → click **Create new tag: v1.0.0**
4. Set **Release title** to `v1.0.0`
5. Write a description (e.g. "Initial release")
6. Drag and drop `dist\ProjectTrackingTool.exe` into the upload box
7. Click **Publish release**

Done! The app is now published. Give the `.exe` to your team.

---

## Part 3 — Releasing Updates (every future update)

This is your workflow every time you make changes:

### Step 1 — Make Your Code Changes

Edit whatever files need changing.

### Step 2 — Bump the Version Number

Open `version.py` and increment the version:

```python
# Was:
__version__ = "1.0.0"

# Change to (bug fix):
__version__ = "1.0.1"

# Or for new features:
__version__ = "1.1.0"
```

**Version numbering guide:**
- `1.0.0` → `1.0.1` — Bug fix or small tweak
- `1.0.0` → `1.1.0` — New feature added
- `1.0.0` → `2.0.0` — Major redesign

### Step 3 — Build the New Executable

Double-click `build.bat`. Test `dist\ProjectTrackingTool.exe`.

### Step 4 — Push Code to GitHub

```
git add .
git commit -m "v1.0.1 - fixed watermark visibility"
git push
```

### Step 5 — Create a New GitHub Release

1. Go to your repository → **Releases** → **Draft a new release**
2. Tag: `v1.0.1`
3. Title: `v1.0.1`
4. Write what changed in the description box — this shows in the
   **"What's New?"** popup your users see in the app
5. Upload `dist\ProjectTrackingTool.exe`
6. Click **Publish release**

### What Happens for Your Users

Next time they open the app, they'll see a green banner at the bottom:

> 🆕 Update available — v1.0.1 is ready. You're on v1.0.0.
> [What's New?]  [Install & Restart]  [✕]

They click **Install & Restart**, the new exe downloads and replaces
the old one, and the app restarts automatically. No IT involvement,
no file sharing.

---

## Quick Reference

| Task | Command |
|------|---------|
| Build exe | Double-click `build.bat` |
| Push code to GitHub | `git add . && git commit -m "message" && git push` |
| Check current version | Look at `version.py` |
| See what version the exe is | Help → About (or check the title bar) |

---

## Troubleshooting

**"git is not recognized"**
→ Restart Command Prompt after installing Git.

**"pyinstaller is not recognized"**
→ Run `pip install pyinstaller` again and restart Command Prompt.

**Build fails with import errors**
→ Make sure all `.py` files and asset files are in the same folder.

**Update banner never appears**
→ Check that `GITHUB_OWNER` and `GITHUB_REPO` in `updater.py` exactly
  match your GitHub username and repository name (case-sensitive).
→ Check that you published the release (not left it as Draft) on GitHub.
→ Check that you uploaded the `.exe` as a release asset.

**"Update can only be applied to a compiled .exe build"**
→ This is expected when running from PyCharm/source. Updates only work
  on the built `.exe` — this is intentional.
