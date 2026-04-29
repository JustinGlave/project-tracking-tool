"""
updater.py — GitHub-based auto-updater for Project Tracking Tool.

How it works
------------
1. On startup the GUI calls check_for_update() in a background thread.
2. That function hits the GitHub Releases API and compares the latest tag
   against the local __version__ string.
3. If a newer version exists it returns an UpdateInfo object; the GUI shows
   a banner with an "Install & Restart" button.
4. When the user clicks the button, download_and_apply() is called:
      a. Downloads the auto-updater .zip to a temp file.
      b. Validates that the zip contains the full PyInstaller one-folder app.
      c. Writes a small PowerShell updater plus a .bat wrapper that waits for
         this process to exit, extracts the whole app folder over the install
         folder, then relaunches it.
      d. Launches the .bat and calls sys.exit() — Windows takes it from there.

Configuration
-------------
Set GITHUB_OWNER and GITHUB_REPO to match your GitHub account and repository.
The updater looks for ProjectTrackingTool.zip, falling back to a non-full-install .zip asset.
"""

from __future__ import annotations

import os
import sys
import subprocess
import tempfile
import urllib.request
import urllib.error
import json
import logging
import zipfile
from dataclasses import dataclass
from typing import Optional
from pathlib import Path

from version import __version__

logger = logging.getLogger(__name__)

# ── CHANGE THESE to match your GitHub account / repo name ─────────────────────
GITHUB_OWNER = "JustinGlave"
GITHUB_REPO  = "project-tracking-tool"
# ──────────────────────────────────────────────────────────────────────────────

RELEASES_API = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
)
REQUEST_TIMEOUT = 8  # seconds


@dataclass
class UpdateInfo:
    current_version: str
    latest_version:  str
    download_url:    str
    release_notes:   str


class UpdatePackageError(RuntimeError):
    """Raised when the downloaded update package is missing required files."""


def _parse_version(tag: str) -> tuple[int, ...]:
    """Convert 'v1.2.3', 'V1.2.3', or '1.2.3' to (1, 2, 3) for comparison."""
    cleaned = tag.lstrip("vV").strip()
    try:
        return tuple(int(part) for part in cleaned.split("."))
    except ValueError:
        return (0,)


def check_for_update() -> Optional[UpdateInfo]:
    """
    Query the GitHub Releases API.
    Returns an UpdateInfo if a newer version is available, otherwise None.
    Safe to call from a background thread — never raises, logs errors instead.
    """
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={"Accept": "application/vnd.github+json",
                     "User-Agent": "ProjectTrackingTool"},
        )
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = json.loads(resp.read().decode())

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None

        if _parse_version(latest_tag) <= _parse_version(__version__):
            return None  # already up to date

        # Find the update zip asset (not the full install zip)
        assets = data.get("assets", [])
        exe_asset = next(
            (a for a in assets
             if a.get("name", "").lower() == "projecttrackingtool.zip"),
            None,
        )
        # Fallback: any zip that isn't the full install
        if exe_asset is None:
            exe_asset = next(
                (a for a in assets
                 if a.get("name", "").lower().endswith(".zip")
                 and "fullinstall" not in a.get("name", "").lower()),
                None,
            )
        if exe_asset is None:
            logger.warning("New release %s found but no .zip asset attached.", latest_tag)
            return None

        return UpdateInfo(
            current_version = __version__,
            latest_version  = latest_tag.lstrip("vV"),
            download_url    = exe_asset["browser_download_url"],
            release_notes   = data.get("body", "").strip(),
        )

    except urllib.error.URLError as exc:
        logger.debug("Update check failed (network): %s", exc)
        return None
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, AttributeError) as exc:
        logger.warning("Update check failed: %s", exc)
        return None


def _validate_update_zip(zip_path: Path) -> None:
    """Ensure the updater package contains a full one-folder PyInstaller build."""
    try:
        with zipfile.ZipFile(zip_path) as zf:
            names = {name.replace("\\", "/").lstrip("/") for name in zf.namelist()}
    except zipfile.BadZipFile as exc:
        raise UpdatePackageError(
            "The downloaded update package is not a valid zip file.\n"
            "Please download the installer manually from GitHub."
        ) from exc

    flat_exe = "ProjectTrackingTool.exe"
    nested_exe = "ProjectTrackingTool/ProjectTrackingTool.exe"
    has_exe = flat_exe in names or nested_exe in names
    has_internal = any(
        name.startswith("_internal/") or name.startswith("ProjectTrackingTool/_internal/")
        for name in names
    )

    if not has_exe:
        raise UpdatePackageError(
            "The downloaded update package does not contain ProjectTrackingTool.exe.\n"
            "Please download the installer manually from GitHub."
        )
    if not has_internal:
        raise UpdatePackageError(
            "The downloaded update package is incomplete: the _internal runtime folder is missing.\n"
            "Please download the installer manually from GitHub."
        )


def _ps_literal(value: Path | str) -> str:
    """Return a PowerShell single-quoted string literal."""
    return "'" + str(value).replace("'", "''") + "'"


def _build_update_powershell_script(zip_path: Path, install_dir: Path, exe_path: Path) -> str:
    """Build the PowerShell script that performs the file replacement."""
    return f"""$ErrorActionPreference = 'Stop'
$zipPath = {_ps_literal(zip_path)}
$installDir = {_ps_literal(install_dir)}
$exePath = {_ps_literal(exe_path)}
$logPath = Join-Path $env:TEMP 'ProjectTrackingTool_update.log'

"Starting update from $zipPath" | Out-File -FilePath $logPath -Encoding utf8
if (-not (Test-Path -LiteralPath $zipPath)) {{
    throw "Update package was not found: $zipPath"
}}
if (-not (Test-Path -LiteralPath $installDir)) {{
    throw "Install folder was not found: $installDir"
}}

$stage = Join-Path ([IO.Path]::GetTempPath()) ('ptt_update_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $stage | Out-Null

try {{
    Expand-Archive -LiteralPath $zipPath -DestinationPath $stage -Force
    $payload = $stage
    $nested = Join-Path $stage 'ProjectTrackingTool'
    if (Test-Path -LiteralPath (Join-Path $nested 'ProjectTrackingTool.exe')) {{
        $payload = $nested
    }}

    if (-not (Test-Path -LiteralPath (Join-Path $payload 'ProjectTrackingTool.exe'))) {{
        throw "Update package did not contain ProjectTrackingTool.exe."
    }}
    if (-not (Test-Path -LiteralPath (Join-Path $payload '_internal'))) {{
        throw "Update package did not contain the _internal runtime folder."
    }}

    Get-ChildItem -LiteralPath $payload -Force | Copy-Item -Destination $installDir -Recurse -Force

    if (-not (Test-Path -LiteralPath $exePath)) {{
        throw "Updated executable was not found after copy: $exePath"
    }}

    "Update files copied successfully." | Out-File -FilePath $logPath -Append -Encoding utf8
}}
finally {{
    if (Test-Path -LiteralPath $stage) {{
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }}
    if (Test-Path -LiteralPath $zipPath) {{
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }}
}}
"""


def _build_update_batch(pid: int, ps_path: Path, exe_path: Path) -> str:
    ps_str = str(ps_path)
    exe_str = str(exe_path)
    return f"""@echo off
setlocal
set "LOG=%TEMP%\\ProjectTrackingTool_update.log"
echo Waiting for Project Tracking Tool to close... > "%LOG%"
:wait
tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)
powershell -NoProfile -ExecutionPolicy Bypass -File "{ps_str}" >> "%LOG%" 2>&1
if errorlevel 1 (
    echo Update failed. See "%LOG%" for details. >> "%LOG%"
    start "" "{exe_str}"
    del "{ps_str}" >nul 2>nul
    del "%~f0"
    exit /b 1
)
start "" "{exe_str}"
del "{ps_str}" >nul 2>nul
del "%~f0"
"""


def download_and_apply(info: UpdateInfo, progress_callback=None) -> None:
    """
    Download the new zip, extract it over the current install, and restart.

    progress_callback(bytes_done, total_bytes) is called during download
    so the GUI can show a progress bar. Pass None to skip.

    Raises RuntimeError if anything goes wrong so the caller can show
    an error dialog rather than silently failing.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "Update can only be applied to a compiled build.\n"
            "You're running from source, so use git pull/build locally or download the installer from GitHub."
        )

    current_exe = Path(sys.executable).resolve()
    install_dir = current_exe.parent

    # Download zip to system temp
    tmp_fd, tmp_zip_str = tempfile.mkstemp(suffix=".zip")
    tmp_zip = Path(tmp_zip_str)

    try:
        req = urllib.request.Request(
            info.download_url,
            headers={"User-Agent": "ProjectTrackingTool"},
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            total = int(resp.headers.get("Content-Length", 0))
            done  = 0
            chunk = 64 * 1024
            with open(tmp_fd, "wb") as fh:
                while True:
                    block = resp.read(chunk)
                    if not block:
                        break
                    fh.write(block)
                    done += len(block)
                    if progress_callback:
                        progress_callback(done, total)

        # Verify download is complete
        if total > 0 and tmp_zip.stat().st_size < total:
            tmp_zip.unlink(missing_ok=True)
            raise RuntimeError(
                f"Download incomplete: got {tmp_zip.stat().st_size} of {total} bytes.\n"
                "Please try again or download manually from GitHub."
            )

    except RuntimeError:
        raise
    except (OSError, urllib.error.URLError, ValueError) as exc:
        try:
            tmp_zip.unlink(missing_ok=True)
        except OSError:
            logger.exception("Failed to remove incomplete update download: %s", tmp_zip)
        raise RuntimeError(f"Download failed: {exc}") from exc

    try:
        _validate_update_zip(tmp_zip)
    except RuntimeError:
        try:
            tmp_zip.unlink(missing_ok=True)
        except OSError:
            logger.exception("Failed to remove invalid update download: %s", tmp_zip)
        raise

    # Write scripts that wait for this process to exit, extract the full app
    # folder over the install dir, then relaunch.
    pid = os.getpid()
    ps_fd, ps_path_str = tempfile.mkstemp(suffix=".ps1")
    bat_fd, bat_path_str = tempfile.mkstemp(suffix=".bat")
    ps_path = Path(ps_path_str)
    bat_path = Path(bat_path_str)

    with open(ps_fd, "w", encoding="utf-8") as fh:
        fh.write(_build_update_powershell_script(tmp_zip, install_dir, current_exe))
    with open(bat_fd, "w", encoding="utf-8") as fh:
        bat_content = _build_update_batch(pid, ps_path, current_exe)
        fh.write(bat_content)

    subprocess.Popen(
        ["cmd.exe", "/c", str(bat_path)],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        close_fds=True,
    )
    sys.exit(0)
