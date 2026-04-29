"""
updater.py — GitHub-based auto-updater.

How it works
------------
1. On startup the GUI calls check_for_update() in a background thread.
2. That function hits the GitHub Releases API and compares the latest tag
   against the local __version__ string.
3. If a newer version exists it returns an UpdateInfo object; the GUI shows
   a banner with an "Install & Restart" button.
4. When the user clicks the button, download_and_apply() is called:
      a. Downloads the new .zip to a temp file.
      b. Extracts the exe over the current install via a batch script.
      c. Launches the batch and calls sys.exit() — Windows takes it from there.

Configuration
-------------
Set GITHUB_OWNER and GITHUB_REPO to match your GitHub account and repository.
Set ZIP_ASSET_NAME to match the zip file uploaded to your GitHub release
(this is the auto-updater zip produced by build.bat, not the full install zip).
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
from dataclasses import dataclass
from typing import Optional
from pathlib import Path

from version import __version__

logger = logging.getLogger(__name__)

# ── CHANGE THESE ──────────────────────────────────────────────────────────────
GITHUB_OWNER   = "JustinGlave"          # your GitHub username
GITHUB_REPO    = "your-repo-name"       # your repository name
ZIP_ASSET_NAME = "YourAppName.zip"      # the auto-updater zip asset name
EXE_NAME       = "YourAppName.exe"      # the exe inside the zip
# ──────────────────────────────────────────────────────────────────────────────

RELEASES_API   = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT = 8


@dataclass
class UpdateInfo:
    current_version: str
    latest_version:  str
    download_url:    str
    release_notes:   str


def _parse_version(tag: str) -> tuple[int, ...]:
    cleaned = tag.lstrip("vV").strip()
    try:
        return tuple(int(part) for part in cleaned.split("."))
    except ValueError:
        return (0,)


def check_for_update() -> Optional[UpdateInfo]:
    """
    Query the GitHub Releases API.
    Returns UpdateInfo if a newer version is available, otherwise None.
    Safe to call from a background thread — never raises.
    """
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={"Accept": "application/vnd.github+json", "User-Agent": EXE_NAME},
        )
        with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
            data = json.loads(resp.read().decode())

        latest_tag = data.get("tag_name", "")
        if not latest_tag:
            return None
        if _parse_version(latest_tag) <= _parse_version(__version__):
            return None

        assets = data.get("assets", [])
        zip_asset = next(
            (a for a in assets if a.get("name", "").lower() == ZIP_ASSET_NAME.lower()),
            None,
        )
        if zip_asset is None:
            logger.warning("New release %s found but asset %s not attached.", latest_tag, ZIP_ASSET_NAME)
            return None

        return UpdateInfo(
            current_version=__version__,
            latest_version=latest_tag.lstrip("vV"),
            download_url=zip_asset["browser_download_url"],
            release_notes=data.get("body", "").strip(),
        )

    except urllib.error.URLError as exc:
        logger.debug("Update check failed (network): %s", exc)
        return None
    except (json.JSONDecodeError, KeyError, TypeError, ValueError, AttributeError) as exc:
        logger.warning("Update check failed: %s", exc)
        return None


def download_and_apply(info: UpdateInfo, progress_callback=None) -> None:
    """
    Download the new zip, extract the exe over the current install, and restart.

    progress_callback(bytes_done, total_bytes) is called during download.
    Raises RuntimeError on failure so the caller can show an error dialog.
    """
    if not getattr(sys, "frozen", False):
        raise RuntimeError(
            "Update can only be applied to a compiled build.\n"
            "You're running from source — pull the latest code from GitHub instead."
        )

    current_exe = Path(sys.executable).resolve()

    tmp_fd, tmp_zip_str = tempfile.mkstemp(suffix=".zip")
    tmp_zip = Path(tmp_zip_str)

    try:
        req = urllib.request.Request(
            info.download_url,
            headers={"User-Agent": EXE_NAME},
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

        if total > 0 and tmp_zip.stat().st_size < total:
            tmp_zip.unlink(missing_ok=True)
            raise RuntimeError("Download incomplete. Please try again.")

    except RuntimeError:
        raise
    except (OSError, urllib.error.URLError, ValueError) as exc:
        try:
            tmp_zip.unlink(missing_ok=True)
        except OSError:
            logger.exception("Failed to remove incomplete update download: %s", tmp_zip)
        raise RuntimeError(f"Download failed: {exc}") from exc

    pid = os.getpid()
    bat_fd, bat_path_str = tempfile.mkstemp(suffix=".bat")
    bat_path = Path(bat_path_str)
    exe_str = str(current_exe)
    zip_str = str(tmp_zip)
    bat_content = f"""@echo off
:wait
tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)
powershell -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.IO.Compression.FileSystem; $zip = [System.IO.Compression.ZipFile]::OpenRead('{zip_str}'); $entry = $zip.Entries | Where-Object {{ $_.Name -eq '{EXE_NAME}' }} | Select-Object -First 1; if ($entry) {{ [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, '{exe_str}', $true) }}; $zip.Dispose()"
del "{zip_str}"
start "" "{exe_str}"
del "%~f0"
"""
    with open(bat_fd, "w") as fh:
        fh.write(bat_content)

    subprocess.Popen(
        ["cmd.exe", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NO_WINDOW,
        close_fds=True,
    )
    sys.exit(0)
