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
      a. Downloads the new .exe to a temp file.
      b. Writes a tiny .bat script that waits for this process to exit,
         overwrites the exe, then relaunches it.
      c. Launches the .bat and calls sys.exit() — Windows takes it from there.

Configuration
-------------
Set GITHUB_OWNER and GITHUB_REPO to match your GitHub account and repository.
The updater looks for a release asset whose name ends with .exe.
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

        # Find the .zip asset
        assets = data.get("assets", [])
        exe_asset = next(
            (a for a in assets if a.get("name", "").lower().endswith(".zip")),
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
    except Exception as exc:
        logger.warning("Update check failed: %s", exc)
        return None


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
            "You're running from source — pull the latest code from GitHub instead."
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
    except Exception as exc:
        tmp_zip.unlink(missing_ok=True)
        raise RuntimeError(f"Download failed: {exc}") from exc

    # Write batch script that waits for this process to exit,
    # extracts the zip over the install dir, then relaunches.
    pid = os.getpid()
    bat_fd, bat_path_str = tempfile.mkstemp(suffix=".bat")
    bat_path = Path(bat_path_str)
    bat_content = f"""@echo off
:wait
tasklist /FI "PID eq {pid}" 2>nul | find "{pid}" >nul
if not errorlevel 1 (
    timeout /t 1 /nobreak >nul
    goto wait
)
powershell -Command "Expand-Archive -Path '{tmp_zip}' -DestinationPath '{install_dir}' -Force"
del "{tmp_zip}"
start "" "{current_exe}"
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