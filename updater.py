"""
updater.py — GitHub-based auto-updater for Project Tracking Tool.

Wave 8b B3 hybrid facade. As of 2026-05-28, this module delegates the
generic update-check + full-folder download/apply path to
``phoenix_commons.updater`` while preserving Job-Tracker-specific
naming constants and the regression-test surface
(``UpdatePackageError`` / ``_validate_update_zip`` /
``_build_update_powershell_script``) that ``tests/test_regressions.py``
imports by name.

ADR-003 production payload asymmetry — Project Tracking Tool ships a
**full-folder** updater payload. ``download_and_apply`` therefore calls
commons with ``expected_internal=True`` (commons' default); commons
handles the zip validation (checking ``<EXE_NAME>`` exists at zip root
or inside a top-level folder named after the exe stem, AND
``_internal/`` runtime folder is present) and uses its PowerShell +
batch wrapper to extract the full folder over the install directory,
then relaunch.

How it works
------------
1. On startup the GUI calls ``check_for_update()`` in a background
   thread.
2. That function hits the GitHub Releases API (via commons) and
   compares the latest tag against the local ``__version__`` string.
3. If a newer version exists it returns an ``UpdateInfo`` object; the
   GUI shows a banner with an "Install & Restart" button.
4. When the user clicks the button, ``download_and_apply()`` is called.
   commons downloads, validates, and runs a PowerShell-driven
   full-folder replacement before relaunching.

Preserved-local logic (MIGRATION_RULES § 1 hybrid facade)
---------------------------------------------------------
- ``GITHUB_OWNER`` / ``GITHUB_REPO`` — naming constants passed to the
  commons facade.

- ``UpdatePackageError``: re-exported from
  ``phoenix_commons.updater.installer.UpdatePackageError`` (identity
  preserved) — kept available as ``updater.UpdatePackageError`` because
  ``tests/test_regressions.py`` imports it by name.

- ``_validate_update_zip`` / ``_build_update_powershell_script``: kept
  at module level for the ``tests/test_regressions.py`` regression
  baseline. commons has equivalent private helpers
  (``_validate_update_zip`` / ``_build_full_folder_powershell``) but
  with different signatures; keeping ours local means the test contract
  doesn't reach into commons internals.

- ``_parse_version`` / ``_ps_literal`` / ``_build_update_batch``: kept
  local — used by the preserved helpers above. commons has internal
  equivalents but they're private.

UpdateInfo + UpdatePackageError identity contract
-------------------------------------------------
``updater.UpdateInfo is phoenix_commons.updater.UpdateInfo`` and
``updater.UpdatePackageError is phoenix_commons.updater.installer.UpdatePackageError``
— both verified by the B3 commit's identity check. Callers that did
``from updater import UpdateInfo`` or ``from updater import
UpdatePackageError`` continue to work unchanged.
"""

from __future__ import annotations

import logging
import zipfile
from pathlib import Path
from typing import Optional

# Commons facade imports — these provide the generic update-check and
# full-folder download/apply implementation. UpdateInfo + UpdatePackageError
# are re-exported here (identity preserved) so callers that do
# ``from updater import UpdateInfo`` or ``from updater import UpdatePackageError``
# don't change.
from phoenix_commons.updater import UpdateInfo
from phoenix_commons.updater import check_for_update as _commons_check_for_update
from phoenix_commons.updater import download_and_apply as _commons_download_and_apply
from phoenix_commons.updater.installer import UpdatePackageError

from version import __version__

logger = logging.getLogger(__name__)

# ── Project Tracking Tool release contract ──────────────────────────────────
GITHUB_OWNER = "JustinGlave"
GITHUB_REPO = "project-tracking-tool"
EXE_NAME = "ProjectTrackingTool.exe"
ZIP_ASSET_NAME = "ProjectTrackingTool.zip"
# ────────────────────────────────────────────────────────────────────────────

__all__ = [
    "UpdateInfo",
    "UpdatePackageError",
    "check_for_update",
    "download_and_apply",
    "GITHUB_OWNER",
    "GITHUB_REPO",
    "EXE_NAME",
    "ZIP_ASSET_NAME",
    # Helpers below are preserved-local for the
    # tests/test_regressions.py regression baseline. commons has
    # equivalent private helpers; ours stay independently exercised.
    "_parse_version",
    "_validate_update_zip",
    "_build_update_powershell_script",
]


# ─── Preserved-local helpers (test surface) ──────────────────────────────────

def _parse_version(tag: str) -> Optional[tuple[int, ...]]:
    """Convert ``'v1.2.3'``, ``'V1.2.3'``, or ``'1.2.3'`` to ``(1, 2, 3)``.

    Returns ``None`` if the tag is empty or unparseable so callers can
    skip the comparison rather than treating the version as ``(0,)`` —
    which would incorrectly suppress every update check whenever the
    local ``__version__`` is also unparseable.

    Preserved-local for ``tests/test_regressions.py`` (commons has its
    own private ``_parse_version`` in ``phoenix_commons.updater.client``
    with slightly different fail-soft semantics — ``(0,)`` vs
    ``None``).
    """
    cleaned = tag.lstrip("vV").strip()
    if not cleaned:
        return None
    try:
        return tuple(int(part) for part in cleaned.split("."))
    except ValueError:
        return None


def _validate_update_zip(zip_path: Path) -> None:
    """Ensure the updater package contains a full one-folder PyInstaller build.

    Preserved-local for ``tests/test_regressions.py`` — commons has
    ``phoenix_commons.updater.installer._validate_update_zip(zip_path,
    exe_name, *, expected_internal=True)`` with a different signature;
    keeping ours local means tests don't need to learn the commons
    signature.
    """
    try:
        with zipfile.ZipFile(zip_path) as zf:
            names = {name.replace("\\", "/").lstrip("/") for name in zf.namelist()}
    except zipfile.BadZipFile as exc:
        raise UpdatePackageError(
            "The downloaded update package is not a valid zip file.\n"
            "Please download the installer manually from GitHub."
        ) from exc

    flat_exe = EXE_NAME
    nested_exe = f"ProjectTrackingTool/{EXE_NAME}"
    has_exe = flat_exe in names or nested_exe in names
    has_internal = any(
        name.startswith("_internal/") or name.startswith("ProjectTrackingTool/_internal/")
        for name in names
    )

    if not has_exe:
        raise UpdatePackageError(
            f"The downloaded update package does not contain {EXE_NAME}.\n"
            "Please download the installer manually from GitHub."
        )
    if not has_internal:
        raise UpdatePackageError(
            "The downloaded update package is incomplete: the _internal runtime folder is missing.\n"
            "Please download the installer manually from GitHub."
        )


def _ps_literal(value: "Path | str") -> str:
    """Return a PowerShell single-quoted string literal."""
    return "'" + str(value).replace("'", "''") + "'"


def _build_update_powershell_script(zip_path: Path, install_dir: Path, exe_path: Path) -> str:
    """Build the PowerShell script that performs the file replacement.

    Preserved-local for ``tests/test_regressions.py`` — commons has
    ``_build_full_folder_powershell(zip_path, install_dir, exe_path,
    exe_name)`` with a 4-arg signature; keeping ours local means tests
    don't need to adapt to the commons signature.
    """
    return f"""$ErrorActionPreference = 'Stop'
$zipPath = {_ps_literal(zip_path)}
$installDir = {_ps_literal(install_dir)}
$exePath = {_ps_literal(exe_path)}

Write-Output "Starting update from $zipPath"
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

    Write-Output "Update files copied successfully."
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


# ─── Public API — hybrid facade around phoenix_commons.updater ───────────────

def check_for_update() -> Optional[UpdateInfo]:
    """Query the GitHub Releases API and return an :class:`UpdateInfo` when newer.

    Wave 8b B3 facade. Delegates to
    :func:`phoenix_commons.updater.check_for_update` with the Project
    Tracking Tool release contract baked in (owner = ``JustinGlave``,
    repo = ``project-tracking-tool``, asset = ``ProjectTrackingTool.zip``).

    Safe to call from a background thread — never raises. commons logs
    network failures at DEBUG and payload-parse problems at WARNING.

    Semantic-narrowing note: the retired local implementation had a
    fallback that picked "any non-``fullinstall`` .zip asset" when the
    canonical name wasn't found. commons does exact-match only. No
    current release ships under a non-canonical name; if a future
    release ever does, operator can re-add the fallback locally without
    touching commons.
    """
    return _commons_check_for_update(
        owner=GITHUB_OWNER,
        repo=GITHUB_REPO,
        current_version=__version__,
        zip_asset_name=ZIP_ASSET_NAME,
    )


def download_and_apply(info: UpdateInfo, progress_callback=None) -> None:
    """Download the update zip, validate the full-folder payload, apply, and relaunch.

    Wave 8b B3 facade. Delegates to
    :func:`phoenix_commons.updater.download_and_apply` with the
    Project Tracking Tool release contract baked in:

    - ``exe_name=EXE_NAME`` (``ProjectTrackingTool.exe``) — the entry
      commons looks for in the zip and the basename used to construct
      the PowerShell extraction script.
    - ``expected_internal=True`` per **ADR-003** — Project Tracking
      Tool ships a full-folder updater zip (exe + ``_internal/``
      runtime folder). commons validates both at zip root (or inside a
      top-level folder named ``ProjectTrackingTool/``).
    - ``progress_callback(bytes_done, total_bytes)`` — forwarded
      verbatim for GUI progress-bar driving.

    Raises :class:`RuntimeError` (or :class:`UpdatePackageError`, a
    subclass) on any failure so the caller can show an error dialog
    rather than silently fail.
    """
    _commons_download_and_apply(
        info,
        exe_name=EXE_NAME,
        expected_internal=True,
        progress_callback=progress_callback,
    )
