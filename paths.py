"""Project Tracking Tool — local paths facade.

Re-exports commons path helpers and binds the tool-specific source-tree
``base`` for ``resource_path`` so call sites can use the historical
``resource_path(filename) -> Path`` shape (byte-identical to the
``_resource_path`` helper this file retires at Wave 8b B2).

Wave 8b B2 — created 2026-05-27 (operator-approved early-open override).

Why the wrapper?

  ``phoenix_commons.paths.resource_path(filename, base=None)`` returns
  ``Path(filename)`` in source mode when no ``base`` is provided, which
  is cwd-relative and brittle. The retired ``_resource_path`` always
  resolved against ``Path(__file__).with_name(filename)`` (next to the
  caller's module — repo root, since the function lived in
  ``project_tracker_gui.py``). Binding
  ``base = Path(__file__).resolve().parent`` here preserves that
  behavior — ``paths.py`` lives at repo root alongside the main GUI
  module, so the resolved base is identical.

Frozen-mode behavior is unchanged (commons returns ``_MEIPASS / filename``
when ``is_frozen()`` is true, regardless of ``base``).

Return type is ``Path`` (not ``str``) to preserve byte-identity with the
retired helper for any caller that may stringify, concatenate, or pass
to Qt constructors.

Preserved-local helpers in ``project_tracker_gui.py``:

  - ``_app_data_path()`` — Job-Tracker-specific data-file path with
    legacy-location migration logic (no commons equivalent).
  - ``_backup_data_file(data_path)`` — timestamped backup with 10-file
    retention (Job-Tracker-specific; no commons equivalent).
"""

from __future__ import annotations

from pathlib import Path

from phoenix_commons.paths import is_frozen, user_data_dir
from phoenix_commons.paths import resource_path as _commons_resource_path

__all__ = ["is_frozen", "user_data_dir", "resource_path"]

_TOOL_ROOT: Path = Path(__file__).resolve().parent


def resource_path(filename: str) -> Path:
    """Resolve a bundled-resource path. Frozen-aware via commons.

    Returns the same ``Path`` a caller would have gotten from the retired
    ``_resource_path`` helper in ``project_tracker_gui.py``.
    """
    return _commons_resource_path(filename, base=_TOOL_ROOT)
