# Changelog

All notable changes to **Project Tracking Tool** (working repo:
"Job Tracker") are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [1.8.6] — 2026-05-30

Wave 8b commons retrofit + release hardening + starter_package
removal — no functional changes.

### Changed
- **Wave 8b commons retrofit complete** (merged 2026-05-28,
  commit `6a0d60b`). Migrated to commons-backed pattern per
  ADR-015 (`phoenix-commons` git submodule + editable install).
  Theme + widgets + paths + updater now facade through
  `phoenix_commons` rather than local duplicates. Local QSS
  overlay preserved (two-layer compose pattern) for app-specific
  selectors (`#StatCard`, `#taskToolsButton`, `#FinDataMeta`,
  `#ResizeHandle`, `#PassBadge`, etc.). **AppId absence preserved
  byte-for-byte** in `installer.iss` (per Decision #8 hard rule —
  v1.6.0..v1.8.5 users have AppName-hashed upgrade detection).
  Full-folder updater payload contract preserved
  (`expected_internal=True`). Detailed reports under
  `phoenix-commons/docs/ui-platform-baseline-v1/WAVE_8B_*.md`
  + `PHASE_8B_JOB_TRACKER_REPORT.md`.
- **Build pipeline hardened** per FROZEN_BUILD_BASELINE
  (Wave 8b B8, merged 2026-05-28). `build.bat` now enforces
  Python 3.12 soft-warn + Step 0 full cleanup +
  `--noupx` + `--collect-all=phoenix_commons` + 8× stdlib
  `--exclude-module` flags, on top of existing sanity-check
  pipeline (README version + py_compile + unittest discover +
  post-build zip layout verify). S1-safe profile per ADR-014.

### Removed
- `starter_package/` directory — historical Phoenix-tool scaffold
  that was bundled in this repo but never imported at runtime.
  Its updater + GUI patterns were ported into commons during
  Phase 1/3 of the platform rollout. Deleted at Wave 8b B7 per
  Decision #2.

### Added
- CHANGELOG.md (this file) — Operational Hardening Sprint
  2026-05-19.

## [1.8.5] — 2026-05-12

Bug-fix release covering 12 verified findings from a code audit.

### Fixed
- Backend / storage layer fixes (12 items from the audit; see
  commit `bc81037` for the full list).
- User session storage compatibility hardened (commit `1f556b0`).
- Release build script hardened (commit `4540311`).

### See also
- Earlier 1.8.x patch releases (1.8.0–1.8.4) — see git tags. Not
  reproduced here per the Phoenix Tools CHANGELOG policy ("current
  release + retrofit milestone only" — full history in git log).
