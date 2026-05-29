# Changelog

All notable changes to **Project Tracking Tool** (working repo:
"Job Tracker") are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- CHANGELOG.md (this file) — Operational Hardening Sprint
  2026-05-19.

### Changed
- Wave 8b retrofit to commons-backed pattern in progress on
  branch `phase-8b-job-tracker-retrofit`. B1-B5 facades complete
  (commons submodule, paths/updater/theme/widget facades);
  `version.py` unchanged at 1.8.5 (tag-skip per Decision #1).

### Removed
- `starter_package/` directory — historical Phoenix-tool scaffold
  that was bundled in this repo but never imported at runtime.
  Its updater + GUI patterns were ported into commons during
  Phase 1/3 of the platform rollout. Deleted at Wave 8b B7 per
  Decision #2.

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
