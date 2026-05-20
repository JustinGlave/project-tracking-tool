# Changelog

All notable changes to **Project Tracking Tool** (working repo:
"Job Tracker") are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- CHANGELOG.md (this file) — Operational Hardening Sprint
  2026-05-19.

### Pending
- Phase 8b retrofit to commons-backed pattern per
  MIGRATION_RULES.md § Migration order. Largest surface of any
  Phoenix tool retrofit; `starter_package/` deletion planned in
  the same PR.

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
