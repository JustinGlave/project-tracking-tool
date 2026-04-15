# Security Policy

## Supported Versions

Only the latest release is actively supported with security fixes.

| Version | Supported |
|---------|-----------|
| Latest  | Yes       |
| Older   | No        |

## Reporting a Vulnerability

If you discover a security vulnerability in this project, **do not open a public issue**.

Report it privately to the repository owner:

- **Email:** justinglave@gmail.com
- **Subject:** `[PTT Security] Brief description`

Please include:
- A description of the vulnerability
- Steps to reproduce it
- Potential impact

You can expect a response within 3 business days. If the issue is confirmed, a fix will be prioritized for the next release.

## Scope

Key security considerations for this application:

- **User credentials** are stored as bcrypt hashes — plain-text passwords are never written to disk
- **Project data** is stored in a local JSON file — ensure the data folder has appropriate OS-level access controls in shared environments
- **Auto-updater** downloads only from the official GitHub releases page of this repository
