# Changelog

All notable changes to this project are documented in this file.
This format follows Keep a Changelog, and dates use ISO 8601 (YYYY-MM-DD).

## 2025-10-11 (0.3.0)
### Added
- Convert to a PowerShell module (WindowLayout.psm1 + WindowLayout.psd1).
- Export functions for DPI, enumeration, capture/apply, and window movement.
- Wrapper script remains for convenience.

### Changed
- Documentation updated for module usage.

## 2025-10-11 (0.2.1)
### Fixed
- Out-GridView compatibility on Windows PowerShell 5.x by piping Select-Object before Out-GridView.

## 2025-10-11 (0.2.0)
### Changed
- Generalize naming and defaults (non-ToLiss specific).
- Default layout filename to WindowLayout.json.
- Rename folder to 	ools/WindowLayout/ and script to WindowLayout.ps1.
- Generalize OGV prompt; remove ToLiss mentions.


## 2025-10-11 (0.4.0)
### Added
- New high-level commands: Save-WindowLayout and Apply-WindowLayout.
- Updated manifest to export new commands.
- Updated README with module command usage.

## 2025-10-11 (0.4.1)
### Changed
- Use approved verbs: Export-WindowLayout and Restore-WindowLayout.
- Provide legacy aliases: Save-WindowLayout, Apply-WindowLayout.
### Changed
- Remove export of unapproved verbs (Capture-Layout, Apply-Layout) to avoid import warnings; keep them internal.
