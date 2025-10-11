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

## 2025-10-11 (0.4.2)
### Fixed
- Guard Add-Type in module to avoid TYPE_ALREADY_EXISTS on re-import.

## 2025-10-11 (0.4.3)
### Added
- Parameter validation on Set-Window, Export-WindowLayout, and Restore-WindowLayout.
- -WhatIf/-Confirm support on Set-Window and top-level commands.
- Directory creation and guarded file write for export.
- JSON entry validation when restoring (fields and numeric coercion).

### Changed
- More verbose diagnostics via Write-Verbose; improved warnings on skips.

## 2025-10-11 (0.4.4)
### Added
- Checkbox-based window picker (WinForms) in Select-WindowsInteractive; falls back to Out-GridView/console.

## 2025-10-11 (0.4.9)
### Fixed
- Clean exports to approved verbs only even when importing .psm1 directly.
- Remove dark-mode attempts; stabilize Forms picker and OGV defaults.
- Ensure Cancel exits cleanly without fallback.
- DPI-safe Forms dialog with bottom panel and Select All/Select None buttons.

### Changed
- Default picker is OGV; Forms can be used with -Picker Forms.
- Updated documentation and manifest alignment.
