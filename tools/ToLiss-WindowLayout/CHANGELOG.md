# Changelog

All notable changes to this project are documented in this file.
This format follows Keep a Changelog, and dates use ISO 8601 (YYYY-MM-DD).

## Unreleased
### Added
- Comment-based help for script and functions.
- Support for multiple `-LayoutPath` values on `apply` (applies each provided file).
- ASCII-only handling and normalization for separators and header text.
- `-FirstOnly` switch to `Set-Window` to only position the first matched window.

### Changed
- `Select-WindowsInteractive` now returns original objects directly via Out-GridView.
- `Get-OpenWindows` skips zero-area windows to reduce clutter.
- `Apply-Layout` JSON parsing wrapped in try/catch with clear warnings.
- More robust title prefix extraction via `Suggest-TitleLikeSimple`.

### Removed
- Obsolete regex-based `Suggest-TitleLike` helper.

## Unreleased
### Changed
- Generalize references from ToLiss-specific to generic Windows apps.
- Default layout filename changed to WindowLayout.json.
- Out-GridView prompt generalized.
