# Changelog

All notable changes to this toolset are documented here.

## 2025-10-30

- Fix: Out-GridView usage on Windows PowerShell 5.x by selecting columns before piping to `Out-GridView` (removes unsupported `-Property` parameter).
- Improve: More reliable window sizing when applying layouts. Adds verification and up to 4 retries with a 1px nudge to force apps/Windows to recalculate frame metrics.
- Fix: When `StripTitleBar` is true, reapply target bounds after starting the overlay so final size matches JSON on the first run (previously could require a second apply).
