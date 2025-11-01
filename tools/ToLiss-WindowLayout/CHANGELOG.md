# Changelog

All notable changes to this toolset are documented here.

## 2025-10-30

- Fix: Out-GridView usage on Windows PowerShell 5.x by selecting columns before piping to `Out-GridView` (removes unsupported `-Property` parameter).
- Improve: More reliable window sizing when applying layouts. Adds verification and up to 4 retries with a 1px nudge to force apps/Windows to recalculate frame metrics.
- Fix: When `StripTitleBar` is true, reapply target bounds after starting the overlay so final size matches JSON on the first run (previously could require a second apply).

- Fix: Ignore TYPE_ALREADY_EXISTS when loading Win32 helper type to allow reruns in same session.
- Improve: Reapply bounds for stripped windows up to 3 passes with delays to ensure final outer size matches JSON even on slower style transitions.
- Tweak: Quieter console output with per-entry summaries (positioned/stripped/resized) and no duplicate 'Placed' lines during retries/reapply.
## 2025-11-01

- Refactor: Version Win32 helper to `Win32NativeV2` to avoid stale in-session types; add style APIs and `StripTitleBarKeepBounds`.
- Improve: 3-phase apply (position → optional strip → reapply) with handle-targeted sizing and verification; stability checks and `SWP_FRAMECHANGED` on retries.
- Fix: PS5.1 parsing and session reruns (no duplicate type errors); remove ternary usage; quieter per-entry summaries.
- Overlay: Add per-edge cover with `Cover` and `TopCoverExtra` (JSON: `BorderCover`, `BorderTopCoverExtra`) to hide 1–2 px seams; bump overlay types to V5.
- Docs: README updated with new options and behavior.
