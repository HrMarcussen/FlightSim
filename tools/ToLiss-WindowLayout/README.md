ToLiss Window Layout + Black Border Overlays

This toolset captures and reapplies pop-out window layouts (e.g., ToLiss) and can draw click-through black borders to hide gray edges or remove title bars.

Components
- `ToLiss-WindowLayout.ps1` — capture/apply layout JSON and manage overlays
- `Add-BlackBorderOverlay.ps1` — overlay engine (one per window)
- `Start-OverlayLayout.ps1` — helper to start overlays from a JSON file

Quick Start
- Capture windows and create JSON:
  - `./ToLiss-WindowLayout.ps1 -Action capture -LayoutPath ./TolissWindowLayout.json`
- Edit JSON overlay fields per entry (optional):
  - `BorderThickness` (int): uniform thickness on left/right/bottom; 0 disables overlay
  - `BorderTopExtra` (int): extra pixels for top only (to cover title bars)
  - `StripTitleBar` (bool): remove caption/frame from the target window; keeps bounds
  - `Follow` (bool): overlay follows the window when it moves/resizes
  - `BorderCover` (int, optional): overlap (px) to shrink the inner overlay hole on all sides to hide 1px seams; default 2
  - `BorderTopCoverExtra` (int, optional): extra overlap (px) added on the top side only (e.g., use 1–2 if a faint line shows only on top)
- Apply layout and start overlays:
  - `./ToLiss-WindowLayout.ps1 -Action apply -LayoutPath ./TolissWindowLayout.json`

Other Commands
- Launch overlays only (no movement):
  - `./ToLiss-WindowLayout.ps1 -Action overlays -LayoutPath ./TolissWindowLayout.json`
- Stop all running overlays:
  - `./ToLiss-WindowLayout.ps1 -Action stop-overlays`

Notes
- Overlays are hidden PowerShell helper processes, one per window, with very low CPU usage. They update only when window bounds change.
- Click-through ensures you can interact with the instrument beneath the border.
- DPI aware best-effort; values are raw screen pixels.
- Title matching uses `-like '*TitleLike*'`. Capture suggests stable prefixes.

Troubleshooting
- No overlay? Ensure `BorderThickness > 0` or `StripTitleBar = true` in JSON and that the title substring matches the live window’s title.
- Multi-monitor/scaling: the scripts enable per-monitor DPI awareness; if alignment seems off, confirm the monitor’s scaling and window positions in the JSON.
- Faint top seam under the border? Set `BorderCover: 0–1` and add `BorderTopCoverExtra: 1–2` on entries where only the top edge shows a line.

Compatibility

- Windows PowerShell 5.x: `Out-GridView` does not support the `-Property` parameter. The script shapes columns using `Select-Object` before `Out-GridView` for compatibility.

Recent Fixes

- More reliable sizing when applying layouts: multi-pass with stability checks and `SWP_FRAMECHANGED`, handle-targeted verification, and 3-phase apply (position → optional strip → reapply).
- Quieter console output: per-entry summaries (positioned/stripped/resized/ok|mismatch).
- Overlay seam control: `BorderCover` and `BorderTopCoverExtra` to hide 1px compositor/DPI seams.

Version

- 0.3.0 (2025-11-01)
