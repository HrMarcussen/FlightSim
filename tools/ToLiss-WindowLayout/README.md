ToLiss Window Layout + Black Border Overlays

This toolset captures and reapplies pop‑out window layouts (e.g., ToLiss) and can draw click‑through black borders to hide gray edges or remove title bars.

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
- Apply layout and start overlays:
  - `./ToLiss-WindowLayout.ps1 -Action apply -LayoutPath ./TolissWindowLayout.json`

Other Commands
- Launch overlays only (no movement):
  - `./ToLiss-WindowLayout.ps1 -Action overlays -LayoutPath ./TolissWindowLayout.json`
- Stop all running overlays:
  - `./ToLiss-WindowLayout.ps1 -Action stop-overlays`

Notes
- Overlays are hidden PowerShell helper processes, one per window, with very low CPU usage. They update only when window bounds change.
- Click‑through ensures you can interact with the instrument beneath the border.
- DPI aware best‑effort; values are raw screen pixels.
- Title matching uses `-like '*TitleLike*'`. Capture suggests stable prefixes.

Troubleshooting
- No overlay? Ensure `BorderThickness > 0` or `StripTitleBar = true` in JSON and that the title substring matches the live window’s title.
- Multi‑monitor/scaling: the scripts enable per‑monitor DPI awareness; if alignment seems off, confirm the monitor’s scaling and window positions in the JSON.
