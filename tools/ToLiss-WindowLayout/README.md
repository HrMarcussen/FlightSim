# Window Layout (PowerShell)

Capture and apply window layouts for any Windows desktop apps (e.g., X-Plane, MSFS) using Win32 APIs.
Per‑monitor DPI awareness is enabled where supported.

## Requirements
- Windows (uses `user32.dll` Win32 APIs)
- PowerShell 5.1+ or PowerShell 7+
- Optional: `Out-GridView` for GUI selection; falls back to console if unavailable

## Usage

Capture a layout (prompts to select windows):

```powershell
# Save to default path
& $PSCommandPath -Action capture

# Save to a specific path
& $PSCommandPath -Action capture -LayoutPath .\A321.json
```

Apply one or more layouts:

```powershell
# Apply from default path
& $PSCommandPath -Action apply

# Apply from multiple JSON files
& $PSCommandPath -Action apply -LayoutPath .\A321.json, .\A346.json
```

Additional options:
- `-FirstOnly` on `Set-Window` to only move the first matched window.

## Notes
- Uses `EnumWindows`, `GetWindowText`, `GetWindowRect`, `SetWindowPos`, `ShowWindowAsync`.
- Skips zero-area windows when capturing to avoid hidden/minimized artifacts.
- JSON parsing is wrapped in try/catch with clear warnings on errors.
- Title suggestions trim common separators (e.g., ` - `, ` | `, `: `).

