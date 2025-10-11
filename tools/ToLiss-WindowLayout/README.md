# ToLiss Window Layout (PowerShell)

Capture and apply window layouts for ToLiss/FlightSim (or any app) using Win32 APIs.
Per‑monitor DPI awareness is enabled where supported.

## Requirements
- Windows (uses `user32.dll` Win32 APIs)
- PowerShell 5.1+ or PowerShell 7+
- Optional: `Out-GridView` for GUI selection; falls back to console if unavailable

## Usage

Capture a layout (prompts to select windows):

```powershell
# Save to default path
.\ToLiss-WindowLayout.ps1 -Action capture

# Save to a specific path
.\ToLiss-WindowLayout.ps1 -Action capture -LayoutPath .\Toliss-A321.json
```

Apply one or more layouts:

```powershell
# Apply from default path
.\ToLiss-WindowLayout.ps1 -Action apply

# Apply from multiple JSON files
.\ToLiss-WindowLayout.ps1 -Action apply -LayoutPath .\Toliss-A321.json, .\Toliss-A346.json
```

Additional options:
- `-FirstOnly` on `Set-Window` to only move the first matched window.

## Notes
- Uses `EnumWindows`, `GetWindowText`, `GetWindowRect`, `SetWindowPos`, `ShowWindowAsync`.
- Skips zero-area windows when capturing to avoid hidden/minimized artifacts.
- JSON parsing is wrapped in try/catch with clear warnings on errors.
- Title suggestions trim common separators (e.g., ` - `, ` | `, `: `).
