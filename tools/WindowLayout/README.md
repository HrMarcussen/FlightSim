# Window Layout

Capture and apply window layouts for Windows desktop apps using Win32 APIs. Per‑monitor DPI awareness is enabled where supported.

## Requirements
- Windows (uses `user32.dll` Win32 APIs)
- PowerShell 5.1+ or PowerShell 7+
- Optional: `Out-GridView` for GUI selection; falls back to console if unavailable

## Usage (as a module)

```powershell
# Import module
Import-Module (Join-Path $PSScriptRoot "WindowLayout.psd1")

# Capture interactively to default path in current folder
Export-WindowLayout

# Capture to specific file
Export-WindowLayout -LayoutPath .\MyLayout.json

# Apply one or more layouts
Restore-WindowLayout -LayoutPath .\MyLayout.json, .\OtherLayout.json
```

## Usage (as a script)

```powershell
# Capture with defaults
.\WindowLayout.ps1 -Action capture

# Apply from default path
.\WindowLayout.ps1 -Action apply

# Apply from specific files
.\WindowLayout.ps1 -Action apply -LayoutPath .\WindowLayout.json, .\Another.json
```

### Notes
- Functions exported: `Enable-PerMonitorDpi`, `Get-OpenWindows`, `Set-Window`, `Select-WindowsInteractive`, `Export-WindowLayout`, `Restore-WindowLayout`.
- Skips zero-area windows when capturing to avoid hidden/minimized artifacts.
- Title suggestions trim common separators (e.g., ` - `, ` | `, `: `).
- JSON parsing is wrapped in try/catch with warnings.


> Legacy aliases: Save-WindowLayout, Apply-WindowLayout

