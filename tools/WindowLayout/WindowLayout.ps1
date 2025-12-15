# WindowLayout.ps1
<#
.SYNOPSIS
Captures and applies window layouts for any Windows desktop apps (e.g., X-Plane, MSFS).

.VERSION
0.3.0

.DESCRIPTION
Thin wrapper that imports the WindowLayout module and routes actions.

.PARAMETER LayoutPath
Path(s) to JSON layout file(s) to read or write.
For capture, if multiple are provided, the first is used.

.PARAMETER Action
"capture" to interactively select and save windows; "apply" to position windows from JSON.

.EXAMPLE
PS> .\WindowLayout.ps1 -Action capture
Interactively select open windows and save layout to WindowLayout.json

.EXAMPLE
PS> .\WindowLayout.ps1 -Action apply
Apply positions/sizes from WindowLayout.json
#>
param(
  [string[]]$LayoutPath = "WindowLayout.json",
  [ValidateSet("capture", "apply")] [string]$Action = "capture"
)

Import-Module -Force -ErrorAction Stop (Join-Path -Path $PSScriptRoot -ChildPath 'WindowLayout.psd1')

$script:LayoutPath = $LayoutPath
Enable-PerMonitorDpi
switch ($Action) {
  "capture" { Export-WindowLayout }
  "apply" { Restore-WindowLayout }
}
