# Start-OverlayLayout.ps1
# Launches one black-border overlay per window entry from a JSON layout.

param(
  [Parameter(Mandatory)] [string]$LayoutPath,
  [int]$Thickness = 8,
  [int]$TopExtra = 0,
  [switch]$StripTitleBar,
  [switch]$Follow = $true,
  [switch]$ApplyLayout
)

$ErrorActionPreference = 'Stop'

# Resolve paths
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$overlayScript = Join-Path $here 'Add-BlackBorderOverlay.ps1'
$layoutTool    = Join-Path $here 'ToLiss-WindowLayout.ps1'

if (-not (Test-Path $overlayScript)) { throw "Overlay script not found: $overlayScript" }
if (-not (Test-Path $LayoutPath))   { throw "Layout file not found: $LayoutPath" }

# Optionally position windows first (does not start overlays)
if ($ApplyLayout -and (Test-Path $layoutTool)) {
  Write-Host "Applying window positions from $LayoutPath"
  & $layoutTool -Action apply -LayoutPath $LayoutPath
}

# Read layout entries
try {
  $layout = Get-Content -Path $LayoutPath -Raw | ConvertFrom-Json -ErrorAction Stop
} catch {
  throw "Failed to parse JSON '$LayoutPath': $($_.Exception.Message)"
}

if (-not $layout -or $layout.Count -eq 0) { throw "Layout is empty: $LayoutPath" }

$launched = @()
foreach ($w in $layout) {
  if (-not $w.TitleLike) { continue }
  $args = @(
    '-NoProfile','-ExecutionPolicy','Bypass','-File',"$overlayScript",
    '-TitleLike',"$($w.TitleLike)",
    '-Thickness',"$Thickness",
    '-TopExtra',"$TopExtra",
    '-TimeoutSec','30'
  )
  if ($StripTitleBar) { $args += '-StripTitleBar' }
  if ($Follow)        { $args += '-Follow' }

  # Hide the helper console windows completely
  $p = Start-Process -FilePath 'powershell.exe' -ArgumentList $args -WindowStyle Hidden -PassThru
  $launched += [pscustomobject]@{ TitleLike = $w.TitleLike; PID = $p.Id }
}

Write-Host "Started $($launched.Count) overlay process(es)."
$launched | Format-Table -AutoSize

# Tip to stop overlays:
Write-Host "To stop an overlay: Stop-Process -Id <PID>"
