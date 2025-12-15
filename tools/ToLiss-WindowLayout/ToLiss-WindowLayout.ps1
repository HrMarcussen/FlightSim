# ToLiss-WindowLayout.ps1
<#
.SYNOPSIS
Captures and applies window layouts for any Windows desktop apps (e.g., X-Plane, MSFS).

.DESCRIPTION
Enumerates visible top-level windows and saves their positions/sizes to JSON,
then restores them using Win32 APIs. Includes best-effort per-monitor DPI awareness.

.PARAMETER LayoutPath
Path(s) to JSON layout file(s) to read or write.
For capture, if multiple are provided, the first is used.

.PARAMETER Action
"capture" to interactively select and save windows; "apply" to position windows from JSON and optionally start black border overlays per entry.

EXAMPLE
PS> .\WindowLayout.ps1 -Action capture
Interactively select open windows and save layout to WindowLayout.json

.EXAMPLE
PS> .\WindowLayout.ps1 -Action apply
Apply positions/sizes from WindowLayout.json
#>
param(
  [string[]]$LayoutPath = "WindowLayout.json",
  [ValidateSet("capture", "apply", "overlays", "stop-overlays")] [string]$Action = "capture"
)

# Import shared common module
$commonModule = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent)) 'Common\FlightSim.Common.psm1'
if (Test-Path $commonModule) {
  Import-Module $commonModule -ErrorAction Stop
}
else {
  Write-Warning "FlightSim.Common module not found at $commonModule"
}

# Script directory for locating helper scripts (PS5.1-safe)
if (-not $script:ToolDir) {
  if ($PSCommandPath) { $script:ToolDir = Split-Path -Parent $PSCommandPath }
  if (-not $script:ToolDir) { $script:ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
}

<#
.SYNOPSIS
Position and resize windows matching a partial title.
#>
function Set-Window {
  param(
    [string]$TitleLike,
    [IntPtr]$Handle,
    [Parameter(Mandatory)] [int]$X,
    [Parameter(Mandatory)] [int]$Y,
    [Parameter(Mandatory)] [int]$Width,
    [Parameter(Mandatory)] [int]$Height,
    [switch]$FirstOnly,
    [int]$TimeoutSec = 20,
    [switch]$Quiet,
    [switch]$ReturnResults
  )
  if (-not $Handle -and [string]::IsNullOrWhiteSpace($TitleLike)) { throw "Provide either -TitleLike or -Handle" }
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  $targets = @()
  if ($Handle) {
    # Build a single target from handle, enrich with title for logs
    [FlightSim.Common.Win32Native+RECT]$r0 = New-Object 'FlightSim.Common.Win32Native+RECT'
    $ok0 = [FlightSim.Common.Win32Native]::GetWindowRect($Handle, [ref]$r0)
    if ($ok0) {
      $len = [FlightSim.Common.Win32Native]::GetWindowTextLength($Handle)
      $title = ''
      if ($len -gt 0) { $sb = New-Object System.Text.StringBuilder ($len + 1); [void][FlightSim.Common.Win32Native]::GetWindowText($Handle, $sb, $sb.Capacity); $title = $sb.ToString() }
      $targets = @([pscustomobject]@{ Handle = $Handle; Title = $title })
    }
  }
  else {
    do {
      $targets = Get-OpenWindows | Where-Object { $_.Title -like "*$TitleLike*" }
      if ($targets) {
        if ($FirstOnly) { $targets = @($targets | Select-Object -First 1) }
        break
      }
      Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)
    if (-not $targets) {
      Write-Warning "Window '$TitleLike' not found within ${TimeoutSec}s."
      return @()
    }
  }

  $results = @()
  foreach ($t in $targets) {
    [void][FlightSim.Common.Win32Native]::ShowWindowAsync($t.Handle, [FlightSim.Common.Win32Native]::SW_RESTORE)

    $flags = [FlightSim.Common.Win32Native]::SWP_NOZORDER -bor [FlightSim.Common.Win32Native]::SWP_NOACTIVATE
    $attempt = 0
    $placed = $false
    do {
      $attempt++
      $w = $Width; $h = $Height; $x = $X; $y = $Y
      $flagsTry = ($flags -bor [FlightSim.Common.Win32Native]::SWP_FRAMECHANGED)
      if ($attempt -ge 3) {
        # Nudge size and force non-client frame recalculation
        [void][FlightSim.Common.Win32Native]::SetWindowPos($t.Handle, [FlightSim.Common.Win32Native]::HWND_TOP, $x, $y, ($w + 1), $h, $flagsTry)
        Start-Sleep -Milliseconds 80
      }

      $ok = [FlightSim.Common.Win32Native]::SetWindowPos($t.Handle, [FlightSim.Common.Win32Native]::HWND_TOP, $x, $y, $w, $h, $flagsTry)
      if (-not $ok) { break }

      Start-Sleep -Milliseconds 100
      # Verify current rect
      [FlightSim.Common.Win32Native+RECT]$r = New-Object 'FlightSim.Common.Win32Native+RECT'
      [void][FlightSim.Common.Win32Native]::GetWindowRect($t.Handle, [ref]$r)
      $cw = [Math]::Max(0, $r.Right - $r.Left)
      $ch = [Math]::Max(0, $r.Bottom - $r.Top)
      if ([Math]::Abs($cw - $w) -le 1 -and [Math]::Abs($ch - $h) -le 1) {
        # Require stability across a short delay to avoid races with style changes
        Start-Sleep -Milliseconds 150
        [FlightSim.Common.Win32Native+RECT]$r2 = New-Object 'FlightSim.Common.Win32Native+RECT'
        [void][FlightSim.Common.Win32Native]::GetWindowRect($t.Handle, [ref]$r2)
        $cw2 = [Math]::Max(0, $r2.Right - $r2.Left)
        $ch2 = [Math]::Max(0, $r2.Bottom - $r2.Top)
        if ([Math]::Abs($cw2 - $w) -le 1 -and [Math]::Abs($ch2 - $h) -le 1) {
          $placed = $true
          break
        }
      }
    } while ($attempt -lt 6)

    if (-not $Quiet) {
      if ($placed) { Write-Host "Placed '$($t.Title)' -> $X,$Y ${Width}x${Height}" }
      else { Write-Warning "Failed to precisely size '$($t.Title)' (tried $attempt)." }
    }
    # Collect final rect
    [FlightSim.Common.Win32Native+RECT]$rf = New-Object 'FlightSim.Common.Win32Native+RECT'
    [void][FlightSim.Common.Win32Native]::GetWindowRect($t.Handle, [ref]$rf)
    $fw = [Math]::Max(0, $rf.Right - $rf.Left)
    $fh = [Math]::Max(0, $rf.Bottom - $rf.Top)
    $results += [pscustomobject]@{ Handle = $t.Handle; Title = $t.Title; X = $rf.Left; Y = $rf.Top; Width = $fw; Height = $fh }
  }
  if ($ReturnResults) { return , $results } else { return }
}

<#
.SYNOPSIS
Interactive picker (Out-GridView or console) to choose windows to capture.
#>
function Select-WindowsInteractive {
  $all = Get-OpenWindows | Sort-Object Title

  $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
  if ($ogv) {
    $picked = $all |
    Select-Object Title, Class, X, Y, Width, Height |
    Out-GridView -Title "Select windows to capture, then click OK" -PassThru
    if (-not $picked) { return @() }
    return $picked
  }

  Write-Host "`nSelect windows (comma-separated indices):`n"
  $idx = 0
  $indexed = $all | ForEach-Object {
    "{0,3}.  {1}  [{2}]  @ {3},{4}  {5}x{6}" -f $idx, $_.Title, $_.Class, $_.X, $_.Y, $_.Width, $_.Height
    $idx++
  }
  $indexed | ForEach-Object { Write-Host $_ }
  $ans = Read-Host "`nEnter indices (e.g. 0,3,4)"
  if ([string]::IsNullOrWhiteSpace($ans)) { return @() }
  $want = $ans -split '[^0-9]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
  $sel = foreach ($i in $want) { if ($i -ge 0 -and $i -lt $all.Count) { $all[$i] } }
  return $sel
}


<#
Return ASCII separators used to split window titles.
#>
function Get-AsciiSeparators { @(' - ', ' | ', ': ') }

<#
.SYNOPSIS
Extract a stable prefix from a full window title.
.DESCRIPTION
Splits on common separators (e.g., " - ", " | ", ": ") and
returns the part before the separator; otherwise caps to 20 chars.
#>
function Get-TitleSuggestion([string]$title) {
  if ([string]::IsNullOrWhiteSpace($title)) { return "" }
  foreach ($sep in (Get-AsciiSeparators)) {
    $idx = $title.IndexOf($sep)
    if ($idx -gt 0) { return $title.Substring(0, $idx).Trim() }
  }
  if ($title.Length -le 20) { return $title }
  return $title.Substring(0, [Math]::Min(20, $title.Length))
}

<#
.SYNOPSIS
Interactively select windows and save their positions to JSON.
#>
function Save-WindowLayout {
  $picked = Select-WindowsInteractive
  if (-not $picked -or $picked.Count -eq 0) {
    Write-Warning "Nothing selected."
    return
  }

  $layout = @()
  foreach ($w in $picked) {
    $default = Get-TitleSuggestion $w.Title
    $titleInput = Read-Host ("TitleLike for '{0}' (Enter to accept: {1})" -f $w.Title, $default)
    if ([string]::IsNullOrWhiteSpace($titleInput)) { $titleInput = $default }

    $layout += [pscustomobject]@{
      TitleLike       = $titleInput
      X               = [int]$w.X
      Y               = [int]$w.Y
      Width           = [int]$w.Width
      Height          = [int]$w.Height
      BorderThickness = 8
      BorderTopExtra  = 0
      StripTitleBar   = $false
      Follow          = $true
    }
  }

  $targetPath = if ($LayoutPath -is [array]) {
    if ($LayoutPath.Count -gt 0) {
      if ($LayoutPath.Count -gt 1) { Write-Warning "Multiple LayoutPath values provided; using first: $($LayoutPath[0])" }
      $LayoutPath[0]
    }
    else { ".\\WindowLayout.json" }
  }
  else { $LayoutPath }
  $layout | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -Path $targetPath
  Write-Host "Saved layout ($($layout.Count) entries) -> $targetPath"
}

<#
.SYNOPSIS
Start a black border overlay process for a layout entry if configured.
#>
function Start-OverlayForEntry {
  param([Parameter(Mandatory)]$Entry, [switch]$SkipStripTitleBar)
  $toolDir = $script:ToolDir
  if (-not $toolDir) { $toolDir = Split-Path -Parent $PSCommandPath }
  $overlayScript = if ($toolDir) { Join-Path $toolDir 'Add-BlackBorderOverlay.ps1' } else { $null }
  if (-not (Test-Path $overlayScript)) { return }

  # PS5.1-compatible null/exists checks (no '??' operator)
  $th = 0
  if ($Entry.PSObject.Properties.Match('BorderThickness').Count -gt 0 -and $null -ne $Entry.BorderThickness) {
    $th = [int]$Entry.BorderThickness
  }
  $tx = 0
  if ($Entry.PSObject.Properties.Match('BorderTopExtra').Count -gt 0 -and $null -ne $Entry.BorderTopExtra) {
    $tx = [int]$Entry.BorderTopExtra
  }
  $cv = 2
  if ($Entry.PSObject.Properties.Match('BorderCover').Count -gt 0 -and $null -ne $Entry.BorderCover) {
    $cv = [int]$Entry.BorderCover
  }
  $tcx = 0
  if ($Entry.PSObject.Properties.Match('BorderTopCoverExtra').Count -gt 0 -and $null -ne $Entry.BorderTopCoverExtra) {
    $tcx = [int]$Entry.BorderTopCoverExtra
  }
  $st = $false
  if ($Entry.PSObject.Properties.Match('StripTitleBar').Count -gt 0 -and $null -ne $Entry.StripTitleBar) {
    $st = [bool]$Entry.StripTitleBar
  }
  $fw = $true
  if ($Entry.PSObject.Properties.Match('Follow').Count -gt 0 -and $null -ne $Entry.Follow) {
    $fw = [bool]$Entry.Follow
  }

  if ($th -le 0 -and -not $st) { return }
  if ([string]::IsNullOrWhiteSpace($Entry.TitleLike)) { return }

  $launchArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "$overlayScript",
    '-TitleLike', "$($Entry.TitleLike)",
    '-Thickness', "$th",
    '-TopExtra', "$tx",
    '-Cover', "$cv",
    '-TopCoverExtra', "$tcx",
    '-TimeoutSec', '30'
  )
  if ($st -and -not $SkipStripTitleBar) { $launchArgs += '-StripTitleBar' }
  if ($fw) { $launchArgs += '-Follow' }

  # Build a single argument string for robust quoting on PS5.1
  $argString = @()
  foreach ($a in $launchArgs) {
    if ($a -match '\s') { $argString += ('"' + $a + '"') } else { $argString += $a }
  }
  $argString = ($argString -join ' ')

  Start-Process -FilePath 'powershell.exe' -ArgumentList $argString -WindowStyle Hidden -WorkingDirectory $toolDir | Out-Null
}

<#
.SYNOPSIS
Read JSON layout and position matching windows.
.DESCRIPTION
Supports applying multiple layout files in one run.
#>
function Restore-WindowLayout {
  $paths = @($LayoutPath)
  if (-not $paths -or $paths.Count -eq 0) {
    Write-Warning "No layout path(s) provided."
    return
  }

  $totalApplied = 0
  foreach ($path in $paths) {
    if (-not (Test-Path $path)) {
      Write-Warning "Layout file not found: $path"
      continue
    }
    try {
      $layout = Get-Content -Path $path -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
      Write-Warning "Failed to parse layout JSON '$path': $($_.Exception.Message)"
      continue
    }
    if (-not $layout) {
      Write-Warning "Layout is empty: $path"
      continue
    }
    $count = 0
    $entries = @($layout)
    $wasStripped = @{}
    $reapplyCount = @{}
    $handles = @{}
    $finalOk = @{}
    # Phase 1: position all windows first
    foreach ($w in $entries) {
      $res = Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 20 -FirstOnly -Quiet -ReturnResults
      if ($res -and $res.Count -gt 0) { $handles["$($w.TitleLike)"] = $res[0].Handle }
      $count++
    }
    # Phase 2: synchronously strip title bars (if requested) before overlays
    foreach ($w in $entries) {
      $key = "$($w.TitleLike)"
      $wasStripped[$key] = $false
      if ($w -and ($w.PSObject.Properties.Match('StripTitleBar').Count -gt 0)) {
        $doStrip = $false
        try { $doStrip = [bool]$w.StripTitleBar } catch {}
        if ($doStrip -and $handles.ContainsKey($key)) {
          [FlightSim.Common.Win32Native]::StripTitleBarKeepBounds($handles[$key], [int]$w.X, [int]$w.Y, [int]$w.Width, [int]$w.Height)
          Start-Sleep -Milliseconds 150
          $wasStripped[$key] = $true
        }
      }
      $reapplyCount[$key] = 0
    }
    # Phase 2.5: start overlays (skip stripping here if already done)
    foreach ($w in $entries) {
      $key = "$($w.TitleLike)"
      $skip = $false; if ($wasStripped.ContainsKey($key) -and $wasStripped[$key]) { $skip = $true }
      Start-OverlayForEntry -Entry $w -SkipStripTitleBar:$skip
    }
    # Phase 3: if any stripped, re-apply final bounds to match JSON exactly
    $needsReapply = $false
    foreach ($w in $entries) {
      if ($w -and ($w.PSObject.Properties.Match('StripTitleBar').Count -gt 0)) {
        try { if ([bool]$w.StripTitleBar) { $needsReapply = $true; break } } catch {}
      }
    }
    if ($needsReapply) {
      # Give stripped windows a moment to settle, then reapply a few times
      foreach ($w in $entries) {
        if ($w -and ($w.PSObject.Properties.Match('StripTitleBar').Count -gt 0)) {
          $isStripped = $false
          try { $isStripped = [bool]$w.StripTitleBar } catch {}
          if (-not $isStripped) { continue }
          $okNow = $false
          for ($i = 0; $i -lt 5; $i++) {
            Start-Sleep -Milliseconds (200 + ($i * 250))
            if ($handles.ContainsKey("$($w.TitleLike)")) {
              Set-Window -Handle $handles["$($w.TitleLike)"] -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 10 -Quiet
            }
            else {
              Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 10 -FirstOnly -Quiet
            }
            $reapplyCount["$($w.TitleLike)"] = $reapplyCount["$($w.TitleLike)"] + 1
            # Verify actual rect and require stability across two reads
            $okOnce = $false
            for ($p = 0; $p -lt 2; $p++) {
              Start-Sleep -Milliseconds 150
              if ($handles.ContainsKey("$($w.TitleLike)") ) { $cur = Get-WindowRectByHandle -Handle $handles["$($w.TitleLike)"] }
              else { $cur = Get-OpenWindows | Where-Object { $_.Title -like "*$($w.TitleLike)*" } | Select-Object -First 1 }
              if ($cur) {
                $cw = [int]$cur.Width; $ch = [int]$cur.Height
                if ([Math]::Abs($cw - [int]$w.Width) -le 1 -and [Math]::Abs($ch - [int]$w.Height) -le 1) {
                  if ($okOnce) { $okNow = $true; break }
                  $okOnce = $true
                  continue
                }
              }
              $okOnce = $false
            }
            if ($okNow) { break }
          }
          $finalOk["$($w.TitleLike)"] = $okNow
        }
      }
    }
    $totalApplied += $count
    foreach ($w in $entries) {
      $key = "$($w.TitleLike)"
      $msg = "Applied: '$($w.TitleLike)' - positioned"
      if ($wasStripped.ContainsKey($key) -and $wasStripped[$key]) { $msg += "; titlebar stripped" }
      if ($reapplyCount.ContainsKey($key) -and $reapplyCount[$key] -gt 0) { $msg += "; resized x$($reapplyCount[$key])" }
      if ($finalOk.ContainsKey($key)) {
        if ($finalOk[$key]) { $msg += "; ok" }
        else { $msg += "; mismatch" }
      }
      Write-Host $msg
    }
    Write-Host "Applied $count entry(ies) from $path"
  }

  if ($totalApplied -eq 0) {
    Write-Warning "No layout entries applied from provided path(s)."
  }
}

<#
.SYNOPSIS
Launch overlays only from a layout without moving windows.
#>
function Start-OverlayProcess {
  $paths = @($LayoutPath)
  if (-not $paths -or $paths.Count -eq 0) { return }
  foreach ($path in $paths) {
    if (-not (Test-Path $path)) { continue }
    try { $layout = Get-Content -Path $path -Raw | ConvertFrom-Json -ErrorAction Stop } catch { continue }
    foreach ($w in $layout) { Start-OverlayForEntry -Entry $w }
  }
}

<#
.SYNOPSIS
Stop all running black border overlay processes.
#>
function Stop-AllOverlays {
  $procs = Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*Add-BlackBorderOverlay.ps1*' }
  if ($procs) {
    $procs | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
    Write-Host "Stopped $($procs.Count) overlay process(es)."
  }
  else {
    Write-Host "No overlay processes found."
  }
}

# --- MAIN ---
Enable-PerMonitorDpi
switch ($Action) {
  "capture" { Save-WindowLayout }
  "apply" { Restore-WindowLayout }
  "overlays" { Start-OverlayProcess }
  "stop-overlays" { Stop-AllOverlays }
}
