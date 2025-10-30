# WindowLayout.ps1
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
  [ValidateSet("capture","apply","overlays","stop-overlays")] [string]$Action = "capture"
)
# --- Single, self-contained type (no duplicate 'using' issues) ---
$code = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class Win32Native {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int  GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    // DPI awareness (best effort; older OS will ignore)
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);

    public static readonly IntPtr HWND_TOP = IntPtr.Zero;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_FRAMECHANGED = 0x0020;
    public const int  SW_RESTORE = 9;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@

try {
  Add-Type -TypeDefinition $code -Language CSharp -ErrorAction Stop | Out-Null
} catch {
  if ($_.FullyQualifiedErrorId -notlike 'TYPE_ALREADY_EXISTS*') { throw }
}

<#
.SYNOPSIS
Enable best-effort per-monitor DPI awareness (v2 if supported).
#>
function Enable-PerMonitorDpi {
  try { [void][Win32Native]::SetProcessDpiAwarenessContext([IntPtr]::new(-4)) } catch {
    try { [void][Win32Native]::SetProcessDPIAware() } catch {}
  }
}

# Script directory for locating helper scripts (PS5.1-safe)
if (-not $script:ToolDir) {
  if ($PSCommandPath) { $script:ToolDir = Split-Path -Parent $PSCommandPath }
  if (-not $script:ToolDir) { $script:ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
}

<#
.SYNOPSIS
Enumerate open top-level windows.

.PARAMETER VisibleOnly
Return only visible windows (default).
#>
function Get-OpenWindows {
  param([switch]$VisibleOnly = $true)
  $list = New-Object System.Collections.Generic.List[object]
  [Win32Native]::EnumWindows({
    param([IntPtr]$h, [IntPtr]$p)
    if ($VisibleOnly -and -not [Win32Native]::IsWindowVisible($h)) { return $true }

    $len = [Win32Native]::GetWindowTextLength($h)
    if ($len -le 0) { return $true }
    $sb = New-Object System.Text.StringBuilder ($len + 1)
    [void][Win32Native]::GetWindowText($h, $sb, $sb.Capacity)
    $title = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($title)) { return $true }

    $csb = New-Object System.Text.StringBuilder 256
    [void][Win32Native]::GetClassName($h, $csb, $csb.Capacity)
    $class = $csb.ToString()

    [Win32Native+RECT]$r = New-Object 'Win32Native+RECT'
    [void][Win32Native]::GetWindowRect($h, [ref]$r)
    $ww = [Math]::Max(0, $r.Right - $r.Left)
    $hh = [Math]::Max(0, $r.Bottom - $r.Top)
    if ($ww -le 0 -or $hh -le 0) { return $true }
    $obj = [pscustomobject]@{
      Handle = $h
      Title  = $title
      Class  = $class
      X      = $r.Left
      Y      = $r.Top
      Width  = $ww
      Height = $hh
    }
    $list.Add($obj) | Out-Null
    return $true
  }, [IntPtr]::Zero) | Out-Null
  $list
}

<#
.SYNOPSIS
Position and resize windows matching a partial title.

.PARAMETER TitleLike
Substring to match window titles with PowerShell -like (wildcards added automatically).

.PARAMETER X
Left coordinate.

.PARAMETER Y
Top coordinate.

.PARAMETER Width
Window width.

.PARAMETER Height
Window height.

.PARAMETER FirstOnly
If set, only position the first matched window.

.PARAMETER TimeoutSec
How long to wait for a matching window to appear.
#>
function Set-Window {
  param(
    [Parameter(Mandatory)] [string]$TitleLike,
    [Parameter(Mandatory)] [int]$X,
    [Parameter(Mandatory)] [int]$Y,
    [Parameter(Mandatory)] [int]$Width,
    [Parameter(Mandatory)] [int]$Height,
    [switch]$FirstOnly,
    [int]$TimeoutSec = 20,
    [switch]$Quiet
  )
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
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
    return
  }

  foreach ($t in $targets) {
    [void][Win32Native]::ShowWindowAsync($t.Handle, [Win32Native]::SW_RESTORE)

    $flags = [Win32Native]::SWP_NOZORDER -bor [Win32Native]::SWP_NOACTIVATE
    $attempt = 0
    $placed = $false
    do {
      $attempt++
      $w = $Width; $h = $Height; $x = $X; $y = $Y
      $flagsTry = $flags
      if ($attempt -ge 3) {
        # Nudge size and force non-client frame recalculation
        $flagsTry = $flagsTry -bor [Win32Native]::SWP_FRAMECHANGED
        [void][Win32Native]::SetWindowPos($t.Handle, [Win32Native]::HWND_TOP, $x, $y, ($w + 1), $h, $flagsTry)
        Start-Sleep -Milliseconds 80
      }

      $ok = [Win32Native]::SetWindowPos($t.Handle, [Win32Native]::HWND_TOP, $x, $y, $w, $h, $flagsTry)
      if (-not $ok) { break }

      Start-Sleep -Milliseconds 100
      # Verify current rect
      [Win32Native+RECT]$r = New-Object 'Win32Native+RECT'
      [void][Win32Native]::GetWindowRect($t.Handle, [ref]$r)
      $cw = [Math]::Max(0, $r.Right - $r.Left)
      $ch = [Math]::Max(0, $r.Bottom - $r.Top)
      if ([Math]::Abs($cw - $w) -le 1 -and [Math]::Abs($ch - $h) -le 1) {
        $placed = $true
        break
      }
    } while ($attempt -lt 4)

    if (-not $Quiet) {
      if ($placed) { Write-Host "Placed '$($t.Title)' -> $X,$Y ${Width}x${Height}" }
      else         { Write-Warning "Failed to precisely size '$($t.Title)' (tried $attempt)." }
    }
  }
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
      Select-Object Title,Class,X,Y,Width,Height |
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
  $sel  = foreach ($i in $want) { if ($i -ge 0 -and $i -lt $all.Count) { $all[$i] } }
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
function Suggest-TitleLikeSimple([string]$title) {
  if ([string]::IsNullOrWhiteSpace($title)) { return "" }
  $separators = @(' - ', ' | ', ': ')
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
function Capture-Layout {
  $picked = Select-WindowsInteractive
  if (-not $picked -or $picked.Count -eq 0) {
    Write-Warning "Nothing selected."
    return
  }

  $layout = @()
  foreach ($w in $picked) {
    $default = Suggest-TitleLikeSimple $w.Title
    $input = Read-Host ("TitleLike for '{0}' (Enter to accept: {1})" -f $w.Title, $default)
    if ([string]::IsNullOrWhiteSpace($input)) { $input = $default }

    $layout += [pscustomobject]@{
      TitleLike = $input
      X         = [int]$w.X
      Y         = [int]$w.Y
      Width     = [int]$w.Width
      Height    = [int]$w.Height
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
    } else { ".\\WindowLayout.json" }
  } else { $LayoutPath }
  $layout | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -Path $targetPath
  Write-Host "Saved layout ($($layout.Count) entries) -> $targetPath"
}

<#
.SYNOPSIS
Start a black border overlay process for a layout entry if configured.
#>
function Start-OverlayForEntry {
  param([Parameter(Mandatory)]$Entry)
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

  $args = @(
    '-NoProfile','-ExecutionPolicy','Bypass','-File',"$overlayScript",
    '-TitleLike',"$($Entry.TitleLike)",
    '-Thickness',"$th",
    '-TopExtra',"$tx",
    '-TimeoutSec','30'
  )
  if ($st) { $args += '-StripTitleBar' }
  if ($fw) { $args += '-Follow' }

  # Build a single argument string for robust quoting on PS5.1
  $argString = @()
  foreach ($a in $args) {
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
function Apply-Layout {
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
    } catch {
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
    # Phase 1: position all windows first
    foreach ($w in $entries) {
      Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 20 -Quiet
      $count++
    }
    # Phase 2: start overlays (may strip title bars)
    foreach ($w in $entries) {
      Start-OverlayForEntry -Entry $w
      $key = "${($w.TitleLike)}"
      $wasStripped[$key] = $false
      if ($w -and ($w.PSObject.Properties.Match('StripTitleBar').Count -gt 0)) {
        try { if ([bool]$w.StripTitleBar) { $wasStripped[$key] = $true } } catch {}
      }
      $reapplyCount[$key] = 0
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
          for ($i=0; $i -lt 3; $i++) {
            Start-Sleep -Milliseconds (200 + ($i*200))
            Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 10 -Quiet
            $reapplyCount["${($w.TitleLike)}"] = $reapplyCount["${($w.TitleLike)}"] + 1
          }
        }
      }
    }
    $totalApplied += $count
    foreach ($w in $entries) {
      $key = "${($w.TitleLike)}"
      $msg = "Applied: '${($w.TitleLike)}' — positioned"
      if ($wasStripped.ContainsKey($key) -and $wasStripped[$key]) { $msg += "; titlebar stripped" }
      if ($reapplyCount.ContainsKey($key) -and $reapplyCount[$key] -gt 0) { $msg += "; resized x$($reapplyCount[$key])" }
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
function Apply-OverlaysOnly {
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
  } else {
    Write-Host "No overlay processes found."
  }
}

# --- MAIN ---
Enable-PerMonitorDpi
switch ($Action) {
  "capture" { Capture-Layout }
  "apply"   { Apply-Layout }
  "overlays" { Apply-OverlaysOnly }
  "stop-overlays" { Stop-AllOverlays }
}
