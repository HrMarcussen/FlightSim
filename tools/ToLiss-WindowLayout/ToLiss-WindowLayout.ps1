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

public static class Win32NativeV2 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int  GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll", EntryPoint = "GetWindowLong")] private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")] private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", EntryPoint = "SetWindowLong")] private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")] private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    // DPI awareness (best effort; older OS will ignore)
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);

    public static readonly IntPtr HWND_TOP = IntPtr.Zero;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const uint SWP_FRAMECHANGED = 0x0020;
    public const uint SWP_SHOWWINDOW = 0x0040;
    public const int  SW_RESTORE = 9;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    public const int WS_SYSMENU = 0x00080000;
    public const int WS_MINIMIZEBOX = 0x00020000;
    public const int WS_MAXIMIZEBOX = 0x00010000;

    private static IntPtr GetWindowLongPtrSafe(IntPtr hWnd, int nIndex) {
        if (IntPtr.Size == 8) return GetWindowLongPtr64(hWnd, nIndex);
        return new IntPtr(GetWindowLong32(hWnd, nIndex));
    }
    private static IntPtr SetWindowLongPtrSafe(IntPtr hWnd, int nIndex, IntPtr newVal) {
        if (IntPtr.Size == 8) return SetWindowLongPtr64(hWnd, nIndex, newVal);
        return new IntPtr(SetWindowLong32(hWnd, nIndex, newVal.ToInt32()));
    }

    public static void StripTitleBarKeepBounds(IntPtr hWnd, int x, int y, int w, int h) {
        try {
            IntPtr stylePtr = GetWindowLongPtrSafe(hWnd, GWL_STYLE);
            long style = stylePtr.ToInt64();
            style &= ~(long)(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
            SetWindowLongPtrSafe(hWnd, GWL_STYLE, new IntPtr(style));
            SetWindowPos(hWnd, IntPtr.Zero, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW | SWP_FRAMECHANGED);
        } catch { }
    }
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
  try { [void][Win32NativeV2]::SetProcessDpiAwarenessContext([IntPtr]::new(-4)) } catch {
    try { [void][Win32NativeV2]::SetProcessDPIAware() } catch {}
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
  [Win32NativeV2]::EnumWindows({
    param([IntPtr]$h, [IntPtr]$p)
    if ($VisibleOnly -and -not [Win32NativeV2]::IsWindowVisible($h)) { return $true }

    $len = [Win32NativeV2]::GetWindowTextLength($h)
    if ($len -le 0) { return $true }
    $sb = New-Object System.Text.StringBuilder ($len + 1)
    [void][Win32NativeV2]::GetWindowText($h, $sb, $sb.Capacity)
    $title = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($title)) { return $true }

    $csb = New-Object System.Text.StringBuilder 256
    [void][Win32NativeV2]::GetClassName($h, $csb, $csb.Capacity)
    $class = $csb.ToString()

    [Win32NativeV2+RECT]$r = New-Object 'Win32NativeV2+RECT'
    [void][Win32NativeV2]::GetWindowRect($h, [ref]$r)
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
    [Win32NativeV2+RECT]$r0 = New-Object 'Win32NativeV2+RECT'
    $ok0 = [Win32NativeV2]::GetWindowRect($Handle, [ref]$r0)
    if ($ok0) {
      $len = [Win32NativeV2]::GetWindowTextLength($Handle)
      $title = ''
      if ($len -gt 0) { $sb = New-Object System.Text.StringBuilder ($len + 1); [void][Win32NativeV2]::GetWindowText($Handle, $sb, $sb.Capacity); $title = $sb.ToString() }
      $targets = @([pscustomobject]@{ Handle=$Handle; Title=$title })
    }
  } else {
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
    [void][Win32NativeV2]::ShowWindowAsync($t.Handle, [Win32NativeV2]::SW_RESTORE)

    $flags = [Win32NativeV2]::SWP_NOZORDER -bor [Win32NativeV2]::SWP_NOACTIVATE
    $attempt = 0
    $placed = $false
    do {
      $attempt++
      $w = $Width; $h = $Height; $x = $X; $y = $Y
      $flagsTry = ($flags -bor [Win32NativeV2]::SWP_FRAMECHANGED)
      if ($attempt -ge 3) {
        # Nudge size and force non-client frame recalculation
        [void][Win32NativeV2]::SetWindowPos($t.Handle, [Win32NativeV2]::HWND_TOP, $x, $y, ($w + 1), $h, $flagsTry)
        Start-Sleep -Milliseconds 80
      }

      $ok = [Win32NativeV2]::SetWindowPos($t.Handle, [Win32NativeV2]::HWND_TOP, $x, $y, $w, $h, $flagsTry)
      if (-not $ok) { break }

      Start-Sleep -Milliseconds 100
      # Verify current rect
      [Win32NativeV2+RECT]$r = New-Object 'Win32NativeV2+RECT'
      [void][Win32NativeV2]::GetWindowRect($t.Handle, [ref]$r)
      $cw = [Math]::Max(0, $r.Right - $r.Left)
      $ch = [Math]::Max(0, $r.Bottom - $r.Top)
      if ([Math]::Abs($cw - $w) -le 1 -and [Math]::Abs($ch - $h) -le 1) {
        # Require stability across a short delay to avoid races with style changes
        Start-Sleep -Milliseconds 150
        [Win32NativeV2+RECT]$r2 = New-Object 'Win32NativeV2+RECT'
        [void][Win32NativeV2]::GetWindowRect($t.Handle, [ref]$r2)
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
      else         { Write-Warning "Failed to precisely size '$($t.Title)' (tried $attempt)." }
    }
    # Collect final rect
    [Win32NativeV2+RECT]$rf = New-Object 'Win32NativeV2+RECT'
    [void][Win32NativeV2]::GetWindowRect($t.Handle, [ref]$rf)
    $fw = [Math]::Max(0, $rf.Right - $rf.Left)
    $fh = [Math]::Max(0, $rf.Bottom - $rf.Top)
    $results += [pscustomobject]@{ Handle=$t.Handle; Title=$t.Title; X=$rf.Left; Y=$rf.Top; Width=$fw; Height=$fh }
  }
  if ($ReturnResults) { return ,$results } else { return }
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
.SYNOPSIS
Get window rect by handle.
#>
function Get-WindowRectByHandle {
  param([Parameter(Mandatory)][IntPtr]$Handle)
  [Win32NativeV2+RECT]$rr = New-Object 'Win32NativeV2+RECT'
  if (-not [Win32NativeV2]::GetWindowRect($Handle, [ref]$rr)) { return $null }
  $ww = [Math]::Max(0, $rr.Right - $rr.Left)
  $hh = [Math]::Max(0, $rr.Bottom - $rr.Top)
  return [pscustomobject]@{ X=$rr.Left; Y=$rr.Top; Width=$ww; Height=$hh }
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
  if ($st -and -not $SkipStripTitleBar) { $args += '-StripTitleBar' }
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
          [Win32NativeV2]::StripTitleBarKeepBounds($handles[$key], [int]$w.X, [int]$w.Y, [int]$w.Width, [int]$w.Height)
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
          for ($i=0; $i -lt 5; $i++) {
            Start-Sleep -Milliseconds (200 + ($i*250))
            if ($handles.ContainsKey("$($w.TitleLike)")) {
              Set-Window -Handle $handles["$($w.TitleLike)"] -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 10 -Quiet
            } else {
              Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 10 -FirstOnly -Quiet
            }
            $reapplyCount["$($w.TitleLike)"] = $reapplyCount["$($w.TitleLike)"] + 1
            # Verify actual rect and require stability across two reads
            $okOnce = $false
            for ($p=0; $p -lt 2; $p++) {
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
