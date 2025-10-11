# Toliss-WindowLayout.ps1
param(
  [string]$LayoutPath = ".\TolissWindowLayout.json",
  [ValidateSet("capture","apply")] [string]$Action = "capture"
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
    public const int  SW_RESTORE = 9;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@

Add-Type -TypeDefinition $code -Language CSharp -PassThru | Out-Null

function Enable-PerMonitorDpi {
  try { [void][Win32Native]::SetProcessDpiAwarenessContext([IntPtr]::new(-4)) } catch {
    try { [void][Win32Native]::SetProcessDPIAware() } catch {}
  }
}

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
    $obj = [pscustomobject]@{
      Handle = $h
      Title  = $title
      Class  = $class
      X      = $r.Left
      Y      = $r.Top
      Width  = [Math]::Max(0, $r.Right - $r.Left)
      Height = [Math]::Max(0, $r.Bottom - $r.Top)
    }
    $list.Add($obj) | Out-Null
    return $true
  }, [IntPtr]::Zero) | Out-Null
  $list
}

function Set-Window {
  param(
    [Parameter(Mandatory)] [string]$TitleLike,
    [Parameter(Mandatory)] [int]$X,
    [Parameter(Mandatory)] [int]$Y,
    [Parameter(Mandatory)] [int]$Width,
    [Parameter(Mandatory)] [int]$Height,
    [int]$TimeoutSec = 20
  )
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  do {
    $targets = Get-OpenWindows | Where-Object { $_.Title -like "*$TitleLike*" }
    if ($targets) { break }
    Start-Sleep -Milliseconds 250
  } while ((Get-Date) -lt $deadline)

  if (-not $targets) {
    Write-Warning "Window '$TitleLike' not found within ${TimeoutSec}s."
    return
  }

  foreach ($t in $targets) {
    [void][Win32Native]::ShowWindowAsync($t.Handle, [Win32Native]::SW_RESTORE)
    $ok = [Win32Native]::SetWindowPos($t.Handle, [Win32Native]::HWND_TOP, $X, $Y, $Width, $Height,
      [Win32Native]::SWP_NOZORDER -bor [Win32Native]::SWP_NOACTIVATE)
    if ($ok) { Write-Host "Placed '$($t.Title)' -> $X,$Y ${Width}x${Height}" }
    else     { Write-Warning "Failed to position '$($t.Title)'." }
  }
}

function Select-WindowsInteractive {
  $all = Get-OpenWindows | Sort-Object Title

  $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
  if ($ogv) {
    $picked = $all | Select-Object Title,Class,X,Y,Width,Height | Out-GridView -Title "Select Toliss (and other) windows, then click OK" -PassThru
    if (-not $picked) { return @() }
    $picked = foreach ($p in $picked) {
      $all | Where-Object { $_.Title -eq $p.Title -and $_.Class -eq $p.Class -and $_.X -eq $p.X -and $_.Y -eq $p.Y }
    }
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

function Suggest-TitleLike([string]$title) {
  if ($title -match '^(.*?)(\s+-\s+| – )') { return $Matches[1] }
  if ($title.Length -le 20) { return $title }
  return $title.Substring(0, [Math]::Min(20,$title.Length))
}

function Capture-Layout {
  $picked = Select-WindowsInteractive
  if (-not $picked -or $picked.Count -eq 0) {
    Write-Warning "Nothing selected."
    return
  }

  $layout = @()
  foreach ($w in $picked) {
    $default = Suggest-TitleLike $w.Title
    $input = Read-Host ("TitleLike for '{0}' (Enter to accept: {1})" -f $w.Title, $default)
    if ([string]::IsNullOrWhiteSpace($input)) { $input = $default }

    $layout += [pscustomobject]@{
      TitleLike = $input
      X         = [int]$w.X
      Y         = [int]$w.Y
      Width     = [int]$w.Width
      Height    = [int]$w.Height
    }
  }

  $layout | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -Path $LayoutPath
  Write-Host "Saved layout ($($layout.Count) entries) -> $LayoutPath"
}

function Apply-Layout {
  if (-not (Test-Path $LayoutPath)) {
    Write-Warning "Layout file not found: $LayoutPath"
    return
  }
  $layout = Get-Content $LayoutPath -Raw | ConvertFrom-Json
  if (-not $layout) {
    Write-Warning "Layout is empty."
    return
  }
  foreach ($w in $layout) {
    Set-Window -TitleLike $w.TitleLike -X $w.X -Y $w.Y -Width $w.Width -Height $w.Height -TimeoutSec 20
  }
}

# --- MAIN ---
Enable-PerMonitorDpi
switch ($Action) {
  "capture" { Capture-Layout }
  "apply"   { Apply-Layout }
}
