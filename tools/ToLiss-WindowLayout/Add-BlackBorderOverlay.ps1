<#
.SYNOPSIS
Creates a topmost, click-through black border overlay around a target window.

.DESCRIPTION
Matches a window by title substring, optionally strips its title bar, and draws
an always-on-top, click-through black border around it. The overlay can follow
the window if it moves or resizes. Designed for sim pop-outs (e.g., ToLiss).

.PARAMETER TitleLike
Substring to match the window title (case-insensitive, like '*TitleLike*').

.PARAMETER Thickness
Uniform border thickness (pixels) for left, right, and bottom.

.PARAMETER TopExtra
Extra pixels to add only to the top border (to cover captions).

.PARAMETER Follow
If set, the overlay polls and follows the window when it moves/resizes.

.PARAMETER StripTitleBar
Remove caption/frame from the target window while keeping same bounds.

.PARAMETER TimeoutSec
How long to wait for the target window to appear.

.EXAMPLE
PS> .\Add-BlackBorderOverlay.ps1 -TitleLike "ToLiss Captain Left" -Thickness 8 -TopExtra 28 -Follow

.EXAMPLE
PS> .\Add-BlackBorderOverlay.ps1 -TitleLike "ToLiss ND" -StripTitleBar -Thickness 8 -Follow
#>

param(
  [Parameter(Mandatory)] [string]$TitleLike,
  [int]$Thickness = 8,
  [int]$TopExtra = 0,
  [int]$Cover = 2,
  [int]$TimeoutSec = 20,
  [switch]$Follow,
  [switch]$StripTitleBar
)

if ($Thickness -lt 1) { $Thickness = 1 }

$pinvoke = @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;

// V4 names to avoid stale type collisions between runs
public static class OverlayNativeV4 {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int  GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);
    [DllImport("user32.dll", EntryPoint = "GetWindowLong")] private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")] private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll", EntryPoint = "SetWindowLong")] private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);
    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")] private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const UInt32 SWP_NOMOVE = 0x0002;
    public const UInt32 SWP_NOSIZE = 0x0001;
    public const UInt32 SWP_NOACTIVATE = 0x0010;
    public const UInt32 SWP_SHOWWINDOW = 0x0040;
    public const UInt32 SWP_NOZORDER = 0x0004;
    public const UInt32 SWP_FRAMECHANGED = 0x0020;
    public const int GWL_STYLE = -16;
    public const int WS_CAPTION = 0x00C00000;
    public const int WS_THICKFRAME = 0x00040000;
    public const int WS_SYSMENU = 0x00080000;
    public const int WS_MINIMIZEBOX = 0x00020000;
    public const int WS_MAXIMIZEBOX = 0x00010000;

    public static void EnablePerMonitorDpi() {
        try { SetProcessDpiAwarenessContext(new IntPtr(-4)); }
        catch { try { SetProcessDPIAware(); } catch { } }
    }

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

public class BorderOverlayFormV4 : Form {
    private int _tL, _tT, _tR, _tB;
    private int _cover;

    public BorderOverlayFormV4(int x, int y, int w, int h, int tLeft, int tTop, int tRight, int tBottom, int cover) {
        this.StartPosition = FormStartPosition.Manual;
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.TopMost = true;
        this.BackColor = Color.Black;
        this.Opacity = 1.0; // solid black
        _tL = Math.Max(1, tLeft);
        _tT = Math.Max(1, tTop);
        _tR = Math.Max(1, tRight);
        _tB = Math.Max(1, tBottom);
        _cover = Math.Max(0, cover);
        SetBounds(x, y, w, h);
        UpdateRegion(w, h);
    }

    protected override CreateParams CreateParams {
        get {
            const int WS_EX_TOOLWINDOW = 0x00000080;
            const int WS_EX_TRANSPARENT = 0x00000020; // mouse click-through
            // layered not required for Region cut-out and can interfere with painting
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_TOOLWINDOW | WS_EX_TRANSPARENT;
            return cp;
        }
    }

    protected override void OnShown(EventArgs e) {
        base.OnShown(e);
        ForceTopMost();
    }

    public void ForceTopMost() {
        OverlayNativeV4.SetWindowPos(this.Handle, OverlayNativeV4.HWND_TOPMOST, 0, 0, 0, 0,
            OverlayNativeV4.SWP_NOMOVE | OverlayNativeV4.SWP_NOSIZE | OverlayNativeV4.SWP_NOACTIVATE | OverlayNativeV4.SWP_SHOWWINDOW);
        this.TopMost = true;
        this.BringToFront();
    }

    public void UpdateBoundsAndThickness(int x, int y, int w, int h, int tLeft, int tTop, int tRight, int tBottom, int cover) {
        _tL = Math.Max(1, tLeft);
        _tT = Math.Max(1, tTop);
        _tR = Math.Max(1, tRight);
        _tB = Math.Max(1, tBottom);
        _cover = Math.Max(0, cover);
        this.Bounds = new Rectangle(x, y, Math.Max(1, w), Math.Max(1, h));
        UpdateRegion(this.Width, this.Height);
    }

    private void UpdateRegion(int w, int h) {
        var outer = new Rectangle(0, 0, Math.Max(1, w), Math.Max(1, h));
        var r = new Region(outer);
        int innerW = Math.Max(0, w - (_tL + _tR));
        int innerH = Math.Max(0, h - (_tT + _tB));
        int innerX = Math.Max(0, _tL + _cover);
        int innerY = Math.Max(0, _tT + _cover);
        innerW = Math.Max(0, innerW - (2 * _cover));
        innerH = Math.Max(0, innerH - (2 * _cover));
        if (innerW > 0 && innerH > 0) {
            var inner = new Rectangle(innerX, innerY, innerW, innerH);
            r.Exclude(inner); // cut a hole so only the border shows
        }
        this.Region = r;
    }
}
"@;

Add-Type -AssemblyName System.Windows.Forms, System.Drawing | Out-Null
# Always try to add; ignore 'already exists' errors so reruns work
try {
  Add-Type -TypeDefinition $pinvoke -Language CSharp -ReferencedAssemblies System.Windows.Forms, System.Drawing -ErrorAction Stop | Out-Null
} catch {
  if ($_.FullyQualifiedErrorId -notlike 'TYPE_ALREADY_EXISTS*') { throw }
}

# Ensure DPI awareness so WinForms coordinates match GetWindowRect
[OverlayNativeV4]::EnablePerMonitorDpi()

function Get-OpenWindows {
  $list = New-Object System.Collections.Generic.List[object]
  [OverlayNativeV4]::EnumWindows({
    param([IntPtr]$h, [IntPtr]$p)
    if (-not [OverlayNativeV4]::IsWindowVisible($h)) { return $true }
    $len = [OverlayNativeV4]::GetWindowTextLength($h)
    if ($len -le 0) { return $true }
    $sb = New-Object System.Text.StringBuilder ($len + 1)
    [void][OverlayNativeV4]::GetWindowText($h, $sb, $sb.Capacity)
    $title = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($title)) { return $true }

    [OverlayNativeV4+RECT]$r = New-Object 'OverlayNativeV4+RECT'
    [void][OverlayNativeV4]::GetWindowRect($h, [ref]$r)
    $ww = [Math]::Max(0, $r.Right - $r.Left)
    $hh = [Math]::Max(0, $r.Bottom - $r.Top)
    if ($ww -le 0 -or $hh -le 0) { return $true }

    $list.Add([pscustomobject]@{
      Handle = $h
      Title  = $title
      X      = $r.Left
      Y      = $r.Top
      Width  = $ww
      Height = $hh
    }) | Out-Null
    return $true
  }, [IntPtr]::Zero) | Out-Null
  $list
}

function Get-WindowRectByHandle {
  param([Parameter(Mandatory)][IntPtr]$Handle)
  try {
    [OverlayNativeV4+RECT]$r = New-Object 'OverlayNativeV4+RECT'
    $ok = [OverlayNativeV4]::GetWindowRect($Handle, [ref]$r)
    if (-not $ok) { return $null }
    $ww = [Math]::Max(0, $r.Right - $r.Left)
    $hh = [Math]::Max(0, $r.Bottom - $r.Top)
    if ($ww -le 0 -or $hh -le 0) { return $null }
    return [pscustomobject]@{ X = $r.Left; Y = $r.Top; Width = $ww; Height = $hh }
  } catch { return $null }
}

$deadline = (Get-Date).AddSeconds($TimeoutSec)
$target = $null
do {
  $target = Get-OpenWindows | Where-Object { $_.Title -like "*$TitleLike*" } | Select-Object -First 1
  if ($target) { break }
  Start-Sleep -Milliseconds 200
} while ((Get-Date) -lt $deadline)

if (-not $target) {
  Write-Warning "Window '$TitleLike' not found within ${TimeoutSec}s."
  return
}

Write-Host "Overlaying '$($target.Title)' with ${Thickness}px black border. Press Ctrl+C to stop."

if ($StripTitleBar) {
  # Remove caption/frame but keep same outer bounds
  [OverlayNativeV4]::StripTitleBarKeepBounds($target.Handle, [int]$target.X, [int]$target.Y, [int]$target.Width, [int]$target.Height)
}

# Thickness per edge (thicker top when TopExtra>0)
$tL = [Math]::Max(1, $Thickness)
$tT = [Math]::Max(1, $Thickness + $TopExtra)
$tR = [Math]::Max(1, $Thickness)
$tB = [Math]::Max(1, $Thickness)

# Create initial overlay and ensure it paints by pumping messages
$overlay = New-Object BorderOverlayFormV4 @($target.X, $target.Y, $target.Width, $target.Height, $tL, $tT, $tR, $tB, $Cover)
$null = $overlay.Show()
[System.Windows.Forms.Application]::DoEvents()
Start-Sleep -Milliseconds 50
$overlay.ForceTopMost()

if ($Follow) {
  try {
    while ($true) {
      $rect = Get-WindowRectByHandle -Handle $target.Handle
      if (-not $rect) { break }
      if ($rect.X -ne $overlay.Left -or $rect.Y -ne $overlay.Top -or $rect.Width -ne $overlay.Width -or $rect.Height -ne $overlay.Height) {
        $overlay.UpdateBoundsAndThickness($rect.X, $rect.Y, $rect.Width, $rect.Height, $tL, $tT, $tR, $tB, $Cover)
        $overlay.ForceTopMost()
      }
      [System.Windows.Forms.Application]::DoEvents()
      Start-Sleep -Milliseconds 100
    }
  } finally {
    if ($overlay) { $overlay.Close() }
  }
} else {
  # Keep overlay alive until the target window closes or user stops the script
  try {
    while ($true) {
      $rect = Get-WindowRectByHandle -Handle $target.Handle
      if (-not $rect) { break }
      [System.Windows.Forms.Application]::DoEvents()
      Start-Sleep -Milliseconds 250
    }
  } finally {
    if ($overlay) { $overlay.Close() }
  }
}
