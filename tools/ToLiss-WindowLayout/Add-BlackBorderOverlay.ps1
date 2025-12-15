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
#>
param(
  [Parameter(Mandatory)] [string]$TitleLike,
  [int]$Thickness = 8,
  [int]$TopExtra = 0,
  [int]$Cover = 2,
  [int]$TopCoverExtra = 0,
  [int]$TimeoutSec = 20,
  [switch]$Follow,
  [switch]$StripTitleBar
)

if ($Thickness -lt 1) { $Thickness = 1 }

# Import shared common module
$commonModule = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent)) 'Common\FlightSim.Common.psm1'
if (Test-Path $commonModule) {
  Import-Module $commonModule -ErrorAction Stop
}
else {
  Write-Warning "FlightSim.Common module not found at $commonModule"
}

$pinvoke = @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using FlightSim.Common;

public static class OverlayNativeV5 {
     [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    public const UInt32 SWP_NOMOVE = 0x0002;
    public const UInt32 SWP_NOSIZE = 0x0001;
    public const UInt32 SWP_NOACTIVATE = 0x0010;
    public const UInt32 SWP_SHOWWINDOW = 0x0040;
}

public class BorderOverlayFormV5 : Form {
    private int _tL, _tT, _tR, _tB;
    private int _cL, _cT, _cR, _cB;

    public BorderOverlayFormV5(int x, int y, int w, int h, int tLeft, int tTop, int tRight, int tBottom, int cLeft, int cTop, int cRight, int cBottom) {
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
        _cL = Math.Max(0, cLeft);
        _cT = Math.Max(0, cTop);
        _cR = Math.Max(0, cRight);
        _cB = Math.Max(0, cBottom);
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
        OverlayNativeV5.SetWindowPos(this.Handle, OverlayNativeV5.HWND_TOPMOST, 0, 0, 0, 0,
            OverlayNativeV5.SWP_NOMOVE | OverlayNativeV5.SWP_NOSIZE | OverlayNativeV5.SWP_NOACTIVATE | OverlayNativeV5.SWP_SHOWWINDOW);
        this.TopMost = true;
        this.BringToFront();
    }

    public void UpdateBoundsAndThickness(int x, int y, int w, int h, int tLeft, int tTop, int tRight, int tBottom, int cLeft, int cTop, int cRight, int cBottom) {
        _tL = Math.Max(1, tLeft);
        _tT = Math.Max(1, tTop);
        _tR = Math.Max(1, tRight);
        _tB = Math.Max(1, tBottom);
        _cL = Math.Max(0, cLeft);
        _cT = Math.Max(0, cTop);
        _cR = Math.Max(0, cRight);
        _cB = Math.Max(0, cBottom);
        this.Bounds = new Rectangle(x, y, Math.Max(1, w), Math.Max(1, h));
        UpdateRegion(this.Width, this.Height);
    }

    private void UpdateRegion(int w, int h) {
        var outer = new Rectangle(0, 0, Math.Max(1, w), Math.Max(1, h));
        var r = new Region(outer);
        int innerW = Math.Max(0, w - (_tL + _tR));
        int innerH = Math.Max(0, h - (_tT + _tB));
        int innerX = Math.Max(0, _tL + _cL);
        int innerY = Math.Max(0, _tT + _cT);
        innerW = Math.Max(0, innerW - (_cL + _cR));
        innerH = Math.Max(0, innerH - (_cT + _cB));
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
  Add-Type -TypeDefinition $pinvoke -Language CSharp -ReferencedAssemblies System.Windows.Forms, System.Drawing, $commonModule -ErrorAction Stop | Out-Null
}
catch {
  if ($_.FullyQualifiedErrorId -notlike 'TYPE_ALREADY_EXISTS*') { throw }
}

# Ensure DPI awareness so WinForms coordinates match GetWindowRect
Enable-PerMonitorDpi

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
  [FlightSim.Common.Win32Native]::StripTitleBarKeepBounds($target.Handle, [int]$target.X, [int]$target.Y, [int]$target.Width, [int]$target.Height)
}

# Thickness per edge (thicker top when TopExtra>0)
$tL = [Math]::Max(1, $Thickness)
$tT = [Math]::Max(1, $Thickness + $TopExtra)
$tR = [Math]::Max(1, $Thickness)
$tB = [Math]::Max(1, $Thickness)

# Create initial overlay and ensure it paints using a timer instead of a busy loop
$overlay = New-Object BorderOverlayFormV5 @($target.X, $target.Y, $target.Width, $target.Height, $tL, $tT, $tR, $tB, [Math]::Max(0, $Cover), [Math]::Max(0, $Cover + $TopCoverExtra), [Math]::Max(0, $Cover), [Math]::Max(0, $Cover))
$overlay.Show()
[System.Windows.Forms.Application]::DoEvents()
Start-Sleep -Milliseconds 50
$overlay.ForceTopMost()

if ($Follow) {
  try {
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200 # Poll every 200ms
    $timer.Add_Tick({
        param($sender, $e)
        $rect = Get-WindowRectByHandle -Handle $target.Handle
        if (-not $rect) { 
          $overlay.Close(); 
          [System.Windows.Forms.Application]::Exit()
          return 
        }
        
        if ($rect.X -ne $overlay.Left -or $rect.Y -ne $overlay.Top -or $rect.Width -ne $overlay.Width -or $rect.Height -ne $overlay.Height) {
          $overlay.UpdateBoundsAndThickness($rect.X, $rect.Y, $rect.Width, $rect.Height, $tL, $tT, $tR, $tB, [Math]::Max(0, $Cover), [Math]::Max(0, $Cover + $TopCoverExtra), [Math]::Max(0, $Cover), [Math]::Max(0, $Cover))
          $overlay.ForceTopMost()
        }
      })
    $timer.Start()
     
    # Start message loop
    [System.Windows.Forms.Application]::Run($overlay)
     
  }
  finally {
    if ($overlay) { $overlay.Close() }
  }
}
else {
  # Keep overlay alive until the target window closes or user stops the script
  try {
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # Poll slower if not following updates strictly
    $timer.Add_Tick({
        param($sender, $e)
        $rect = Get-WindowRectByHandle -Handle $target.Handle
        if (-not $rect) {
          $overlay.Close(); 
          [System.Windows.Forms.Application]::Exit()
        }
      })
    $timer.Start()
    [System.Windows.Forms.Application]::Run($overlay)
  }
  finally {
    if ($overlay) { $overlay.Close() }
  }
}
