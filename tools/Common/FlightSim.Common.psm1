<#
.SYNOPSIS
Shared common module for FlightSim tools.
Exports Win32 native types and common helper functions.
#>

$__typeLoaded = $true; try { [void][FlightSim.Common.Win32Native] } catch { $__typeLoaded = $false }
if (-not $__typeLoaded) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Text;
    using System.Runtime.InteropServices;

    namespace FlightSim.Common {
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
            
            [DllImport("user32.dll", EntryPoint = "GetWindowLong")] private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);
            [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")] private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);
            [DllImport("user32.dll", EntryPoint = "SetWindowLong")] private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);
            [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")] private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

            [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
            [DllImport("user32.dll")] public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);
            [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

            public static readonly IntPtr HWND_TOP = IntPtr.Zero;
            public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
            
            public const uint SWP_NOSIZE = 0x0001;
            public const uint SWP_NOMOVE = 0x0002;
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
    }
"@ -Language CSharp
}

function Enable-PerMonitorDpi {
    try { [void][FlightSim.Common.Win32Native]::SetProcessDpiAwarenessContext([IntPtr]::new(-4)) } catch {
        try { [void][FlightSim.Common.Win32Native]::SetProcessDPIAware() } catch {}
    }
}

function Get-OpenWindows {
    [CmdletBinding()] param([switch]$VisibleOnly = $true)
    $list = New-Object System.Collections.Generic.List[object]
    [FlightSim.Common.Win32Native]::EnumWindows({
        param([IntPtr]$h, [IntPtr]$p)
        if ($VisibleOnly -and -not [FlightSim.Common.Win32Native]::IsWindowVisible($h)) { return $true }
    
        $len = [FlightSim.Common.Win32Native]::GetWindowTextLength($h)
        if ($len -le 0) { return $true }
        $sb = New-Object System.Text.StringBuilder ($len + 1)
        [void][FlightSim.Common.Win32Native]::GetWindowText($h, $sb, $sb.Capacity)
        $title = $sb.ToString()
        if ([string]::IsNullOrWhiteSpace($title)) { return $true }
    
        $csb = New-Object System.Text.StringBuilder 256
        [void][FlightSim.Common.Win32Native]::GetClassName($h, $csb, $csb.Capacity)
        $class = $csb.ToString()
    
        [FlightSim.Common.Win32Native+RECT]$r = New-Object 'FlightSim.Common.Win32Native+RECT'
        [void][FlightSim.Common.Win32Native]::GetWindowRect($h, [ref]$r)
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

function Get-WindowRectByHandle {
    param([Parameter(Mandatory)][IntPtr]$Handle)
    try {
        [FlightSim.Common.Win32Native+RECT]$r = New-Object 'FlightSim.Common.Win32Native+RECT'
        $ok = [FlightSim.Common.Win32Native]::GetWindowRect($Handle, [ref]$r)
        if (-not $ok) { return $null }
        $ww = [Math]::Max(0, $r.Right - $r.Left)
        $hh = [Math]::Max(0, $r.Bottom - $r.Top)
        if ($ww -le 0 -or $hh -le 0) { return $null }
        return [pscustomobject]@{ X = $r.Left; Y = $r.Top; Width = $ww; Height = $hh }
    } catch { return $null }
}

Export-ModuleMember -Function Enable-PerMonitorDpi, Get-OpenWindows, Get-WindowRectByHandle
