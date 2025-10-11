<#
WindowLayout PowerShell module
Exports commands to capture and restore window layouts on Windows.
#>

# Guard native type definition so re-imports are safe
$__typeLoaded = $true; try { [void][Win32Native] } catch { $__typeLoaded = $false }
if (-not $__typeLoaded) {
  Add-Type -TypeDefinition @"
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

    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);
    [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

    public static readonly IntPtr HWND_TOP = IntPtr.Zero;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public const int  SW_RESTORE = 9;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
}
"@ -Language CSharp
}

function Enable-PerMonitorDpi {
  [CmdletBinding()] param()
  try { [void][Win32Native]::SetProcessDpiAwarenessContext([IntPtr]::new(-4)) } catch {
    try { [void][Win32Native]::SetProcessDPIAware() } catch {}
  }
}

function Get-OpenWindows {
  [CmdletBinding()] param([switch]$VisibleOnly = $true)
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

function Set-Window {
  [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TitleLike,
    [Parameter(Mandatory)][int]$X,
    [Parameter(Mandatory)][int]$Y,
    [Parameter(Mandatory)][ValidateRange(1,[int]::MaxValue)][int]$Width,
    [Parameter(Mandatory)][ValidateRange(1,[int]::MaxValue)][int]$Height,
    [switch]$FirstOnly,
    [ValidateRange(1,600)][int]$TimeoutSec = 20
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
    if ($PSCmdlet.ShouldProcess($t.Title, "Move+Resize to $X,$Y ${Width}x${Height}")) {
      [void][Win32Native]::ShowWindowAsync($t.Handle, [Win32Native]::SW_RESTORE)
      $ok = [Win32Native]::SetWindowPos($t.Handle, [Win32Native]::HWND_TOP, $X, $Y, $Width, $Height,
        [Win32Native]::SWP_NOZORDER -bor [Win32Native]::SWP_NOACTIVATE)
      if ($ok) { Write-Verbose ("Placed '{0}' -> {1},{2} {3}x{4}" -f $t.Title,$X,$Y,$Width,$Height) }
      else     { Write-Warning "Failed to position '$($t.Title)'." }
    }
  }
}

function Show-WindowPickerCheckbox {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][System.Collections.IEnumerable]$Items
  )
  try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop; Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch { return $null }

  $scriptBlock = {
    param($Items)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select windows to capture'
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size (1100, 700)
    $form.MinimumSize = New-Object System.Drawing.Size (900, 560)
    $form.AutoScaleMode = 'Dpi'

    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.CheckBoxes = $true
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.Dock = 'Fill'

    [void]$list.Columns.Add('Title', 500)
    [void]$list.Columns.Add('Class', 240)
    [void]$list.Columns.Add('Location', 160)
    [void]$list.Columns.Add('Size', 160)

    foreach ($it in $Items) {
      $loc = '{0},{1}' -f $it.X, $it.Y
      $siz = '{0}x{1}' -f $it.Width, $it.Height
      $lv  = New-Object System.Windows.Forms.ListViewItem($it.Title)
      [void]$lv.SubItems.Add($it.Class)
      [void]$lv.SubItems.Add($loc)
      [void]$lv.SubItems.Add($siz)
      $lv.Tag = $it
      [void]$list.Items.Add($lv)
    }

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Bottom'
    $panel.Height = 56

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'Select All'
    $btnAll.AutoSize = $true
    $btnAll.Location = New-Object System.Drawing.Point (12, 12)
    $btnAll.Add_Click({ foreach($i in $list.Items){ $i.Checked = $true } })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Select None'
    $btnNone.AutoSize = $true
    $btnNone.Location = New-Object System.Drawing.Point (120, 12)
    $btnNone.Add_Click({ foreach($i in $list.Items){ $i.Checked = $false } })

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Size = New-Object System.Drawing.Size (75, 30)
    $btnOk.Anchor = 'Bottom,Right'
    $btnOk.Location = New-Object System.Drawing.Point (850, 12)
    $btnOk.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK; $form.Close() })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size (75, 30)
    $btnCancel.Anchor = 'Bottom,Right'
    $btnCancel.Location = New-Object System.Drawing.Point (930, 12)
    $btnCancel.Add_Click({ $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

    $form.Add_Resize({
      $btnOk.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - 190, 12)
      $btnCancel.Location = New-Object System.Drawing.Point ($form.ClientSize.Width - 100, 12)
    })

    $panel.Controls.AddRange(@($btnAll,$btnNone,$btnOk,$btnCancel))
    $form.Controls.Add($list)
    $form.Controls.Add($panel)
    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    $dlg = $form.ShowDialog()
    if ($dlg -ne [System.Windows.Forms.DialogResult]::OK) { return @() }
    $checked = @()
    foreach($i in $list.Items){ if($i.Checked){ $checked += $i.Tag } }
    return ,$checked
  }

  if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $result = New-Object System.Collections.ArrayList
    $t = New-Object System.Threading.Thread({ param($s)
      $sb=$s[0]; $items=$s[1]; $out=$s[2]
      $res = & $sb $items
      if ($res -ne $null) { [void]$out.AddRange(@($res)) }
    })
    $t.SetApartmentState('STA'); $t.IsBackground = $true
    $t.Start(@($scriptBlock, $Items, $result))
    $t.Join()
    return @($result.ToArray())
  } else {
    return & $scriptBlock $Items
  }
}

function Select-WindowsInteractive {
  [CmdletBinding()]
  param(
    [ValidateSet('OGV','Forms','Console','Auto')][string]$Picker = 'Auto',
    [switch]$ForceLightTheme
  )
  $all = Get-OpenWindows | Sort-Object Title

  if ($Picker -eq 'Forms') { $picked = Show-WindowPickerCheckbox -Items $all  } elseif ($Picker -eq 'OGV') { $picked = $null } else { $picked = Show-WindowPickerCheckbox -Items $all  }
  if ($picked -ne $null) { return $picked }

  $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
  if ($ogv) {
    $picked = $all | Select-Object Title,Class,X,Y,Width,Height | Out-GridView -Title "Select windows to capture, then click OK" -PassThru
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

function Get-AsciiSeparators { @(' - ', ' | ', ': ') }

function Suggest-TitleLikeSimple([string]$title) {
  if ([string]::IsNullOrWhiteSpace($title)) { return "" }
  foreach ($sep in (Get-AsciiSeparators)) {
    $idx = $title.IndexOf($sep)
    if ($idx -gt 0) { return $title.Substring(0, $idx).Trim() }
  }
  if ($title.Length -le 20) { return $title }
  return $title.Substring(0, [Math]::Min(20, $title.Length))
}

function Capture-Layout {
  [CmdletBinding()]
  param(
    [ValidateSet('OGV','Forms','Console','Auto')][string]$Picker = 'Auto',
    [switch]$ForceLightTheme
  )
  $picked = Select-WindowsInteractive -Picker $Picker 
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
    }
  }

  $targetPath = if ($script:LayoutPath -is [array]) {
    if ($script:LayoutPath.Count -gt 0) {
      if ($script:LayoutPath.Count -gt 1) { Write-Warning "Multiple LayoutPath values provided; using first: $($script:LayoutPath[0])" }
      $script:LayoutPath[0]
    } else { "WindowLayout.json" }
  } else { $script:LayoutPath }
  if (-not $targetPath) { $targetPath = "WindowLayout.json" }

  try {
    $dir = Split-Path -Path $targetPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $layout | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -Path $targetPath
    Write-Verbose ("Saved layout to {0}" -f $targetPath)
    Write-Host "Saved layout ($($layout.Count) entries) -> $targetPath"
  } catch {
    Write-Error "Failed to write layout '$targetPath': $($_.Exception.Message)"
  }
}

function Apply-Layout {
  [CmdletBinding()] param()
  $paths = @($script:LayoutPath)
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
    foreach ($w in $layout) {
      if (-not $w.TitleLike) { Write-Warning "Missing TitleLike in entry; skipping."; continue }
      try { $x=[int]$w.X; $y=[int]$w.Y; $wth=[int]$w.Width; $hgt=[int]$w.Height }
      catch { Write-Warning "Invalid numeric values in entry for '$($w.TitleLike)'; skipping."; continue }
      if ($wth -le 0 -or $hgt -le 0) { Write-Warning "Non-positive size for '$($w.TitleLike)'; skipping."; continue }
      Set-Window -TitleLike $w.TitleLike -X $x -Y $y -Width $wth -Height $hgt -TimeoutSec 20
      $count++
    }
    $totalApplied += $count
    Write-Host "Applied $count layout entrie(s) from $path"
  }

  if ($totalApplied -eq 0) {
    Write-Warning "No layout entries applied from provided path(s)."
  }
}

function Export-WindowLayout {
  [CmdletBinding(SupportsShouldProcess=$true)]
  param(
    [Parameter()][ValidateNotNullOrEmpty()][string[]]$LayoutPath = "WindowLayout.json",
    [ValidateSet('OGV','Forms','Console')][string]$Picker = 'OGV'
  )
  $script:LayoutPath = $LayoutPath
  Enable-PerMonitorDpi
  Capture-Layout -Picker $Picker
}

