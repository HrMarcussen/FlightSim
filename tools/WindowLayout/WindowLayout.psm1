<#
WindowLayout PowerShell module
Exports commands to capture and restore window layouts on Windows.
#>

# Import shared common module
$commonModule = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent)) 'Common\FlightSim.Common.psm1'
if (Test-Path $commonModule) {
  Import-Module $commonModule -ErrorAction Stop
}
else {
  Write-Warning "FlightSim.Common module not found at $commonModule"
}

function Set-Window {
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TitleLike,
    [Parameter(Mandatory)][int]$X,
    [Parameter(Mandatory)][int]$Y,
    [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$Width,
    [Parameter(Mandatory)][ValidateRange(1, [int]::MaxValue)][int]$Height,
    [switch]$FirstOnly,
    [ValidateRange(1, 600)][int]$TimeoutSec = 20
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
      [void][FlightSim.Common.Win32Native]::ShowWindowAsync($t.Handle, [FlightSim.Common.Win32Native]::SW_RESTORE)
      $ok = [FlightSim.Common.Win32Native]::SetWindowPos($t.Handle, [FlightSim.Common.Win32Native]::HWND_TOP, $X, $Y, $Width, $Height,
        [FlightSim.Common.Win32Native]::SWP_NOZORDER -bor [FlightSim.Common.Win32Native]::SWP_NOACTIVATE)
      if ($ok) { Write-Verbose ("Placed '{0}' -> {1},{2} {3}x{4}" -f $t.Title, $X, $Y, $Width, $Height) }
      else { Write-Warning "Failed to position '$($t.Title)'." }
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
      $lv = New-Object System.Windows.Forms.ListViewItem($it.Title)
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
    $btnAll.Add_Click({ foreach ($i in $list.Items) { $i.Checked = $true } })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'Select None'
    $btnNone.AutoSize = $true
    $btnNone.Location = New-Object System.Drawing.Point (120, 12)
    $btnNone.Add_Click({ foreach ($i in $list.Items) { $i.Checked = $false } })

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

    $panel.Controls.AddRange(@($btnAll, $btnNone, $btnOk, $btnCancel))
    $form.Controls.Add($list)
    $form.Controls.Add($panel)
    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    $dlg = $form.ShowDialog()
    if ($dlg -ne [System.Windows.Forms.DialogResult]::OK) { return @() }
    $checked = @()
    foreach ($i in $list.Items) { if ($i.Checked) { $checked += $i.Tag } }
    return , $checked
  }

  if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $result = New-Object System.Collections.ArrayList
    $t = New-Object System.Threading.Thread({ param($s)
        $sb = $s[0]; $items = $s[1]; $out = $s[2]
        $res = & $sb $items
        if ($null -ne $res) { [void]$out.AddRange(@($res)) }
      })
    $t.SetApartmentState('STA'); $t.IsBackground = $true
    $t.Start(@($scriptBlock, $Items, $result))
    $t.Join()
    return @($result.ToArray())
  }
  else {
    return & $scriptBlock $Items
  }
}

function Select-WindowsInteractive {
  [CmdletBinding()]
  param(
    [ValidateSet('OGV', 'Forms', 'Console')][string]$Picker = 'OGV'
  )
  $all = Get-OpenWindows | Sort-Object Title

  if ($Picker -eq 'Forms') { $picked = Show-WindowPickerCheckbox -Items $all } elseif ($Picker -eq 'OGV') { $picked = $null } else { $picked = Show-WindowPickerCheckbox -Items $all }
  if ($null -ne $picked) { return $picked }

  $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
  if ($ogv) {
    $picked = $all | Select-Object Title, Class, X, Y, Width, Height | Out-GridView -Title "Select windows to capture, then click OK" -PassThru
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

function Get-AsciiSeparators { @(' - ', ' | ', ': ') }

function Get-TitleSuggestion([string]$title) {
  if ([string]::IsNullOrWhiteSpace($title)) { return "" }
  foreach ($sep in (Get-AsciiSeparators)) {
    $idx = $title.IndexOf($sep)
    if ($idx -gt 0) { return $title.Substring(0, $idx).Trim() }
  }
  if ($title.Length -le 20) { return $title }
  return $title.Substring(0, [Math]::Min(20, $title.Length))
}

function Save-WindowLayout {
  [CmdletBinding()]
  param(
    [ValidateSet('OGV', 'Forms', 'Console')][string]$Picker = 'OGV'
  )
  $picked = Select-WindowsInteractive -Picker $Picker 
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
      TitleLike = $titleInput
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
    }
    else { "WindowLayout.json" }
  }
  else { $script:LayoutPath }
  if (-not $targetPath) { $targetPath = "WindowLayout.json" }

  try {
    $dir = Split-Path -Path $targetPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $layout | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 -Path $targetPath
    Write-Verbose ("Saved layout to {0}" -f $targetPath)
    Write-Host "Saved layout ($($layout.Count) entries) -> $targetPath"
  }
  catch {
    Write-Error "Failed to write layout '$targetPath': $($_.Exception.Message)"
  }
}

function Restore-WindowLayout {
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
    foreach ($w in $layout) {
      if (-not $w.TitleLike) { Write-Warning "Missing TitleLike in entry; skipping."; continue }
      try { $x = [int]$w.X; $y = [int]$w.Y; $wth = [int]$w.Width; $hgt = [int]$w.Height }
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
  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [Parameter()][ValidateNotNullOrEmpty()][string[]]$LayoutPath = "WindowLayout.json",
    [ValidateSet('OGV', 'Forms', 'Console')][string]$Picker = 'OGV'
  )
  $script:LayoutPath = $LayoutPath
  Enable-PerMonitorDpi
  Save-WindowLayout -Picker $Picker
}


Export-ModuleMember -Function Enable-PerMonitorDpi, Get-OpenWindows, Set-Window, Export-WindowLayout, Restore-WindowLayout
