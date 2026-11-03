<#
.SYNOPSIS
    Joro RGB - lightweight Synapse-free colour/brightness app for the Razer Joro (Bluetooth).

.DESCRIPTION
    Runs on the built-in Windows PowerShell + WinForms (no compiled binaries, so it works with
    Smart App Control enabled). Bluetooth work runs on a background runspace so the window stays
    responsive. Settings: %APPDATA%\JoroRGB\settings.json

.PARAMETER Apply
    Headless: apply the saved colour/brightness and exit (used by the sign-in entry).

.PARAMETER Install
    Create a Start menu shortcut ("Joro RGB") and exit.
#>
[CmdletBinding()]
param([switch]$Apply, [switch]$Install)

$ErrorActionPreference = 'Stop'
$CorePath = Join-Path $PSScriptRoot 'JoroCore.ps1'
$DataDir = Join-Path $env:APPDATA 'JoroRGB'
$SettingsPath = Join-Path $DataDir 'settings.json'
$IconPath = Join-Path $DataDir 'joro.ico'
$RunKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$RunValue = 'JoroRGB'
# conhost --headless avoids a console window flashing up.
$Launch = '"{0}\System32\conhost.exe" --headless "{0}\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -File "{1}"' -f $env:SystemRoot, $PSCommandPath

function Read-Settings {
    $s = @{ Color = 'FFFFFF'; Brightness = 100 }
    if (Test-Path -LiteralPath $SettingsPath) {
        try {
            $j = Get-Content -LiteralPath $SettingsPath -Raw | ConvertFrom-Json
            if ($j.Color -match '^[0-9A-Fa-f]{6}$') { $s.Color = $j.Color.ToUpper() }
            if ($j.Brightness -is [int] -or $j.Brightness -is [long]) { $s.Brightness = [Math]::Max(0, [Math]::Min(100, [int]$j.Brightness)) }
        }
        catch { }
    }
    $s
}

function Save-Settings($s) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
    @{ Color = $s.Color; Brightness = $s.Brightness } | ConvertTo-Json | Set-Content -LiteralPath $SettingsPath -Encoding UTF8
}

# --------------------------------------------------------------------------- headless modes

if ($Apply) {
    . $CorePath
    $s = Read-Settings
    # Bluetooth may not be ready right after sign-in, so retry for about a minute.
    for ($i = 0; $i -lt 12; $i++) {
        try { Set-JoroLighting $s.Color $s.Brightness; Close-Joro; exit 0 }
        catch { Close-Joro; Start-Sleep -Seconds 5 }
    }
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

function New-ColorIcon([System.Drawing.Color]$Color, [int]$Size = 32) {
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $pad = [int]($Size / 8)
    $keys = New-Object System.Drawing.Rectangle $pad, ($pad * 2), ($Size - 2 * $pad), ($Size - 4 * $pad)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(40, 40, 40))), $keys)
    $glow = New-Object System.Drawing.SolidBrush $Color
    $k = [int](($keys.Width - 4) / 4)
    for ($row = 0; $row -lt 2; $row++) {
        for ($col = 0; $col -lt 4; $col++) {
            $g.FillRectangle($glow, $keys.X + 2 + $col * $k, $keys.Y + 2 + $row * [int](($keys.Height - 4) / 2), $k - 2, [int](($keys.Height - 4) / 2) - 2)
        }
    }
    $g.Dispose()
    $icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $bmp.Dispose()
    $icon
}

if ($Install) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
    $icon = New-ColorIcon ([System.Drawing.Color]::FromArgb(0, 255, 120)) 48
    $fs = [System.IO.File]::Create($IconPath)
    try { $icon.Save($fs) } finally { $fs.Dispose() }
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Programs')) 'Joro RGB.lnk'
    $shell = New-Object -ComObject WScript.Shell
    $lnk = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\conhost.exe'
    $lnk.Arguments = '--headless "{0}\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -File "{1}"' -f $env:SystemRoot, $PSCommandPath
    $lnk.WorkingDirectory = $PSScriptRoot
    $lnk.IconLocation = "$IconPath,0"
    $lnk.Description = 'Set Razer Joro colour and brightness without Synapse'
    $lnk.Save()
    Write-Output "Created $lnkPath"
    exit 0
}

# --------------------------------------------------------------------------- GUI

$mutex = New-Object System.Threading.Mutex($false, 'Local\JoroRGB')
if (-not $mutex.WaitOne(0)) { exit 0 }

# Crisp text on high-DPI screens. Emitted in memory (no compiled file on disk).
try {
    $asm = [AppDomain]::CurrentDomain.DefineDynamicAssembly((New-Object System.Reflection.AssemblyName 'JoroNative'), 'Run')
    $tb = $asm.DefineDynamicModule('JoroNative').DefineType('JoroNative', 'Public, Class')
    $null = $tb.DefinePInvokeMethod('SetProcessDPIAware', 'user32.dll', 'Public, Static, PinvokeImpl',
        [System.Reflection.CallingConventions]::Standard, [bool], [Type[]]@(),
        [System.Runtime.InteropServices.CallingConvention]::Winapi, [System.Runtime.InteropServices.CharSet]::Auto)
    $null = $tb.DefinePInvokeMethod('DwmSetWindowAttribute', 'dwmapi.dll', 'Public, Static, PinvokeImpl',
        [System.Reflection.CallingConventions]::Standard, [int], [Type[]]@([IntPtr], [int], [int].MakeByRefType(), [int]),
        [System.Runtime.InteropServices.CallingConvention]::Winapi, [System.Runtime.InteropServices.CharSet]::Auto)
    $script:Native = $tb.CreateType()
    $null = $script:Native::SetProcessDPIAware()
}
catch { $script:Native = $null }

# Theme: background from skycode-workspace.code-workspace (#111113); text/borders use
# VS Code's default dark-theme shades so the app sits alongside the editor.
$Theme = @{
    Back    = [System.Drawing.ColorTranslator]::FromHtml('#111113')
    Surface = [System.Drawing.ColorTranslator]::FromHtml('#1C1C20')
    Hover   = [System.Drawing.ColorTranslator]::FromHtml('#2A2A30')
    Border  = [System.Drawing.ColorTranslator]::FromHtml('#3C3C42')
    Text    = [System.Drawing.ColorTranslator]::FromHtml('#CCCCCC')
    Dim     = [System.Drawing.ColorTranslator]::FromHtml('#8B8B92')
    Error   = [System.Drawing.ColorTranslator]::FromHtml('#F14C4C')
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$settings = Read-Settings

# Background runspace keeps one Bluetooth connection open between changes.
$runspace = [runspacefactory]::CreateRunspace()
$runspace.ApartmentState = 'MTA'
$runspace.Open()
$worker = [powershell]::Create()
$worker.Runspace = $runspace
$state = @{ Handle = $null; Pending = $null; Job = $null }

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Joro RGB'
$form.FormBorderStyle = 'FixedSingle'
$form.MaximizeBox = $false
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.AutoScaleMode = 'None'
$form.BackColor = $Theme.Back
$form.ForeColor = $Theme.Text
$form.Add_HandleCreated({
    # Dark title bar (DWMWA_USE_IMMERSIVE_DARK_MODE = 20; 19 on older Windows 10 builds).
    if ($script:Native) {
        $on = 1
        if ($script:Native::DwmSetWindowAttribute($form.Handle, 20, [ref]$on, 4) -ne 0) {
            $null = $script:Native::DwmSetWindowAttribute($form.Handle, 19, [ref]$on, 4)
        }
    }
})

$scale = 1.0
$g = $form.CreateGraphics(); $scale = $g.DpiX / 96.0; $g.Dispose()
function S([int]$v) { [int][Math]::Round($v * $scale) }
$form.ClientSize = New-Object System.Drawing.Size (S 360), (S 300)

function Add-Control($ctl, [int]$x, [int]$y, [int]$w, [int]$h) {
    $ctl.Location = New-Object System.Drawing.Point (S $x), (S $y)
    $ctl.Size = New-Object System.Drawing.Size (S $w), (S $h)
    if ($ctl -is [System.Windows.Forms.Button]) {
        $ctl.FlatStyle = 'Flat'
        $ctl.BackColor = $Theme.Surface
        $ctl.ForeColor = $Theme.Text
        $ctl.FlatAppearance.BorderColor = $Theme.Border
        $ctl.FlatAppearance.MouseOverBackColor = $Theme.Hover
        $ctl.FlatAppearance.MouseDownBackColor = $Theme.Border
    }
    elseif ($ctl -is [System.Windows.Forms.TextBox]) {
        $ctl.BorderStyle = 'FixedSingle'
        $ctl.BackColor = $Theme.Surface
        $ctl.ForeColor = $Theme.Text
    }
    elseif ($ctl -is [System.Windows.Forms.CheckBox]) {
        $ctl.FlatStyle = 'Flat'
        $ctl.ForeColor = $Theme.Text
        $ctl.FlatAppearance.BorderColor = $Theme.Border
        $ctl.FlatAppearance.CheckedBackColor = $Theme.Surface
    }
    else {
        $ctl.BackColor = $Theme.Back
        $ctl.ForeColor = $Theme.Text
    }
    $form.Controls.Add($ctl)
    $ctl
}

function To-Color([string]$hex) {
    [System.Drawing.Color]::FromArgb([Convert]::ToInt32($hex, 16) -bor 0xFF000000)
}
function To-Hex([System.Drawing.Color]$c) { '{0:X2}{1:X2}{2:X2}' -f $c.R, $c.G, $c.B }

$swatch = Add-Control (New-Object System.Windows.Forms.Panel) 14 14 64 64
$swatch.Cursor = 'Hand'
$swatch.Add_Paint({
    param($sender, $e)
    $pen = New-Object System.Drawing.Pen $Theme.Border
    $e.Graphics.DrawRectangle($pen, 0, 0, $sender.Width - 1, $sender.Height - 1)
    $pen.Dispose()
})

$pick = Add-Control (New-Object System.Windows.Forms.Button) 92 14 140 30
$pick.Text = 'Choose colour...'

$hexLabel = Add-Control (New-Object System.Windows.Forms.Label) 92 55 20 22
$hexLabel.Text = '#'
$hexLabel.TextAlign = 'MiddleRight'
$hexLabel.ForeColor = $Theme.Dim
$hexBox = Add-Control (New-Object System.Windows.Forms.TextBox) 112 54 80 24
$hexBox.MaxLength = 6
$hexBox.CharacterCasing = 'Upper'

$presetColors = 'FF0000', 'FF6000', 'FFDC00', '00FF00', '00FFFF', '0000FF', 'A000FF', 'FFFFFF'
$tips = New-Object System.Windows.Forms.ToolTip
for ($i = 0; $i -lt $presetColors.Count; $i++) {
    $b = Add-Control (New-Object System.Windows.Forms.Button) (14 + $i * 42) 92 36 36
    $b.BackColor = To-Color $presetColors[$i]
    $b.FlatAppearance.MouseOverBackColor = $b.BackColor
    $b.FlatAppearance.MouseDownBackColor = $b.BackColor
    $b.Tag = $presetColors[$i]
    $b.Cursor = 'Hand'
    $tips.SetToolTip($b, '#' + $presetColors[$i])
    $b.Add_Click({ param($sender) Set-UiColor $sender.Tag $true })
}

$brightLabel = Add-Control (New-Object System.Windows.Forms.Label) 14 142 120 20
$brightLabel.Text = 'Brightness'
$pctLabel = Add-Control (New-Object System.Windows.Forms.Label) 270 142 76 20
$pctLabel.TextAlign = 'TopRight'
$pctLabel.ForeColor = $Theme.Dim

$slider = Add-Control (New-Object System.Windows.Forms.TrackBar) 8 162 344 36
$slider.Minimum = 0
$slider.Maximum = 100
$slider.TickStyle = 'None'
$slider.SmallChange = 5
$slider.LargeChange = 10

$applyBtn = Add-Control (New-Object System.Windows.Forms.Button) 14 210 100 30
$applyBtn.Text = 'Apply'
$form.AcceptButton = $applyBtn

$startup = Add-Control (New-Object System.Windows.Forms.CheckBox) 130 213 220 24
$startup.Text = 'Apply at Windows sign-in'
$startup.Checked = [bool](Get-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue)

$status = Add-Control (New-Object System.Windows.Forms.Label) 14 252 334 40
$status.ForeColor = $Theme.Dim
$status.Text = 'Pick a colour or press Apply. Razer Synapse must be closed.'

$debounce = New-Object System.Windows.Forms.Timer
$debounce.Interval = 300
$poll = New-Object System.Windows.Forms.Timer
$poll.Interval = 100

function Set-Status([string]$text, [bool]$isError = $false) {
    $status.Text = $text
    $status.ForeColor = if ($isError) { $Theme.Error } else { $Theme.Dim }
}

function Start-LightingJob([hashtable]$job) {
    $worker.Commands.Clear()
    $null = $worker.AddScript('param($c, $b) Set-JoroLighting $c $b').AddArgument($job.Color).AddArgument($job.Brightness)
    $state.Job = $job
    $state.Handle = $worker.BeginInvoke()
    if (-not $poll.Enabled) { $poll.Start() }
}

function Submit-Lighting([string]$color, [int]$brightness = -1) {
    Save-Settings $settings
    $job = @{ Color = $color; Brightness = $brightness }
    if ($state.Handle) {
        # Busy: keep only the latest requested state.
        if (-not $state.Pending) { $state.Pending = @{ Color = ''; Brightness = -1 } }
        if ($color) { $state.Pending.Color = $color }
        if ($brightness -ge 0) { $state.Pending.Brightness = $brightness }
        return
    }
    Set-Status 'Applying...'
    Start-LightingJob $job
}

function Set-UiColor([string]$hex, [bool]$apply) {
    $settings.Color = $hex
    $c = To-Color $hex
    $swatch.BackColor = $c
    if ($hexBox.Text -ne $hex) { $hexBox.Text = $hex }
    $old = $form.Icon
    $form.Icon = New-ColorIcon $c
    if ($old) { $old.Dispose() }
    if ($apply) { Submit-Lighting $hex }
}

$poll.Add_Tick({
    if (-not $state.Handle -or -not $state.Handle.IsCompleted) { return }
    $job = $state.Job
    try {
        $worker.EndInvoke($state.Handle) | Out-Null
        $parts = @()
        if ($job.Color) { $parts += "#$($job.Color)" }
        if ($job.Brightness -ge 0) { $parts += "$($job.Brightness)%" }
        Set-Status ('Applied ' + ($parts -join ' at '))
    }
    catch {
        $e = $_.Exception
        while ($e.InnerException) { $e = $e.InnerException }
        Set-Status $e.Message $true
    }
    $state.Handle = $null
    if ($state.Pending) {
        $next = $state.Pending
        $state.Pending = $null
        Set-Status 'Applying...'
        Start-LightingJob $next
    }
    else { $poll.Stop() }
})

$debounce.Add_Tick({
    $debounce.Stop()
    Submit-Lighting '' $slider.Value
})

$slider.Add_ValueChanged({
    $pctLabel.Text = "$($slider.Value)%"
    $settings.Brightness = $slider.Value
    $debounce.Stop()
    $debounce.Start()
})

$openPicker = {
    $dlg = New-Object System.Windows.Forms.ColorDialog
    $dlg.FullOpen = $true
    $dlg.Color = To-Color $settings.Color
    if ($dlg.ShowDialog($form) -eq 'OK') { Set-UiColor (To-Hex $dlg.Color) $true }
    $dlg.Dispose()
}
$pick.Add_Click($openPicker)
$swatch.Add_Click($openPicker)

$hexBox.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Enter') {
        $e.SuppressKeyPress = $true
        if ($hexBox.Text -match '^[0-9A-F]{6}$') { Set-UiColor $hexBox.Text $true }
        else { Set-Status 'Enter a 6-digit hex colour, e.g. FF8800.' $true }
    }
})

$applyBtn.Add_Click({
    if ($hexBox.Text -match '^[0-9A-F]{6}$' -and $hexBox.Text -ne $settings.Color) { Set-UiColor $hexBox.Text $false }
    Submit-Lighting $settings.Color $settings.Brightness
})

$startup.Add_Click({
    try {
        if ($startup.Checked) {
            Save-Settings $settings
            Set-ItemProperty -Path $RunKey -Name $RunValue -Value ($Launch + ' -Apply')
            Set-Status 'Saved colour will be applied at sign-in.'
        }
        else {
            Remove-ItemProperty -Path $RunKey -Name $RunValue -ErrorAction SilentlyContinue
            Set-Status 'Sign-in apply turned off.'
        }
    }
    catch {
        $startup.Checked = -not $startup.Checked
        Set-Status $_.Exception.Message $true
    }
})

$form.Add_FormClosing({
    $poll.Stop(); $debounce.Stop()
    try {
        if (-not $state.Handle) {
            $worker.Commands.Clear()
            $null = $worker.AddScript('Close-Joro').Invoke()
        }
    }
    catch { }
})

# Load the Bluetooth core on the worker; the first Apply connects.
$null = $worker.AddScript(". '$($CorePath -replace "'", "''")'").Invoke()
if ($worker.HadErrors) { [System.Windows.Forms.MessageBox]::Show("Failed to load JoroCore.ps1:`n$($worker.Streams.Error[0])", 'Joro RGB', 'OK', 'Error') | Out-Null; exit 1 }

Set-UiColor $settings.Color $false
$slider.Value = $settings.Brightness
$pctLabel.Text = "$($settings.Brightness)%"
$debounce.Stop()

try {
    [System.Windows.Forms.Application]::Run($form)
}
finally {
    $worker.Dispose()
    $runspace.Dispose()
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
