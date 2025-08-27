Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Paths
$downloadsPath  = [IO.Path]::Combine($env:USERPROFILE,'Downloads')
$processedPath  = Join-Path $downloadsPath 'vv_processed'
if (-not (Test-Path $processedPath)) { New-Item -ItemType Directory -Path $processedPath | Out-Null }

# Tray icon
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon  = [System.Drawing.SystemIcons]::Application
$notify.Text  = 'VV Auto-Launcher'
$notify.Visible = $true

# Menu
$menu = New-Object System.Windows.Forms.ContextMenuStrip
$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem 'Exit'
$menu.Items.Add($exitItem) | Out-Null
$notify.ContextMenuStrip = $menu

# Timer-based poller
$busy = $false
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000  # 1s

$timer.add_Tick({
    if ($busy) { return }
    $busy = $true
    try {
        Get-ChildItem -Path $downloadsPath -Filter '*.vv' -ErrorAction SilentlyContinue | ForEach-Object {
            $file = $_.FullName

            # Wait until file is no longer locked (browser still writing)
            $ready = $false
            for ($i=0; $i -lt 10 -and -not $ready; $i++) {
                try {
                    $fs = [IO.File]::Open($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
                    $fs.Dispose(); $ready = $true
                } catch { Start-Sleep -Milliseconds 200 }
            }
            if (-not $ready) { return }

            # Launch & focus
            $proc = Start-Process "C:\Program Files\VirtViewer v11.0-256\bin\remote-viewer.exe" -ArgumentList $file -PassThru
            Start-Sleep -Milliseconds 1200
            [Microsoft.VisualBasic.Interaction]::AppActivate($proc.Id) | Out-Null

            # Move to processed (so it doesn't re-trigger)
            $dest = Join-Path $processedPath ([IO.Path]::GetFileName($file))
            try { Move-Item -Path $file -Destination $dest -Force } catch { Remove-Item -Path $file -Force }
        }
    } finally { $busy = $false }
})

# Exit click
$exitItem.add_Click({
    $timer.Stop()
    $notify.Visible = $false
    $notify.Dispose()
    [System.Windows.Forms.Application]::Exit()
})

$timer.Start()
[System.Windows.Forms.Application]::Run()
