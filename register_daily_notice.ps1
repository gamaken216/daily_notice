# Daily Notice - Task Scheduler registration script
# Run this script with PowerShell as Administrator

$taskName = "DailyNotice"

# Automatically get the folder where this script is located
$scriptDir = $PSScriptRoot
$scriptPath = Join-Path $scriptDir "daily_notice.py"

# Automatically find Python / pythonw
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = $pythonPath -replace "python\.exe$", "pythonw.exe"

# Remove existing task if present
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action   = New-ScheduledTaskAction -Execute $pythonwPath -Argument ('"' + $scriptPath + '"')
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Daily Notice at PC startup"

Write-Host "Task registered successfully. Auto-starts on next login." -ForegroundColor Green
Write-Host "Script path: $scriptPath" -ForegroundColor Cyan
