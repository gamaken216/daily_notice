# Daily Notice - タスクスケジューラー登録スクリプト
# このスクリプトを管理者権限の PowerShell で実行してください

$taskName = "DailyNotice"

# このスクリプトが置かれているフォルダを自動取得
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $scriptDir "daily_notice.py"

# Python / pythonw のパスを自動取得
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = $pythonPath -replace "python\.exe$", "pythonw.exe"

# 既存タスクがあれば削除
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action   = New-ScheduledTaskAction -Execute $pythonwPath -Argument ('"' + $scriptPath + '"')
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Daily Notice at PC startup"

Write-Host "✅ タスクの登録が完了しました。次回ログオン時から自動起動します。" -ForegroundColor Green
Write-Host "   スクリプトパス: $scriptPath" -ForegroundColor Cyan
