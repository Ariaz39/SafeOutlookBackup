# TEST SCRIPT - Remove after validation
# Creates a one-time test task for today at 12:55 AM

$action = New-ScheduledTaskAction `
    -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -Argument "-NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"C:\Users\CONAMBIENTESAS\ScriptOutlookV2\backup-outlook.ps1`"" `
    -WorkingDirectory "C:\Users\CONAMBIENTESAS\ScriptOutlookV2"

$trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Tuesday, Friday `
    -At "13:15"

$triggerTest = New-ScheduledTaskTrigger `
    -Once `
    -At "12:55"

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# Fix battery settings directly on the object
$settings.DisallowStartIfOnBatteries = $false
$settings.StopIfGoingOnBatteries = $false

$task = New-ScheduledTask `
    -Action $action `
    -Trigger @($triggerTest, $trigger) `
    -Settings $settings `
    -Description "Weekly PST backup to OneDrive - Tuesdays & Fridays 13:15 --ing.ariaz"

# Set compatibility to Windows 10
$task.Settings.Compatibility = 6

Register-ScheduledTask `
    -TaskName "PST_WeeklyBackup_TEST" `
    -InputObject $task `
    -RunLevel Limited `
    -Force

Write-Host "Test task created. Will run at 12:55 AM, then every Tuesday & Friday at 13:15" -ForegroundColor Yellow