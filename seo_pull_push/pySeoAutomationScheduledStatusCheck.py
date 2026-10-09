# ===== Register the SEO daily task on the pc =====
$Time     = "11:30AM"     # <-- set your preferred daily run time
$Root     = "C:\Users\pc\PycharmProjects\IntelliBI_SEO_Automation"
$Python   = Join-Path $Root ".venv\Scripts\python.exe"
$Runner   = Join-Path $Root "scripts\run_seo_reports.py"
$TaskName = "IntelliBI SEO Automation"

$Action   = New-ScheduledTaskAction -Execute $Python -Argument "`"$Runner`"" -WorkingDirectory $Root
$Trigger  = New-ScheduledTaskTrigger -Daily -At $Time
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "IntelliBI SEO Walk-In Analytics - daily Weekly + Monthly reports" -Force

# Check the Schedule
Get-ScheduledTaskInfo -TaskName "IntelliBI SEO Automation"

# Manually Execute it.
Start-ScheduledTask   -TaskName "IntelliBI SEO Automation"   # optional: run it once now to test