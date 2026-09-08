<#
  setup_schedule.ps1  -  register the daily IntelliBI SEO Walk-In report task.

  Mirrors the IntelliBI_Operations_Automation scheduling approach: it creates a
  Windows Scheduled Task that runs the project's virtual-environment Python
  against scripts/run_seo_reports.py once a day. Weekly + Monthly reports are
  generated EVERY day (current period to date vs previous equal period).

  Run this ONCE, in an elevated PowerShell, from the project root:

      cd "C:\Users\<you>\Documents\IntelliBI Automation\IntelliBI_SEO_Automation"
      powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1

  Re-running updates the existing task in place.
#>

param(
    [string]$Time = "19:00",                       # daily run time (24h, local) - 7:00 PM
    [string]$TaskName = "IntelliBI_SEO_WalkIn_Report"
)

$ErrorActionPreference = "Stop"

# Project root = parent of this script's folder (portable; no hard-coded path).
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

# Prefer the project virtual-env Python; fall back to venv one level up, else PATH python.
$PyCandidates = @(
    (Join-Path $ProjectRoot ".venv\Scripts\python.exe"),
    (Join-Path (Split-Path -Parent $ProjectRoot) ".venv\Scripts\python.exe")
)
$Python = $null
foreach ($p in $PyCandidates) { if (Test-Path $p) { $Python = $p; break } }
if (-not $Python) {
    Write-Warning "No .venv Python found. Create one first:  python -m venv .venv ; .\.venv\Scripts\pip install -r requirements.txt"
    $Python = "python"
}

$Runner = Join-Path $ProjectRoot "scripts\run_seo_reports.py"
if (-not (Test-Path $Runner)) { throw "Runner not found: $Runner" }

Write-Host "Python : $Python"
Write-Host "Runner : $Runner"
Write-Host "Time   : $Time daily"
Write-Host "Task   : $TaskName"

$Action  = New-ScheduledTaskAction -Execute $Python -Argument "`"$Runner`"" -WorkingDirectory $ProjectRoot
$Trigger = New-ScheduledTaskTrigger -Daily -At $Time
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger `
    -Settings $Settings -Description "IntelliBI SEO Walk-In Analytics - daily Weekly + Monthly reports" `
    -Force | Out-Null

Write-Host ""
Write-Host "Registered. Verify with:  Get-ScheduledTaskInfo -TaskName $TaskName"
Write-Host "Run once now with:         Start-ScheduledTask -TaskName $TaskName"
