<#
================================================================================
  IntelliBI Operations Automation — Task Scheduler registrar
  ------------------------------------------------------------------------------
  Registers ONE scheduled task that fires at 10:00, 11:00 and 12:00. The pipeline
  must succeed only ONCE per day:

    * 10:00 is the normal run.
    * 11:00 and 12:00 are FALLBACK/RETRY windows — scripts/run_scheduled.py
      (invoked with --once-per-day) checks a per-day success marker and EXITS
      immediately if the pipeline already succeeded earlier today, so the
      fallback windows never produce an extra daily run.
    * A failed early window leaves the marker unset, so the next window retries.
    * StartWhenAvailable also catches a run missed because the machine was off.

  Overlap protection: MultipleInstances = IgnoreNew, plus an OS file lock in the
  wrapper.

  RUN THIS ONCE, from an **elevated (Administrator) PowerShell**, after deployment:

        powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1

  All paths are derived from this script's own location — nothing to edit.
================================================================================
#>
$ErrorActionPreference = "Stop"

$proj    = Split-Path -Parent $PSScriptRoot
$wrapper = Join-Path $PSScriptRoot "run_scheduled.py"
$py      = Join-Path $proj ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { $py = "python" }

$taskName = "IntelliBI Operations Automation"
$times    = @("10:00","11:00","12:00")

Write-Host "Project : $proj"
Write-Host "Python  : $py"
Write-Host "Task    : $taskName  @ $($times -join ', ')  (succeed once/day; 11:00 & 12:00 are retries)"

$action   = New-ScheduledTaskAction -Execute $py `
              -Argument "`"$wrapper`" --label ops --once-per-day" -WorkingDirectory $proj
$triggers = foreach ($t in $times) { New-ScheduledTaskTrigger -Daily -At $t }
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
              -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
              -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType S4U -RunLevel Highest

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $triggers `
    -Settings $settings -Principal $principal -Force | Out-Null

Write-Host "OK - '$taskName' registered (retry-until-success, once per day, overlap-protected)."
Write-Host "Verify:  Get-ScheduledTask -TaskName '$taskName' | Get-ScheduledTaskInfo"
