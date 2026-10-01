<#
================================================================================
  IntelliBI Operations Automation — Task Scheduler registrar
  ------------------------------------------------------------------------------
  Registers TWO scheduled tasks, both through scripts/run_scheduled.py:

  1. "IntelliBI Operations Automation"  — MORNING batch (scripts/run_all.py:
     Layer-1 refresh + Layer-2 reports, incl. the Batch Coordinator daily task
     report) at 10:30, retry-only window 11:30.
  2. "IntelliBI Operations Automation - Evening" — EVENING batch
     (scripts/run_evening_reports.py: Coordinator Task Performance report) at
     19:00, retry-only window 20:00 — after the Coordinator has worked the day's
     task list.

  Each batch must succeed only ONCE per day:

    * 10:30 is the normal run.
    * 11:30 is a FALLBACK/RETRY window — scripts/run_scheduled.py
      (invoked with --once-per-day) checks a per-day success marker and EXITS
      immediately if the pipeline already succeeded earlier today, so the
      fallback window never produces an extra daily run.
    * A failed early window leaves the marker unset, so the next window retries.
      (Times are staggered off the Sales schedule so the two projects never
       run at the same time.)
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
$times    = @("10:30","11:30")

Write-Host "Project : $proj"
Write-Host "Python  : $py"
Write-Host "Task    : $taskName  @ $($times -join ', ')  (succeed once/day; 11:30 is the retry)"

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

# ── Evening batch: Coordinator Task Performance (19:00, 20:00 retry-only) ─────
# Own label => own lock + own once-per-day marker, so it never blocks / is never
# skipped by the morning batch.
$eveTask  = "IntelliBI Operations Automation - Evening"
$eveTimes = @("19:00","20:00")
Write-Host ""
Write-Host "Task    : $eveTask  @ $($eveTimes -join ', ')  (succeed once/day; 20:00 is the retry)"
$eveAction   = New-ScheduledTaskAction -Execute $py `
                 -Argument "`"$wrapper`" --label ops_evening --once-per-day --entry run_evening_reports.py" `
                 -WorkingDirectory $proj
$eveTriggers = foreach ($t in $eveTimes) { New-ScheduledTaskTrigger -Daily -At $t }
$eveSettings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
                 -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
                 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $eveTask -Action $eveAction -Trigger $eveTriggers `
    -Settings $eveSettings -Principal $principal -Force | Out-Null

Write-Host "OK - '$eveTask' registered (retry-until-success, once per day, overlap-protected)."
Write-Host "Verify:  Get-ScheduledTask -TaskName '$eveTask' | Get-ScheduledTaskInfo"
