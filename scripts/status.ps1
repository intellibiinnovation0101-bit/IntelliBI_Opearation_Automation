<#
================================================================================
  IntelliBI Operations Automation — schedule status dashboard
  ------------------------------------------------------------------------------
  Shows the scheduled task's next run, last run + decoded result, missed count,
  whether today's once-per-day run has already SUCCEEDED, and the tail of the
  most recent scheduler log.

        powershell -ExecutionPolicy Bypass -File scripts\status.ps1
================================================================================
#>
$ErrorActionPreference = "SilentlyContinue"
$proj     = Split-Path -Parent $PSScriptRoot
$taskName = "IntelliBI Operations Automation"
$logDir   = Join-Path $proj "logs"
$marker   = Join-Path $proj "cache\scheduler\ops_last_success.txt"

function Decode-Result($code) {
  switch ($code) {
    0          { "SUCCESS (0)" }
    267008     { "Ready (never run yet)" }
    267009     { "RUNNING NOW" }
    267010     { "Disabled" }
    267011     { "Not yet run" }
    267012     { "No more scheduled runs" }
    267014     { "Last run terminated" }
    2147750687 { "Skipped - instance already running" }
    $null      { "n/a" }
    default    { "FAILED (0x{0:X})" -f $code }
  }
}

Write-Host "==================================================================="
Write-Host " IntelliBI Operations Automation - schedule status" -ForegroundColor Cyan
Write-Host "==================================================================="

$task = Get-ScheduledTask -TaskName $taskName
if (-not $task) {
  Write-Host "Task '$taskName' is NOT registered. Run scripts\setup_schedule.ps1." -ForegroundColor Yellow
} else {
  $i = $task | Get-ScheduledTaskInfo
  Write-Host ("State        : {0}" -f $task.State)
  Write-Host ("Last run     : {0}" -f $i.LastRunTime)
  Write-Host ("Last result  : {0}" -f (Decode-Result $i.LastTaskResult))
  Write-Host ("Next run     : {0}" -f $i.NextRunTime)
  Write-Host ("Missed runs  : {0}" -f $i.NumberOfMissedRuns)
  Write-Host  "Trigger times: 10:00 normal, 11:00 & 12:00 retry-only (daily)"
}

Write-Host ""
$today = (Get-Date).ToString("yyyy-MM-dd")
if (Test-Path $marker) {
  $succDate = (Get-Content $marker -Raw).Trim()
  if ($succDate -eq $today) {
    Write-Host ("Today ({0}) : ALREADY SUCCEEDED - 11:00/12:00 will skip." -f $today) -ForegroundColor Green
  } else {
    Write-Host ("Today ({0}) : not yet succeeded (last success {1}) - next window will run/retry." -f $today,$succDate) -ForegroundColor Yellow
  }
} else {
  Write-Host ("Today ({0}) : no success recorded yet - next window will run." -f $today) -ForegroundColor Yellow
}

Write-Host ""
Write-Host "--- latest scheduler log (logs\run_scheduled.log) ------------------" -ForegroundColor DarkGray
$log = Join-Path $logDir "run_scheduled.log"
if (Test-Path $log) { Get-Content $log -Tail 15 } else { Write-Host "(no run_scheduled.log yet)" }
Write-Host ""
Write-Host "Tip: History for every firing is in Task Scheduler (taskschd.msc) -> History tab."
