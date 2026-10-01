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

$tasks = @(
  @{ Name = $taskName;                       Marker = (Join-Path $proj "cache\scheduler\ops_last_success.txt") },
  @{ Name = "IntelliBI Operations Automation - Evening"; Marker = (Join-Path $proj "cache\scheduler\ops_evening_last_success.txt") }
)
$today = (Get-Date).ToString("yyyy-MM-dd")
foreach ($t in $tasks) {
  Write-Host ""
  Write-Host ("[{0}]" -f $t.Name) -ForegroundColor Cyan
  $task = Get-ScheduledTask -TaskName $t.Name
  if (-not $task) {
    Write-Host "  NOT registered. Run scripts\setup_schedule.ps1." -ForegroundColor Yellow
    continue
  }
  $i = $task | Get-ScheduledTaskInfo
  $times = ($task.Triggers | ForEach-Object { ([datetime]$_.StartBoundary).ToString("HH:mm") }) -join ", "
  Write-Host ("  State        : {0}" -f $task.State)
  Write-Host ("  Last run     : {0}" -f $i.LastRunTime)
  Write-Host ("  Last result  : {0}" -f (Decode-Result $i.LastTaskResult))
  Write-Host ("  Next run     : {0}" -f $i.NextRunTime)
  Write-Host ("  Missed runs  : {0}" -f $i.NumberOfMissedRuns)
  Write-Host ("  Trigger times: {0}  (first = normal run, later = retry-only)" -f $times)
  if (Test-Path $t.Marker) {
    $succDate = (Get-Content $t.Marker -Raw).Trim()
    if ($succDate -eq $today) {
      Write-Host ("  Today ({0}) : ALREADY SUCCEEDED - retry windows will skip." -f $today) -ForegroundColor Green
    } else {
      Write-Host ("  Today ({0}) : not yet succeeded (last success {1})." -f $today,$succDate) -ForegroundColor Yellow
    }
  } else {
    Write-Host ("  Today ({0}) : no success recorded yet." -f $today) -ForegroundColor Yellow
  }
}

Write-Host ""
Write-Host "--- latest scheduler log (logs\run_scheduled.log) ------------------" -ForegroundColor DarkGray
$log = Join-Path $logDir "run_scheduled.log"
if (Test-Path $log) { Get-Content $log -Tail 15 } else { Write-Host "(no run_scheduled.log yet)" }
Write-Host ""
Write-Host "Tip: History for every firing is in Task Scheduler (taskschd.msc) -> History tab."
