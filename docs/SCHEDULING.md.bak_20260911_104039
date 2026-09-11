# Scheduling — IntelliBI Operations Automation

The Operations pipeline must **succeed exactly once per day**. It is scheduled at
three times, where the later two are **fallback/retry** windows, not extra runs:

| 10:00 | 11:00 | 12:00 |
|-------|-------|-------|
| normal daily run | retry **only if** 10:00 did not succeed | retry **only if** neither 10:00 nor 11:00 succeeded |

- If **10:00 succeeds** → 11:00 and 12:00 skip.
- If **10:00 fails but 11:00 succeeds** → 12:00 skips.
- If all three fail → you have three failure e-mails and the run is retried next day.

Each real run executes the full `scripts/run_all.py` (Layer 1 in parallel →
Layer 2 with dependency gating) and e-mails the detailed summary log to
`info@intellibiinnovationstechnologies.in`.

## One-time setup (on the target machine, after deployment)

Open **PowerShell as Administrator**, `cd` into the project folder, and run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1
```

That registers a single Task Scheduler task named **"IntelliBI Operations
Automation"** with the three daily triggers. Paths are derived automatically.

## The once-per-day success mechanism

`scripts/run_scheduled.py --label ops --once-per-day` (the wrapper the task calls):

1. Reads the per-day **success marker** `cache/scheduler/ops_last_success.txt`
   (contains the date of the last successful run). If it equals **today**, the
   trigger logs *"already completed successfully today — skipping"* and exits.
   This is what makes 11:00 / 12:00 pure fallbacks.
2. Otherwise it takes the overlap lock (`cache/scheduler/ops.lock`) — if a
   previous run is still going, it skips (no double-run).
3. It launches `scripts\run_all.py`. **Only if that exits 0** does it write
   today's date into the success marker. A failed run leaves the marker unset, so
   the next window retries.

`StartWhenAvailable` is also enabled, so a run missed because the machine was off
starts as soon as it powers on (still subject to the once-per-day gate).

> Note: the windows are one hour apart. Keep a normal run comfortably under an
> hour; if a 10:00 run is still going at 11:00, the 11:00 trigger correctly skips
> (overlap protection) and the 10:00 run is allowed to finish.

## Verify / manage

```powershell
Get-ScheduledTask -TaskName "IntelliBI Operations Automation" | Get-ScheduledTaskInfo
Start-ScheduledTask -TaskName "IntelliBI Operations Automation"     # run now (respects the daily gate)
type cache\scheduler\ops_last_success.txt                          # today's date once it has succeeded
Unregister-ScheduledTask -TaskName "IntelliBI Operations Automation"
```

Force a re-run today (e.g. after fixing data): delete the marker, then start the
task —
```bat
del cache\scheduler\ops_last_success.txt
```

Manual test without the scheduler:
```bat
.venv\Scripts\python.exe scripts\run_scheduled.py --label ops --once-per-day
```

## Change history
- 2026-08-24 — Added scheduling (10:00 normal + 11:00/12:00 retry-until-success, once per day, overlap-protected) via `run_scheduled.py` + `setup_schedule.ps1`.
