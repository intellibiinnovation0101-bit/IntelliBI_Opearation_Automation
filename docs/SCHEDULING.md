# Scheduling — IntelliBI Operations Automation

The Operations pipeline must **succeed exactly once per day**. It is scheduled at
two times, where the later one is a **fallback/retry** window, not an extra run.
The times are staggered off the Sales schedule (11:00 / 14:00 / 17:00 / 18:45 /
21:00 / 23:00) so the two projects never run at the same time:

| 10:30 | 11:30 |
|-------|-------|
| normal daily run | retry **only if** 10:30 did not succeed |

- If **10:30 succeeds** → 11:30 skips.
- If **both fail** → you have two failure e-mails and the run is retried next day.

Each real run executes the full `scripts/run_all.py` (Layer 1 in parallel →
Layer 2 with dependency gating) and e-mails the detailed summary log to
`info@intellibiinnovationstechnologies.in`. Layer 2 includes the **Batch
Coordinator daily task report** (`co-ordinator reports/`), so the Coordinator's
list is ready each morning.

## Evening batch — Coordinator Task Performance

A second task, **"IntelliBI Operations Automation - Evening"**, runs
`scripts/run_evening_reports.py` (the Coordinator Task Performance report) after
the working day, using the same once-per-day mechanism with its own label:

| 19:00 | 20:00 |
|-------|-------|
| normal evening run | retry **only if** 19:00 did not succeed |

`scripts/run_scheduled.py --label ops_evening --once-per-day --entry run_evening_reports.py`.
The marker is `cache/scheduler/ops_evening_last_success.txt` and the lock is
`ops_evening.lock`, so the morning and evening batches never block or skip each
other. It e-mails its own summary, like the morning batch.

## One-time setup (on the target machine, after deployment)

Open **PowerShell as Administrator**, `cd` into the project folder, and run:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1
```

That registers **two** tasks: **"IntelliBI Operations Automation"** (10:30 / 11:30)
and **"IntelliBI Operations Automation - Evening"** (19:00 / 20:00). Paths are
derived automatically. `scripts\status.ps1` shows both.

## The once-per-day success mechanism

`scripts/run_scheduled.py --label ops --once-per-day` (the wrapper the task calls):

1. Reads the per-day **success marker** `cache/scheduler/ops_last_success.txt`
   (contains the date of the last successful run). If it equals **today**, the
   trigger logs *"already completed successfully today — skipping"* and exits.
   This is what makes 11:30 a pure fallback.
2. Otherwise it takes the overlap lock (`cache/scheduler/ops.lock`) — if a
   previous run is still going, it skips (no double-run).
3. It launches `scripts\run_all.py`. **Only if that exits 0** does it write
   today's date into the success marker. A failed run leaves the marker unset, so
   the next window retries.

`StartWhenAvailable` is also enabled, so a run missed because the machine was off
starts as soon as it powers on (still subject to the once-per-day gate).

> Note: the windows are one hour apart. Keep a normal run comfortably under an
> hour; if a 10:30 run is still going at 11:30, the 11:30 trigger correctly skips
> (overlap protection) and the 10:30 run is allowed to finish.

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
.venv\Scripts\python.exe scripts\run_scheduled.py --label ops_evening --once-per-day --entry run_evening_reports.py
```

## Change history
- 2026-08-24 — Added scheduling (10:00 normal + 11:00/12:00 retry-until-success, once per day, overlap-protected) via `run_scheduled.py` + `setup_schedule.ps1`.
- 2026-10-01 — Added the Batch Coordinator report to the morning batch (Layer 2) and the evening batch (19:00 + 20:00 retry, `run_evening_reports.py`, Coordinator Task Performance report); `run_scheduled.py` gained `--entry`. Re-run `setup_schedule.ps1` as Administrator to apply.
- 2026-09-11 — Changed schedule to 10:30 normal + 11:30 retry (removed 12:00), staggered off the Sales schedule so the two projects never run at the same time. Re-run `setup_schedule.ps1` as Administrator to apply.
