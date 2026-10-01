# pyCoordinatorTaskPerformanceReport.py

Location: `co-ordinator reports/`. Output: `IntelliBI_Coordinator_Task_Performance_Report_<Daily|Weekly|Monthly|Manual>_<period>`. This is a native Google Sheet saved in the **same reporting-period folder** as the Batch Coordinator report of that type and period (layout below), versioned like it (`- Version N`). A local copy goes to `output/reports/coordinator_performance/<type folder>/<period folder>/`.

## Purpose
Management view of the Batch Coordinator's **task completion and timeliness**:
Tasks Generated → Completed → Pending → Completion % → Timely Completion % → progress over time → overall status. It covers all task groups together and each group separately.

The report is **read-only**. It changes nothing in `pyBatchCoordinatorDailyAttendanceReport.py`, its tabs, its follow-up columns or its versioning.

## Intended rhythm & scheduling
1. **Morning batch (10:30, retry 11:30):** `pyBatchCoordinatorDailyAttendanceReport` builds the day's task list as part of `scripts/run_all.py`. The Coordinator works through it, filling Action Taken / Comment / Done?. Follow-Up DateTime stamps itself.
2. **Evening batch (19:00, retry 20:00):** `scripts/run_evening_reports.py` runs this report. Both batches are registered by `scripts/setup_schedule.ps1` and run through `scripts/run_scheduled.py`, which provides the lock, the once-per-day success marker (label `ops_evening`) and the summary e-mail. `scripts/status.ps1` shows both tasks.

## Report periods (configuration at the top of the script)
```
GENERATE_AUTO    = True
GENERATE_DAILY   = True
GENERATE_WEEKLY  = False
GENERATE_MONTHLY = False
GENERATE_MANUAL  = False
DAILY_DATE            = None  # "YYYY-MM-DD"
WEEKLY_REFERENCE_DATE = None  # any date within required week
MONTHLY_MONTH         = None  # 1-12
MONTHLY_YEAR          = None
MANUAL_START_DATE     = "2026-08-21"
MANUAL_END_DATE       = "2026-09-22"
```
This is the same scheme as `pyLeadFollowUpAnalysisReport.py`. The shared logic is `coordinator_periods.plan_jobs`, which the Batch Coordinator report uses with the same week and month definitions.

- **AUTO** (the flags are ignored):
  - Daily runs every run, for today.
  - Weekly runs every **Monday**, for the previous completed Mon–Sun week.
  - Monthly runs on the **last day** of the month, for the 1st to the last day. Month lengths and leap years are handled automatically.
- **Flag mode** (`GENERATE_AUTO = False`): each flag works independently. The dates pin a specific day, week or month; `None` means today, this week or this month.
- **Manual** needs both dates with Start ≤ End. A missing or invalid date, or Start > End, skips the Manual report and makes the run exit 1.

**Only the period's own tasks are measured.** A task's period is its report day: the day of the Coordinator task list it appeared on. A Daily report evaluates only that day's tasks. Weekly evaluates only the 7 days Mon–Sun, Monthly only the calendar month, and Manual only Start…End. Earlier pending or overdue tasks are never pulled in. **Days on List / Pending 2+ Days** count consecutive report days *within the period*. The daily dashboard shows "Attempted, Not Done" in place of that tile.

## Drive layout (shared with the Batch Coordinator report — `coordinator_periods.py`)
```
<coordinator root 1BEokUc7Np7iBVSMwIyMAgZUa0mrecT-h>
  Daily Coordinator Reports/   Daily 01-Oct-2026/                   task report + performance report
  Weekly Coordinator Reports/  Weekly 21-Sep-2026 to 27-Sep-2026/   weekly roll-up + performance report
  Monthly Coordinator Reports/ Monthly Sep-2026/                     monthly roll-up + performance report
  Manual Coordinator Reports/  Manual 21-Aug-2026 to 22-Sep-2026/    manual roll-up + performance report
```
- Folders are found by name and created when missing. The folder names follow the `pyLeadFollowUpAnalysisReport` convention.
- The performance report reads task lists **only** from the Daily period folders, plus the legacy `<root>/YYYY-MM-DD/` folders created before this layout. It reads only files named like the daily task report.
- Weekly, Monthly and Manual roll-ups, and earlier performance reports, are never read as task data.
- To move the legacy folders' files into the layout (file ids and links are kept): `python "co-ordinator reports/coordinator_periods.py" --migrate-legacy`, then add `--apply`.

## Source of truth
The Coordinator daily reports are the only place Coordinator actions are recorded. Every version of every day **in the reporting period** is read through a Drive export: `<root>/Daily Coordinator Reports/Daily DD-Mon-YYYY/IntelliBI_Batch_Coordinator_Daily_Attendance_Report_* [- Version N]`, plus the legacy `<root>/YYYY-MM-DD/` folders.
- Each file is parsed once per revision and cached in `cache/coordinator_performance/` (key = file id + modifiedTime). Re-runs only fetch files that changed, including a past day's sheet the Coordinator updated later.
- Weekly/Monthly roll-up files, the older "Coordinator Attendance Tasks" sheets and every other file are ignored.

## Task groups
The groups are defined in `TASK_GROUPS`:
- **Tab names:** current names plus earlier ones, e.g. "Learner Wise Validation".
- **Stable identity:** the fields that identify one task. Each field is taken from the row's own column first, then from the section banner above it (`Tech Name: … · Duration: …`, the banner title).
- **Display fields:** the "who" shown for each task.

| Group | Identity |
|---|---|
| Learner Attendance Follow-Ups | Tech Name, Duration, Student Name, Phone |
| Learner Assignment Follow-Ups | Tech Name, assignment (sub-banner title), Email, Student Name |
| Learner Admission Formalities | Email ID, Student Name |
| Wise & Interview Feedback Validation | validation section + the section's record fields |
| Instructor Follow-Ups | Tech Name, Duration, Instructor, Session Date |
| Learner Instructor Interview Reminder | Interview Time, Batch Name (Class), Candidate Name |

A new task tab that carries the follow-up columns is picked up automatically with a generic identity, and a warning is logged. Add a `TASK_GROUPS` entry to give it a precise identity. Rank, `#`, row order and figures are never part of the identity, because they change from run to run.

## Rules

| Measure | Definition |
|---|---|
| Task | One flagged row on one report day's list. |
| Versions of a day | Merged into one list: the latest version's tasks, plus any task actioned in an earlier version that dropped off later. A task is counted **once per day**. An un-actioned task that dropped off before the final run is not counted; it is shown in coverage as "Resolved Before Final Run". |
| Completed | `Follow-Up Done? = Yes` in **any** version of that day. Entries are not carried between versions. Completion time = earliest valid Follow-Up DateTime among the Yes entries. |
| Pending | Not completed. **Open (due today)** while the report day runs, **Missed** once it is over. Done? = No, or an Action/Comment without Yes, is shown as *Attempted* but stays pending. |
| On time / Late | Each tab is the action list for its report day ("needs a follow-up today"), so a task is due by 11:59 PM IST that day. The task-specific dates (assignment deadline, admission expiry, interview date) are all on or after that day. On time = completed on the report day; Late = completed on a later day. |
| Completion % / Pending % | Completed ÷ Tasks / Pending ÷ Tasks |
| Timely Completion % | On time ÷ every task whose timing can be judged. Open tasks count as not (yet) on time, so an evening report never shows 100% timely while work is still open. A Yes without a valid DateTime (e.g. `#REF!`) is completed but excluded from timing. |
| Time to Complete | Follow-Up DateTime − the time the task first appeared in that day's report (file creation time). The median is taken over on-time completions. |
| Reporting period | A task belongs to its report day. Only tasks whose report day is inside the period are counted. |
| Next day | An item still listed the next day is that day's new task (fresh follow-up columns). **Days on List** = consecutive report days *within the period* the same item stayed pending (the accumulation signal). |
| Status | On track ≥ 75%, Watch ≥ 45%, Behind below. These are the same bands as the Counsellor Follow-Up Trend. A group's status is the worse of its Completion % and Timely %. |
| Not trackable | Reports created before the follow-up columns existed (before 30-Sep-2026, Version 6) cannot show completion. Their tasks are listed in coverage and are **never** counted as pending. |

**DateTime caveat.** Follow-Up DateTime stamps the **first** time Done? is set, whether to Yes or No. A task first set to No and later to Yes keeps the earlier time. Clearing Done? clears the stamp.

## Tabs
| Tab | Shows |
|---|---|
| Dashboard | 8 KPI tiles, a plain-language assessment, a scorecard per task group (tasks, completed, pending, completion %, pending %, timely %, status, on time / late / no time / missed / open, median time to complete, unique items, pending 2+ days), a Needs Management Attention list, and an outcome-by-group chart |
| Progress Trend | Daily: hour-by-hour tasks generated vs completed vs open. Period: day-wise outcome per report day, completion % / timely % trend, and completion % per group per day. Both: completions by hour of day, and time-to-complete buckets. |
| Pending & Overdue | Every pending task, Missed first, with days on list, attempted flag and last action / comment |
| Task Register | The de-duplicated audit trail: one row per task, with the versions it was seen in and recorded in |
| Data Coverage & Rules | Per report day: versions, trackable vs not, resolved-before-final-run, recorded in >1 version, Yes without valid time, Yes/No conflicts. Also lists the rules above. |

## Run
```
python scripts\run_evening_reports.py                                                 # as scheduled (with summary e-mail)
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py"                 # planned jobs (AUTO)
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --no-upload     # local copy only
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --date 2026-09-30
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --from 2026-09-01 --to 2026-09-30
```
Settings are at the top of the script: `GENERATE_*` and the period dates, `UPLOAD_TO_DRIVE`, `STATUS_ON_TRACK` / `STATUS_WATCH`, `TTC_BUCKETS`.

Exit code: 0 when every planned report was delivered. 1 when a report failed (the other reports still run) or the configuration is invalid.

## Reuse
- Folder, file-name prefix, follow-up column names, banner colours, the `ds_*` design system and the versioned `upload_report()` all come from `pyBatchCoordinatorDailyAttendanceReport` (imported as `BC`). If those change there, this report follows.
- Retries use `common/api_retry.py`.

## Verification (no Google access needed)
- `python ops_validation\verify_coordinator_task_performance.py` builds real Coordinator tabs with the Coordinator report's own builders, fills follow-ups across 2 days and 3 versions, and checks every task rule above against hand-computed answers.
- `python ops_validation\verify_coordinator_reporting_periods.py` runs end to end against an in-memory Drive. It checks:
  - period selection: AUTO / flag mode, Monday weekly, month-ends including leap years, Manual validation;
  - folder names; both reports landing in the same period folder; re-runs creating versions;
  - no leakage from earlier periods; legacy-folder discovery;
  - Morning-batch membership and the evening scheduler entry, labels and once-per-day gate.

## Change history
- 2026-10-01:
  - Period scoping by task origin (`scope_ledger`) and the shared `coordinator_periods` planner, with Manual validation.
  - Output moved into the shared Daily / Weekly / Monthly / Manual period folders.
  - Evening scheduling at 19:00 via `scripts/run_evening_reports.py`.
  - Exit codes added.
  - Configuration flags set as specified; `HISTORY_START_DATE` and the interim "Coordinator Performance" folder removed.
- 2026-09-30: first version.
