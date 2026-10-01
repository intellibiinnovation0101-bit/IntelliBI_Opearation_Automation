# pyCoordinatorTaskPerformanceReport.py

Location: `co-ordinator reports/`. Output: `IntelliBI_Coordinator_Task_Performance_Report_<Daily|Weekly|Monthly|Manual>_<period>`. This is a native Google Sheet saved in the **same reporting-period folder** as the Batch Coordinator report of that type and period (layout below), versioned like it (`- Version N`). **Google Drive is the only place the report is stored.** The workbook is built in memory and uploaded from memory, so there is no local copy and no temporary file. Each run also removes report copies that earlier versions left under `output/reports/coordinator_performance/`. It removes only this report's own `.xlsx` files and the folders they leave empty, and never touches other files. A file that can't be removed is logged and doesn't fail the run.

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
| Dashboard | One 7-column grid. **Overall Coordinator Performance**: 7 KPI tiles — Tasks Generated, Completed, Pending, Completion %, Timely Completion %, Median Time to Complete, and Pending 2+ Days (on a Daily report: Attempted, Not Done). **Task Group Scorecard**: Task Group, Tasks, Completed, Pending, Completion %, Pending %, Status, with a total row. **Completed vs Pending by Task Group** chart, directly under the scorecard (reads the scorecard's Completed and Pending columns). |
| Progress Trend | **WITHIN-DAY PROGRESS — <report day or period dates>** first, on every report type: the Generated vs Completed vs Open hour-by-hour table (IST) with its chart directly underneath. Daily = that day. Weekly / Monthly / Manual = the period's report days combined by hour of day (each day's own hour-by-hour figures summed). Periods then add DAY-WISE PROGRESS (table, with the outcome chart and the completion % / timely % chart side by side underneath) and COMPLETION % BY TASK GROUP AND DAY. Every chart sits under its table inside the 10-column grid. |
| Task Register | The de-duplicated audit trail: one row per task, with the versions it was seen in and recorded in. Also the place to review pending / overdue tasks: filter Status = Missed / Open (due today) and sort Days on List. A pending task with "Recorded in" filled was attempted but not done. |
| Data Coverage & Rules | Per report day: versions, trackable vs not, resolved-before-final-run, recorded in >1 version, Yes without valid time, Yes/No conflicts. Also lists the rules above. |

## E-mail
One e-mail **per report** goes from `info@intellibiinnovationstechnologies.in` to `info@intellibiinnovationstechnologies.in` and `intellibihropsb2ch@gmail.com`. The subject is `<Type> Coordinator Task Performance Report - <period>`.
- The layout follows the Sales lead-performance e-mail (`pyConsolidatedLeadPerformanceReport.build_email_body`):
  - navy header with the reporting period, then "Hello Team";
  - Task Volume cards, then Completion & Timeliness cards;
  - **Performance vs Goals**: first **Overall Completion %**, which is the period's own total, the same figure as the Dashboard's "All task groups" row. Then one bar per task group that has tasks in the reported period, in the existing order. Each bar shows
    `completed / tasks · Completion %` against the `COMPLETION_BENCHMARK` (95% Task Completion) marker.
    Green when Completion % >= 95%, red below. Groups with no tasks in the period are not shown; a period
    with no tasks at all shows "No tasks were generated in this period.";
  - the Task Groups table;
  - an "Open …" button plus a text link; no attachment, so people work in the live sheet.
- `SEND_EMAIL = True` sends it. `False` builds and uploads the reports exactly the same, but sends nothing. The `--no-email` command-line option does the same as `False`.
- Sending goes through the Operations Gmail account in `credentials/email_config.py`, via the shared `co-ordinator reports/coordinator_email.py`:
  - recipients are validated, so a missing comma never sends to a fused address;
  - temporary SMTP errors are retried through `common/api_retry.py`;
  - a sender other than the configured account is refused.
- A failed e-mail is logged and does not fail the run, because the reports are already in Drive.
- Performance vs Goals reuses the period's own scorecard figures (`run_jobs` → `groups`, the same
  `summarise()` rows as the Dashboard), so every percentage reconciles with the Dashboard's Completion %
  (both shown to 0.1%; the green/red decision uses that shown value). Nothing is recalculated.
- Group names in that section come from `EMAIL_GROUP_LABELS` (e-mail only). Task-group identity, the
  Dashboard, the Task Groups table and every other tab keep the registry names. A group not listed in
  `EMAIL_GROUP_LABELS` is shown under its registry name. When two groups share a display name, the bar's
  note line shows the registry name so they can be told apart.

## Run
```
python scripts\run_evening_reports.py                                                 # as scheduled (with summary e-mail)
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py"                 # planned jobs (AUTO)
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --no-upload     # dry run: built in memory, nothing saved
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --date 2026-09-30
python "co-ordinator reports\pyCoordinatorTaskPerformanceReport.py" --from 2026-09-01 --to 2026-09-30
```
Settings are at the top of the script: `GENERATE_*` and the period dates, `UPLOAD_TO_DRIVE`, `SEND_EMAIL` / `EMAIL_*`, `COMPLETION_BENCHMARK`, `EMAIL_GROUP_LABELS`, `STATUS_ON_TRACK` / `STATUS_WATCH`.

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
- `python ops_validation\verify_coordinator_email_dashboard.py` checks both e-mails (SEND_EMAIL on/off,
  sender, recipients) and the Dashboard layout. For Performance vs Goals it checks the 95% benchmark, the display
  names, the green/red threshold, that only applicable groups are shown, and that the figures equal the
  Dashboard scorecard.

## Status row colours
Rows with a Status column are tinted by that status. The tint is a lighter shade of the colour the Status chip already uses (`ROW_TINT`). Text stays dark, and the chip stays stronger.

| Where | Status → row colour |
|---|---|
| Task Register (task-level) | **Green = completed** (on time, late, or time not recorded). **Orange = attention required**: Open (due today), still within its report day. **Red = overdue / urgent**: Missed (report day over, not done), or still open while the same item has been pending `CARRIED_DAYS` (2) or more consecutive report days. That is the existing "Pending 2+ Days" / Days on List rule, not a new threshold. The whole row is coloured; the Status chip is a stronger shade of the same colour, and the first cell has a coloured left edge. The colour key is in the tab's guide line. |
| Dashboard scorecard | On track → green · Watch → amber · Behind → red · No tasks → grey. The task-group chip, the Completion % band pill, the Status chip and the "All task groups" total row keep their own styling. |

Highlights with their own purpose stay on top: the red bold Days on List (2+), the red bold Pending count, headers and section banners. Values, formats, filters and sorting are unchanged.

## Change history
- 2026-10-01 (e-mail): Performance vs Goals now starts with **Overall Completion %**. It uses the period's existing summary (`summarise()` → the Dashboard's "All task groups" Completion %) with the same 95% benchmark and green/red rule. The task-group bars follow unchanged, and a period with no tasks still shows only "No tasks were generated in this period."
- 2026-10-01 (severity colours): Task Register rows switched to Green / Orange / Red by severity (`TASK_ROW_COLORS`, `_task_severity`). Severity comes from each task's existing status and Days on List. The Dashboard is unchanged. Every value, format, filter and freeze pane was checked unchanged on all four report types.
- 2026-10-01 (status colours): Task Register and Dashboard scorecard rows are tinted by Status. This is presentation only, and every cell value was checked unchanged on all four report types.
- 2026-10-01 (simplification):
  - Dashboard: the Needs Management Attention section is removed and not replaced. The chart now sits directly under the scorecard, at the grid's full width.
  - **Pending & Overdue tab removed.** Every field it showed is also in the Task Register, which was checked on live data (159 pending tasks, 0 differences):
    - Status, Task Group, Report Day, Item, Context, What Was Flagged, Generated At and Days on List appear with the same values.
    - Last Action Taken / Last Comment = Action Taken / Follow-Up Comment.
    - Attempted? = "Recorded in Versions" is filled, which uses the same rule.
  - No formula, chart or e-mail figure read from the removed tab. The Task Register guide now explains how to review pending tasks (filter Status, sort Days on List). The e-mail's closing sentence was updated to match.
  - Calculations, Task Register data, Progress Trend, Drive and scheduling are unchanged.
- 2026-10-01 (storage):
  - Drive only: no local report copy. `SAVE_LOCAL_COPY` and `OUTPUT_DIR` were removed, and `run_jobs()` no longer takes `save_dir`.
  - The workbook goes straight from memory to Drive, so no temporary file is needed and nothing is left behind if the upload or the e-mail fails.
  - `cleanup_legacy_local_copies()` runs at the start of every run and removes earlier local copies and their empty period folders.
  - `--no-upload` / `UPLOAD_TO_DRIVE = False` is now a pure dry run.
  - The parsed-source cache in `cache/coordinator_performance/` (JSON of the task sheets, not reports) is unchanged.
- 2026-10-01 (Progress Trend):
  - WITHIN-DAY PROGRESS — <report day / period> is now the first section on every report type, with the Generated vs Completed vs Open hour-by-hour table and its chart directly underneath. For Weekly / Monthly / Manual, the period's report days are combined by hour of day.
  - Removed the "When are tasks completed?" (completions by hour of day) and "How quickly are tasks completed?" (time-to-complete buckets) sections, their charts, and their display-only code (`completion_hour_profile`, `ttc_profile`, `TTC_BUCKETS`). The median time to complete on the Dashboard and in the e-mail is unchanged.
  - Layout tidied: every chart sits under its table inside the 10-column grid, with no gaps reserved for side charts. The day-wise charts moved from beside the table (columns L–AE) to underneath it.
  - Task calculations, the other tabs, the e-mail, Drive and scheduling are unchanged; this was checked cell for cell against the previous build on live data.
- 2026-10-02:
  - E-mail added (`SEND_EMAIL`, `EMAIL_SENDER`, `EMAIL_RECIPIENTS`; shared `coordinator_email.py`).
  - Dashboard simplified:
    - removed the Missed (Overdue) tile and the ASSESSMENT section;
    - the scorecard now shows Task Group, Tasks, Completed, Pending, Completion %, Pending % and Status;
    - the dashboard is a 7-column grid; the chart now shows Completed vs Pending and a print area is set.
  - Calculations and other tabs are unchanged.
  - E-mail **Performance vs Goals** reworked: per-task-group Completion % against a 95% Task Completion
    benchmark (green >= 95%, red below), only for groups with tasks in the period, with short e-mail-only
    names (`COMPLETION_BENCHMARK`, `EMAIL_GROUP_LABELS`). It replaces the two overall bars that used the 75%
    band; the overall Completion % and Timely % stay in the cards. The Task Groups table now shows
    Completion % to 0.1%, like the Dashboard.
- 2026-10-01:
  - Period scoping by task origin (`scope_ledger`) and the shared `coordinator_periods` planner, with Manual validation.
  - Output moved into the shared Daily / Weekly / Monthly / Manual period folders.
  - Evening scheduling at 19:00 via `scripts/run_evening_reports.py`.
  - Exit codes added.
  - Configuration flags set as specified; `HISTORY_START_DATE` and the interim "Coordinator Performance" folder removed.
- 2026-09-30: first version.
