# pyCoordinatorTaskPerformanceReport.py

Location: `co-ordinator reports/`. Output: `IntelliBI_Coordinator_Task_Performance_Report_<Daily|Weekly|Monthly|Manual>_<period>`. This is a native Google Sheet saved in the **same reporting-period folder** as the Batch Coordinator report of that type and period (layout below), versioned like it (`- Version N`). **Google Drive is the only place the report is stored.** The workbook is built in memory and uploaded from memory, so there is no local copy and no temporary file. Each run also removes report copies that earlier versions left under `output/reports/coordinator_performance/`. It removes only this report's own `.xlsx` files and the folders they leave empty, and never touches other files. A file that can't be removed is logged and doesn't fail the run.

## Purpose
Director-level view of the Batch Coordinator's work in two dimensions, side by side: **Effort → Completion → Outcome**.
- **Effort**: Tasks Generated → Completed → Pending → Completion % → Timely Completion % → progress over time.
- **Outcome (Actual Performance %)**: the result each task group is meant to produce, taken from that area's own report (see *Actual performance* below).

It covers all task groups together and each group separately, and answers six questions on the Dashboard: how much work was assigned, how much was completed, what actual result was achieved, whether performance is improving, whether effort is translating into outcomes, and which area needs management attention.

The report is **read-only**. It changes nothing in `pyCoordinatorTaskListReport.py`, its tabs, its follow-up columns or its versioning.

## Intended rhythm & scheduling
1. **Morning batch (10:30, retry 11:30):** `pyCoordinatorTaskListReport` builds the day's task list as part of `scripts/run_all.py`. The Coordinator works through it, filling Action Taken / Comment / Done?. Follow-Up DateTime stamps itself.
2. **Evening batch (19:00, retry 20:00):** `scripts/run_evening_reports.py` first refreshes the outcome sources (`pyZohoSignatureStatusRefresh`, `pyStudentPaymentClassesStudentEnrolled` with its API cache capped at 30 minutes via `INTELLIBI_CACHE_MAX_AGE_SECONDS`, `pyAssignmentSubmissions`), then runs this report. A failed refresh does not stop the report; its status is passed on in `INTELLIBI_EVENING_REFRESH` and shown in the Data Coverage & Rules tab. Both batches are registered by `scripts/setup_schedule.ps1` and run through `scripts/run_scheduled.py`, which provides the lock, the once-per-day success marker (label `ops_evening`) and the summary e-mail. `scripts/status.ps1` shows both tasks.

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

## Actual performance (outcome)
Calculated in `co-ordinator reports/coordinator_outcomes.py`. Every measure reuses the calculation of the report that owns it.

| Area | Actual Performance % | Daily | Weekly / Monthly |
|---|---|---|---|
| Attendance | **Overall Att %** of the Attendance report's Session Summary — `pyAttendaceFeedbackReport.overall_att_kpi()`, the function the Session Summary itself now calls | window yesterday 12:00 PM → time of the day's task list, daily formula | sessions dated in the period, period (de-duplicated) formula |
| Assignment | **Total Submission %** read from the corresponding generated **Assignment Submission Performance report** (Summary tab, by label). Not generated yet / unreadable → calculated now with that report's own `build_dataset` + `summarise` for the same period, marked **PROVISIONAL** with the reason | the Assignment Daily report generated that morning (deadlines of the day before) | the Assignment report of the same period (the Monthly one is made on the 1st, so a Monthly Coordinator report shows the provisional figure) |
| Admission | Resolved ÷ applicable: learners on the day's list whose form is **signed** at the evening re-check (`pyAdmissionFormalitiesReport.build_rows` on refreshed signers). Done? = Yes does not count | the day's list | sum of the period's days |
| Wise & IV Feedback | Resolved issues ÷ applicable issues, issue level (each failing field, failed check or missing interview feedback). Wise validation re-run on refreshed data; interview feedback re-checked in the Interview Consolidate Sheet | the day's list | sum of the period's days |
| Instructor | Sessions without escalation ÷ held sessions (window yesterday 12:00 PM → task-list time; not-yet-held sessions excluded). One session counts once; reasons (`review_instructor_sessions` codes) are shown separately | the day | sum of the period's days |
| Interview | **Overall Interview Attendance %** (Interview Consolidated Report rule: not absent/skipped AND scored or published) | yesterday → today | Monday → today / 1st → today |

- *No longer applicable* (learner inactive, record deleted) is shown but leaves the denominator; *Not checked* (source unreadable, or no stored check for a past day) is excluded and reported, never guessed. A student whose **name** was corrected is matched by Joined On + batch and counts as resolved.
- The evening state cannot be re-created later, so each **delivered Daily** report stores every checked item on the reporting PC in `cache/coordinator_outcome_checks/<YYYY-MM-DD>.json` (`OUTCOME_CHECKS_DIR`; not shown in the report). Weekly / Monthly reports add those stored checks up; the current day is checked live. Only a live evening check is stored (a re-run for a past day or a `--no-upload` dry run never overwrites it). Keep this folder: deleting it makes those days "not checked" in later Weekly / Monthly reports. Days before 07-Oct-2026 are read from the old "Outcome Detail" tab of that day's Daily report.
- Targets (`OUTCOME_TARGETS`): Attendance 75% (the Attendance report's green band), Assignment 80%, Admission 75%, Wise 75%, Instructor 90%, Interview 80%. Near target = within `OUTCOME_NEAR_BAND` (15) points.
- **Effort → Outcome quadrant**: high effort = Task Completion % ≥ `EFFORT_HIGH_PCT` (75); good outcome = Actual ≥ target. High·Good *Effort is paying off*, High·Low *Effort not converting*, Low·Good *Check if tasks needed*, Low·Low *Needs management attention*. No tasks in the period: *On target, no tasks needed* / *Below target, no tasks raised*.
- `CHECK_OUTCOMES = False` turns all of this off (effort-only report, original e-mail goals section).

## Tabs
Every tab's title band ends with `  |  Generated On: dd-Mon-yyyy hh:mm AM/PM` — the run time in IST, same wording as the Assignment Submission Performance report (`stamp_generated_on`). On a narrow tab a long title wraps onto a second line.

| Tab | Shows |
|---|---|
| Dashboard | 11-column grid. A one-line headline ("Coordinator completed X% of required actions (n of N) — actual result: …"). **Overall Coordinator Performance**: 8 uniform KPI cards in 2 rows of 4 equal-width cards (label + value, no footer text) — the 7 effort tiles — Tasks Generated, Completed, Pending, Completion %, Timely Completion %, Median Time to Complete, Pending 2+ Days (Daily: Attempted, Not Done) — plus **Average Actual %** (the simple average of the measured areas' Actual Performance %, the same figure as the scorecard total). **Task Group Scorecard**: Task Group, Tasks, Completed, Pending, Completion %, Pending %, Status (effort, rows tinted by Status) + Actual Performance % (cell note = measure and basis, e.g. the Assignment report it was read from), Target, Performance / Impact, Effort → Outcome, with a total row. **Effort vs Outcome by Task Group**: one clustered bar chart (Task Completion % vs Actual Performance %, 0–100%, value labels) read straight from the scorecard columns. |
| Effort vs Outcome Trend | **OVERALL — DAY BY DAY** only (Daily: the last 7 days; periods: every day): Day, Tasks Generated, Tasks Completed, Task Completion %, Areas Measured, Average Actual %, and under it one line chart **Task Completion % vs Average Actual % — Day by Day** (blue = effort, orange = outcome, value labels, 0–100% scale, gaps on days with nothing measured). |
| Progress Trend | **WITHIN-DAY PROGRESS — <report day or period dates>** first, on every report type: the Generated vs Completed vs Open hour-by-hour table (IST) with its chart directly underneath. Daily = that day. Weekly / Monthly / Manual = the period's report days combined by hour of day (each day's own hour-by-hour figures summed). Periods then add DAY-WISE PROGRESS (table, with the outcome chart and the completion % / timely % chart side by side underneath) and COMPLETION % BY TASK GROUP AND DAY. Every chart sits under its table inside the 10-column grid. |
| Task Register | The de-duplicated audit trail: one row per task, with the versions it was seen in and recorded in. Also the place to review pending / overdue tasks: filter Status = Missed / Open (due today) and sort Days on List. A pending task with "Recorded in" filled was attempted but not done. |
| Data Coverage & Rules | Read top to bottom: **At a glance** (report days read, versions read, tasks tracked, data notes) → **A. What data is covered** (per report day: versions, task groups / tasks tracked and not trackable, data-quality checks; amber = worth a look; total row) → **B. Where the actual performance comes from** (per area: measure, Actual %, target, basis, source report, notes incl. resolved / pending / not applicable / not checked and the evening-refresh status) → **C. Rules & definitions** in six groups: 1 data source & task, 2 reporting periods, 3 effort rules, 4 how each area's actual performance is measured, 5 targets & verdicts, 6 important limitations (amber). Rule texts unchanged. |

## E-mail
One e-mail **per report** goes from `info@intellibiinnovationstechnologies.in` to `info@intellibiinnovationstechnologies.in` and `intellibihropsb2ch@gmail.com`. The subject is `<Type> Coordinator Task Performance Report - <period>`.
- The layout follows the Sales lead-performance e-mail (`pyConsolidatedLeadPerformanceReport.build_email_body`):
  - navy header with the reporting period, then "Hello Team";
  - Task Volume cards, then Completion & Timeliness cards;
  - **Performance vs Goals · Effort → Outcome** (when outcomes are measured): the headline, then **Overall Completion %**, then per task group an *Effort* bar (Task Completion % vs 95%) and an *Outcome* bar (Actual Performance % vs the area's target, green / amber / red = on / near / below target, with the basis and the change vs the previous period) and the Effort → Outcome chip; then **Needs Management Attention**. The Task Groups table shows Tasks Generated, Tasks Completed, Task Completion %, Actual Performance % and Performance / Impact. Every figure is the Dashboard's own.
  - Effort-only runs (`CHECK_OUTCOMES = False`) keep the earlier section: first **Overall Completion %**, which is the period's own total, the same figure as the Dashboard's "All task groups" row. Then one bar per task group that has tasks in the reported period, in the existing order. Each bar shows
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
Settings are at the top of the script: `GENERATE_*` and the period dates, `UPLOAD_TO_DRIVE`, `SEND_EMAIL` / `EMAIL_*`, `COMPLETION_BENCHMARK`, `EMAIL_GROUP_LABELS`, `STATUS_ON_TRACK` / `STATUS_WATCH`, `CHECK_OUTCOMES`, `OUTCOME_TARGETS`, `OUTCOME_NEAR_BAND`, `EFFORT_HIGH_PCT`, `TREND_DAYS_DAILY`.

Run the report by hand in the evening through `scripts\run_evening_reports.py` so the sources are refreshed first; running the script alone re-checks against whatever the morning collectors loaded.

Exit code: 0 when every planned report was delivered. 1 when a report failed (the other reports still run) or the configuration is invalid.

## Reuse
- Folder, file-name prefix, follow-up column names, banner colours, the `ds_*` design system and the versioned `upload_report()` all come from `pyCoordinatorTaskListReport` (imported as `BC`). If those change there, this report follows.
- Retries use `common/api_retry.py`.

## Verification (no Google access needed)
- `python ops_validation\verify_coordinator_outcomes.py` (synthetic data only) checks every outcome: the shared Overall Att % against both Session Summary formulas, instructor session-level % (10 held, 2 escalated = 80%, reasons separate), admission 4 → 1 pending = 75%, Wise issue level from the real tab layout (fixed / still pending / no longer applicable / renamed / not checked), interview attended rule and windows, assignment = the source's `summarise()`, the Daily check stored on the PC and read back so a period = the sum of its Dailies, Dashboard and e-mail reconciliation, quadrants, headline, the six questions, and the evening refresh / cache cap.
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
- 2026-10-07 (Generated On) — Every tab's header now ends with "Generated On: <run date & time, IST>" (Daily / Weekly / Monthly / Manual), like the Assignment Submission Performance report.
- 2026-10-07 (Data Coverage & Rules) — Tab reorganised for reading: at-a-glance cards, sections A (coverage, with column groups and a total row) / B (outcome sources) / C (rules in six numbered groups, limitations highlighted), sized row heights and wrapped text. Coverage figures and rule texts unchanged.
- 2026-10-07 (KPI cards) — Overall Coordinator Performance shown as 2 rows of 4 equal-width cards with uniform borders and spacing; the description line under each value removed. Values unchanged.
- 2026-10-07 (Outcome Detail) — The **Outcome Detail** tab was removed from the Daily report. The evening checks it held (needed by Weekly / Monthly Admission and Wise results) are now stored on the reporting PC in `cache/coordinator_outcome_checks/`; earlier reports' tabs are still read for their days.
- 2026-10-07 (Trend tab) — Effort vs Outcome Trend reduced to OVERALL — DAY BY DAY ("Areas On Target" column removed) and its improved line chart; the by-area tables / chart and the previous-period table were removed from the tab. Calculations unchanged (Δ vs previous period still feeds the e-mail).
- 2026-10-07 (Dashboard) — Back to the Overall Coordinator Performance tiles and the original 7-column Task Group Scorecard, extended with **Average Actual %** (tile) and **Actual Performance % / Target / Performance / Impact / Effort → Outcome** (scorecard). "What this report answers" and the Dashboard's "Needs management attention" list removed (the e-mail keeps its attention table); the chart's data table removed — the improved chart reads the scorecard. Calculations unchanged.
- 2026-10-07 (fix) — **Assignment Total Submission % was blank** on the Daily report: it measured assignments due on the report day itself, which usually has none passed yet. It now reads the Assignment Submission Performance report that corresponds to the period (Daily D → the Assignment Daily report generated on D, i.e. deadlines of D-1; Weekly / Monthly → same period), and falls back to a clearly labelled provisional figure when that report does not exist yet. Read-only Drive lookup, cached per file revision.
- 2026-10-07 — **Effort → Outcome.** The report now shows each task group's actual result next to the Coordinator's effort (see *Actual performance*): new Dashboard (headline, effort + outcome tiles, Effort → Outcome scorecard with targets, quadrants and Δ vs previous period, the six questions, areas needing management attention, effort-vs-outcome chart), new **Effort vs Outcome Trend** tab, new **Outcome Detail** tab (Daily), outcome sources and rules in Data Coverage & Rules, and an upgraded e-mail Performance vs Goals. The evening batch refreshes the sources first. Supporting changes, each verified to give identical output: `overall_att_kpi()` extracted in `pyAttendaceFeedbackReport.py` (both Session Summary builders call it); `review_instructor_session(s)` extracted in `pyCoordinatorTaskListReport.py` (Instructor Follow-Ups tab identical on 150 synthetic days); `load_wise_validation(return_sources=True)`. Parsed-version cache bumped to `_v2` (tasks now keep their row cells). E-mail label fix: Wise & Interview Feedback Validation was shown as a second "Learner Admission Formalities"; it is now "Wise & Interview Feedback". Task completion rules, Progress Trend, Task Register, Drive layout and scheduling times are unchanged.
- 2026-10-03 — **Progress Trend, historical days:** WITHIN-DAY PROGRESS no longer shows "No trackable tasks" when the day's tasks exist but none of them was available during the day. This happens when the day's task list was generated only after the day had ended, for example a re-run with `DAILY_DATE` set. The section now says so, with the task count and when the list was first generated, and points to the Dashboard and Task Register where the tasks are counted. When some tasks were on the timeline, a note under the chart lists tasks generated after the report day, completions recorded after it, and completions without a valid time. The hour-by-hour figures were already based on the selected day only, and are unchanged. Covered by `ops_validation/verify_coordinator_trend_history.py`.
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
