# pyCoordinatorTaskListReport.py

*Formerly `pyBatchCoordinatorDailyAttendanceReport.py`, renamed 2026-10-01. Only the file name changed; the Drive report name `IntelliBI_Batch_Coordinator_Daily_Attendance_Report`, the e-mail subject and the morning-batch job label "Batch Coordinator Daily Report" are unchanged, so earlier reports, versioning and the Task Performance report keep working.*

Location: `co-ordinator reports/`. Output: `IntelliBI_Batch_Coordinator_Daily_Attendance_Report_<window>` (native Google Sheet) plus Weekly / Monthly / Manual roll-ups, saved in the shared Coordinator layout (see **Drive layout & scheduling** below). Runs in the Operations **Morning batch** (`scripts/run_reports_action.py`).

## Drive layout & scheduling
- Folders come from `co-ordinator reports/coordinator_periods.py`, shared with `pyCoordinatorTaskPerformanceReport.py`, so both reports for the same type and period land in one folder:
  `<coordinator root>/Daily Coordinator Reports/Daily 01-Oct-2026/`, `Weekly Coordinator Reports/Weekly 21-Sep-2026 to 27-Sep-2026/`, `Monthly Coordinator Reports/Monthly Sep-2026/`, `Manual Coordinator Reports/Manual 21-Aug-2026 to 22-Sep-2026/`.
- `upload_report()` accepts the (type folder, period folder) pair and keeps its never-overwrite versioning (`- Version N`) per report base name.
- Reports made before this layout are in `<root>/YYYY-MM-DD/`. To move them across (files keep their ids/links), run `python "co-ordinator reports/coordinator_periods.py" --migrate-legacy` for a dry run, then add `--apply`.
- Morning batch: Layer-2 job "Batch Coordinator Daily Report", gated on all four Layer-1 refreshes. AUTO mode produces Daily every run (for `DAILY_DATE` when it is set, otherwise today in IST), Weekly on Monday (previous Mon–Sun) and Monthly on the last day of the month. Keep `DAILY_DATE = None` for scheduled runs: a pinned date makes every run rebuild that same day. The script now exits 1 when a planned report fails, so the pipeline records it and the 11:30 window retries.

## Purpose
One action list per day for the Batch Coordinator. Every tab answers, row by row:
**who / what is affected → the figures behind the flag → Why Flagged → the action to take.**

## Tabs (daily)
| Tab | Lists | Source logic reused (unchanged) |
|---|---|---|
| Learner Attendance Follow-Ups | learners who trip an attendance / feedback / low-rating follow-up today, ranked per technology | `pyAttendaceFeedbackReport` Student Detail + till-date aggregates (`build_aggregates`) |
| Learner Assignment Follow-Ups | pending assignment submissions by technology → assignment → learner, most-urgent reminder first | `pyAssignmentSubmissionEmailReminder` defaulter/reminder logic |
| Learner Admission Formalities | learners whose admission form / e-signature is pending | `pyAdmissionFormalitiesReport` matching + status classification |
| Wise & Interview Feedback Validation | failed Wise checks (student / course / instructor) + interviews without feedback | `pyWiseDataValidationReport` + Interview Consolidate Sheet |
| Instructor Follow-Ups | sessions needing a coordinator touch (late start, not 5 min early, cancelled, underrun, instructor feedback missing, low student-feedback rate) | AR session classifier, `AR.teacher_feedback_session_ids()` |
| Learner Instructor Interview Reminder | interviews due a MESSAGE (2 days prior) or CALL (1 day prior) today | interview schedule files (Interview_Helper / feedback tabs) |

## Report design system (presentation only)
All tabs share one visual language, implemented by the `ds_*` helpers at the top of the script:
- **Row 1** navy title band (tab purpose | period). **Row 2** guide strip: how to read the tab, the colour key, and the record count.
- **Headers** dark navy; wide tabs (Learner Attendance, Instructor) carry a grouped super-header (STUDENT / DAILY / OVERALL / ACTION, SESSION / TIMING / ATTENDANCE / FEEDBACK / ACTION).
- **Body** white / pale-zebra rows, thin light borders, no gridlines. Colour is used only where it carries meaning:
  - the **first cell of every flagged row is a priority chip** with a coloured left edge — High = red, Medium = amber, Info = blue, OK = green. The level comes from each tab's own existing signal (learner rank 1–3 / 4–6, reminder stage, admission red/amber classification, Wise Invalid / Missing / Warning, Call- vs Message-level instructor reasons, CALL vs MESSAGE interview day);
  - status and measure cells (Absent / Present, Feedback Given, Valid / Missing / Invalid, reminder stage, Att %, ratings, streaks) use the same pale tints with the thresholds unchanged (`ds_att_bg` / `ds_rating_bg` mirror `AR._att_bg` / `AR._rating_bg`);
  - **Why Flagged** is always the pale-yellow last column: one numbered line per reason (`1)`, `2)` … or `•` when single), the issue label and key figures bold red, and the **action on its own line in bold navy** (`→ Call Learner`). `ds_why_text()` splits the existing reason wording at its "—" / "→" / sentence break; no wording is changed. Row height is computed from the wrapped line count so nothing is cut off.
- **Section banners** (technology / batch / validation section) blue with white text; **sub-banners** (assignment details, interview group) pale blue; **empty states** pale green.
- Freeze panes under the headers, autofilter, fixed sensible column widths, landscape / fit-to-width print setup.
The weekly / monthly roll-up tabs (`build_learner_followups_period`, `build_instructor_followups_period`) use the same helpers.

## Coordinator follow-up tracking columns
Every daily tab ends with four teal input columns the coordinator fills in on the generated Google Sheet (`FOLLOWUP_COLS`; the report never reads them back — they exist only so the follow-up can be recorded next to the flag):

| Column | Behaviour |
|---|---|
| ✎ Action Taken | dropdown, values per tab (`FOLLOWUP_ACTIONS`): Attendance = Call / WhatsApp / Call & WhatsApp / Email / No response / Other; Assignment adds WhatsApp Group; Admission = Form Send / Form Signed / Call / WhatsApp / Call & WhatsApp / Email / Other / Not Applicable; Wise = Corrected / Invalid / Call / WhatsApp / Call & WhatsApp / Email / Other / Not Applicable; Instructor & Interview = Call / WhatsApp / Call & WhatsApp / Email / Other |
| ✎ Follow-Up Comment | free text |
| ✎ Follow-Up Done? | dropdown Yes / No |
| ✎ Follow-Up DateTime | `dd-MMM-yyyy HH:MM:SS` (IST). Stamps itself the moment Follow-Up Done? is set, then holds that value; clearing Follow-Up Done? clears it, re-selecting stamps again |

How the timestamp works: the cell holds `=IF(Done="","",IF(ISERROR(DT),NOW(),IF(OR(DT="",DT=0),NOW(),DT)))` — a self-referencing formula that only evaluates `NOW()` while the cell is still empty. It needs *iterative calculation* to be on, so the script (a) sets it in the workbook (`ds_enable_iterative_calc`, for Excel) and (b) after the upload calls the Sheets API once per file (`_enable_followup_timestamps`) to switch iterative calculation on and pin the sheet's time zone to `Asia/Kolkata` — first with the same impersonated Drive-scope credentials the upload used (the Sheets API accepts the Drive scope, so no extra domain-wide-delegation scope is required), then with the plain service account. If both fail the console shows a loud `FOLLOW-UP DATETIME NOT ENABLED` warning and the fix is manual: File ▸ Settings ▸ Calculation ▸ Iterative calculation = On (Max iterations 1), then clear and re-select Follow-Up Done? on rows already marked. A `#REF!` in the column means exactly that: iterative calculation is off on that sheet. Dropdowns are standard list data validations (they survive the xlsx → Google Sheets conversion). Only flagged rows carry the input cells; banners and empty-state rows do not.

## Sheet access & protection
Every sheet this script uploads (Daily, re-run versions, Weekly / Monthly / Manual roll-ups) is locked straight after upload by `protect_and_share()`. Settings sit at the top of the script:

| Setting | Value |
|---|---|
| `PROTECT_SHEETS` | `True` (set `False` to upload exactly as before, with no protection and no extra share) |
| `COORDINATOR_EDITOR` | `intellibihropsb2ch@gmail.com` |
| `SHEET_FULL_CONTROL` | `[IMPERSONATE_USER]` (info@, the owner). The pipeline service account from `service_account.json` is added automatically. |
| `COORDINATOR_EDITABLE` | Action Taken, Follow-Up Comment, Follow-Up Done? |

What happens, in order:
1. **Find the input cells.** `followup_input_ranges()` reads the workbook just uploaded. Every follow-up row carries the self-stamping Follow-Up DateTime formula, and its three input cells are the three columns to its left. Headers, banners, empty-state rows and Follow-Up DateTime are never included.
2. **Protect every tab.** One atomic Sheets `batchUpdate` adds, per tab, a protected range over the **whole tab** (`warningOnly: false`). Only info@ and the service account are its editors, and the follow-up input cells are its `unprotectedRanges`. Because the whole tab is protected, nobody else can insert or delete rows or columns, rename or delete the tab, edit headers, data or Follow-Up DateTime, or change the protection. Tabs without follow-ups (period roll-ups, for example) are fully read-only for the Coordinator. The batch first removes any earlier protection with the same description, so re-applying never duplicates it.
3. **Check it.** The protections are read back. Each tab must have exactly one, not warning-only, covering the whole tab, with the expected number of editable ranges.
4. **No re-sharing.** `writersCanShare = false` on the file, so editors cannot share it or change its permissions.
5. **Share.** Only now does `intellibihropsb2ch@gmail.com` get **Editor** on the file, with no notification e-mail.

Dropdowns and the Follow-Up DateTime stamp keep working. Data validation stays on the editable cells, and the protected DateTime formula still recalculates when Follow-Up Done? changes. Transient Google errors are retried (`common/api_retry.py`). If any step still fails, the console shows `SHEET NOT PROTECTED / NOT SHARED FOR EDITING` and `[protect] FAILED — …`, and the run summary carries a note. The Coordinator is **not** given edit access to that sheet: they keep the folder's view access. The report itself is still uploaded and e-mailed. Earlier sheets are never modified.

Things to know:
- Any **other** editor of the Coordinator folder (for example `intellibiinnovation0101@gmail.com`) is limited exactly like the Coordinator on these sheets. Add them to `SHEET_FULL_CONTROL` if they need full control.
- Editors can still add a new tab of their own and use filter views. Google does not allow blocking those through protection.
- The owner sees protected ranges under Data ▸ Protect sheets and ranges, described "IntelliBI Coordinator task list - protected (automation)".
- Checked offline by `ops_validation/verify_coordinator_sheet_protection.py`, which tests every cell of real report tabs for the Coordinator, the owner and the service account.

## E-mail
One e-mail **per generated report** goes from `info@intellibiinnovationstechnologies.in` to `info@intellibiinnovationstechnologies.in` and `intellibihropsb2ch@gmail.com`. The subject is `<Type> Batch Coordinator Report - <period>`.
- The layout follows the Sales lead-performance e-mail.
- The Daily e-mail shows the task count per tab (from each tab's own guide count) and asks the Coordinator to record follow-ups in the sheet.
- Weekly / Monthly / Manual roll-ups send their link.
- There is no attachment: the Coordinator records follow-ups in the live sheet.
- `SEND_EMAIL = True` sends it; `False` generates and uploads exactly the same but sends nothing.
- It uses the Operations Gmail account (`credentials/email_config.py`) through `co-ordinator reports/coordinator_email.py`, with recipient validation and SMTP retry.
- A failed e-mail is logged and does not fail the run.

## Storage
Google Drive is the only place this report is stored. Every workbook (Daily, Weekly, Monthly, Manual) is built in an in-memory buffer and uploaded from it (`upload_report()` → `MediaIoBaseUpload`). Nothing is written to `output/reports/` or to any temporary file, and the e-mail carries the Google Sheet link, not an attachment. `ops_validation/verify_coordinator_email_dashboard.py` checks this.

## Change history
- 2026-10-05 — Wise & Interview Feedback Validation: the ✎ Action Taken dropdown has one more option, `Not Applicable` (after the existing ones). Other tabs' dropdowns, follow-up logic and protection are unchanged. Covered by `ops_validation/verify_coordinator_sheet_protection.py`.
- 2026-10-03 — **Fix:** `DAILY_DATE` was ignored when `GENERATE_AUTO = True`. AUTO always used the run date, so `DAILY_DATE = "2026-10-02"` produced a 03-Oct report. The Daily date now comes from one helper, `_daily_report_date()`, in both AUTO and flag mode. A set date drives the data window, as-of aggregates, period label, file name, Drive folder, tab headings and e-mail subject and body. A malformed value stops with a clear error. The e-mail is worded "Today's" only when the report is for today. Weekly, Monthly, Manual and scheduling are unchanged. Covered by `ops_validation/verify_coordinator_daily_date.py`.
- 2026-10-01 — **Renamed** from `pyBatchCoordinatorDailyAttendanceReport.py` to `pyCoordinatorTaskListReport.py`, with every reference updated (scheduler job, performance-report import, verify scripts, run summary, docs, shared `gmail_star`) and the logger renamed `CoordinatorTaskList`. **Sheet access & protection** added (`protect_and_share`, see above): `intellibihropsb2ch@gmail.com` can edit only Action Taken / Follow-Up Comment / Follow-Up Done?, and everything else is protected for everyone but info@ and the service account. Report content, task logic, Drive name and layout, versioning and e-mail are unchanged.
- 2026-10-01 — Reviewed the storage flow: already Drive-only (in-memory build and upload, no local or temporary file). No code change; documented above and covered by the verify script.
- 2026-10-02 — E-mail added (`SEND_EMAIL`, `EMAIL_SENDER`, `EMAIL_RECIPIENTS`). Task generation, Drive layout and versioning unchanged.
- 2026-10-01 — Shared Coordinator Drive layout (Daily / Weekly / Monthly / Manual → reporting-period folder) via `coordinator_periods.py`; `upload_report()` takes nested folder names; `_drive_client()` / `_list_children()` helpers; non-zero exit on a failed report; added to the Morning batch. Task-generation logic unchanged.
- 2026-09-30 — Freeze panes: Learner Admission Formalities through Phone Number (D4), Wise & Interview Feedback Validation first 4 columns (E3), Instructor Follow-Ups through Phone (E5), Learner Instructor Interview Reminder through Phone (F4). Verified that Google Sheets keeps a frozen column split when full-width banner merges cross it. Wise has no autofilter by design: its four sections have different column headers, and a sheet allows one filter range.
- 2026-09-30 — Layout: Learner Attendance Follow-Ups drops the per-row Tech Name / Duration columns (they live in the technology banner, which now also states "N learners pending follow-up") and freezes Rank · Student Name · Phone while scrolling sideways; Learner Assignment Follow-Ups' technology banner now reads "Tech Name: … · Duration: … · N pending" (the assignment sub-banner repeats Duration only when a technology has more than one) and freezes # · Student Name · Email · Phone. Data, grouping, ranking, dropdowns and follow-up logic unchanged (`ds_finish(freeze_after_col=…)`).
- 2026-09-30 — Wise & Interview Feedback Validation: Student section `Joined On` is now shown in IST (`YYYY-MM-DD HH:MM:SS IST`, same style as the Course section's `Created On`); the Wise UTC ISO timestamp is converted for display only (`_wise_joined_on_ist`), the record and validation logic are untouched. Learner Admission Formalities: `Action Taken` dropdown gained `Not Applicable`.
- 2026-09-30 — Added the four coordinator follow-up tracking columns (Action Taken dropdown per tab, Follow-Up Comment, Follow-Up Done? Yes/No, self-stamping Follow-Up DateTime) to all six daily tabs; iterative calculation + IST time zone applied to the uploaded sheet. No report logic changed.
- 2026-09-30 — Report redesign (formatting only): common design system across all six daily tabs and the period roll-ups; priority chips, semantic-only colour, readable multi-reason Why Flagged with the action line, consistent headers/banners/widths/heights. No calculation, filter, threshold, follow-up condition or data selection changed.
- 2026-09-30 — Instructor Follow-Ups "Feedback Given" uses `AR.teacher_feedback_session_ids()` (Teacher_Feedback rows with content only), same rule as the AR Teacher_No_Feedback tab.
