# pyAssignmentSubmissionPerformanceReport.py — Assignment Submission Performance & Non-Submission Reports

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** AssignmentSubmissions (Layer 1, `ops_data_collection/pyAssignmentSubmissions.py`)
**Schedule:** Operations Morning batch (`scripts/run_reports_action.py`; Windows task "IntelliBI Operations Automation", 10:30 with an 11:30 retry). It is skipped when that morning's Assignment Submissions refresh failed.

## Purpose
For each reporting period, one script builds **two Google Sheets from one shared dataset**, so the two always reconcile:

| Report | Drive root | Tabs |
|---|---|---|
| **A. Assignment Submission Performance Report** (management view) | "Assignment Submission Report" `190_1NRLGgwoLpqiiW0fPGZ85ux2QknZl` | **Summary**: KPIs, batch-wise performance · **Assignment Performance**: one row per assignment · **Learner Submission Detail**: per assignment, a `Class: <technology> | Duration: <class duration> | 📝 <Assignment> | Start: MM/DD/YYYY | Deadline: MM/DD/YYYY` banner (Duration = the class's `class_subject`), then its learners with #, Learner, Email, Phone, Submitted At, Status, Marks Obtained, Max Marks, Feedback, Active Learner. This follows the Student_Detail layout of `pyAssignmentSubmissionsReport`. |
| **B. Assignment Non-Submission Report** (follow-up list) | "Assignment Not Submitted Students List" `1PmpsLw-4L4woeznxYC_8bQnFQzHQYA2N` | **Non-Submitters**: one row per learner per missed assignment, with contact details · **By Learner**: one row per learner, all their pending assignments · **Assignment Reconciliation**: the same per-assignment counts as report A |

Folder layout under each root, with period-folder names from `co-ordinator reports/coordinator_periods.py`:
```
Daily Assignment Report/Daily 04-Oct-2026/IntelliBI_Assignment_Submission_Performance_Report_Daily_04-Oct-2026
Weekly Assignment Report/Weekly 28-Sep-2026 to 04-Oct-2026/…_Weekly_28-Sep-2026_to_04-Oct-2026
Monthly Assignment Report/Monthly Sep-2026/…_Monthly_September_2026
Manual Assignment Report/Manual 21-Aug-2026 to 22-Sep-2026/…_Manual_21-Aug-2026_to_22-Sep-2026
```
The Non-Submission report uses the same names with `IntelliBI_Assignment_Non_Submission_Report`. A re-run for a period that already has a report **overwrites it in place**: the same Drive file and link get the new content (`OVERWRITE_EXISTING = True`). If Drive refuses the in-place update, a new file with the same name replaces it and the old copy goes to the Drive trash, where it can be restored. With `OVERWRITE_EXISTING = False`, the old file is kept and the new one is saved as "- Version N". Older "- Version N" copies from earlier runs are left untouched.

## Data and business rules (reused, not redefined)
- **Source:** IntelliBIAssessmentSubmission ▸ `Submissions`, read with `pyAssignmentSubmissionsReport.read_sheet_df`. It has one row per enrolled learner per assignment. The collector writes a row for every enrolled learner, and "enrolled" means the class roster plus everyone who submitted.
- **Expected learners** of an assignment = its rows. **Submitted** = any status other than "Not Submitted". **Not submitted** = status "Not Submitted". This is the counting rule of `pyAssignmentSubmissionsReport` and the "not submitted" test of `pyAssignmentSubmissionEmailReminder`.
- **One row per learner per assignment** (assignment = assessment_id + class_id). Duplicate rows are ignored, and a submitted row wins over a not-submitted one.
- **Timing (information):** On time = `submitted_at` ≤ deadline; Late = after it. Late submissions still count as submitted.
- **Active Learner (information):** the reminder's Students-tab check (`load_active_status_map` / `is_student_active`): Yes, No or Not found. Every non-submitter is listed regardless of this flag, so the reports reconcile.
- **Technology / Batch** = the class's `class_name` / `class_subject`, the same as the Coordinator task list's Tech Name / Duration.

## Colour code (same as `pyAttendaceFeedbackReport.py`)
The report uses the Attendance & Feedback report's own style helpers and palette (`AR.*`):
- Title, navy column headers, blue section banners, alternating row tint and a light-blue **⬛ TOTAL** row.
- KPI cards with the same icons and palettes: 📚 / 👥 blue, 📊 Submission % green at ≥ 75 % and red below, ❌ red, ✅ green.
- **Submission %** cells use the Att % colour code (`AR._att_bg`): dark green ≥ 95 %, light green ≥ 75 %, amber ≥ 60 %, red below. The text is dark green at ≥ 75 % and dark red below.
- **Learner Submission Detail rows** (whole row): Active Learner = No is **amber with strikethrough**, and this takes priority over the status. Otherwise Submitted is **green** and Not Submitted is **red**. This is presentation only; status and activity are unchanged.
- **Non-Submitters rows** use the absent reds.
- The e-mail figure uses the same 75 % / 60 % bands.
- **E-mail "By Technology" table:** coloured like the Attendance e-mail's *Attendance Summary — by Technology*, using its `AR._email_att_color`. Each technology row takes its band's background, and the Submission % is bold in the band colour: green above 70 %, amber 55–70 %, red below 55 %. A bold TOTAL row closes the table: the summed counts, with Submission % worked out from those sums and coloured by the same bands. A short legend sits under the table. Columns, values and order are unchanged.

## Eligibility
An assignment is reported only when:
1. its **deadline date** lies in the reporting period, and
2. its **deadline date and time** (IST, `DD/MM/YYYY HH:MM:SS IST`) had already passed when the report ran.

A deadline without a time counts as 23:59:59 of that day. Assignments with no readable deadline are never reported. Assignments due later in the period are not evaluated. Both counts are printed in the run log (`Not yet due` / `No deadline`).

## Periods (`pyLeadFollowUpAnalysisReport` flag framework, `coordinator_periods` helpers)
| Setting | Behaviour |
|---|---|
| `GENERATE_AUTO = True` | **Daily** every run, for **system date − 1** · **Weekly** on Monday, for the previous Mon–Sun week · **Monthly** on the 1st, for the whole previous calendar month |
| `GENERATE_AUTO = False` | `GENERATE_DAILY` (`DAILY_DATE`, None = system date − 1; a future date is refused) · `GENERATE_WEEKLY` (`WEEKLY_REFERENCE_DATE`, None = this week) · `GENERATE_MONTHLY` (`MONTHLY_MONTH` / `MONTHLY_YEAR`, None = current) · `GENERATE_MANUAL` (`MANUAL_START_DATE`..`MANUAL_END_DATE`) |
| Manual validation | Both dates required · Start ≤ End · **End ≤ system date − 1**. Otherwise the run stops that report with a clear `[CONFIG ERROR]` and exits 1. |

## Delivery
- **Drive only:** both workbooks are built in memory and uploaded as native Google Sheets by the impersonated `info@` (`UPLOAD_TO_DRIVE`). Nothing is written to local disk.
- **E-mail:** one e-mail per period from `info@intellibiinnovationstechnologies.in` to `info@` and `intellibihropsb2ch@gmail.com`, sent through `co-ordinator reports/coordinator_email.send`, the same sender, recipients and layout as the Coordinator Task Performance report. It carries the KPI cards, a by-technology table and links to both reports. Subject: `<Type> Assignment Submission Report - <period>`. `SEND_EMAIL = False` turns it off.
- **Gmail Star (★):** with `STAR_EMAIL_IN_GMAIL = True` (the default), each e-mail that was sent successfully is starred in the **info@** mailbox only, by the shared `common/gmail_star.py` (IMAP `\Flagged` = Gmail's STARRED label), exactly as for the Coordinator reports. There is one e-mail per Daily / Weekly / Monthly / Manual period, and it carries both the Performance and the Non-Submission report. Other recipients receive it normally. A failed send is never starred. A starring problem prints one `[Email] ★ not starred — …` warning and never changes the e-mail, the reports or the exit code.
- **Re-runs:** every run (AUTO, flag or Manual) rebuilds its periods, overwrites the existing reports and sends the e-mail again. This includes the 11:30 retry of the Morning batch when the 10:30 run already succeeded.
- **Retries:** Google reads go through `common/api_retry.py`. The file upload itself is not retried, so a lost response can never leave two copies. A failed period never stops the other periods, and the script exits 1.

## Run
```
python ops_reports_action\pyAssignmentSubmissionPerformanceReport.py            # planned reports
python ops_reports_action\pyAssignmentSubmissionPerformanceReport.py --dry-run  # read live, build, print; no upload, no e-mail
```
Run summary in the pipeline e-mail: Reports uploaded, Eligible assignments, Submitted, Not submitted, E-mailed.

## Verification
`python ops_validation\verify_assignment_submission_reports.py` runs offline against in-memory Google Drive and Sheets. It covers every period rule and Manual error, deadline passed / not passed / boundary / missing, submitted, not submitted, late and duplicate rows, reconciliation of the two reports, the Drive folders and file names under both roots, the in-place overwrite on a re-run (same files and links, new content), the replace-and-trash fallback, `OVERWRITE_EXISTING = False` versioning and the e-mail.

## Known limits
- Statuses are those in the Submissions sheet **when the report runs**. A historical period therefore shows today's view; for example, a late submission made after the period is counted as submitted (Late).
- An assignment with no rows in `Submissions` (no enrolled learners) does not appear.

## Change history
- 2026-10-05 — Learner Submission Detail: the assignment banner is reordered and now shows the class duration: `Class: SQL | Duration: 16-Sep-2026 To Current Date | 📝 T SQL Assignment | Start: 09/28/2026 | Deadline: 10/04/2026`. Duration is the class's `class_subject` from the Submissions data, the same value the reminder and Coordinator reports show as the class duration. Nothing else changed.
- 2026-10-05 — E-mail only: the By Technology table ends with a bold TOTAL row, as in the Attendance & Feedback e-mail. Assignments, Expected, Submitted and Not Submitted are summed, and Submission % is Submitted ÷ Expected of those sums, not an average of the technology percentages. The row is colour-coded by its own band. Technology rows and all calculations are unchanged.
- 2026-10-05 — A re-run now overwrites the period's existing reports in place (same file and link) instead of skipping them. `SKIP_IF_ALREADY_GENERATED` was replaced by `OVERWRITE_EXISTING = True`, and every run re-sends the e-mail.
- 2026-10-05 — Learner Submission Detail: the filter is kept (from the first column-header row to the last learner row) and freeze panes are removed. When a filter is applied, the 📝 banner rows are filtered like any other row.
- 2026-10-05 — Learner Submission Detail rebuilt as assignment blocks: a 📝 banner (assignment, class, start, deadline), then only #, Learner, Email, Phone, Submitted At, Status, Marks Obtained, Max Marks, Feedback and Active Learner. Rows are colour-coded with inactive (amber + strikethrough) first, then submitted (green) and not submitted (red). Feedback comes from `evaluation_feedback`. The tab's filter covers all the blocks. Calculations and other tabs are unchanged.
- 2026-10-05 — Assignment Performance tab: freeze panes removed, so no rows or columns are frozen. Other tabs are unchanged.
- 2026-10-05 — Summary tab: the NOTES section and its display-only code were removed; the tab ends with the Batch-wise table. Calculations, other tabs, the e-mail and the run summary are unchanged.
- 2026-10-05 — Every tab's existing row-1 header now ends with `  |  Generated On: DD-Mon-YYYY HH:MM AM/PM`, the report's generation time in IST. There is no new row and no styling change.
- 2026-10-05 — E-mail only: the By Technology table is colour-coded like the Attendance & Feedback e-mail (see Colour code). No calculation or content change.
- 2026-10-03 — Applied the colour code of `pyAttendaceFeedbackReport.py` (see above). The Submission % bands are now 95 / 75 / 60 instead of 80 / 50. Data, counts and layout are unchanged.
- 2026-10-03 — Summary tab: the Technology-wise Performance table was removed. The Batch-wise table, which also shows the technology, the KPIs and the notes are unchanged.
- 2026-10-03 — Gmail Star (★) on the report e-mail, in the info@ mailbox only, through the shared `common/gmail_star.py` (`STAR_EMAIL_IN_GMAIL = True`). Nothing else changed.
- 2026-10-03 — Created.
