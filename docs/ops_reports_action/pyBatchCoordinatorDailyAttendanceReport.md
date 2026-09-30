# pyBatchCoordinatorDailyAttendanceReport.py

Location: `co-ordinator reports/`. Output: `IntelliBI_Batch_Coordinator_Daily_Attendance_Report_<window>` (native Google Sheet in the coordinator Drive folder, date-wise) plus Weekly / Monthly / Manual roll-ups.

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
| ✎ Action Taken | dropdown, values per tab (`FOLLOWUP_ACTIONS`): Attendance = Call / WhatsApp / Call & WhatsApp / Email / No response / Other; Assignment adds WhatsApp Group; Admission = Form Send / Form Signed / Call / WhatsApp / Call & WhatsApp / Email / Other; Wise = Corrected / Invalid / Call / WhatsApp / Call & WhatsApp / Email / Other; Instructor & Interview = Call / WhatsApp / Call & WhatsApp / Email / Other |
| ✎ Follow-Up Comment | free text |
| ✎ Follow-Up Done? | dropdown Yes / No |
| ✎ Follow-Up DateTime | `dd-MMM-yyyy HH:MM:SS` (IST). Stamps itself the moment Follow-Up Done? is set, then holds that value; clearing Follow-Up Done? clears it, re-selecting stamps again |

How the timestamp works: the cell holds `=IF(Done="","",IF(OR(DT="",DT=0),NOW(),DT))` — a self-referencing formula that only evaluates `NOW()` while the cell is still empty. It needs *iterative calculation* to be on, so the script (a) sets it in the workbook (`ds_enable_iterative_calc`, for Excel) and (b) after the upload calls the Sheets API once per file (`_enable_followup_timestamps`) to switch iterative calculation on and pin the sheet's time zone to `Asia/Kolkata`. If that API call ever fails the log says so and the fix is manual: File ▸ Settings ▸ Calculation ▸ Iterative calculation = On. Dropdowns are standard list data validations (they survive the xlsx → Google Sheets conversion). Only flagged rows carry the input cells; banners and empty-state rows do not.

## Change history
- 2026-09-30 — Added the four coordinator follow-up tracking columns (Action Taken dropdown per tab, Follow-Up Comment, Follow-Up Done? Yes/No, self-stamping Follow-Up DateTime) to all six daily tabs; iterative calculation + IST time zone applied to the uploaded sheet. No report logic changed.
- 2026-09-30 — Report redesign (formatting only): common design system across all six daily tabs and the period roll-ups; priority chips, semantic-only colour, readable multi-reason Why Flagged with the action line, consistent headers/banners/widths/heights. No calculation, filter, threshold, follow-up condition or data selection changed.
- 2026-09-30 — Instructor Follow-Ups "Feedback Given" uses `AR.teacher_feedback_session_ids()` (Teacher_Feedback rows with content only), same rule as the AR Teacher_No_Feedback tab.
