# pyAttendaceFeedbackReport.py — Attendance & Feedback Report

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** SessionAttendance, StudentPayment (Layer 1)

## Purpose
Reads live attendance/feedback data, builds a styled Excel report, and e-mails it.

## Input sources
- Google Sheet **IntellBIAttendance** (`Sessions`, `Attendance`, `Student_Feedback`).
- Google Sheet **IntelliBIStudentInfo** (`Students` — for phone numbers).

## Output
- Styled Excel workbook with sheets: **Session Summary**, **Student Detail**, **Absent & At-Risk**, **Feedback Rating**. Written to `output/reports/`, e-mailed via Gmail.

## Main business logic
- Report period controlled by top-of-file RUN CONFIGURATION: `report_type` = daily / weekly / fortnightly / monthly / quarterly / yearly (or a specific `report_date`), `start_date`/`end_date`, `send_email`.
- Per-session attendance metrics; per-student per-session breakdown; absent & at-risk action list; session-level feedback stats.
- **Instructor feedback "received" rule** (`teacher_feedback_session_ids()`, shared with the Batch Coordinator report's *Instructor Follow-Ups*): a session counts as having instructor feedback only when its `Teacher_Feedback` row carries content — non-blank `topics_covered` or `comments`. The LMS creates a `teacherFeedback` record the moment the instructor marks the session *Completed* (`session_status = COMPLETED`, `created_at` = session end) before any feedback is written; that placeholder row (blank topics + comments, shown in the LMS as "Instructor Feedback: Pending") is NOT feedback. `Teacher_No_Feedback` (daily list, grouped counts, period status) and the `[Data]` log line ("with feedback content / completion-only placeholders") all use this one rule.

## Important fields / columns
`session_id`, `class_id`, `student_id`, attendance status, present/absent counts, attendance %, feedback rating, phone.

## Dependencies
- Project: `common/utils.py` (`get_sheets_service`), `credentials/email_config.py` (`GMAIL_SENDER`, `GMAIL_APP_PASS`), `credentials/service_account.json`.
- Packages: `openpyxl`, `google-api-python-client`, `google-auth`.

## Execution flow
Read sheets → compute metrics → build styled workbook → save to `output/reports/` → e-mail.

## Email / report behaviour
Sends the report via Gmail SMTP (`GMAIL_SENDER`). Recipient list and `send_email` toggle are near the top of the script (preserved as-is).

## Configuration used
`google.service_account_file`; email sender in `credentials/email_config.py`. Depends on fresh IntellBIAttendance + IntelliBIStudentInfo (Layer 1).

## Known issues
- If Layer 1 attendance refresh failed, `run_all.py` skips this report to avoid stale numbers.

## Future improvements
- Move recipient list + report_type default into `config.yaml`.

## Change history
- 2026-09-30 — `read_sheet_df()` retries transient Google errors (socket time-out / connection reset — e.g. `WinError 10060` on the OAuth token endpoint — and HTTP 429/5xx) with back-off via `common/api_retry.py` (`pipeline.api_retry_attempts` / `api_retry_base_wait_sec` in `config.yaml`), so a network blip at the first read no longer aborts the run. Permanent errors still raise at once. Also covers `pyBatchCoordinatorDailyAttendanceReport.py`, which reads through the same function.
- 2026-09-30 — Instructor feedback detection fixed: `Teacher_No_Feedback` used to treat ANY `Teacher_Feedback` row as "feedback given", so sessions whose instructor had only marked them Completed (placeholder row with blank topics/comments — e.g. Power BI, 30-Sep-2026) never appeared. New shared `teacher_feedback_session_ids()` requires feedback content; `pyBatchCoordinatorDailyAttendanceReport.py` (daily + period Instructor Follow-Ups, "Feedback Given") now calls the same helper instead of its own presence check. No session that has real feedback changes; only completion-only placeholders are now reported as missing (43 of 104 sessions in Sep-2026 at the time of the fix). Verify offline: `python ops_validation\verify_teacher_feedback_detection.py`.
- 2026-08-24 — Moved into Operations project; service-account path + email import made project-root-relative (`from email_config import ...`). No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAttendaceFeedbackReport.log`.

- **Reports generated** — count for the current run
- **E-mailed** — Yes/No — whether it was sent/uploaded this run
- **Sessions covered** — volume handled this run (context, not a change count)
- **Students covered** — volume handled this run (context, not a change count)
- **Absent** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
