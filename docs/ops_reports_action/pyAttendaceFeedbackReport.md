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
- 2026-08-24 — Moved into Operations project; service-account path + email import made project-root-relative (`from email_config import ...`). No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAttendaceFeedbackReport.log`.

- **Reports generated** — count for the current run
- **E-mailed** — Yes/No — whether it was sent/uploaded this run
- **Sessions covered** — volume handled this run (context, not a change count)
- **Students covered** — volume handled this run (context, not a change count)
- **Absent** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
