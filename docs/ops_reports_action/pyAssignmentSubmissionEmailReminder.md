# pyAssignmentSubmissionEmailReminder.py — Assignment Submission Email Reminder

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** AssignmentSubmissions (Layer 1)

## Purpose
Sends deadline-approaching reminder e-mails to students with pending submissions, and a batch-wise PDF summary to staff. This is an **action** script (sends student e-mails), not just a report.

## Input sources
- Google Sheet **IntelliBIAssessmentSubmission** (`Submissions`).

## Output
- Per-student reminder e-mails (consolidated: one e-mail lists all of a student's pending assignments at the most urgent level).
- A batch-wise **PDF** report (reminder level + pending students per batch), uploaded to Google Drive.
- A staff summary e-mail; console batch/subject-wise summary.

## Main business logic — reminder ladder
For each `submission_status = "Not Submitted"`, compare `submission_deadline` to today:
- 2 days before → **1st Reminder** (gentle nudge)
- 1 day before → **2nd Reminder** (urgent warning)
- deadline day → **Final Reminder** (last chance)
- 1 day after → **Missed Deadline** (overdue follow-up; sent once, notes 3 reminders already sent)

Consolidates multiple pending assignments per student into one e-mail using the most-urgent level.

## Important fields / columns
`student_id`, `student_email`, `student_name`, `batch`, `assessment_title`, `submission_deadline`, `submission_status`.

## Dependencies
- Project: `common/utils.py`, `credentials/email_config.py`, `credentials/service_account.json`.
- Packages: `openpyxl`, `reportlab` (PDF), `google-api-python-client`, `google-auth`.

## Execution flow
Read Submissions → classify reminder levels → consolidate per student → send student e-mails → build batch PDF → upload → send staff summary.

## Email / report behaviour
- Student reminders sent via Gmail (`GMAIL_SENDER`).
- `CC_RECIPIENTS = []`; `STAFF_SUMMARY_RECIPIENTS = [intellibihropsb2ch@gmail.com, info@intellibiinnovationstechnologies.in]`.
- RUN CONFIGURATION (top of file): `report_date`, `send_email`, `dry_run`, `generate_pdf`, `include_overdue_report`, `overdue_all`. **Use `dry_run` when testing so no student e-mails go out.**

## Configuration used
`google.service_account_file`; email sender in `credentials/email_config.py`. Recipient/CC lists preserved in-script.

## Known issues
- Because it sends real student e-mails, always verify with `dry_run=True` first when changing logic. Skipped by `run_all.py` if the Layer-1 submissions refresh failed.

## Future improvements
- Move reminder-day thresholds + staff recipients into `config.yaml`.

## Change history
- 2026-08-24 — Moved into Operations project; service-account + email import made project-root-relative. No business-logic / reminder-behaviour change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAssignmentSubmissionEmailReminder.log`.

- **Students to notify** — volume handled this run (context, not a change count)
- **Reminders sent** — count for the current run
- **Reminders failed** — records/items that errored this run
- **Pending assignments** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
