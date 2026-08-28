# pyStudentProfileReport.py — Unified Student Profile Report

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** SessionAttendance, AssignmentSubmissions, StudentPayment (Layer 1)

## Purpose
Produces ONE individual PDF per student combining attendance, assignment submissions, and interview feedback — plus a single MASTER Excel Student Progress Report for the CEO/admin.

## Input sources
- Attendance pipeline sheets (IntellBIAttendance).
- Assignment Submissions sheets (IntelliBIAssessmentSubmission).
- Consolidated Interview sheet (interview feedback & scores).

## Output
- One **PDF per student** + one **master Excel** workbook, written to a month-wise folder (e.g. `output/StudentProgressReport_May_2026/`).
- Master workbook e-mailed to the admin address only.

## Main business logic — "Cumulative + month highlight"
Each PDF shows the student's full all-time history plus a "This Month" highlight strip for the most-recent completed calendar month. Combines the three domains (attendance %, assignment submission status, interview scores) into one profile.

## Important fields / columns
`student_id`, `student_name`, attendance %, assignments submitted/pending, interview scores, month highlights.

## Dependencies
- Project: `common/utils.py`, `credentials/email_config.py` (uses the **DIGITAL** sender `intellibidigital@gmail.com`), `credentials/service_account.json`.
- Packages: `reportlab` (PDF), `openpyxl`, `google-api-python-client`, `google-auth`.

## Execution flow
Read attendance + assignment + interview sheets → per student build PDF → build master workbook → e-mail master to admin.

## Email / report behaviour
Uses `GMAIL_SENDER_DIGITAL` (imported as `GMAIL_SENDER`). Master workbook e-mailed to the admin address (in-script). `--output` CLI flag overrides the output folder.

## Configuration used
`google.service_account_file`; DIGITAL email sender in `credentials/email_config.py`.

## Known issues
- Depends on all three source domains being fresh; `run_all.py` skips it if any of its three Layer-1 dependencies failed.
- Generates many PDFs — can be slow for large cohorts.

## Future improvements
- Parallelise per-student PDF generation; move output root to `output/reports/` via a config key.

## Change history
- 2026-08-24 — Moved into Operations project; service-account + email import made project-root-relative (uses DIGITAL sender). No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyStudentProfileReport.log`.

- **Students processed** — volume handled this run (context, not a change count)
- **Failed** — records/items that errored this run
- **Master report e-mailed** — Yes/No — whether it was sent/uploaded this run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
