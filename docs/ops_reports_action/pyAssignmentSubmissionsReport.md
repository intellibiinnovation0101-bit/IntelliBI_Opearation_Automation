# pyAssignmentSubmissionsReport.py — Assignment Submissions Report

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** AssignmentSubmissions (Layer 1)

## Purpose
Builds a styled Excel report of assignment submissions, run in AUTO mode off the trigger written by `pyAssignmentSubmissions.py`, or MANUAL mode for a date range.

## Input sources
- Google Sheet **IntelliBIAssessmentSubmission** (`Submissions`, `Assessment_Assigned`, `Assessment_Not_Assigned`).
- Trigger file `credentials/.assignment_report_trigger.json` (AUTO mode).

## Output
- Excel workbook with sheets: **Assignment_Summary**, **Class_Breakdown**, **Student_Detail**, **Not_Assigned**. Written to `output/reports/`, e-mailed / uploaded.

## Main business logic
- **AUTO** (default): if the trigger file exists → generate for the past 1 year, deadline ≤ today; if no trigger → exit without generating (no changes detected).
- **MANUAL**: user date range, bypasses the trigger; filters by `submission_start_date` in range + deadline ≤ today.

## Important fields / columns
`assessment_id`, `student_id`, `class_id`, `assessment_title`, `submission_status`, `submission_start_date`, `submission_deadline`, `score`.

## Dependencies
- Project: `common/utils.py`, `credentials/email_config.py`, `credentials/service_account.json`.
- Packages: `openpyxl`, `google-api-python-client`, `google-auth`.

## Execution flow
Check trigger (auto) → read Submissions → compute per-class/per-student metrics → build workbook → save → e-mail.

## Email / report behaviour
Sends via Gmail (`GMAIL_SENDER`). CLI flags: `--type auto|manual`, `--start`, `--end`, `--no-email`.

## Configuration used
`google.service_account_file`; email sender in `credentials/email_config.py`. Trigger path resolved in `credentials/`.

## Known issues
- Shares the trigger file with the Layer-1 submissions job; if that job did not run, AUTO mode exits without a report (correct behaviour).

## Future improvements
- Externalise the reporting window (past-1-year) into `config.yaml`.

## Change history
- 2026-08-24 — Moved into Operations project; service-account + trigger + email import made project-root-relative. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAssignmentSubmissionsReport.log`.

- **New submissions** — count for the current run
- **Updated submissions** — count for the current run
- **Reports generated** — count for the current run
- **E-mailed** — Yes/No — whether it was sent/uploaded this run
- *Note:* the e-mail may instead show “No new submissions — report skipped (nothing to report)”.

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
