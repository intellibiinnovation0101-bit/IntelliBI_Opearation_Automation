# pyAdmissionFormalitiesReport.py — Admission Formalities Reports

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** StudentPayment, ZohoSignatureStatus (Layer 1)

## Purpose
Two independent admission-formalities reports built on the existing plumbing (does not modify any other report).

## Input sources
- Google Sheet **IntelliBIStudentInfo**, tab `Students`.
- Google Sheet **"Zoho Sign - Signers"**, tab `Signers` (produced by `pyZohoSignatureStatusRefresh.py`).

## Output
- **Report 1 — Historical**: consolidated admission-form / signature status of ALL active students.
- **Report 2 — Current (daily)**: only students still PENDING admission formalities (joined on/after `CURRENT_REPORT_START_DATE`, not yet signed), with row colour-coding.

## Main business logic — matching (case-insensitive, trimmed)
Per student: match on **Email ID** first, then **Student Name** as fallback; first match wins. Joins student roster to Zoho signer status.

## Important fields / columns
`student_name`, `email`, join date, `recipient_status`, `sent_date`, `signed_date`, `expiry_date`, pending flag.

## Dependencies
- Project: `common/utils.py` (`get_sheets_service`, `ensure_tab_exists`, `_normalize`, `sort_sheet_by_column`, retry helpers), `credentials/service_account.json`.
- Packages: `openpyxl`, `google-api-python-client`, `google-auth`.

## Execution flow
Read Students + Signers → match per student → build Historical + Current reports → write to output sheets (service account needs EDITOR on both).

## Email / report behaviour
Writes to Google Sheets output tabs (colour-coded); see script for any e-mail step.

## Configuration used
`google.service_account_file`. **Prerequisite:** run `pyZohoSignatureStatusRefresh.py` first so the Signers tab is current — enforced by the Layer-1 dependency gate.

## Known issues
- Name-fallback matching can mis-match students with identical names lacking an email; email match is preferred.

## Future improvements
- Externalise `CURRENT_REPORT_START_DATE` and output sheet ids into `config.yaml`.

## Change history
- 2026-08-24 — Moved into Operations project; service-account path made project-root-relative. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAdmissionFormalitiesReport.log`.

- **Students processed** — volume handled this run (context, not a change count)
- **Inserted** — count for the current run
- **Updated** — count for the current run
- **Unchanged** — already up to date / no change this run
- **Failed** — records/items that errored this run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
