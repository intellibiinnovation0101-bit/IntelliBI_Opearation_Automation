# pyStudentPaymentClassesStudentEnrolled.py — Student Info & Payments Pipeline

**Layer 1 — Operations Data Collection** · `ops_data_collection/`

## Purpose
Loads the student master data into the **IntelliBIStudentInfo** Google Sheet in one run: the class-enrolment roster (with instructors) and the students-with-payments dataset.

## Input sources
- **Wise API** (`from wise_config import HEADERS`):
  - Pipeline A — `GET /institutes/{id}/sessions` → `ClassLearnerTeacherEnrolled` (one row per class × enrolled student, with instructor(s)).
  - Pipeline B — `GET /institutes/v3/{id}/students` → `Students` + `Payments`.
- Class→instructor map fetched live (`_fetch_class_instructor_map`).

## Output
- Google Sheet **IntelliBIStudentInfo** — `SHEET ID 1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA`.
- Tabs: `ClassLearnerTeacherEnrolled`, `Students`, plus payments and Instructor tabs.

## Main business logic
- Each pipeline is **independent** — a failure in one does not block the others.
- **Failure protection**: if any paginated fetch fails mid-way, that pipeline aborts and **nothing** is written to its tab (no partial writes).
- Instructor tab loaded with **SCD Type-2** (`utils.upsert_scd2`) to keep history.
- Upserts via `utils.upsert_rows` (insert new / update changed / skip unchanged) with automatic header + column re-alignment.

## Important fields / columns
`Student_ID`, `Student_Name`, `Email`, `Phone`, `Class_ID`, `Instructor_Name`, payment status/amount/date fields, enrolment date.

## Dependencies
- Project: `common/utils.py` (`get_sheets_service`, `upsert_rows`, `upsert_scd2`, `overwrite_rows`, `sort_sheet_by_column`, date helpers), `credentials/wise_config.py` (HEADERS), `credentials/service_account.json`.
- Packages: `requests`, `google-api-python-client`, `google-auth`.

## Execution flow
1. Auth to Sheets (service account) + Wise API headers.
2. Pipeline A: fetch sessions → build class×student rows (+ instructor) → upsert.
3. Pipeline B: fetch students v3 (paginated) → flatten students + payments → upsert.
4. Instructors: SCD-2 upsert.

## Email / report behaviour
None — this is a data-refresh job (no e-mail).

## Configuration used
- `google.service_account_file` → `GOOGLE_SERVICE_ACCOUNT_FILE`.
- Wise credentials in `credentials/wise_config.py`.
- Cache under the project `cache/student_payment/` (routed from the old `.cache`).

## Known issues
- Depends on Wise API availability/quota; a mid-pagination failure aborts that pipeline (by design) so the tab keeps its previous good data.
- Instructor names depend on the live Wise session data being populated.

## Future improvements
- Move the institute id and Wise base URL into `config.yaml`.
- Add per-tab processed-row counts to the run summary.

## Change history
- 2026-08-24 — Moved into `IntelliBI_Operations_Automation/ops_data_collection/`; paths/cache/credentials made project-root-relative; Wise import switched to `wise_config`. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyStudentPaymentClassesStudentEnrolled.log`.

- **Students (full refresh)** — full dataset re-written this run (snapshot refresh)
- **Payment records** — count for the current run
- **Class-enrolment rows** — count for the current run
- **Instructors** — count for the current run
- **Students flagged removed** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
