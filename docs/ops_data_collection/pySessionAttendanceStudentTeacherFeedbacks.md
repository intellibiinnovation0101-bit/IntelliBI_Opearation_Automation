# pySessionAttendanceStudentTeacherFeedbacks.py — Session Attendance & Feedback Pipeline

**Layer 1 — Operations Data Collection** · `ops_data_collection/`

## Purpose
Incrementally loads session attendance and student/teacher feedback into the **IntellBIAttendance** Google Sheet. (Ported from the n8n workflow `wfIntelliBISessionAttendanceFeedbacks`.)

## Input sources
- **Wise API** (`from wise_config import HEADERS`) — sessions, attendance, and feedback endpoints.
- Class→instructor map fetched live (`_fetch_class_instructor_map`).

## Output
- Google Sheet **IntellBIAttendance** — `SHEET ID 1TqDjq4gAyo32eRNMbuLd6uu0eCNZb7h1j5YH-q68AhU`.
- Tabs written: `Sessions`, `Attendance`, `Student_Feedback`, `Teacher_Feedback`, `Sessions_No_TF` (sessions missing teacher feedback), `Watermark_Attendance` (internal state).

## Main business logic — incremental (watermark) load
- **First run** (no watermark): fetches all data from 2020-01-01 and stores the max watermark value **per sheet**.
- **Next runs**: fetches from the watermark date onward, filters to sessions **strictly after** the stored watermark, appends only new rows, then advances the watermark to the new max.
- Each sheet has its **own independent watermark**, so tabs can be at different positions.
- `Sessions_No_TF` highlights sessions with no teacher feedback submitted.

## Important fields / columns
`session_id`, `class_id`, `Instructor_Name`, `student_id`, attendance status, feedback ratings/comments, `sync_key | load_type | last_sync_time | total_synced` (watermark tab).

## Dependencies
- Project: `common/utils.py`, `credentials/wise_config.py`, `credentials/service_account.json`.
- Packages: `requests`, `google-api-python-client`, `google-auth`.

## Execution flow
1. Read each tab's watermark (create `Watermark_Attendance` if missing).
2. Fetch sessions/attendance/feedback from the watermark date.
3. Filter to strictly-new records, append, update watermark.
4. Step 5c (every run, watermark-neutral): re-fetch the last `BACKFILL_LOOKBACK_DAYS` (7) days FRESH from the LMS, append any attendance/feedback rows that are missing, and **refresh in place** any existing Attendance row whose facts changed since it was first written (`reconcile_attendance_rows`). Only the fact columns in `ATTENDANCE_FACT_COLUMNS` (session start/end, duration, attendance %, first join, last leave, status, student name, email) are refreshed, matched by column NAME; a fresh blank never erases a value and an existing "Present" is never downgraded. When rows were refreshed, the Sessions tab is re-aligned to them (steps 5b/5b2 re-run).
   - One-off repair of older rows: `python ops_data_collection/pySessionAttendanceStudentTeacherFeedbacks.py --backfill-days 60` (any N; watermarks are not touched).

## Email / report behaviour
None — data-refresh job.

## Configuration used
- `google.service_account_file`; Wise credentials in `credentials/wise_config.py`.
- Cache under project `cache/session_attendance/`.

## Known issues
- **Resolved 2026-09-26 — a session showed 0 Present although the LMS showed attendance.** Root cause: the scheduled run synced the session BEFORE the LMS had finalised its attendance (same-day run while the class was on / minutes after it ended), so every participant was written with duration 0 = "Absent". The tab is append-only and Step 5c deduplicated on `session_id+student_id`, so those rows were frozen forever and the report (which simply counts `status == Present`) kept showing 0. Fix: Step 5c now refreshes existing rows whose facts changed (see Execution flow, step 4). Verification script: `python ops_validation\verify_attendance_reconcile.py`.
- A report generated while a session is still live (or within minutes of its end, before the LMS publishes durations) can only show the attendance known at that moment; the next pipeline run corrects the sheet automatically.
- Watermark logic assumes source timestamps are monotonic; out-of-order backfills in the source may need a manual watermark reset (clear the `Watermark_Attendance` tab to force a full reload).

## Future improvements
- Configurable initial backfill date and institute id via `config.yaml`.

## Change history
- 2026-09-26 — Step 5c refreshes stale Attendance rows in place (`reconcile_attendance_rows`, `ATTENDANCE_FACT_COLUMNS`), re-aligns Sessions afterwards, new `--backfill-days N` option; verification script `ops_validation/verify_attendance_reconcile.py`. No change to the watermark engine, feedback tabs, report script or output format.
- 2026-08-24 — Moved into Operations project; portability + cache/credentials paths made project-root-relative; Wise import switched to `wise_config`. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pySessionAttendanceStudentTeacherFeedbacks.log`.

- **New Sessions** — count for the current run
- **New Attendance rows** — count for the current run
- **New Student Feedback** — count for the current run
- **New Teacher Feedback** — count for the current run
- **Source sessions fetched** — volume handled this run (context, not a change count)
- **Suspended students found** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
