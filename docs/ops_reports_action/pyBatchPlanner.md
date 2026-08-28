# pyBatchPlanner.py — Batch Planner Intelligence System

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** *(independent — reads local `data_inputs/` workbooks)*

## Purpose
Analyses enrolment, course structure, technology mapping, session schedules, dependencies and timing preferences to recommend **what technology batch to start next, for which students, in which timing slot**.

## Input sources — local workbooks in `data_inputs/` (via `BP_BASE_DIR`)
| File | Sheet(s) |
|------|----------|
| `CourseAndTechnologiesMappingDocument.xlsx` | `CourseAndTechnologiesMappingDoc` |
| `TechnologyMappingDocuments.xlsx` | `TechnologyMapping` |
| `IntellBIAttendance.xlsx` | `Sessions` |
| `IntelliBIStudentInfo.xlsx` | `Students`, `ClassLearnerTeacherEnrolled` |

## Output
- **Batch_Planner_Report.xlsx** (IntelliBI-styled) uploaded to Google Drive folder `1goPKLAlbL-cEC8r9x7P0a4xAeg1pvDmN` (impersonating `info@intellibiinnovationstechnologies.in`).
- Sheets: **Student_Progress_Report**, **Batch_Planning_Report** (technology × timing), **Batch_Recommendation** (ranked what-to-start-next), **Bottlenecks_Insights**, **Assumptions**.

## Main business logic
- `START_BUFFER_DAYS` (2–5 day buffer after a running batch ends).
- `ACTIVE_STUDENTS_ONLY`; dummy/free students & dummy shortforms excluded.
- `NON_PLANNABLE_KEYWORDS` (placement, project, interview, real-time domain, learning journey, sample course, business communication) never planned as standalone batches.
- Ranks demand vs bottleneck vs idle capacity to recommend the next batch.

## Important fields / columns
Student, technology, course, timing slot, enrolment/attendance status, demand count, recommendation rank.

## Dependencies
- Project: `credentials/service_account.json` (Drive upload). Packages: `pandas`, `numpy`, `openpyxl`, `google-api-python-client`, `google-auth`.

## Execution flow
Load 4 workbooks → build student progress → compute demand/bottlenecks → rank recommendations → build styled workbook → upload to Drive.

## Email / report behaviour
No e-mail; delivers via Drive upload (`GDRIVE_UPLOAD=True`).

## Configuration used
- `batch_planner.base_dir` → `BP_BASE_DIR` (default `data_inputs/`).
- `google.service_account_file`.

## Known issues
- Reads **local** workbooks, not the live sheets — refresh the four files in `data_inputs/` before running for the latest data.
- `IntellBIAttendance.xlsx` is large (~4 MB); keep it current.

## Future improvements
- Optionally read the Sessions/Students data directly from the live Google Sheets so `data_inputs/` refresh is not required.

## Change history
- 2026-08-24 — Moved into Operations project; `BP_BASE_DIR` defaulted to `data_inputs/`; service-account path made project-root-relative. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyBatchPlanner.log`.

- **Students analysed** — volume handled this run (context, not a change count)
- **Batch-plan rows** — count for the current run
- **Report uploaded** — Yes/No — whether it was sent/uploaded this run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
