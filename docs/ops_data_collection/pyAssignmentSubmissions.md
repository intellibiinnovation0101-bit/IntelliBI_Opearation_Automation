# pyAssignmentSubmissions.py — Assignment Submissions Pipeline

**Layer 1 — Operations Data Collection** · `ops_data_collection/`

## Purpose
Collects every student assignment (assessment) submission across all classes into the **IntelliBIAssessmentSubmission** Google Sheet, and drops a trigger file so the Assignment Submissions Report knows new data is available.

## Input sources
- **Wise API** (`from wise_config import HEADERS`):
  - `/institutes/{INSTITUTE_ID}/classes` (all classes)
  - `/user/classes/{id}/contentTimeline` (per class → find assessments)
  - `/user/getAssessment/{assessment_id}` (details + submissions)

## Output
- Google Sheet **IntelliBIAssessmentSubmission** — `SHEET ID 1E_pOuZfw4BUhQ1bRMmuc8lPDDJRtoXuGkP-qtHC96HU`, tab **`Submissions`**.
- Trigger file **`.assignment_report_trigger.json`** (in `credentials/`) signalling the report that changes occurred.

## Main business logic
1. Fetch all classes → dedupe.
2. Per class: fetch content timeline; recursively walk it collecting every entity whose `entityType/type/contentType == 'assessment'`.
3. Dedupe assessment IDs across classes.
4. Per assessment: fetch details + submissions → flatten to one row per student submission.
5. **Upsert** keyed on `assessment_id + student_id` (insert new / update changed / skip unchanged) — safe to re-run anytime.

## Important fields / columns
`assessment_id`, `student_id`, `class_id`, `assessment_title`, `submission_status`, `submission_start_date`, `submission_deadline`, `score`, `submitted_at`.

## Dependencies
- Project: `common/utils.py` (`get_sheets_service`, `upsert_rows`, `clean_sheet`, date helpers), `credentials/wise_config.py`, `credentials/service_account.json`.
- Packages: `requests`, `google-api-python-client`, `google-auth`.

## Execution flow
Classes → timelines → assessments → submissions → flatten → upsert → write trigger.

## Email / report behaviour
None directly; its trigger file drives `pyAssignmentSubmissionsReport.py`.

## Configuration used
- `google.service_account_file`; Wise credentials. Cache under project `cache/` (assessments subfolder).

## Known issues
- Deep content-timeline nesting relies on the recursive walk; unusual timeline shapes could miss an assessment (defensive dedupe mitigates duplicates).

## Future improvements
- Externalise `INSTITUTE_ID` and endpoints to `config.yaml`.

## Change history
- 2026-08-24 — Moved into Operations project; trigger + cache + credentials paths made project-root-relative; Wise import switched to `wise_config`. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyAssignmentSubmissions.log`.

- **Submissions extracted** — volume handled this run (context, not a change count)
- **New submissions** — count for the current run
- **Updated submissions** — count for the current run
- **Unchanged** — already up to date / no change this run
- **New assignments detected** — count for the current run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
