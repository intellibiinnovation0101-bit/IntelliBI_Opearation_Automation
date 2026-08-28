# pyStudentAdditionalNote.py — Student Additional Note Generator

**Layer 2 — Operations Reports & Actions** · `ops_reports_action/`
**Depends on:** *(independent)*

## Purpose
Reads student admission responses from one Google Sheet, generates formatted note data, and writes it into the destination "Note List" tab — **without disturbing an IMPORTRANGE-driven student list**.

## Input sources
- Source Google Sheet: student admission responses.
- Destination Google Sheet: "Note List" tab whose columns A:D (student_id / name / email / phone) are an **IMPORTRANGE** array-formula spill.

## Output
- Writes ONLY the `Generated_Data` and `IsRecordUpdated` columns of the Note List tab.

## Main business logic — IMPORTRANGE-safe writes ⚠
- **Never** inserts/deletes rows and **never** writes into columns A:D (doing so collides with the IMPORTRANGE spill and makes the imported list vanish with `#REF!`).
- Because A:D spill live, students shift row position when the source gains/loses/reorders students. Each run **re-aligns** both `Generated_Data` and `IsRecordUpdated` to the current names, recovering each student's existing `IsRecordUpdated` value from the sheet's own prior state (candidate name embedded in each note) and moving it to follow its student. The value is never invented or changed (`_build_isrecordupdated_owner_map()`).

## Important fields / columns
A:D (IMPORTRANGE spill — read-only), `Generated_Data`, `IsRecordUpdated`.

## Dependencies
- `gspread`, `google-auth`; `credentials/service_account.json`.

## Execution flow
Read admission responses + current Note List → build owner map for IsRecordUpdated → generate note text → re-align + write only Generated_Data + IsRecordUpdated.

## Email / report behaviour
None — sheet-writing action.

## Configuration used
`google.service_account_file` (service account = `CREDENTIALS_FILE`, now resolved in `credentials/`).

## Known issues
- **Critical**: any change that inserts/deletes rows or touches A:D will break the destination list. Preserve the IMPORTRANGE-safe contract exactly.

## Future improvements
- Add a self-check that aborts if columns A:D are detected as stored (non-formula) values.

## Change history
- 2026-08-24 — Moved into Operations project; service-account path made project-root-relative. No business-logic change; IMPORTRANGE-safe contract preserved.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyStudentAdditionalNote.log`.

- **Responses processed** — volume handled this run (context, not a change count)
- **Matched to students** — count for the current run
- **No response** — already up to date / no change this run
- *Note:* the e-mail may instead show “No changes to write this run”.

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
