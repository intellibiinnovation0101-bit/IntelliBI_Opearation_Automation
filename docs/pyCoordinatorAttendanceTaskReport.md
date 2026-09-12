# pyCoordinatorAttendanceTaskReport.py

**Layer:** 2 — Operations Reports & Action (`ops_reports_action/`)
**Purpose:** Daily Batch Coordinator **Attendance** Task & Performance Monitoring
System. Produces a technology-wise action list of the students who *genuinely*
need follow-up (fusing today's/latest performance with till-date history), lets
the coordinator record the follow-up taken, and gives Management a live view of
follow-up completion — the future source for the coordinator's KRA.

```
Daily Monitoring -> Intelligent Follow-Up Identification -> Coordinator Action
                 -> Follow-Up Tracking -> Management Performance Monitoring
```

## Output
A **day-wise sub-folder** under Google Drive folder
`1BEokUc7Np7iBVSMwIyMAgZUa0mrecT-h`, each containing **one native Google Sheet**
(`Coordinator Attendance Tasks — YYYY-MM-DD`). Previous days are never
overwritten. Re-running the same day **preserves** coordinator entries already
typed in (merged by `student_id + technology`).

Tabs:

| Tab | Audience | Contents |
|-----|----------|----------|
| **Summary** | Management | Per-technology + overall KPIs; **live** completion formulas (completed / pending / completion % / critical unattended) that update as the coordinator fills the Outcome column. |
| **Follow-Up Tasks** | Coordinator | Flagged students only. Protected read-only diagnostics + editable entry columns (dropdowns, conditional formatting). Sorted Technology → Priority. |
| **Master Data** | Management / analytics | Every applicable student with full metrics + score + band + Requires-Follow-Up. Clean KRA feed. |
| **Guide** | Everyone | Legend, scoring model, how-to. |

## Data sources (read-only)
- `IntellBIAttendance` (`1TqDjq4g…q68AhU`) — `Attendance`, `Student_Feedback`
- `IntelliBIStudentInfo` (`1Eq7Q3Go…MWzVA`) — `Students`
  (`phone`, `Batch_Timing`, `is_attendance_required`, `Is_Deleted`)

"Technology" = the `course_name` field. Attendance maths mirror the validated
`pyAttendaceFeedbackReport.py` (Present ÷ distinct sessions; **suspended** and
`is_attendance_required != Y` rows excluded), so this report never conflicts
with the existing calculation.

## How a student is flagged — Attention Score (0–100)
Transparent weighted sum of risk signals (higher = more attention needed):

| Signal | Default weight |
|--------|----------------|
| Overall attendance shortfall | 30 |
| Absent in latest session | 15 |
| Consecutive-absence streak | 20 |
| Recent 5-session trend | 10 |
| Feedback participation | 12 |
| Feedback rating (/10) | 8 |
| No feedback despite attending | 5 |

Bands: **Critical ≥ 65**, **High ≥ 45**, **Medium ≥ 30**, below = not flagged.
**Hard triggers** also force a flag regardless of score: attendance < 60 %,
3+ absences in a row, absent-again with history < 75 %, zero feedback despite
attending, or a low average rating. A student absent once but with excellent
history and no streak scores low and is intentionally **left off**; a student
present today but with weak history/feedback **is** flagged. Each row carries a
dynamic **Why Flagged** and **Recommended Action**.

## Coordinator entry columns
`Follow-Up Done?` · `Action Taken` · `Follow-Up Comment` (required) · `Outcome`
· `Coordinator` · `Follow-Up Date` · `Next Action`. Diagnostic columns are
protected (warning-only — never blocks automation).

## Configuration
All weights / thresholds / band cut-offs / dropdown option lists live in
`config/coordinator_attendance_config.json` (**optional** — the script
self-defaults). No hardcoded students, technologies, dates or sessions.

Run-block constants at the top of the script: `SEND_EMAIL` (default **False**),
`REPORT_DATE` (pin the report day; default derives from the latest data),
`VERBOSE`.

## Auth
Service account `credentials/service_account.json` impersonating
`info@intellibiinnovationstechnologies.in` (domain-wide delegation), scopes
`spreadsheets` + `drive` — same pattern as `pyAttendaceFeedbackReport.py`, so
created files are owned by `info@` and inherit access from the shared folder.

## Run
```bat
.venv\Scripts\python.exe ops_reports_action\pyCoordinatorAttendanceTaskReport.py
REM self-test (no Google calls):
.venv\Scripts\python.exe ops_reports_action\pyCoordinatorAttendanceTaskReport.py --selftest
```

## Pipeline wiring (optional)
Not auto-discovered. To include it in Layer 2, add to the `JOBS` list in
`scripts/run_reports_action.py`:
```python
("pyCoordinatorAttendanceTaskReport", "Batch Coordinator Attendance Tasks",
    ["pySessionAttendanceStudentTeacherFeedbacks", "pyStudentPaymentClassesStudentEnrolled"]),
```

## Dependencies
`google-api-python-client`, `google-auth` (already in the project). No pandas/numpy.

## Change history
- 2026-09-11 — Initial version: technology-wise daily attendance follow-up
  identification (weighted rule-based Attention Score), day-wise native Google
  Sheet output, coordinator entry + protected diagnostics, live management
  completion tracking, same-day entry preservation, config-driven tuning.
