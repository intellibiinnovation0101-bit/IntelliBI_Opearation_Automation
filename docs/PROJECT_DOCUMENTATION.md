# IntelliBI Operations Automation — Project Documentation

A clean, layered, **portable** operations pipeline that refreshes the source
datasets (student info, attendance, assignments, Zoho Sign) and then produces the operations reports, reminders and
warnings — with centralized configuration, centralized logging, per-script
dependency gating, and a one-command launcher that e-mails a run summary.

Every path resolves from the project root at run time (`common/paths.py`), so the
whole `IntelliBI_Operations_Automation/` folder can be copied to another machine,
configured once, and scheduled — **without editing any source-code paths**.

---

## 1. Architecture

```
Layer 1  ops_data_collection      4 refresh jobs, ALL in PARALLEL
   |
   |  verify success per job
   v
Layer 2  ops_reports_action       reports/actions in PARALLEL where safe;
                                   a report whose Layer-1 dependency FAILED is
                                   SKIPPED, independent ones still run
```

### Layer 1 — Operations Data Collection / Refresh (`ops_data_collection/`)

| Script | Refreshes |
|--------|-----------|
| `pyStudentPaymentClassesStudentEnrolled.py` | IntelliBIStudentInfo (Students, Payments, ClassLearnerTeacherEnrolled) |
| `pySessionAttendanceStudentTeacherFeedbacks.py` | IntellBIAttendance (Sessions, Attendance, Feedback) |
| `pyAssignmentSubmissions.py` | IntelliBIAssessmentSubmission (Submissions) + report trigger |
| `pyZohoSignatureStatusRefresh.py` | Zoho Sign → "Zoho Sign - Signers" sheet |

### Layer 2 — Operations Reports & Actions (`ops_reports_action/`)

| Script | Depends on (Layer 1) |
|--------|----------------------|
| `pyAttendaceFeedbackReport.py` | SessionAttendance, StudentPayment |
| `pyAssignmentSubmissionsReport.py` | AssignmentSubmissions |
| `pyAssignmentSubmissionEmailReminder.py` | AssignmentSubmissions |
| `pyBatchPlanner.py` | *(independent — reads local `data_inputs/`)* |
| `pyStudentProfileReport.py` | SessionAttendance, AssignmentSubmissions, StudentPayment |
| `pyAdmissionFormalitiesReport.py` | StudentPayment, ZohoSignatureStatus |
| `pyStudentAdditionalNote.py` | *(independent)* |

Per-script docs are in `docs/ops_data_collection/` and `docs/ops_reports_action/`.

---

## 2. Folder structure

```
IntelliBI_Operations_Automation/
├── ops_data_collection/     Layer 1 entry scripts
├── ops_reports_action/      Layer 2 entry scripts
├── common/                  shared code + portability layer
│   ├── paths.py             PROJECT_ROOT + all folders (pathlib)
│   ├── _bootstrap.py        sys.path + env defaults + config.yaml (imported first)
│   ├── config_loader.py     reads config/config.yaml -> environment
│   ├── logging_utils.py     centralized logger factory (logs/ + console)
│   ├── common_utils.py      subprocess runner, record-count parsing, summary e-mail
│   ├── google_sheet_utils.py facade over the shared Google Sheets helpers
│   ├── instructor_utils.py  reusable class -> instructor mapping helper
│   └── utils.py, interakt_common.py   the pipeline's shared modules
├── config/                  config.yaml, logging_config.yaml
├── credentials/             secrets & session state (git-ignored)
├── data_inputs/             Batch Planner source workbooks (git-ignored)
├── cache/  temp/  logs/     working folders (git-ignored, auto-created)
├── output/reports/          generated report files
├── output/exports/          CSV/XLSX exports (generic)
├── scripts/                 run_data_collection.py, run_reports_action.py, run_all.py
├── docs/                    this documentation
├── requirements.txt  .gitignore  README.md  run_all.bat
```

---

## 3. Portability

Every script's first project import is `common/_bootstrap.py` (auto-inserted at
the top, **before** any other project import). It discovers `PROJECT_ROOT` from
its own file location, puts `common/` and `credentials/` on `sys.path`, seeds the
environment-variable defaults the scripts already honour (service account,
`BP_BASE_DIR`, exports, temp), loads `config/config.yaml`, and creates the
working folders. There are **no absolute machine paths** in the code.

---

## 4. Configuration

One file: **`config/config.yaml`** (read by `common/config_loader.py`), with
sections for `google`, `email`, `batch_planner`, and `pipeline`. Precedence: OS environment > `config.yaml` > project
path defaults. Deeply-embedded per-report settings (Drive folder IDs, recipient
lists, CC/BCC) remain inside the report scripts to preserve exact behaviour and
are documented in `config.yaml`'s reference block.

### Email — centralized sender

The Gmail **sender** identity is centralized in `credentials/email_config.py`
(`GMAIL_SENDER`/`GMAIL_APP_PASS` for the primary Operations account; the
`_DIGITAL` pair for the secondary account used by the Student Profile report).
Report recipient/CC/BCC lists stay in each report because they intentionally
differ per report; the reminder/warning behaviours are preserved as-is.

---

## 5. Credentials

All secrets live in **`credentials/`** (git-ignored). Required:
`service_account.json`, `email_config.py`, `wise_config.py` (Wise API), plus the
Zoho session files. See `credentials/README.md`.

---

## 6. Execution & failure handling

```bat
python scripts\run_data_collection.py     REM Layer 1 (4 jobs in parallel)
python scripts\run_reports_action.py      REM Layer 2 (parallel; standalone = no gating)
python scripts\run_all.py                 REM full pipeline (or run_all.bat)
```

`run_all.py`:
- runs all Layer-1 jobs in parallel and records each job's success/failure and record counts;
- logs failures clearly and names the failed job;
- runs Layer 2 in parallel, but **skips any report whose Layer-1 dependency failed** (per-script dependency map in `run_reports_action.py`; honours `pipeline.stop_on_failure`), so dependent reports never run on stale/incomplete data — while independent reports still run;
- never silently ignores a failed source refresh — it is surfaced in the log and the summary e-mail;
- e-mails a summary (per-script status, timings, record counts, logs attached) to `email.log_recipients` (default `info@intellibiinnovationstechnologies.in`);
- exits `0` only if every executed job succeeded and nothing was skipped.

---

## 7. Logging & completion e-mail

Two clearly separated channels:

**Detailed technical logs → the `logs/` folder (for developers).**
`common/logging_utils.py` writes every script's full output to both the console
and `logs/<name>.log`: start/end, API calls & errors, exceptions/tracebacks,
warnings, processing detail, record-level issues, **retry information**, and
execution duration. Nothing is removed from the logs — they remain the
troubleshooting source of truth. The Zoho collector's own file log
(`zoho_sign_sync.log`) is routed into `logs/` too. Cache → `cache/`, scratch →
`temp/`. No machine-specific locations.

**Meaningful business summary → the completion e-mail (for management/ops).**
`scripts/run_all.py` sends ONE concise e-mail built by
`common/exec_summary.py` + `common/common_utils.build_summary_html`. It does
**not** contain debug/technical noise. It shows, top to bottom:

- **Pipeline / Status / Started / Completed / Duration** and a one-line result
  (N succeeded • N failed • N skipped).
- **Execution Summary** — a section per script with *business* KPIs for that
  script (see the table below). Counts are for the **current run** (new /
  updated / skipped / sent / generated), parsed from each script's own end-of-run
  summary — never the whole existing dataset.
- **Action Required** — shown only when something failed or was skipped: the
  process, a short business-readable reason (e.g. "Google Sheets API rate limit
  (429) reached"), and retry status. Full tracebacks stay in the log file, which
  is **attached** to the e-mail.

`common/exec_summary.py` derives these KPIs by reading what each script already
prints — **no business logic was changed** to produce them.

### Status criteria
- **SUCCESS** — every executed script succeeded and nothing was skipped.
- **PARTIAL** — at least one script succeeded, but something failed or was
  skipped (e.g. a report skipped because its Layer-1 dependency failed).
- **FAILED** — every executed script failed.

### Per-script e-mail KPIs (Operations)

| Script | E-mail KPIs (current-run) |
|--------|---------------------------|
| Student Info & Payments | Students (full refresh), Payment records, Class-enrolment rows, Instructors, Students flagged removed |
| Sessions & Attendance | New Sessions, New Attendance rows, New Student/Teacher Feedback, Source sessions fetched, Suspended students found |
| Assignment Submissions (load) | Submissions extracted, New / Updated / Unchanged submissions, New assignments detected |
| Zoho Sign Status | Records fetched, Inserted (new), Updated, Unchanged/skipped, Failed |
| Attendance & Feedback Report | Reports generated, E-mailed, Sessions/Students covered, Absent |
| Assignment Submissions Report | New / Updated submissions, Reports generated, E-mailed (or "skipped — no changes") |
| Assignment Reminders | Students to notify, Reminders sent, Reminders failed, Pending assignments |
| Batch Planner | Students analysed, Batch-plan rows, Report uploaded |
| Student Profile Reports | Students processed, Failed, Master report e-mailed |
| Admission Formalities Report | Students processed, Inserted, Updated, Unchanged, Failed |
| Student Additional Note | Responses processed, Matched, No response (or "no changes to write") |

A zero-valued KPI is omitted to keep the e-mail clean (except a few where 0 is
itself meaningful, e.g. New Sessions or Reports generated).

---

## 8. Scheduling (Windows Task Scheduler)

Create a task → Action **Start a program** → `run_all.bat`, with **Start in** set
to the project folder (no absolute paths needed). "Run whether user is logged on
or not." The same task works after copying the folder to another machine.

---

## 9. Dependencies & deployment

Python 3.10+ and `requirements.txt`:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Then put secrets in `credentials/`, refresh the Batch Planner workbooks in
`data_inputs/`, review `config/config.yaml`, and run. See `README.md`.

---

## 10. Documentation policy

Every main `.py` script has a matching `.md` in `docs/`. **When a script changes,
update its `.md` in the same change** (Change History section at the bottom of
each doc).
