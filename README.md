# IntelliBI Operations Automation

A **clean, layered, portable** operations pipeline: refresh the source datasets
(student info, attendance, assignments, Zoho Sign), then generate the operations reports, reminders and warnings — with
centralized config, centralized logging, per-script dependency gating, and a
one-command launcher that e-mails a run summary.

Everything resolves from the project root at run time, so you can **copy this
whole folder to another machine, configure it once, and schedule it** without
editing any source-code paths.

```
Layer 1  ops_data_collection    4 refresh jobs (parallel)
   |
Layer 2  ops_reports_action     reports / reminders / actions (parallel where safe;
                                 a report whose Layer-1 dependency failed is skipped)
         + co-ordinator reports/pyBatchCoordinatorDailyAttendanceReport.py
                                 (the Coordinator's daily task list)
   ...
Evening  scripts/run_evening_reports.py
                                 co-ordinator reports/pyCoordinatorTaskPerformanceReport.py
                                 (Coordinator completion + timeliness, after the working day)
```

Full details: `docs/PROJECT_DOCUMENTATION.md` and the per-script docs under `docs/`.

## 1. Requirements
- **Python 3.10+**, the packages in `requirements.txt`
- A Google **service account** with access to the sheets/folders
- **Wise API** credentials, a **Zoho Sign** refresh token

## 2. Install
```bat
cd IntelliBI_Operations_Automation
python -m venv .venv
.venv\Scripts\activate            REM (source .venv/bin/activate on Linux/mac)
pip install -r requirements.txt
```

## 3. Configure once
1. **Secrets** → `credentials/` (see `credentials/README.md`):
   - `service_account.json`
   - `email_config.py` (copy `email_config.example.py`; add Gmail app passwords — primary + digital accounts)
   - `wise_config.py` (copy `wise_config.example.py`; add the Wise API key)
   - `zoho_credentials.json`
2. **Batch Planner inputs** → put the four workbooks in `data_inputs/` (see `data_inputs/README.md`).
3. **Settings** → edit `config/config.yaml` (service account, e-mail summary recipients, `stop_on_failure`, per-script timeout).
4. Everything else (cache/logs/output/temp) is created automatically.

## 4. Run
```bat
python scripts\run_data_collection.py     REM Layer 1 (4 jobs, parallel)
python scripts\run_reports_action.py      REM Layer 2 (parallel; standalone = no gating)
python scripts\run_all.py                 REM full pipeline  (or run_all.bat)
```

`run_all.py` runs Layer 1 in parallel, verifies each job, then runs Layer 2 in
parallel while **skipping any report whose Layer-1 dependency failed** (so no
report runs on stale/incomplete data). It logs each script to `logs/`, names any
failed job, and e-mails a summary (status, timings, record counts, logs
attached) to `email.log_recipients` (default
`info@intellibiinnovationstechnologies.in`).

Testing an action script? Use its dry-run switch first (e.g. the Assignment
reminder's `dry_run=True`) so no student e-mails go out.

## 5. Logging
Every script logs to console **and** `logs/<name>.log`; cache goes to `cache/`,
scratch to `temp/`, outputs to `output/`. No machine-specific locations. Log
level: `SALES_LOG_LEVEL` env or `config/logging_config.yaml`.

## 6. Schedule (Windows Task Scheduler)
From an **elevated PowerShell** in the project folder, once per machine:
```
powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1
```
It registers two tasks, both through `scripts\run_scheduled.py` (overlap lock +
once-per-day success gate; the later time of each pair is a retry-only window):

| Task | Times | Runs |
|---|---|---|
| IntelliBI Operations Automation (Morning batch) | 10:30, retry 11:30 | `scripts\run_all.py` — Layer 1 + Layer 2, incl. the Batch Coordinator daily task report |
| IntelliBI Operations Automation - Evening | 19:00, retry 20:00 | `scripts\run_evening_reports.py` — Coordinator Task Performance report |

`powershell -ExecutionPolicy Bypass -File scripts\status.ps1` shows both tasks.
The same setup works verbatim after copying the folder to another machine.

## 7. Deploy to a new machine
1. Copy the whole folder. 2. venv + `pip install -r requirements.txt`.
3. Put secrets in `credentials/` and workbooks in `data_inputs/`.
4. Review `config/config.yaml`. 5. `python scripts\run_all.py`, then run `scripts\setup_schedule.ps1` (section 6).

No source-code path edits are ever required.

## 8. Docs policy
Each `.py` has a matching `.md` in `docs/`. **Update the `.md` in the same change
whenever you change a script** (see each doc's Change History).
