"""
================================================================================
  IntelliBI Operations Automation
  COORDINATOR REPORTING PERIODS & DRIVE LAYOUT
  (co-ordinator reports / coordinator_periods.py)
  ------------------------------------------------------------------------------
  One shared definition, used by BOTH Coordinator reports, of:

    * the reporting periods (Daily / Weekly / Monthly / Manual) and the
      AUTO / manual-flag job selection — the same scheme as
      pyLeadFollowUpAnalysisReport.py (Mon–Sun weeks; Weekly = previous completed
      week on Monday; Monthly = 1st → last day, generated on the last day);

    * the Google Drive folder each report is saved into:

        <Coordinator root 1BEokUc7Np7iBVSMwIyMAgZUa0mrecT-h>
          ├── Daily Coordinator Reports
          │     └── Daily 01-Oct-2026
          │           ├── IntelliBI_Coordinator_Task_List_Report_Daily_…  (before 09-Oct-2026:
          │           │       IntelliBI_Batch_Coordinator_Daily_Attendance_Report_…)
          │           └── IntelliBI_Coordinator_Task_Performance_Report_Daily_…
          ├── Weekly Coordinator Reports
          │     └── Weekly 21-Sep-2026 to 27-Sep-2026
          ├── Monthly Coordinator Reports
          │     └── Monthly Sep-2026
          └── Manual Coordinator Reports
                └── Manual 21-Aug-2026 to 22-Sep-2026

      The period-folder names follow the IntelliBI convention already used by
      pyLeadFollowUpAnalysisReport.py (_period_folder_name). Because both reports
      derive the folder from (report type, start, end) through THIS module, the
      Coordinator Task List report and the Task Performance report for the same type
      and period always land in the SAME folder. Folders are resolved by name at
      run time (created when missing); files inside are versioned by the existing
      no-overwrite upload (pyCoordinatorTaskListReport.upload_report).

  Legacy layout: reports created before this change sit directly in
  <root>/YYYY-MM-DD/. They are still read by the performance report, and can be
  moved into the new layout with:
        python "co-ordinator reports/coordinator_periods.py" --migrate-legacy          (dry run)
        python "co-ordinator reports/coordinator_periods.py" --migrate-legacy --apply  (move)
================================================================================
"""
from __future__ import annotations

import re
import calendar
from datetime import date, datetime, timedelta

KINDS = ("Daily", "Weekly", "Monthly", "Manual")

# Report-type folders directly under the Coordinator root folder.
KIND_FOLDERS = {
    "Daily":   "Daily Coordinator Reports",
    "Weekly":  "Weekly Coordinator Reports",
    "Monthly": "Monthly Coordinator Reports",
    "Manual":  "Manual Coordinator Reports",
}

_D = "%d-%b-%Y"


# =============================================================================
#  PERIODS
# =============================================================================
def week_bounds(d: date):
    """Monday → Sunday of the week containing d (same week definition as
    pyLeadFollowUpAnalysisReport.week_bounds and the Coordinator report)."""
    mon = d - timedelta(days=d.weekday())
    return mon, mon + timedelta(days=6)


def month_bounds(year: int, month: int):
    """1st → last calendar day (handles 28/29/30/31-day months and leap years)."""
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def is_month_end(d: date) -> bool:
    return d.day == calendar.monthrange(d.year, d.month)[1]


def period_label(kind: str, start: date, end: date) -> str:
    if kind == "Daily":
        return start.strftime(_D)
    if kind == "Monthly":
        return start.strftime("%B %Y")
    return f"{start.strftime(_D)} to {end.strftime(_D)}"


def make_job(kind: str, start: date, end: date) -> dict:
    return {"kind": kind, "start": start, "end": end,
            "label": period_label(kind, start, end)}


def _parse_date(value, name):
    try:
        return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()
    except (TypeError, ValueError):
        raise ValueError(f"{name} must be 'YYYY-MM-DD' (got {value!r})")


def plan_jobs(today: date, *, auto: bool, daily: bool = False, weekly: bool = False,
              monthly: bool = False, manual: bool = False, daily_date=None,
              weekly_reference_date=None, monthly_month=None, monthly_year=None,
              manual_start=None, manual_end=None):
    """Which reports to build, for which period.

    AUTO (auto=True; the individual flags are ignored, as in
    pyLeadFollowUpAnalysisReport.py):
        Daily   — every run, for `today`
        Weekly  — on Monday, for the previous completed Mon–Sun week
        Monthly — on the last day of the month, for that whole month
    Flag mode (auto=False): each GENERATE_* flag independently; the optional
    dates pin a specific day / week / month; Manual needs both dates with
    start <= end.

    Returns (jobs, errors): errors are configuration problems (e.g. an invalid
    Manual range) — the job is skipped and the caller should report a failure."""
    jobs, errors = [], []
    if auto:
        jobs.append(make_job("Daily", today, today))
        if today.weekday() == 0:                                  # Monday
            jobs.append(make_job("Weekly", *week_bounds(today - timedelta(days=7))))
        if is_month_end(today):
            jobs.append(make_job("Monthly", *month_bounds(today.year, today.month)))
        return jobs, errors

    try:
        if daily:
            d = _parse_date(daily_date, "DAILY_DATE") if daily_date else today
            jobs.append(make_job("Daily", d, d))
        if weekly:
            ref = (_parse_date(weekly_reference_date, "WEEKLY_REFERENCE_DATE")
                   if weekly_reference_date else today)
            jobs.append(make_job("Weekly", *week_bounds(ref)))
        if monthly:
            y, m = int(monthly_year or today.year), int(monthly_month or today.month)
            if not 1 <= m <= 12:
                raise ValueError(f"MONTHLY_MONTH must be 1-12 (got {monthly_month!r})")
            jobs.append(make_job("Monthly", *month_bounds(y, m)))
    except ValueError as exc:
        errors.append(str(exc))

    if manual:
        if not (manual_start and manual_end):
            errors.append("GENERATE_MANUAL is on but MANUAL_START_DATE / MANUAL_END_DATE "
                          "are not both set — Manual report skipped.")
        else:
            try:
                a = _parse_date(manual_start, "MANUAL_START_DATE")
                b = _parse_date(manual_end, "MANUAL_END_DATE")
                if a > b:
                    errors.append(f"MANUAL_START_DATE ({a}) is after MANUAL_END_DATE ({b}) "
                                  f"— Manual report skipped.")
                else:
                    jobs.append(make_job("Manual", a, b))
            except ValueError as exc:
                errors.append(str(exc) + " — Manual report skipped.")
    return jobs, errors


# =============================================================================
#  DRIVE LAYOUT
# =============================================================================
def period_folder_name(kind: str, start: date, end: date) -> str:
    """Reporting-period folder name — the pyLeadFollowUpAnalysisReport convention:
        Daily 01-Oct-2026 | Weekly 21-Sep-2026 to 27-Sep-2026 |
        Monthly Sep-2026  | Manual 21-Aug-2026 to 22-Sep-2026"""
    if kind == "Daily":
        return f"Daily {start.strftime(_D)}"
    if kind == "Monthly":
        return f"Monthly {start.strftime('%b-%Y')}"
    if kind in ("Weekly", "Manual"):
        return f"{kind} {start.strftime(_D)} to {end.strftime(_D)}"
    raise ValueError(f"unknown report type {kind!r}")


def folder_path(kind: str, start: date, end: date) -> tuple:
    """(report-type folder, reporting-period folder) under the Coordinator root."""
    return (KIND_FOLDERS[kind], period_folder_name(kind, start, end))


_DAILY_FOLDER = re.compile(r"^Daily (\d{2}-[A-Za-z]{3}-\d{4})$")
_LEGACY_DAY_FOLDER = re.compile(r"^(\d{4}-\d{2}-\d{2})$")


def daily_folder_date(name: str):
    """Report date of a Daily period folder ('Daily 01-Oct-2026') or of a legacy
    date folder ('2026-10-01'); None for any other folder."""
    m = _DAILY_FOLDER.match(name or "")
    if m:
        try:
            return datetime.strptime(m.group(1), _D).date()
        except ValueError:
            return None
    m = _LEGACY_DAY_FOLDER.match(name or "")
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def resolve_folder(drive, root_id: str, names, find_or_create, cache: dict | None = None) -> str:
    """Folder id of root/names[0]/names[1]/… (created when missing). Re-uses the
    caller's find-or-create helper so lookup/creation stay identical to the
    existing Coordinator report."""
    parent = root_id
    key = (root_id,)
    for n in names:
        key = key + (n,)
        if cache is not None and key in cache:
            parent = cache[key]
            continue
        parent = find_or_create(drive, parent, n)
        if cache is not None:
            cache[key] = parent
    return parent


# =============================================================================
#  ONE-OFF MIGRATION OF THE LEGACY DATE FOLDERS (optional)
# =============================================================================
_VERSION_SUFFIX = re.compile(r"\s*-\s*Version\s*\d+\s*$", re.IGNORECASE)


def classify_legacy_file(name: str, folder_day: date | None):
    """(kind, start, end) a legacy Coordinator file belongs to, or None to leave
    it where it is (e.g. the older 'Coordinator Attendance Tasks' sheets)."""
    base = _VERSION_SUFFIX.sub("", name or "")
    if (base.startswith("IntelliBI_Batch_Coordinator_Daily_Attendance_Report")
            or base.startswith("IntelliBI_Coordinator_Task_List_Report_Daily")) and folder_day:
        return ("Daily", folder_day, folder_day)
    m = (re.match(r"^IntelliBI_Batch_Coordinator_(Weekly|Monthly|Manual)_Follow_Ups_(.+)$", base)
         or re.match(r"^IntelliBI_Coordinator_Task_List_Report_(Weekly|Monthly|Manual)_(.+)$", base))
    if m:
        return _parse_period_label(m.group(1), m.group(2))
    m = re.match(r"^IntelliBI_Coordinator_Task_Performance_Report_(Daily|Weekly|Monthly|Manual)_(.+)$", base)
    if m:
        return _parse_period_label(m.group(1), m.group(2))
    return None


def _parse_period_label(kind, label):
    """Parse the period from a legacy file-name label (underscored forms of
    '21 Sep – 27 Sep 2026', 'September 2026', '01-Sep-2026 to 15-Sep-2026',
    '30-Sep-2026')."""
    toks = [t for t in re.split(r"[_\s]+", label.replace("–", " "))
            if t and t.lower() not in ("-", "to")]
    txt = " ".join(toks)
    try:
        if kind == "Monthly":
            d = datetime.strptime(txt, "%B %Y").date()
            return ("Monthly",) + month_bounds(d.year, d.month)
        dates = re.findall(r"(\d{1,2})[-\s]([A-Za-z]{3})[-\s]?(\d{4})?", txt)
        if not dates:
            return None
        year = next((y for _d, _m, y in reversed(dates) if y), None)
        parsed = [datetime.strptime(f"{d} {m} {y or year}", "%d %b %Y").date() for d, m, y in dates]
        if kind == "Daily":
            return ("Daily", parsed[0], parsed[0])
        if len(parsed) >= 2:
            a, b = parsed[0], parsed[-1]
            if a > b and not dates[0][2]:            # '29 Dec – 04 Jan 2026' spans a year end
                a = a.replace(year=a.year - 1)
            return (kind, a, b)
    except (ValueError, TypeError):
        return None
    return None


def migrate_legacy(apply: bool = False):
    """Move Coordinator reports from <root>/YYYY-MM-DD/ (and the interim
    'Coordinator Performance' folder) into the Daily / Weekly / Monthly / Manual
    hierarchy. Files are MOVED (same file id → existing links keep working),
    never copied or deleted. Dry run unless apply=True."""
    import pyCoordinatorTaskListReport as BC
    drive = BC._drive_client()
    folders = BC._list_children(drive, BC.PARENT_FOLDER_ID, folders_only=True)
    sources = [f for f in folders
               if _LEGACY_DAY_FOLDER.match(f["name"]) or f["name"] == "Coordinator Performance"]
    cache, moved, left = {}, 0, []
    for src in sorted(sources, key=lambda f: f["name"]):
        day = daily_folder_date(src["name"])
        for f in BC._list_children(drive, src["id"]):
            target = classify_legacy_file(f["name"], day)
            if not target:
                left.append(f"{src['name']}/{f['name']}")
                continue
            names = folder_path(*target)
            print(f"{'MOVE' if apply else 'would move'}  {src['name']}/{f['name']}  →  {'/'.join(names)}")
            if apply:
                dest = resolve_folder(drive, BC.PARENT_FOLDER_ID, names, BC._find_or_create_folder, cache)
                drive.files().update(fileId=f["id"], addParents=dest, removeParents=src["id"],
                                     fields="id", supportsAllDrives=True).execute()
            moved += 1
    print(f"\n{moved} file(s) {'moved' if apply else 'to move'}; {len(left)} left in place:")
    for x in left:
        print("   ", x)
    if not apply:
        print("\nDry run only — re-run with --apply to move them.")


if __name__ == "__main__":
    import os
    import sys
    import argparse
    here = os.path.dirname(os.path.abspath(__file__))
    for p in (here, os.path.join(os.path.dirname(here), "common"),
              os.path.join(os.path.dirname(here), "ops_reports_action")):
        if p not in sys.path:
            sys.path.insert(0, p)
    ap = argparse.ArgumentParser(description="Coordinator report folders")
    ap.add_argument("--migrate-legacy", action="store_true")
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    if a.migrate_legacy:
        migrate_legacy(apply=a.apply)
    else:
        ap.print_help()
