"""
================================================================================
  IntelliBI Operations Automation
  ASSIGNMENT SUBMISSION PERFORMANCE  +  ASSIGNMENT NON-SUBMISSION REPORTS
  (ops_reports_action / pyAssignmentSubmissionPerformanceReport.py)
  ------------------------------------------------------------------------------
  Builds, for every planned reporting period, TWO native Google Sheets and assignment-wise PDFs from ONE
  shared dataset, so they always reconcile:

    A. Assignment Submission Performance Report  (management view)
         Summary          — KPIs, batch-wise performance
         Assignment Performance — one row per assignment (expected / submitted /
                            on time / late / not submitted / submission % / marks)
         Learner Submission Detail — per assignment: a 📝 banner, then its learners
       → Drive: "Assignment Submission Report" folder 190_1NRLGgwoLpqiiW0fPGZ85ux2QknZl

    B. Assignment Non-Submission Report  (follow-up list)
         Non-Submitters   — every learner who had not submitted when the
                            deadline passed, with contact details
         By Learner       — one row per learner with all their pending assignments
       → Drive: "Assignment Not Submitted Students List" 1PmpsLw-4L4woeznxYC_8bQnFQzHQYA2N

    C. Assignment-wise Non-Submission PDFs  (one .pdf per eligible assignment)
         A "Class | Duration | Assignment | Start | Deadline" banner, the
         assignment's figures and its learners who had not submitted — exactly its
         Not Submitted rows in A and B (same dataset). The reminder's approach:
         one file per assignment, Assignment_<Class>_<Assignment>_<Kind>_<label>.pdf,
         its PDF look.
       → Drive: B's period folder, beside the Non-Submission Sheet

    Both under:  <root>/<Daily|Weekly|Monthly|Manual> Assignment Report/<period>/

  SOURCE & RULES (reused, not redefined)
    * Data: IntelliBIAssessmentSubmission ▸ Submissions — one row per enrolled
      learner per assignment, written by ops_data_collection/pyAssignmentSubmissions.py
      (enrolled = class roster ∪ everyone who submitted). Read with
      pyAssignmentSubmissionEmailReminder.read_sheet_df.
    * Expected learners of an assignment = its rows; Submitted = status other than
      "Not Submitted"; Not submitted = status "Not Submitted" — the same rule as
      the retired pyAssignmentSubmissionsReport (counts; now in archive/) and
      pyAssignmentSubmissionEmailReminder (who is not submitted). One row per learner per assignment (duplicates
      removed, a submitted row wins).
    * Learner active status (information only): the reminder's Students-tab check
      (pyAssignmentSubmissionEmailReminder.load_active_status_map / is_student_active).

  ELIGIBILITY (both reports)
    An assignment is reported only when
      1. its deadline DATE lies inside the reporting period, AND
      2. its deadline DATE & TIME (IST, "DD/MM/YYYY HH:MM:SS IST") has already
         passed at run time.
    Assignments without a readable deadline are never reported (listed as
    excluded). A deadline with no time is treated as end of that day (23:59:59).

  PERIODS (same flag framework as pyLeadFollowUpAnalysisReport.py, shared helpers
  from co-ordinator reports/coordinator_periods.py)
    GENERATE_AUTO = True:
      Daily   — every run, for SYSTEM DATE − 1 (a 11:45 PM deadline is already past)
      Weekly  — on Monday, for the previous completed Monday–Sunday week
      Monthly — on the 1st, for the whole previous calendar month
    GENERATE_AUTO = False: GENERATE_DAILY / WEEKLY / MONTHLY / MANUAL independently;
      DAILY_DATE (None = system date − 1), WEEKLY_REFERENCE_DATE, MONTHLY_MONTH/YEAR,
      MANUAL_START_DATE..MANUAL_END_DATE (end must be ≤ system date − 1).

  DELIVERY
    Google Drive only (built in memory, uploaded as native Google Sheets, never
    a re-run overwrites that period's report in place — OVERWRITE_EXISTING); one e-mail per period from
    info@ with both links (coordinator_email.send — same sender / recipients /
    style as pyCoordinatorTaskPerformanceReport.py), Starred (★) in the info@
    mailbox after a successful send (common/gmail_star.py). Scheduled in the Operations
    Morning batch (scripts/run_reports_action.py) after the Assignment
    Submissions refresh.

  Run:   python ops_reports_action/pyAssignmentSubmissionPerformanceReport.py
         … --dry-run   (read live data, build both reports in memory, print the
                        summary; nothing uploaded, nothing e-mailed)
================================================================================
"""

# --- IntelliBI Operations Automation portability bootstrap ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402,F401  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR  # noqa: E402
# --- end bootstrap ---

import io
import re
import logging
from collections import OrderedDict, defaultdict
from datetime import date, datetime, time, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT = os.path.dirname(_HERE)
for _p in (_HERE, os.path.join(_PROJECT, "co-ordinator reports")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import openpyxl                                      # noqa: E402
from openpyxl.utils import get_column_letter         # noqa: E402

import pyAssignmentSubmissionEmailReminder as ASG    # noqa: E402  data reader, learner active status
import pyAttendaceFeedbackReport as AR               # noqa: E402  report look & colour code
import coordinator_periods as CP                     # noqa: E402  periods + folder names
import coordinator_email as CE                       # noqa: E402  report e-mail

log = logging.getLogger("AssignmentSubmissionPerformance")

# =============================================================================
#  REPORT GENERATION CONTROL
# =============================================================================
GENERATE_AUTO    = True

GENERATE_DAILY   = True
GENERATE_WEEKLY  = True
GENERATE_MONTHLY = True
GENERATE_MANUAL  = False

DAILY_DATE            = None      # "YYYY-MM-DD"  (None = system date − 1)
WEEKLY_REFERENCE_DATE = None      # any date in the wanted Mon–Sun week (None = this week)
MONTHLY_MONTH         = None      # 1-12          (None = current month)
MONTHLY_YEAR          = None      # e.g. 2026     (None = current year)

MANUAL_START_DATE = "2026-08-21"  # "YYYY-MM-DD"
MANUAL_END_DATE   = "2026-09-22"  # "YYYY-MM-DD"  (must be ≤ system date − 1)

# =============================================================================
#  DELIVERY
# =============================================================================
UPLOAD_TO_DRIVE = True
SEND_EMAIL      = True
# Star (★) each report e-mail in the info@ Gmail mailbox once it is sent
# (common/gmail_star.py through coordinator_email.send — the same as the
# Coordinator reports): only after a successful send, only the sender's own
# mailbox, best-effort — never affects sending or the run.
STAR_EMAIL_IN_GMAIL = True
EMAIL_SENDER     = "info@intellibiinnovationstechnologies.in"
EMAIL_RECIPIENTS = ["info@intellibiinnovationstechnologies.in",
                    "intellibihropsb2ch@gmail.com"]
# A report already in Drive for the same period is OVERWRITTEN in place (same
# file, same link) — every run rebuilds and re-sends. False = keep the old file
# and save the new one as "<name> - Version N".
OVERWRITE_EXISTING = True
# ONE PDF PER ASSIGNMENT with the learners who had not submitted it when the
# deadline passed (same dataset — together they list exactly the Performance
# Report's Not Submitted), uploaded into the period folder of the Non-Submission
# root, beside the Non-Submission Sheet, and linked (folder) in the e-mail. A PDF
# problem never stops the Sheets or the e-mail.
GENERATE_NON_SUBMISSION_PDF = True
# True = a PDF for EVERY eligible assignment (5 eligible → 5 PDFs; one where all
# learners submitted says so). False = only assignments with non-submitters.
PDF_FOR_FULLY_SUBMITTED = True

IMPERSONATE_USER  = "info@intellibiinnovationstechnologies.in"
SERVICE_ACCOUNT_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json")
SUBMISSION_SHEET_ID  = ASG.SUBMISSION_SHEET_ID      # IntelliBIAssessmentSubmission
SUBMISSIONS_TAB      = "Submissions"

PERFORMANCE_ROOT_FOLDER_ID    = "190_1NRLGgwoLpqiiW0fPGZ85ux2QknZl"   # Assignment Submission Report
NON_SUBMISSION_ROOT_FOLDER_ID = "1PmpsLw-4L4woeznxYC_8bQnFQzHQYA2N"   # Assignment Not Submitted Students List
KIND_FOLDERS = {"Daily": "Daily Assignment Report", "Weekly": "Weekly Assignment Report",
                "Monthly": "Monthly Assignment Report", "Manual": "Manual Assignment Report"}
PERFORMANCE_BASENAME    = "IntelliBI_Assignment_Submission_Performance_Report"
NON_SUBMISSION_BASENAME = "IntelliBI_Assignment_Non_Submission_Report"
# earlier PDF layouts, moved to the Drive trash when their period is re-run: the
# single consolidated PDF, and the "Assignment-wise PDFs" subfolder
LEGACY_PDF_BASENAME = "IntelliBI_Assignment_Non_Submission_Learners"
LEGACY_PDF_SUBFOLDER = "Assignment-wise PDFs"

DEADLINE_DEFAULT_TIME = time(23, 59, 59)    # deadline recorded without a time
# Colour code of pyAttendaceFeedbackReport (Att % → Submission %): the % cell is
# dark green ≥ 95, light green ≥ 75, amber ≥ 60, red below (AR._att_bg); KPI and
# e-mail figure green at ≥ 75, red below.
RATE_GOOD, RATE_WATCH = 75.0, 60.0
NOT_SUBMITTED = "Not Submitted"

try:
    from zoneinfo import ZoneInfo
    _IST = ZoneInfo("Asia/Kolkata")
except Exception:                                            # pragma: no cover
    from datetime import timezone
    _IST = timezone(timedelta(hours=5, minutes=30))


def now_ist() -> datetime:
    """Current IST time, naive (the sheet's deadlines are naive IST)."""
    return datetime.now(_IST).replace(tzinfo=None)


# =============================================================================
#  REPORTING PERIODS
# =============================================================================
def plan_jobs(today: date):
    """Which periods to report. Returns (jobs, errors); a job is
    {"kind", "start", "end", "label"} (coordinator_periods.make_job)."""
    yday = today - timedelta(days=1)
    if GENERATE_AUTO:
        jobs = [CP.make_job("Daily", yday, yday)]
        if today.weekday() == 0:                                   # Monday
            jobs.append(CP.make_job("Weekly", *CP.week_bounds(today - timedelta(days=7))))
        if today.day == 1:                                         # 1st of the month
            jobs.append(CP.make_job("Monthly", *CP.month_bounds(yday.year, yday.month)))
        return jobs, []

    errors = []
    if GENERATE_DAILY and DAILY_DATE:
        try:
            d = CP._parse_date(DAILY_DATE, "DAILY_DATE")
            if d > today:
                errors.append(f"DAILY_DATE ({d:%d-%b-%Y}) is in the future — Daily report skipped.")
        except ValueError as exc:
            errors.append(str(exc) + " — Daily report skipped.")
    daily_ok = GENERATE_DAILY and not errors
    jobs, cp_errors = CP.plan_jobs(
        today, auto=False, daily=daily_ok, weekly=GENERATE_WEEKLY, monthly=GENERATE_MONTHLY,
        manual=False, daily_date=DAILY_DATE or yday.isoformat(),
        weekly_reference_date=WEEKLY_REFERENCE_DATE, monthly_month=MONTHLY_MONTH,
        monthly_year=MONTHLY_YEAR)
    errors += cp_errors
    if GENERATE_MANUAL:
        job, err = manual_job(today)
        if err:
            errors.append(err)
        else:
            jobs.append(job)
    return jobs, errors


def manual_job(today: date):
    """(job, None) or (None, error) — Start and End given, Start ≤ End, and End
    no later than system date − 1 (a later end would report an incomplete day)."""
    yday = today - timedelta(days=1)
    if not MANUAL_START_DATE or not MANUAL_END_DATE:
        return None, ("GENERATE_MANUAL is on but MANUAL_START_DATE and MANUAL_END_DATE are not "
                      "both set — Manual report not generated.")
    try:
        a = CP._parse_date(MANUAL_START_DATE, "MANUAL_START_DATE")
        b = CP._parse_date(MANUAL_END_DATE, "MANUAL_END_DATE")
    except ValueError as exc:
        return None, f"{exc} — Manual report not generated."
    if a > b:
        return None, (f"MANUAL_START_DATE ({a:%d-%b-%Y}) is after MANUAL_END_DATE ({b:%d-%b-%Y}) "
                      f"— Manual report not generated.")
    if b > yday:
        return None, (f"MANUAL_END_DATE ({b:%d-%b-%Y}) must be on or before {yday:%d-%b-%Y} "
                      f"(system date − 1); "
                      + ("today" if b == today else "a future date")
                      + " would give an incomplete report — Manual report not generated.")
    return CP.make_job("Manual", a, b), None


# =============================================================================
#  DATES
# =============================================================================
_DT_FORMATS = ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M")


def parse_ist_datetime(val, default_time=DEADLINE_DEFAULT_TIME):
    """'DD/MM/YYYY HH:MM:SS IST' (utils.to_ist_dmy, the collector's format) →
    naive IST datetime. A value with a date but no time gets `default_time`
    (date part parsed by _parse_ist_date).
    None when unreadable."""
    if val is None:
        return None
    s = str(val).replace(" IST", "").strip()
    if s in ("", "nan", "None", "NaT"):
        return None
    for fmt in _DT_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    d = _parse_ist_date(val)
    if d is None or (isinstance(d, float)) or str(d) == "NaT":
        return None
    try:
        return datetime.combine(d, default_time)
    except TypeError:
        return None


def _fmt_dt(dt):
    return dt.strftime("%d-%b-%Y %I:%M %p") if dt else ""


def _fmt_d(d):
    return d.strftime("%d-%b-%Y") if d else ""


# =============================================================================
#  SHARED DATASET  (the ONE source of both reports)
# =============================================================================
def is_not_submitted(status) -> bool:
    """The reminder's test (pyAssignmentSubmissionEmailReminder.find_pending_reminders)."""
    return str(status or "").strip().lower() == NOT_SUBMITTED.lower()


def _parse_ist_date(val):
    """'DD/MM/YYYY HH:MM:SS IST' (or similar) → date, None on failure. Kept
    verbatim from the retired pyAssignmentSubmissionsReport (archive/), so date
    handling is unchanged (it also accepts MM/DD/YYYY, unlike the reminder's)."""
    import pandas as pd
    if not val or str(val).strip() in ("", "nan", "None"):
        return None
    s = str(val).replace(" IST", "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(s, errors="coerce").date()
    except Exception:
        return None


def _safe_float(val):
    """Kept verbatim from the retired pyAssignmentSubmissionsReport (archive/)."""
    import pandas as pd
    try:
        f = float(str(val).strip())
        return f if not pd.isna(f) else None
    except (ValueError, TypeError):
        return None


def _num(v):
    f = _safe_float(v)
    return f


def _learner_key(r):
    sid = str(r.get("student_id", "") or "").strip()
    if sid:
        return "id:" + sid
    em = str(r.get("student_email", "") or "").strip().lower()
    if em:
        return "em:" + em
    return "nm:" + str(r.get("student_name", "") or "").strip().lower()


def _active_label(email, status_map):
    if not status_map:
        return "Unknown"
    ok, reason = ASG.is_student_active({"student_email": email}, status_map)
    if ok:
        return "Yes"
    if "not found" in reason:
        return "Not found"
    if "= N" in reason:
        return "No"
    return "Unknown"


def build_dataset(subs_df, start: date, end: date, now: datetime, status_map=None) -> dict:
    """Eligible assignments of [start, end] and their learner rows.

    Returns {"assignments": [..], "rows": [..], "excluded": {...}, "start", "end",
    "as_of"}: each assignment carries its rows and counts; each row is one learner
    for one assignment with status "Submitted" / "Not Submitted"."""
    excluded = {"not_yet_due": OrderedDict(), "no_deadline": OrderedDict(), "duplicates": 0}
    by_asg = OrderedDict()
    if subs_df is not None and not subs_df.empty:
        for rec in subs_df.to_dict("records"):
            aid = str(rec.get("assessment_id", "") or "").strip()
            if not aid:
                continue
            cid = str(rec.get("class_id", "") or "").strip()
            key = (aid, cid)
            dl = parse_ist_datetime(rec.get("submission_deadline"))
            title = str(rec.get("assessment_title", "") or "").strip() or aid
            if dl is None:
                excluded["no_deadline"].setdefault(key, title)
                continue
            if not (start <= dl.date() <= end):
                continue
            if dl > now:
                excluded["not_yet_due"].setdefault(key, (title, dl))
                continue
            a = by_asg.get(key)
            if a is None:
                a = by_asg[key] = {
                    "assessment_id": aid, "class_id": cid, "title": title,
                    "technology": str(rec.get("class_name", "") or "").strip() or "—",
                    "batch": str(rec.get("class_subject", "") or "").strip() or "—",
                    "start": parse_ist_datetime(rec.get("submission_start_date"), time(0, 0)),
                    "deadline": dl, "max_marks": str(rec.get("maximum_marks", "") or "").strip(),
                    "learners": OrderedDict()}
            lk = _learner_key(rec)
            prev = a["learners"].get(lk)
            if prev is not None:
                excluded["duplicates"] += 1
                if not (is_not_submitted(prev["status_raw"]) and not is_not_submitted(
                        rec.get("submission_status"))):
                    continue                            # keep the first, unless this one is submitted
            sub_at = parse_ist_datetime(rec.get("submitted_at"), time(0, 0))
            not_sub = is_not_submitted(rec.get("submission_status"))
            email = str(rec.get("student_email", "") or "").strip()
            a["learners"][lk] = {
                "assessment_id": aid, "class_id": cid,
                "technology": a["technology"], "batch": a["batch"], "assignment": title,
                "deadline": dl, "learner": str(rec.get("student_name", "") or "").strip() or "—",
                "email": email, "phone": str(rec.get("student_phone", "") or "").strip(),
                "student_id": str(rec.get("student_id", "") or "").strip(),
                "status_raw": rec.get("submission_status"),
                "status": NOT_SUBMITTED if not_sub else "Submitted",
                "submitted_at": None if not_sub else sub_at,
                "timing": ("—" if not_sub else "Not recorded" if sub_at is None
                           else "On time" if sub_at <= dl else "Late"),
                "marks": None if not_sub else _num(rec.get("evaluation_marks")),
                "feedback": "" if not_sub else str(rec.get("evaluation_feedback", "") or "").strip(),
                "active": _active_label(email, status_map),
                "days_overdue": max(0, (now.date() - dl.date()).days),
            }

    assignments, rows = [], []
    for a in by_asg.values():
        lr = list(a.pop("learners").values())
        lr.sort(key=lambda r: (r["status"] == "Submitted", r["learner"].lower()))
        a["rows"] = lr
        a.update(_counts(lr))
        assignments.append(a)
        rows.extend(lr)
    assignments.sort(key=lambda a: (a["technology"].lower(), a["batch"].lower(), a["deadline"],
                                    a["title"].lower()))
    rows.sort(key=lambda r: (r["technology"].lower(), r["batch"].lower(), r["deadline"],
                             r["assignment"].lower(), r["status"] == "Submitted", r["learner"].lower()))
    return {"assignments": assignments, "rows": rows, "excluded": excluded,
            "start": start, "end": end, "as_of": now}


def _counts(rows):
    exp = len(rows)
    sub = sum(1 for r in rows if r["status"] == "Submitted")
    marks = [r["marks"] for r in rows if r["marks"] is not None]
    return {"expected": exp, "submitted": sub, "not_submitted": exp - sub,
            "on_time": sum(1 for r in rows if r["timing"] == "On time"),
            "late": sum(1 for r in rows if r["timing"] == "Late"),
            "rate": round(sub / exp * 100, 1) if exp else None,
            "avg_marks": round(sum(marks) / len(marks), 1) if marks else None}


def summarise(ds) -> dict:
    """Overall KPIs (the same counts both reports show)."""
    s = _counts(ds["rows"])
    s["assignments"] = len(ds["assignments"])
    s["technologies"] = len({a["technology"] for a in ds["assignments"]})
    s["batches"] = len({(a["technology"], a["batch"]) for a in ds["assignments"]})
    pend = [r for r in ds["rows"] if r["status"] == NOT_SUBMITTED]
    s["learners_pending"] = len({_learner_key({"student_id": r["student_id"],
                                               "student_email": r["email"],
                                               "student_name": r["learner"]}) for r in pend})
    s["learners_expected"] = len({_learner_key({"student_id": r["student_id"],
                                                "student_email": r["email"],
                                                "student_name": r["learner"]}) for r in ds["rows"]})
    s["not_yet_due"] = len(ds["excluded"]["not_yet_due"])
    s["no_deadline"] = len(ds["excluded"]["no_deadline"])
    s["duplicates"] = ds["excluded"]["duplicates"]
    return s


def group_rows(ds, keyf):
    """[(key, counts, n_assignments)] grouped by keyf(assignment)."""
    groups = OrderedDict()
    for a in ds["assignments"]:
        groups.setdefault(keyf(a), []).append(a)
    out = []
    for k, alist in groups.items():
        rows = [r for a in alist for r in a["rows"]]
        out.append((k, _counts(rows), len(alist)))
    return out


# =============================================================================
#  WORKBOOK LAYOUT  (pyAttendaceFeedbackReport's helpers, palette and colour
#  code: navy headers, blue section banners, zebra rows, semantic % / status
#  cells, light-blue TOTAL row, KPI cards with icons)
# =============================================================================
PCT_FMT = '0.0"%"'
DT_FMT = "dd-mmm-yyyy hh:mm AM/PM"
D_FMT = "dd-mmm-yyyy"


def _zebra(i):
    """Alternating row tint (pyAttendaceFeedbackReport rows)."""
    return AR.C_ROW_ALT if i % 2 == 0 else AR.C_WHITE


def _pct_cell(ws, row, col, rate):
    """Submission % cell in the Attendance report's Att % colour code
    (AR._att_bg fill; dark-green text at ≥ 75 %, dark-red below)."""
    c = ws.cell(row=row, column=col)
    c.value = rate
    if rate is None:
        c.value = "—"
        c.fill = AR._fill(AR.C_GREY_LITE)
        c.font = AR._font(size=10, color="757575")
    else:
        c.fill = AR._fill(AR._att_bg(rate, "Present"))
        c.font = AR._font(bold=True, size=10,
                          color=AR.C_GREEN_DARK if rate >= RATE_GOOD else AR.C_RED_DARK)
        c.number_format = PCT_FMT
    c.alignment = AR._align("center", "center")
    c.border = AR._border()
    return c


SUB_DT_FMT = "mm/dd/yyyy hh:mm AM/PM"
# Learner rows on Learner Submission Detail (subtle IntelliBI / Attendance shades)
ROW_SUBMITTED, ROW_NOT_SUBMITTED, ROW_INACTIVE = AR.C_GREEN_PALE, AR.C_RED_PALE, AR.C_AMBER


def _assignment_banner(a):
    """'Class: <technology> | Duration: <class duration> | 📝 <Assignment> |
    Start: MM/DD/YYYY | Deadline: MM/DD/YYYY'.
    Duration is the class's `class_subject` (e.g. "16-Sep-2026 To Current Date"),
    the value already held as the assignment's batch and shown as "Class
    Duration" / "Duration" by pyAssignmentSubmissionEmailReminder and
    pyCoordinatorTaskListReport ("—" when blank)."""
    start = a["start"].strftime("%m/%d/%Y") if a["start"] else "—"
    return (f"  Class: {a['technology']}  |  Duration: {a['batch']}  |  📝  {a['title']}  |  "
            f"Start: {start}  |  Deadline: {a['deadline']:%m/%d/%Y}")


def _learner_row_style(x):
    """(row fill, font factory) for one learner row — presentation only.
    Priority: Active Learner = No → amber + strikethrough (whole row);
    otherwise Submitted → green, Not Submitted → red."""
    if x["active"] == "No":
        return ROW_INACTIVE, lambda bold: AR._font(bold=bold, size=10, color="5D4037", strike=True)
    if x["status"] == "Submitted":
        return ROW_SUBMITTED, lambda bold: AR._font(bold=bold, size=10, color=AR.C_GREEN_DARK
                                                    if bold else "000000")
    return ROW_NOT_SUBMITTED, lambda bold: AR._font(bold=bold, size=10, color=AR.C_RED_DARK
                                                    if bold else "000000")


def _marks_value(m):
    if m is None:
        return ""
    return int(m) if float(m).is_integer() else m


def _title(ws, n, ds, text):
    """Row-1 header of every tab: the existing header text with the report's
    generation date & time (IST) appended after a '|' separator."""
    text = f"{text}  |  Generated On: {ds['as_of'].strftime('%d-%b-%Y %I:%M %p')}"
    AR.style_title_row(ws, 1, 1, n, text)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")


def _subtitle(ws, row, n, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.font = AR._font(size=9, color="455A64", italic=True)
    c.alignment = AR._align("left", "center", wrap=True)
    ws.row_dimensions[row].height = 30
    return row + 1


def _empty_band(ws, row, n, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.font = AR._font(bold=True, size=11, color=AR.C_GREEN_DARK)
    c.fill = AR._fill(AR.C_GREEN_PALE)
    c.alignment = AR._align("center", "center", wrap=True)
    c.border = AR._border()
    ws.row_dimensions[row].height = 26
    return row + 1


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _period_text(job, ds):
    return (f"{job['kind']}  ·  {job['label']}" if job["kind"] != "Daily"
            else f"Daily  ·  {job['label']}")


def _eligibility_note(job, ds):
    s, e = ds["start"], ds["end"]
    span = _fmt_d(s) if s == e else f"{_fmt_d(s)} to {_fmt_d(e)}"
    return (f"Assignments whose deadline fell on {span} and had already passed when this report "
            f"was generated ({_fmt_dt(ds['as_of'])} IST). Expected = learners enrolled for the "
            f"assignment; status as recorded in IntelliBIAssessmentSubmission at that time.")


def _totals_row(ws, row, n, label_span, values, pct_col=None):
    """Attendance-report TOTAL row: light-blue band, bold navy text, the label
    merged over the first `label_span` columns; values = {column: value}."""
    for col in range(1, n + 1):
        c = ws.cell(row=row, column=col)
        c.fill = AR._fill(AR.C_BLUE_LITE)
        c.font = AR._font(bold=True, size=10, color=AR.C_NAV)
        c.border = AR._border()
        c.alignment = AR._align("center", "center")
        if col in values:
            c.value = values[col]
    if label_span > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=label_span)
    ws.cell(row=row, column=1).value = "⬛  TOTAL"
    ws.cell(row=row, column=1).alignment = AR._align("left", "center")
    if pct_col and isinstance(values.get(pct_col), float):
        ws.cell(row=row, column=pct_col).number_format = PCT_FMT
    ws.row_dimensions[row].height = 22
    return row + 1


def _kpi_strip(ws, row, kpis, spans):
    """KPI cards (label row + value row) over explicit column spans — the
    pyAssignmentSubmissionsReport KPI style, placed so narrow columns never
    truncate a label."""
    col = 1
    for (label, value, pal), span in zip(kpis, spans):
        lbl_bg, val_bg, txt = pal
        for rr, v, font, bg in ((row, label, AR._font(bold=True, size=9, color=AR.C_WHITE), lbl_bg),
                                (row + 1, value, AR._font(bold=True, size=18, color=txt), val_bg)):
            if span > 1:
                ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + span - 1)
            for cc in range(col, col + span):
                ws.cell(row=rr, column=cc).border = AR._border()
                ws.cell(row=rr, column=cc).fill = AR._fill(bg)
            c = ws.cell(row=rr, column=col)
            c.value, c.font = v, font
            c.alignment = AR._align("center", "center", wrap=True)
        col += span
    ws.row_dimensions[row].height = 22
    ws.row_dimensions[row + 1].height = 36
    return row + 2


def _kpi_pal(rate):
    """Attendance report: the % KPI is green at ≥ 75 %, red below."""
    if rate is None:
        return AR.KPI_BLUE
    return AR.KPI_GREEN if rate >= RATE_GOOD else AR.KPI_RED


def _perf_table(ws, row, n, headers, groups, label_cols):
    """Generic grouped performance table: label columns, then Assignments,
    Expected, Submitted, Not Submitted, Submission %. Returns next row."""
    AR.write_header_row(ws, row, headers, height=30)
    h = row
    row += 1
    tot = {"a": 0, "e": 0, "s": 0}
    for i, (key, c, n_asg) in enumerate(groups, 1):
        bg = _zebra(i)
        vals = [i] + list(key)[:label_cols] + [n_asg, c["expected"], c["submitted"],
                                               c["not_submitted"]]
        for col, v in enumerate(vals, 1):
            AR.style_data_cell(ws, row, col, v, bg=bg,
                               h_align="left" if (1 < col <= 1 + label_cols
                                                  and isinstance(v, str)) else "center",
                               wrap=1 < col <= 1 + label_cols)
        _pct_cell(ws, row, len(vals) + 1, c["rate"])
        tot["a"] += n_asg
        tot["e"] += c["expected"]
        tot["s"] += c["submitted"]
        row += 1
    rate = round(tot["s"] / tot["e"] * 100, 1) if tot["e"] else None
    k = 1 + label_cols
    row = _totals_row(ws, row, n, k, {k + 1: tot["a"], k + 2: tot["e"], k + 3: tot["s"],
                                      k + 4: tot["e"] - tot["s"], k + 5: rate}, pct_col=k + 5)
    AR._border_thick_outer(ws, h, 1, row - 1, n)
    return row


def build_performance_workbook(job, ds):
    s = summarise(ds)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── 1. Summary ───────────────────────────────────────────────────────────
    ws = wb.create_sheet("Summary")
    N = 8
    _title(ws, N, ds, f"Assignment Submission Performance  |  {_period_text(job, ds)}")
    r = _subtitle(ws, 2, N, _eligibility_note(job, ds))
    r = _kpi_strip(ws, r, [
        ("📚  Eligible Assignments", s["assignments"], AR.KPI_BLUE),
        ("👥  Submissions Expected", s["expected"], AR.KPI_BLUE),
        ("📊  Submission %", "—" if s["rate"] is None else f"{s['rate']:.1f}%", _kpi_pal(s["rate"])),
        ("❌  Not Submitted", s["not_submitted"], AR.KPI_RED),
        ("✅  Submitted", s["submitted"], AR.KPI_GREEN),
    ], [2, 1, 2, 2, 1]) + 1
    if not ds["assignments"]:
        r = _empty_band(ws, r, N, "No assignment deadline in this period had passed — "
                                  "nothing to evaluate.")
    else:
        AR.write_section_banner(ws, r, N, "  BATCH-WISE PERFORMANCE", AR.C_BLUE_MID)
        r = _perf_table(ws, r + 1, N, ["#", "Technology", "Batch", "Assignments", "Expected",
                                       "Submitted", "Not Submitted", "Submission %"],
                        group_rows(ds, lambda a: (a["technology"], a["batch"])), 2)
    _widths(ws, [6, 34, 30, 13, 12, 12, 14, 14])
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_NAV

    # ── 2. Assignment Performance ────────────────────────────────────────────
    ws = wb.create_sheet("Assignment Performance")
    H = ["#", "Technology", "Batch", "Assignment", "Start Date", "Deadline (IST)", "Max Marks",
         "Expected", "Submitted", "On Time", "Late", "Not Submitted", "Submission %", "Avg Marks"]
    N = len(H)
    _title(ws, N, ds, f"Assignment Performance  |  {_period_text(job, ds)}")
    _subtitle(ws, 2, N, _eligibility_note(job, ds))
    hr = 3
    AR.write_header_row(ws, hr, H)
    r = hr + 1
    if not ds["assignments"]:
        r = _empty_band(ws, r, N, "No eligible assignments in this period.")
    for i, a in enumerate(ds["assignments"], 1):
        bg = _zebra(i)
        vals = [i, a["technology"], a["batch"], a["title"],
                a["start"].date() if a["start"] else "", a["deadline"],
                a["max_marks"], a["expected"], a["submitted"], a["on_time"], a["late"],
                a["not_submitted"], a["rate"], a["avg_marks"] if a["avg_marks"] is not None else "—"]
        for col, v in enumerate(vals, 1):
            c = AR.style_data_cell(ws, r, col, v, bg=bg,
                                    h_align="left" if col in (2, 3, 4) else "center",
                                    wrap=col in (2, 3, 4))
            if col == 5 and v:
                c.number_format = D_FMT
            elif col == 6:
                c.number_format = DT_FMT
        _pct_cell(ws, r, 13, a["rate"])
        r += 1
    if ds["assignments"]:
        r = _totals_row(ws, r, N, 7, {8: s["expected"], 9: s["submitted"], 10: s["on_time"],
                                      11: s["late"], 12: s["not_submitted"], 13: s["rate"],
                                      14: s["avg_marks"] if s["avg_marks"] is not None else "—"},
                        pct_col=13)
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(N)}{r - 2}"
    _widths(ws, [6, 26, 26, 38, 13, 20, 10, 10, 11, 10, 9, 13, 13, 11])
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_BLUE_MID

    # ── 3. Learner Submission Detail ─────────────────────────────────────────
    #    One block per assignment (pyAssignmentSubmissionsReport Student_Detail
    #    layout): a 📝 assignment banner, the column headers, then its learners.
    ws = wb.create_sheet("Learner Submission Detail")
    H = ["#", "Learner", "Email", "Phone", "Submitted At", "Status", "Marks Obtained",
         "Max Marks", "Feedback", "Active Learner"]
    N = len(H)
    _title(ws, N, ds, f"Learner Submission Detail  |  {_period_text(job, ds)}")
    r = _subtitle(ws, 2, N, _eligibility_note(job, ds)
                  + "  Rows: green = submitted · red = not submitted · amber with strikethrough "
                    "= inactive learner (Active Learner = No).")
    if not ds["assignments"]:
        r = _empty_band(ws, r, N, "No eligible assignments in this period.")
    filter_top = None
    for k, a in enumerate(ds["assignments"]):
        if k:                                          # thin spacer between assignments
            ws.row_dimensions[r].height = 8
            r += 1
        AR.write_section_banner(ws, r, N, _assignment_banner(a), AR.C_BLUE_MID, height=26)
        r += 1
        AR.write_header_row(ws, r, H, height=26)
        first = r
        filter_top = filter_top or r                   # filter starts at the first header row
        r += 1
        for i, x in enumerate(a["rows"], 1):
            bg, font = _learner_row_style(x)
            vals = [i, x["learner"], x["email"], x["phone"], x["submitted_at"] or "", x["status"],
                    _marks_value(x["marks"]), a["max_marks"], x["feedback"], x["active"]]
            for col, v in enumerate(vals, 1):
                c = AR.style_data_cell(ws, r, col, v, bg=bg,
                                       h_align="left" if col in (2, 3, 9) else "center",
                                       wrap=(col == 9))
                c.font = font(col == 6)
                if col == 5 and v:
                    c.number_format = SUB_DT_FMT
            r += 1
        AR._border_thick_outer(ws, first, 1, r - 1, N)
    if filter_top:
        ws.auto_filter.ref = f"A{filter_top}:{get_column_letter(N)}{r - 1}"
    _widths(ws, [6, 26, 32, 16, 21, 15, 15, 11, 40, 14])
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_TEAL
    _print_setup(wb)
    return wb, s


def _print_setup(wb):
    for ws in wb.worksheets:
        try:
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:                                        # pragma: no cover
            pass


def build_non_submission_workbook(job, ds):
    s = summarise(ds)
    pend = [r for r in ds["rows"] if r["status"] == NOT_SUBMITTED]
    per_learner = OrderedDict()
    for x in pend:
        k = _learner_key({"student_id": x["student_id"], "student_email": x["email"],
                          "student_name": x["learner"]})
        per_learner.setdefault(k, []).append(x)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── 1. Non-Submitters (one row per learner per assignment) ───────────────
    ws = wb.create_sheet("Non-Submitters")
    H = ["#", "Technology", "Batch", "Assignment", "Deadline (IST)", "Days Since Deadline",
         "Learner", "Email", "Phone", "Active Learner", "Pending in Period"]
    N = len(H)
    _title(ws, N, ds, f"Assignment Non-Submission  |  {_period_text(job, ds)}")
    r = _subtitle(ws, 2, N, "Learners who had NOT submitted when the assignment deadline passed. "
                  + _eligibility_note(job, ds))
    active_pending = len({k for k, xs in per_learner.items() if xs[0]["active"] == "Yes"})
    r = _kpi_strip(ws, r, [
        ("❌  Pending Submissions", s["not_submitted"], AR.KPI_RED),
        ("👥  Learners Pending", len(per_learner), AR.KPI_AMBER),
        ("📚  Assignments Affected", sum(1 for a in ds["assignments"] if a["not_submitted"]), AR.KPI_BLUE),
        ("✅  Active Learners Pending", active_pending, AR.KPI_BLUE),
    ], [3, 2, 3, 3])
    hr = r + 1
    AR.write_header_row(ws, hr, H)
    r = hr + 1
    if not pend:
        r = _empty_band(ws, r, N, "Every learner submitted every eligible assignment of this period."
                        if ds["assignments"] else
                        "No assignment deadline in this period had passed — nothing to follow up.")
    for i, x in enumerate(pend, 1):
        k = _learner_key({"student_id": x["student_id"], "student_email": x["email"],
                          "student_name": x["learner"]})
        bg = AR.C_RED_LITE if i % 2 else AR.C_RED_PALE          # Attendance report absent rows
        vals = [i, x["technology"], x["batch"], x["assignment"], x["deadline"],
                x["days_overdue"], x["learner"], x["email"], x["phone"], x["active"],
                len(per_learner[k])]
        for col, v in enumerate(vals, 1):
            c = AR.style_data_cell(ws, r, col, v, bg=bg, bold=(col == 7),
                                    h_align="left" if col in (2, 3, 4, 7, 8) else "center",
                                    wrap=col in (2, 3, 4))
            if col == 5:
                c.number_format = DT_FMT
            if col == 10 and v != "Yes":
                c.font = AR._font(size=10, color="757575", italic=True)
        r += 1
    if pend:
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(N)}{r - 1}"
    _widths(ws, [7, 24, 24, 34, 20, 11, 26, 30, 15, 12, 11])
    ws.freeze_panes = ws.cell(row=hr + 1, column=8).coordinate
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_RED_DARK

    # ── 2. By Learner (one row per learner — the follow-up call list) ────────
    ws = wb.create_sheet("By Learner")
    H = ["#", "Learner", "Email", "Phone", "Active Learner", "Pending Assignments",
         "Technology / Batch", "Assignments Not Submitted (deadline)"]
    N = len(H)
    _title(ws, N, ds, f"Non-Submitters by Learner  |  {_period_text(job, ds)}")
    _subtitle(ws, 2, N, "One row per learner with every assignment of this period they had not "
                        "submitted by its deadline — most pending first.")
    hr = 3
    AR.write_header_row(ws, hr, H)
    r = hr + 1
    if not per_learner:
        r = _empty_band(ws, r, N, "No learner has a pending assignment for this period.")
    order = sorted(per_learner.values(), key=lambda xs: (-len(xs), xs[0]["learner"].lower()))
    for i, xs in enumerate(order, 1):
        x0 = xs[0]
        bg = _zebra(i)
        tb = "; ".join(OrderedDict.fromkeys(f"{x['technology']} · {x['batch']}" for x in xs))
        al = "\n".join(f"{x['assignment']} ({x['deadline']:%d-%b %I:%M %p})" for x in xs)
        vals = [i, x0["learner"], x0["email"], x0["phone"], x0["active"], len(xs), tb, al]
        for col, v in enumerate(vals, 1):
            AR.style_data_cell(ws, r, col, v, bg=bg, bold=(col in (2, 6)),
                                h_align="left" if col in (2, 3, 7, 8) else "center",
                                wrap=col in (7, 8))
        ws.row_dimensions[r].height = max(18, 15 * len(xs))
        r += 1
    if per_learner:
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(N)}{r - 1}"
    _widths(ws, [6, 26, 30, 15, 12, 12, 34, 56])
    ws.freeze_panes = ws.cell(row=hr + 1, column=3).coordinate
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_AMBER_DARK

    # ── 3. Assignment Reconciliation (ties to the Performance Report) ────────
    ws = wb.create_sheet("Assignment Reconciliation")
    H = ["#", "Technology", "Batch", "Assignment", "Deadline (IST)", "Expected", "Submitted",
         "Not Submitted (listed)", "Submission %"]
    N = len(H)
    _title(ws, N, ds, f"Reconciliation with the Performance Report  |  "
                                     f"{_period_text(job, ds)}")
    _subtitle(ws, 2, N, "Same assignments and counts as the Assignment Submission Performance "
                        "Report for this period; 'Not Submitted (listed)' = rows on Non-Submitters.")
    hr = 3
    AR.write_header_row(ws, hr, H)
    r = hr + 1
    if not ds["assignments"]:
        r = _empty_band(ws, r, N, "No eligible assignments in this period.")
    for i, a in enumerate(ds["assignments"], 1):
        vals = [i, a["technology"], a["batch"], a["title"], a["deadline"], a["expected"],
                a["submitted"], a["not_submitted"], a["rate"]]
        for col, v in enumerate(vals, 1):
            if col == 9:
                _pct_cell(ws, r, 9, v)
                continue
            c = AR.style_data_cell(ws, r, col, v, bg=_zebra(i),
                                   h_align="left" if col in (2, 3, 4) else "center",
                                   wrap=col in (2, 3, 4))
            if col == 5:
                c.number_format = DT_FMT
        r += 1
    if ds["assignments"]:
        r = _totals_row(ws, r, N, 5, {6: s["expected"], 7: s["submitted"], 8: s["not_submitted"],
                                      9: s["rate"]}, pct_col=9)
    _widths(ws, [6, 26, 26, 38, 20, 11, 11, 14, 13])
    ws.freeze_panes = ws.cell(row=hr + 1, column=5).coordinate
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = AR.C_NAV
    _print_setup(wb)
    return wb, {"pending": s["not_submitted"], "learners": len(per_learner),
                "active_learners": active_pending}


# =============================================================================
#  NON-SUBMISSION PDFs — ONE PDF PER ELIGIBLE ASSIGNMENT
#  The approach of pyAssignmentSubmissionEmailReminder.generate_assignment_pdfs:
#  one file per (class, assignment), named Assignment_<Class>_<Assignment>_… with
#  its _sanitize_filename_part, and the look of its _build_single_assignment_pdf
#  (reportlab, landscape A4, navy #1F3864 header table, row-coloured learner
#  table, small footer). Built from the SAME dataset as both Sheets: a PDF lists
#  exactly that assignment's "Not Submitted" rows, so the PDFs together list
#  exactly the Performance Report's Not Submitted.
# =============================================================================
PDF_NAVY, PDF_META_BG, PDF_GRID = "#1F3864", "#F5F7FA", "#9E9E9E"
PDF_ROW = {"Yes": "#FFEBEE",            # active learner — the reminder's "Final" red
           "No": "#" + ROW_INACTIVE}    # inactive (Students tab = N) — amber + strikethrough,
PDF_INACTIVE_TEXT = "#5D4037"           # as on the Learner Submission Detail tab
PDF_ROW_OTHER = "#F5F5F5"               # active status not found / unknown


def _pdf_esc(v):
    from xml.sax.saxutils import escape
    return escape(str(v if v not in (None, "") else "—"))


def _pdf_dt(dt):
    return dt.strftime("%d-%b-%Y %I:%M %p") if dt else "—"


def _pdf_banner_text(a):
    """'Class | Duration | Assignment | Start | Deadline' — the Learner Submission
    Detail banner, in the same order (📝 replaced by the word: not in PDF fonts)."""
    start = a["start"].strftime("%m/%d/%Y") if a["start"] else "—"
    return (f"Class: {_pdf_esc(a['technology'])}  |  Duration: {_pdf_esc(a['batch'])}  |  "
            f"Assignment: {_pdf_esc(a['title'])}  |  Start: {start}  |  "
            f"Deadline: {a['deadline']:%m/%d/%Y}")


def pdf_assignments(ds):
    """The assignments that get a PDF: every eligible assignment of the period
    (PDF_FOR_FULLY_SUBMITTED), otherwise only those with non-submitters."""
    return [a for a in ds["assignments"] if a["not_submitted"] or PDF_FOR_FULLY_SUBMITTED]


def pdf_filenames(job, ds):
    """{(assessment_id, class_id): file name} —
    Assignment_<Class>_<Assignment>_<Kind>_<label>.pdf. When two
    assignments share class and title (two batches), the class duration is added;
    if still not unique, the class id."""
    san = ASG._sanitize_filename_part
    period = re.sub(r"[^0-9A-Za-z\-]+", "_", job["label"]).strip("_")
    asg = pdf_assignments(ds)

    def stem(a, level):
        parts = [san(a["technology"], 40), san(a["title"], 60)]
        if level >= 1:
            parts.append(san(a["batch"], 40))
        if level >= 2:
            parts.append(san(a["class_id"], 30))
        return "Assignment_" + "_".join(parts)

    out = {}
    for a in asg:
        for level in range(3):
            st = stem(a, level)
            if level == 2 or sum(1 for b in asg if stem(b, level) == st) == 1:
                break
        out[(a["assessment_id"], a["class_id"])] = f"{st}_{job['kind']}_{period}.pdf"
    return out


def _pending_per_learner(ds):
    per = defaultdict(int)
    for r in ds["rows"]:
        if r["status"] == NOT_SUBMITTED:
            per[_learner_key({"student_id": r["student_id"], "student_email": r["email"],
                              "student_name": r["learner"]})] += 1
    return per


def build_assignment_pdf(job, ds, a, pending_per_learner=None) -> bytes:
    """One assignment's PDF: header, the Class | Duration | Assignment | Start |
    Deadline banner, the assignment's figures and the learners who had NOT
    submitted when its deadline passed (active learners first, inactive last)."""
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    per = pending_per_learner if pending_per_learner is not None else _pending_per_learner(ds)
    xs = [r for r in a["rows"] if r["status"] == NOT_SUBMITTED]
    # active learners first, inactive (struck through) last — presentation only
    xs.sort(key=lambda r: ({"Yes": 0, "No": 2}.get(r["active"], 1), r["learner"].lower()))
    C, P = colors.HexColor, Paragraph
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=14 * mm,
                            title=f"IntelliBI Not Submitted — {a['technology']} — {a['title']}",
                            author="IntelliBI Innovations Technologies")
    ss = getSampleStyleSheet()
    st_title = ParagraphStyle("T", parent=ss["Title"], fontSize=15, textColor=C(PDF_NAVY),
                              spaceAfter=4, alignment=1)
    st_sub = ParagraphStyle("S", parent=ss["Normal"], fontSize=10, textColor=C("#555555"),
                            alignment=1, spaceAfter=8)
    st_small = ParagraphStyle("Sm", parent=ss["Normal"], fontSize=9, leading=11)
    st_cell = ParagraphStyle("C", parent=ss["Normal"], fontSize=8, leading=10)
    st_cell_b = ParagraphStyle("CB", parent=st_cell, fontName="Helvetica-Bold")
    st_center = ParagraphStyle("CC", parent=st_cell, alignment=1)
    st_hcell = ParagraphStyle("HC", parent=ss["Normal"], fontSize=8, leading=10,
                              textColor=colors.white, fontName="Helvetica-Bold", alignment=1)
    st_label = ParagraphStyle("L", parent=st_small, textColor=colors.white, fontName="Helvetica-Bold")
    st_banner = ParagraphStyle("B", parent=ss["Normal"], fontSize=10, leading=13,
                               textColor=colors.white, fontName="Helvetica-Bold")

    def grid(t, extra=()):
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("GRID", (0, 0), (-1, -1), 0.4, C(PDF_GRID)),
                               ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                               ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
                              + list(extra)))
        return t

    story = [P("IntelliBI Assignment Submission — Not Submitted Follow-Up", st_title),
             P(f"{_pdf_esc(job['kind'])} Report  |  Period: {_pdf_esc(job['label'])}  |  "
               f"Generated: {_pdf_dt(ds['as_of'])} IST", st_sub),
             grid(Table([[P(_pdf_banner_text(a), st_banner)]], colWidths=[273 * mm]),
                  [("BACKGROUND", (0, 0), (-1, -1), C(PDF_NAVY))]),
             Spacer(1, 6)]

    # ── assignment figures (reminder's two-column meta table) ────────────────
    rate = "—" if a["rate"] is None else f"{a['rate']:.1f}%"
    rate_col = ("#2E7D32" if a["rate"] is not None and a["rate"] >= RATE_GOOD else
                "#E65100" if a["rate"] is not None and a["rate"] >= RATE_WATCH else "#C62828")
    active_n = sum(1 for x in xs if x["active"] == "Yes")
    days = max(0, (ds["as_of"].date() - a["deadline"].date()).days)
    ns_col = "#C62828" if a["not_submitted"] else "#2E7D32"
    meta = [("Not Submitted", f"<font color='{ns_col}'><b>{a['not_submitted']}</b></font> of "
                              f"{a['expected']} learner(s)  —  {active_n} active"),
            ("Submitted", f"<b>{a['submitted']}</b>  (Submission % <font color='{rate_col}'><b>{rate}"
                          f"</b></font>; on time {a['on_time']}, late {a['late']})"),
            ("Deadline Passed", f"{_pdf_dt(a['deadline'])} IST  —  {days} day(s) before this report"),
            ("Maximum Marks", _pdf_esc(a["max_marks"] or "—"))]
    story += [grid(Table([[P(k, st_label), P(v, st_small)] for k, v in meta],
                         colWidths=[45 * mm, 228 * mm], hAlign="LEFT"),
                   [("BACKGROUND", (0, 0), (0, -1), C(PDF_NAVY)),
                    ("BACKGROUND", (1, 0), (1, -1), C(PDF_META_BG))]),
              Spacer(1, 8)]

    if not xs:
        story.append(grid(Table([[P(f"<b>All {a['expected']} learner(s) submitted this assignment — "
                                    f"nothing to follow up.</b>", st_small)]], colWidths=[273 * mm]),
                          [("BACKGROUND", (0, 0), (-1, -1), C("#E8F5E9"))]))
    else:
        rows = [[P(h, st_hcell) for h in ("#", "Learner", "Email", "Phone", "Active Learner",
                                          "Pending in Period", "Follow-Up Notes")]]
        cmds = [("BACKGROUND", (0, 0), (-1, 0), C(PDF_NAVY)), ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ALIGN", (0, 0), (0, -1), "CENTER")]
        for i, x in enumerate(xs, 1):
            k = _learner_key({"student_id": x["student_id"], "student_email": x["email"],
                              "student_name": x["learner"]})
            n = per.get(k, 1)
            pend = f"<b><font color='#C62828'>{n}</font></b>" if n > 1 else str(n)
            cells = [(str(i), st_center), (_pdf_esc(x["learner"]), st_cell_b),
                     (_pdf_esc(x["email"]), st_cell), (_pdf_esc(x["phone"]), st_cell),
                     (_pdf_esc(x["active"]), st_center), (pend, st_center)]
            if x["active"] == "No":                 # inactive: whole row struck through
                cells = [(f"<strike><font color='{PDF_INACTIVE_TEXT}'>{re.sub(r'</?font[^>]*>', '', t)}"
                          f"</font></strike>", st) for t, st in cells]
            rows.append([P(t, st) for t, st in cells] + [P("", st_cell)])
            cmds.append(("BACKGROUND", (0, i), (-1, i), C(PDF_ROW.get(x["active"], PDF_ROW_OTHER))))
        tbl = Table(rows, colWidths=[10 * mm, 52 * mm, 72 * mm, 32 * mm, 24 * mm, 24 * mm, 59 * mm],
                    repeatRows=1, hAlign="LEFT")
        story.append(grid(tbl, cmds))

    def footer(canvas, d):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(C("#888888"))
        canvas.drawString(12 * mm, 7 * mm, f"Automated report — IntelliBI Innovations Technologies  |  "
                                           f"{a['technology']} · {a['title']}  |  {job['kind']} {job['label']}")
        canvas.drawRightString(d.pagesize[0] - 12 * mm, 7 * mm, f"Page {d.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


def build_non_submission_pdfs(job, ds) -> list:
    """One PDF per assignment (pdf_assignments), in the Performance Report's
    order: [{"filename", "pdf_bytes", "technology", "batch", "assignment",
    "assessment_id", "class_id", "pending", "expected"}]."""
    names = pdf_filenames(job, ds)
    per = _pending_per_learner(ds)
    return [{"filename": names[(a["assessment_id"], a["class_id"])],
             "pdf_bytes": build_assignment_pdf(job, ds, a, per),
             "technology": a["technology"], "batch": a["batch"], "assignment": a["title"],
             "assessment_id": a["assessment_id"], "class_id": a["class_id"],
             "pending": a["not_submitted"], "expected": a["expected"]}
            for a in pdf_assignments(ds)]


# =============================================================================
#  GOOGLE DRIVE  (impersonated info@, native Google Sheets, never overwritten)
# =============================================================================
FOLDER_MIME = "application/vnd.google-apps.folder"
SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MIME = "application/pdf"


def _call(fn, what):
    try:
        import api_retry                                        # common/api_retry.py
        return api_retry.call_with_retry(fn, what, log=log.info)
    except ImportError:                                         # pragma: no cover
        return fn()


def _drive_client():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive"]
    ).with_subject(IMPERSONATE_USER)
    return gbuild("drive", "v3", credentials=creds, cache_discovery=False)


def _q(s):
    return str(s).replace("\\", "\\\\").replace("'", "\\'")


def _find_or_create_folder(drive, parent_id, name):
    q = (f"'{parent_id}' in parents and name='{_q(name)}' and mimeType='{FOLDER_MIME}' "
         f"and trashed=false")
    found = _call(lambda: drive.files().list(q=q, fields="files(id,name)", supportsAllDrives=True,
                                             includeItemsFromAllDrives=True).execute(),
                  f"Drive: find folder {name}").get("files", [])
    if found:
        return found[0]["id"]
    meta = {"name": name, "mimeType": FOLDER_MIME, "parents": [parent_id]}
    f = _call(lambda: drive.files().create(body=meta, fields="id", supportsAllDrives=True).execute(),
              f"Drive: create folder {name}")
    log.info("Created Drive folder '%s'", name)
    return f["id"]


def period_folders(job):
    """(report-type folder, reporting-period folder) — the same names under both roots."""
    return (KIND_FOLDERS[job["kind"]],
            CP.period_folder_name(job["kind"], job["start"], job["end"]))


def report_name(basename, job):
    """IntelliBI_..._Report_Daily_04-Oct-2026 / _Weekly_28-Sep-2026_to_04-Oct-2026 /
    _Monthly_September_2026 / _Manual_21-Aug-2026_to_22-Sep-2026."""
    safe = re.sub(r"[^0-9A-Za-z\-]+", "_", job["label"]).strip("_")
    return f"{basename}_{job['kind']}_{safe}"


def _existing(drive, folder_id, base):
    q = f"'{folder_id}' in parents and name contains '{_q(base)}' and trashed=false"
    return _call(lambda: drive.files().list(q=q, fields="files(id,name,webViewLink,mimeType)",
                                            supportsAllDrives=True,
                                            includeItemsFromAllDrives=True).execute(),
                 "Drive: list reports").get("files", [])


def resolve_period_folder(drive, root_id, job):
    parent = root_id
    for name in period_folders(job):
        parent = _find_or_create_folder(drive, parent, name)
    return parent


def _version_of(name, ext=""):
    m = re.search(r"-\s*Version\s*(\d+)\s*" + re.escape(ext) + "$", name or "", re.I)
    return int(m.group(1)) if m else None


def _pick_existing(existing, base, ext=""):
    """The report to overwrite: the one named exactly <base><ext>, otherwise the
    highest '- Version N' (older versions are left untouched)."""
    exact = [f for f in existing if f.get("name") == base + ext]
    if exact:
        return exact[0]
    return max(existing, key=lambda f: _version_of(f.get("name"), ext) or 0)


def _upload(drive, root_id, job, base, data, media_mime, file_mime, link_fmt, ext="",
            folder_id=None, where=None, tag="Uploaded"):
    """Upload <data> as <base><ext> into <root>/<type>/<period>/ (or folder_id).
    A file of this name already there (same type, or its '- Version N') is
    overwritten in place (OVERWRITE_EXISTING: same file id and link, content
    replaced); otherwise a new file is created."""
    from googleapiclient.http import MediaIoBaseUpload
    folder_id = folder_id or resolve_period_folder(drive, root_id, job)
    existing = [f for f in _existing(drive, folder_id, base)
                if f.get("mimeType", file_mime) == file_mime
                and (f.get("name") == base + ext or _version_of(f.get("name"), ext) is not None
                     and re.fullmatch(re.escape(base) + r"\s*-\s*Version\s*\d+\s*" + re.escape(ext),
                                      f.get("name") or ""))]
    media = lambda: MediaIoBaseUpload(io.BytesIO(data), mimetype=media_mime, resumable=False)
    where = where or "/".join(period_folders(job))
    if existing and OVERWRITE_EXISTING:
        old = _pick_existing(existing, base, ext)
        try:
            up = _call(lambda: drive.files().update(fileId=old["id"], media_body=media(),
                                                    fields="id,webViewLink",
                                                    supportsAllDrives=True).execute(),
                       f"Drive: overwrite {old['name']}")
            link = up.get("webViewLink") or link_fmt.format(id=old["id"])
            print(f"[Drive] {tag}: {where}/{old['name']}  (overwrote the existing report)")
            return link
        except Exception as exc:                                # noqa: BLE001
            # content replace refused → new file under the same name, old one to the
            # Drive trash (restorable there), so the folder still holds one report
            log.warning("In-place overwrite of %s failed (%s) — replacing the file.", old["name"], exc)
            up = drive.files().create(body={"name": old["name"], "parents": [folder_id],
                                            "mimeType": file_mime},
                                      media_body=media(), fields="id,webViewLink",
                                      supportsAllDrives=True).execute()
            _call(lambda: drive.files().update(fileId=old["id"], body={"trashed": True},
                                               supportsAllDrives=True).execute(),
                  f"Drive: trash replaced {old['name']}")
            print(f"[Drive] {tag}: {where}/{old['name']}  (replaced the existing report; "
                  f"the old copy is in the Drive trash)")
            return up.get("webViewLink") or link_fmt.format(id=up["id"])
    name = base + ext
    if existing:                                     # OVERWRITE_EXISTING = False
        vers = [v for f in existing for v in [_version_of(f.get("name"), ext)] if v]
        name = f"{base} - Version {max(vers + [1]) + 1}{ext}"
    meta = {"name": name, "parents": [folder_id], "mimeType": file_mime}
    # not retried: a lost response must never leave two copies of the report
    up = drive.files().create(body=meta, media_body=media(), fields="id,webViewLink",
                              supportsAllDrives=True).execute()
    link = up.get("webViewLink") or link_fmt.format(id=up["id"])
    print(f"[Drive] {tag}: {where}/{name}")
    return link


def upload_workbook(drive, root_id, job, basename, wb) -> str:
    """Upload as a native Google Sheet into <root>/<type>/<period>/ (see _upload)."""
    buf = io.BytesIO()
    wb.save(buf)                                     # in memory only — never written to disk
    return _upload(drive, root_id, job, report_name(basename, job), buf.getvalue(), XLSX_MIME,
                   SHEET_MIME, "https://docs.google.com/spreadsheets/d/{id}/edit")


def _find_folder(drive, parent_id, name):
    q = (f"'{parent_id}' in parents and name='{_q(name)}' and mimeType='{FOLDER_MIME}' "
         f"and trashed=false")
    return _call(lambda: drive.files().list(q=q, fields="files(id,name)", supportsAllDrives=True,
                                            includeItemsFromAllDrives=True).execute(),
                 f"Drive: find folder {name}").get("files", [])


def upload_assignment_pdfs(drive, job, pdfs) -> dict:
    """Upload each assignment PDF (PDF file, no conversion) into the period
    folder of the Non-Submission root, beside the Non-Submission Sheet,
    overwriting a same-name PDF in place. With OVERWRITE_EXISTING, assignment
    PDFs of this period that no longer match an assignment, and the earlier
    layouts (single consolidated PDF, "Assignment-wise PDFs" subfolder), go to
    the Drive trash (restorable), so the folder always matches the report.
    Returns {"folder_link", "links"}."""
    folder_id = resolve_period_folder(drive, NON_SUBMISSION_ROOT_FOLDER_ID, job)
    links = {}
    for p in pdfs:
        links[p["filename"]] = _upload(drive, NON_SUBMISSION_ROOT_FOLDER_ID, job, p["filename"][:-4],
                                       p["pdf_bytes"], PDF_MIME, PDF_MIME,
                                       "https://drive.google.com/file/d/{id}/view", ext=".pdf",
                                       folder_id=folder_id, tag="PDF uploaded")
    if OVERWRITE_EXISTING:
        made = {p["filename"] for p in pdfs}
        suffix = "_" + job["kind"] + "_" + re.sub(r"[^0-9A-Za-z\-]+", "_", job["label"]).strip("_") + ".pdf"
        stale = [f for f in _existing(drive, folder_id, "Assignment_")
                 if f.get("mimeType") == PDF_MIME and f.get("name", "").startswith("Assignment_")
                 and f["name"].endswith(suffix) and f["name"] not in made]
        stale += [f for f in _existing(drive, folder_id, report_name(LEGACY_PDF_BASENAME, job))
                  if f.get("mimeType") == PDF_MIME]
        stale += _find_folder(drive, folder_id, LEGACY_PDF_SUBFOLDER)
        for f in stale:
            _call(lambda f=f: drive.files().update(fileId=f["id"], body={"trashed": True},
                                                   supportsAllDrives=True).execute(),
                  f"Drive: trash {f['name']}")
            print(f"[Drive] Moved to trash (no longer in this period's report): {f['name']}")
    return {"folder_link": f"https://drive.google.com/drive/folders/{folder_id}", "links": links}


# =============================================================================
#  E-MAIL  (coordinator_email: same sender, recipients and style as the
#  Coordinator Task Performance report)
# =============================================================================
def _tech_table_html(techs):
    """E-mail 'By Technology' table, colour-coded like pyAttendaceFeedbackReport's
    'Attendance Summary — by Technology': each technology row takes the
    background of its Submission % band and the % cell the band's strong text
    colour in bold — AR._email_att_color (Green > 70 · Amber 55–70 · Red < 55).
    Same columns, values and order as before.
    The last row is TOTAL, as in that table: Assignments, Expected, Submitted and
    Not Submitted summed over the technologies, and Submission % worked out from
    those summed counts (Submitted ÷ Expected), never an average of the
    technology percentages. The whole row is bold and takes the colour of its
    own band."""
    E = CE._E
    heads = ["Technology", "Assignments", "Expected", "Submitted", "Not Submitted", "Submission %"]
    align = ["left"] + ["center"] * 5
    th = "".join(f"<th style='background:{CE.HEADER};color:#fff;padding:7px 8px;font-size:11.5px;"
                 f"text-align:{a};border:1px solid #d0d7e2'>{h}</th>" for h, a in zip(heads, align))
    td = "padding:6px 8px;font-size:12.5px;border:1px solid #d0d7e2;text-align:%s;"
    def _row(label, n, c, rate, bold=False):
        if rate is None:
            bg, tx, pct = "#ffffff", "#5b6b86", "—"
        else:
            bg, tx = AR._email_att_color(rate)
            pct = f"{rate:.1f}%"
        cells = [label, n, c["expected"], c["submitted"], c["not_submitted"]]
        weight = ";font-weight:700" if bold else ""
        return (f"<tr style='background:{bg};color:#1a2a48{weight}'>"
                + "".join(f"<td style='{td % a}'>{v}</td>" for v, a in zip(cells, align))
                + f"<td style='{td % 'center'}color:{tx};font-weight:700'>{pct}</td></tr>")

    trs = [_row(E(k), n, c, c["rate"]) for k, c, n in techs]
    tot = {f: sum(c[f] for _k, c, _n in techs) for f in ("expected", "submitted", "not_submitted")}
    tot_rate = round(tot["submitted"] / tot["expected"] * 100, 1) if tot["expected"] else None
    trs.append(_row("TOTAL", sum(n for _k, _c, n in techs), tot, tot_rate, bold=True))
    legend = ("<div style='font-size:11px;color:#5b6b86;margin-top:5px'>Submission %: "
              "<span style='color:#1B7A34;font-weight:700'>Green &gt; 70%</span> · "
              "<span style='color:#8A6D00;font-weight:700'>Amber 55–70%</span> · "
              "<span style='color:#B02A2A;font-weight:700'>Red &lt; 55%</span></div>")
    return ("<table role='presentation' width='100%' style='border-collapse:collapse;"
            f"border:1px solid #e2e8f0'><tr>{th}</tr>{''.join(trs)}</table>{legend}")


def email_html(job, s, ds, perf_link, nonsub_link, pdf_link=None, pdf_count=None):
    E = CE._E
    rate_col = CE.band_color(s["rate"], RATE_GOOD, RATE_WATCH)
    body = CE.section("Submission Performance") + CE.card_block([
        ("Eligible Assignments", s["assignments"], CE.NAVY),
        ("Submissions Expected", s["expected"], CE.NAVY),
        ("Submission %", "—" if s["rate"] is None else f"{s['rate']:.1f}%", rate_col),
        ("Submitted", s["submitted"], CE.GREEN),
        ("Not Submitted", s["not_submitted"], CE.RED if s["not_submitted"] else CE.GREEN),
        ("Learners Pending", s["learners_pending"], CE.RED if s["learners_pending"] else CE.GREEN),
    ], 3)
    techs = group_rows(ds, lambda a: a["technology"])
    if techs:
        body += CE.section("By Technology") + _tech_table_html(techs)
    else:
        body += ("<p style='margin:12px 0 0;color:#5b6b86'>No assignment deadline in this period "
                 "had passed, so there is nothing to evaluate.</p>")
    if nonsub_link:
        body += (CE.section("Follow-up") +
                 f"<p style='margin:0;line-height:1.5'>The learners who did not submit are listed, "
                 f"with contact details, in the <a href='{E(nonsub_link)}' style='color:{CE.BTN};"
                 f"font-weight:600;text-decoration:none'>Assignment Non-Submission Report</a>.</p>")
    if pdf_link:
        n = pdf_count or 0
        body += ((CE.section("Follow-up") if not nonsub_link else "") +
                 f"<p style='margin:8px 0 0;line-height:1.5'>One printable PDF per assignment "
                 f"({n} file{'s' if n != 1 else ''}) with its learners who did not submit is in the "
                 f"<a href='{E(pdf_link)}' style='color:{CE.BTN};font-weight:600;"
                 f"text-decoration:none'>Non-Submission folder of this period</a>.</p>")
    span = job["label"]
    intro = (f"Please find the <b>{E(job['kind'])}</b> Assignment Submission Performance report for "
             f"<b>{E(span)}</b>. It covers assignments whose deadline fell in this period and had "
             f"passed when the report was generated.")
    return CE.page(f"{job['kind']} Assignment Submission Report", span, intro, body, perf_link,
                   "Assignment Submission Performance Report")


def send_email(job, s, ds, perf_link, nonsub_link, pdf_link=None, pdf_count=None) -> bool:
    subject = f"{job['kind']} Assignment Submission Report - {job['label']}"
    return CE.send(subject, email_html(job, s, ds, perf_link, nonsub_link, pdf_link, pdf_count),
                   EMAIL_RECIPIENTS,
                   sender=EMAIL_SENDER, star=STAR_EMAIL_IN_GMAIL)


# =============================================================================
#  RUN
# =============================================================================
def load_submissions(service):
    df = _call(lambda: ASG.read_sheet_df(service, SUBMISSION_SHEET_ID, SUBMISSIONS_TAB),
               "Sheets: read Submissions")
    print(f"[Data] Submissions rows: {len(df)}")
    return df


def _non_submission_pdfs(job, ds, s, out, drive):
    """Build (and upload) one PDF per assignment into out[...]. Never raises: a
    problem is reported and the Sheets / e-mail of the period still go out.
    Nothing is uploaded unless the PDFs reconcile with the Performance Report."""
    try:
        pdfs = build_non_submission_pdfs(job, ds)
        listed = sum(p["pending"] for p in pdfs)
        if listed != s["not_submitted"] or len(pdfs) != len(pdf_assignments(ds)):
            raise AssertionError(f"PDFs list {listed} learner row(s), the Performance Report "
                                 f"{s['not_submitted']} — not reconciled, not uploaded")
        out["pdfs"] = pdfs
        out["pdf_info"] = {"files": len(pdfs), "pending": listed,
                           "with_pending": sum(1 for p in pdfs if p["pending"])}
        if drive is not None:
            up = upload_assignment_pdfs(drive, job, pdfs)
            out["pdf_link"], out["pdf_links"] = up["folder_link"], up["links"]
    except Exception as exc:                                   # noqa: BLE001
        log.exception("Non-Submission PDFs %s %s failed: %s", job["kind"], job["label"], exc)
        print(f"[PDF FAILED] Non-Submission PDFs — {job['kind']} {job['label']}: {exc}")
        out["pdf_error"] = str(exc)


def run_jobs(jobs, subs_df, now, status_map=None, drive=None, upload=True, email=True):
    """Build (and deliver) both reports for every job. One failing job never
    stops the others. Returns a list of result dicts."""
    results = []
    for job in jobs:
        out = {"job": job}
        try:
            ds = build_dataset(subs_df, job["start"], job["end"], now, status_map)
            perf_wb, s = build_performance_workbook(job, ds)
            ns_wb, ns = build_non_submission_workbook(job, ds)
            assert ns["pending"] == s["not_submitted"], "reports do not reconcile"
            out.update(summary=s, non_submission=ns)
            if upload and drive is not None:
                out["perf_link"] = upload_workbook(drive, PERFORMANCE_ROOT_FOLDER_ID, job,
                                                   PERFORMANCE_BASENAME, perf_wb)
                out["nonsub_link"] = upload_workbook(drive, NON_SUBMISSION_ROOT_FOLDER_ID, job,
                                                     NON_SUBMISSION_BASENAME, ns_wb)
            out["workbooks"] = (perf_wb, ns_wb)
            if GENERATE_NON_SUBMISSION_PDF:
                _non_submission_pdfs(job, ds, s, out, drive if upload else None)
            rate = "—" if s["rate"] is None else f"{s['rate']:.1f}%"
            print(f"\n{'=' * 70}\n  Assignment Submission — {job['kind']} {job['label']}\n"
                  f"  Eligible assignments: {s['assignments']} | Expected: {s['expected']} | "
                  f"Submitted: {s['submitted']} | Not submitted: {s['not_submitted']} | "
                  f"Submission {rate}\n"
                  f"  Learners pending: {s['learners_pending']} | Not yet due: {s['not_yet_due']} | "
                  f"No deadline (whole sheet): {s['no_deadline']}\n"
                  + (f"  Assignment PDFs: {out['pdf_info']['files']} file(s) — "
                     f"{out['pdf_info']['pending']} learner row(s) not submitted in "
                     f"{out['pdf_info']['with_pending']} assignment(s)\n" if out.get("pdf_info") else "")
                  + ("  [Drive] upload OFF — built in memory only\n" if not out.get("perf_link")
                     else "") + "=" * 70)
            if email:
                out["emailed"] = send_email(job, s, ds, out.get("perf_link"), out.get("nonsub_link"),
                                            out.get("pdf_link"), len(out.get("pdfs") or []))
        except Exception as exc:                                   # noqa: BLE001
            log.exception("Assignment report %s %s failed: %s", job["kind"], job["label"], exc)
            print(f"[FAILED] Assignment Submission — {job['kind']} {job['label']}: {exc}")
            out["failed"] = True
            out["error"] = str(exc)
        results.append(out)
    return results


def generate(dry_run=False):
    now = now_ist()
    jobs, errors = plan_jobs(now.date())
    for e in errors:
        log.error("Configuration: %s", e)
        print(f"[CONFIG ERROR] {e}")
    if not jobs:
        print("[Plan] No report planned (check the GENERATE_* flags).")
        return [], errors
    print("[Plan] " + ", ".join(f"{j['kind']} {j['label']}" for j in jobs))
    from utils import get_sheets_service
    service = get_sheets_service(SERVICE_ACCOUNT_FILE)
    subs_df = load_submissions(service)
    status_map = ASG.load_active_status_map(service)
    upload = UPLOAD_TO_DRIVE and not dry_run
    drive = _drive_client() if upload else None
    results = run_jobs(jobs, subs_df, now, status_map, drive, upload=upload,
                       email=SEND_EMAIL and not dry_run)
    return results, errors


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="Assignment Submission Performance & Non-Submission reports")
    ap.add_argument("--dry-run", action="store_true",
                    help="build both reports in memory and print the summary; no upload, no e-mail")
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    results, errors = generate(dry_run=a.dry_run)
    failed = [r for r in results if r.get("failed")]
    print(f"\nAssignment reports built: {len(results) - len(failed)} | failed: {len(failed)} | "
          f"configuration errors: {len(errors)}")
    return 1 if (failed or errors) else 0


if __name__ == "__main__":
    sys.exit(main())
