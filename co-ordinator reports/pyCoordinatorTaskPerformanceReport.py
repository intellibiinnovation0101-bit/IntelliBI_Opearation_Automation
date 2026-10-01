"""
================================================================================
  IntelliBI Operations Automation
  COORDINATOR TASK PERFORMANCE & PROGRESS REPORT
  (co-ordinator reports / pyCoordinatorTaskPerformanceReport.py)
  ------------------------------------------------------------------------------
  WHAT IT ANSWERS
    How much of the Coordinator's generated work is completed, and is it being
    completed at the right time?
        Tasks Generated -> Completed -> Pending -> Completion % ->
        Timely Completion % -> Progress over time -> Overall performance
    overall AND per task group (Learner Attendance / Assignment / Admission
    Formalities / Wise & Interview Feedback Validation / Instructor Follow-Ups /
    Interview Reminders — and any future group that carries the follow-up
    columns).

  SOURCE OF TRUTH  (read-only — nothing in the existing system is changed)
    The Batch Coordinator daily reports produced by
    pyBatchCoordinatorDailyAttendanceReport.py are the ONLY place Coordinator
    actions are recorded (Action Taken / Follow-Up Comment / Follow-Up Done? /
    Follow-Up DateTime). They live in
        <coordinator folder>/Daily Coordinator Reports/Daily DD-Mon-YYYY/
    (and, for reports made before that layout, <coordinator folder>/YYYY-MM-DD/)
    as native Google Sheets; a re-run the same day adds "<name> - Version N" and
    never touches earlier versions. This script reads every daily report version
    of the days in the reporting period (Drive export, cached per file until the
    file changes) and rebuilds a de-duplicated task ledger from them.

  REPORTING PERIOD = TASK ORIGIN
    A task belongs to the report day whose Coordinator list it appeared on
    (the Daily folder / report date). Each report measures ONLY the tasks that
    originated inside its own period:
        Daily   — that day's tasks
        Weekly  — tasks of the 7 report days Mon–Sun
        Monthly — tasks of the calendar month
        Manual  — tasks of MANUAL_START_DATE … MANUAL_END_DATE
    Earlier days' pending / overdue tasks are never pulled in, and "Days on
    List" counts consecutive report days WITHIN the period only.

  HOW A TASK IS COUNTED (see the report's "Data Coverage & Rules" tab)
    * A task = one flagged row on one report day's action list, identified by
      the task group's STABLE fields (learner / assignment / session / record —
      never rank, row number or volatile figures).
    * All versions of the same day are merged: the day's task list is the
      latest version's list + any task the Coordinator actioned in an earlier
      version of that day. A task is counted ONCE per day however many
      versions show it (no double counting across versions).
    * Completed = "Follow-Up Done?" = Yes in ANY version of that day (follow-up
      entries are not carried between versions, so the latest version alone
      would lose work). Completion time = the earliest valid Follow-Up
      DateTime among the "Yes" entries.
    * An item still on the list the next day is that day's new task (each
      daily list is "today's" action list, with fresh follow-up columns); the
      "Days on List" / unique-item views show the accumulation.
    * A day/group whose report had no follow-up columns (every report before
      the columns were introduced) cannot show completion — it is reported as
      "not trackable" and is NEVER counted as pending.

  TIMELINESS (from the system itself, not an invented SLA)
    Every task tab is generated as the action list for its report day ("needs a
    follow-up today"), and the task-specific dates on the lists (assignment
    deadline, admission expiry, interview date) are all on/after that day, so:
        due              = end of the report day (IST)
        On time          = completed on the report day
        Late             = completed on a later day
        Missed           = not completed and the report day is over
        Open (due today) = not completed, report day still running
        Timely %         = On time / (all tasks except "Yes" without a valid time)
                           — open tasks count as not (yet) on time
    A "Yes" without a valid DateTime (e.g. the #REF! stamps before iterative
    calculation was enabled) counts as completed but is excluded from every
    timing figure. NOTE: Follow-Up DateTime stamps the FIRST time Follow-Up
    Done? is set (Yes or No) — a task first set to No and later to Yes keeps
    the earlier stamp.

  OUTPUT
    One native Google Sheet per report, in the SAME reporting-period folder as
    the Batch Coordinator report of that type and period (coordinator_periods):
        <coordinator folder>/<Daily|Weekly|Monthly|Manual> Coordinator Reports/
                             <reporting-period folder>/
    versioned exactly like the Coordinator report ("- Version N"). Google Drive
    is the ONLY place a report is stored: the workbook is built in memory and
    uploaded from memory — no local copy, no temporary file.
    Tabs: Dashboard | Progress Trend | Task Register |
          Data Coverage & Rules

  RUN
    python "co-ordinator reports/pyCoordinatorTaskPerformanceReport.py"
    Scheduled daily at 19:00 by scripts/setup_schedule.ps1 (evening batch,
    scripts/run_evening_reports.py), after the Coordinator has worked the day's
    list. Periods are chosen by the GENERATE_* flags below (same scheme as
    pyLeadFollowUpAnalysisReport.py: AUTO = Daily every run, Weekly on Monday
    for the previous Mon–Sun week, Monthly on the last day of the month).
    Exit code 0 = every planned report delivered, 1 = something failed.
================================================================================
"""
from __future__ import annotations

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT = os.path.dirname(_HERE)
for _p in (_HERE, os.path.join(_PROJECT, "common"),
           os.path.join(_PROJECT, "ops_reports_action")):
    if os.path.isdir(_p) and _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import _bootstrap  # noqa: F401  (sys.path + env defaults + config.yaml)
except Exception:                                              # pragma: no cover
    pass

import io
import re
import json
import logging
import unicodedata
from collections import defaultdict, OrderedDict
from datetime import datetime, date, timedelta, time, timezone
from statistics import median

import openpyxl
from openpyxl.utils import get_column_letter as _gcl
from openpyxl.chart import BarChart, LineChart, Reference

# The Coordinator report is the task generator: its constants (folder, file
# name, follow-up columns, banner colours) and its design system are reused so
# this report can never drift from what the Coordinator actually sees.
import pyBatchCoordinatorDailyAttendanceReport as BC
import coordinator_periods as CP              # shared periods + Drive layout
import coordinator_email as CE                # shared e-mail (Operations Gmail account)

AR = BC.AR
try:
    import api_retry                                    # common/api_retry.py
except Exception:                                       # pragma: no cover
    api_retry = None

log = logging.getLogger("CoordinatorTaskPerformance")

# =============================================================================
#  RUN CONFIGURATION
# =============================================================================
REPORT_BASENAME    = "IntelliBI_Coordinator_Task_Performance_Report"
UPLOAD_TO_DRIVE    = True        # False = dry run: built in memory only, nothing saved anywhere
VERBOSE            = True

# ── E-mail (sent after the report(s) are generated) ─────────────────────────
# True  → generate the report(s) and e-mail the result + Google Sheet link(s)
# False → generate / upload exactly the same, but send NO e-mail
# Sent through the Operations Gmail account in credentials/email_config.py.
SEND_EMAIL       = True
EMAIL_SENDER     = "info@intellibiinnovationstechnologies.in"
EMAIL_RECIPIENTS = ["info@intellibiinnovationstechnologies.in",
                    "intellibihropsb2ch@gmail.com"]
# Star (★) each report e-mail in the sending Gmail account once it is sent
# (common/gmail_star.py; best-effort — never affects sending or the run).
STAR_EMAIL_IN_GMAIL = True

# "Performance vs Goals" (e-mail body only): one bar per task group that has
# tasks in the reported period — its Completion % (the SAME figure as the
# Dashboard scorecard) against this benchmark. >= benchmark → green, below → red.
COMPLETION_BENCHMARK = 95.0

# Short, management-friendly names used ONLY in the e-mail's Performance vs
# Goals section. Keyed by the TASK_GROUPS "name" below — task-group identity,
# the Dashboard and every other tab keep the registry names. A group missing
# here is shown under its registry name.
EMAIL_GROUP_LABELS = {
    "Learner Attendance Follow-Ups":         "Learner Attendance",
    "Learner Assignment Follow-Ups":         "Learner Assignment",
    "Learner Admission Formalities":         "Learner Admission Formalities",
    "Wise & Interview Feedback Validation":  "Learner Admission Formalities",
    "Instructor Follow-Ups":                 "Instructor Instructions",
    "Learner Instructor Interview Reminder": "Interview Reminder",
}

# =============================================================================
#  REPORT GENERATION CONTROL  (same scheme as pyLeadFollowUpAnalysisReport.py)
# =============================================================================
#  GENERATE_AUTO = True  → scheduler-friendly selection by the run date; the
#  GENERATE_* flags and dates below are IGNORED for that run:
#      • Daily   — every run, for the current day
#      • Weekly  — every Monday, for the previous completed week (Mon → Sun)
#      • Monthly — on the last calendar day of the month, for that whole month
#  GENERATE_AUTO = False → each GENERATE_* flag works independently (several can
#  be True). The dates pin a specific period; None = today / this week / this
#  month. Manual needs both dates (Start <= End).
GENERATE_AUTO    = False

GENERATE_DAILY   = True
GENERATE_WEEKLY  = False
GENERATE_MONTHLY = False
GENERATE_MANUAL  = False

DAILY_DATE            = None  # "YYYY-MM-DD"
WEEKLY_REFERENCE_DATE = None  # any date within required week
MONTHLY_MONTH         = None  # 1-12
MONTHLY_YEAR          = None
MANUAL_START_DATE     = "2026-08-21"
MANUAL_END_DATE       = "2026-09-22"

# Performance status bands — the SAME bands IntelliBI already uses for follow-up
# progress (pyLeadFollowUpAnalysisReport._ctrend_status): >= 75 On track,
# >= 45 Watch, else Behind.
STATUS_ON_TRACK = 75.0
STATUS_WATCH    = 45.0

CACHE_DIR = os.path.join(_PROJECT, "cache", "coordinator_performance")
# Where earlier versions of this script kept a local copy of every report. The
# report is no longer written to disk; each run removes any such leftover copies
# (only this report's own .xlsx files, then the folders they leave empty).
LEGACY_OUTPUT_DIR = os.path.join(_PROJECT, "output", "reports", "coordinator_performance")

# =============================================================================
#  TASK GROUP REGISTRY
#  ---------------------------------------------------------------------------
#  One entry per Coordinator task tab. `tabs` = tab names (current + earlier
#  names); `identity` = the STABLE fields that identify one task, resolved from
#  the row's own column first, then from the section banners above it
#  ("Tech Name: …", "Duration: …", the banner title = "Section" / "Subsection").
#  A future task tab that carries the follow-up columns is picked up
#  automatically even without an entry here (generic identity, logged); add an
#  entry to give it a precise identity and a short name.
# =============================================================================
TASK_GROUPS = [
    {"key": "attendance", "name": "Learner Attendance Follow-Ups", "short": "Attendance",
     "tabs": ["Learner Attendance Follow-Ups"],
     "identity": ["Tech Name", "Duration", "Student Name", "Phone"],
     "who": ["Student Name", "Phone"]},
    {"key": "assignment", "name": "Learner Assignment Follow-Ups", "short": "Assignment",
     "tabs": ["Learner Assignment Follow-Ups"],
     "identity": ["Tech Name", "Subsection", "Email", "Student Name"],
     "who": ["Student Name", "Email"]},
    {"key": "admission", "name": "Learner Admission Formalities", "short": "Admission",
     "tabs": ["Learner Admission Formalities"],
     "identity": ["Email ID", "Student Name"],
     "who": ["Student Name", "Email ID"]},
    {"key": "wise", "name": "Wise & Interview Feedback Validation", "short": "Wise & IV Feedback",
     "tabs": ["Wise & Interview Feedback Validation", "Learner Wise Validation"],
     "identity": ["Section", "Student Name", "Batch Name", "Course Title", "Course Subtitle",
                  "Instructor ID", "Instructor Name", "Interview Start Date",
                  "Interviewer Name", "Batch Title / Duration"],
     "who": ["Student Name", "Course Title", "Instructor Name", "Interviewer Name"]},
    {"key": "instructor", "name": "Instructor Follow-Ups", "short": "Instructor",
     "tabs": ["Instructor Follow-Ups"],
     "identity": ["Tech Name", "Duration", "Instructor", "Session Date"],
     "who": ["Instructor"]},
    {"key": "interview", "name": "Learner Instructor Interview Reminder", "short": "Interview Reminder",
     "tabs": ["Learner Instructor Interview Reminder"],
     "identity": ["Interview Time", "Batch Name (Class)", "Candidate Name"],
     "who": ["Candidate Name"]},
]
# columns never used for a generic (unregistered) identity — they change between runs
_VOLATILE_HEADERS = {"#", "rank", "why flagged"}

FU_ACTION, FU_COMMENT, FU_DONE, FU_DT = BC.FOLLOWUP_COLS     # the four follow-up columns
HEADER_ANCHOR = "Why Flagged"                                # present in every task tab

# =============================================================================
#  SMALL HELPERS
# =============================================================================
_IST = timezone(timedelta(hours=5, minutes=30))


def _now_ist() -> datetime:
    return AR._ist_now()


def _norm(s) -> str:
    """Comparison key: NFKC, no edit marker / emoji / placeholders, collapsed
    whitespace, case-folded."""
    t = unicodedata.normalize("NFKC", "" if s is None else str(s))
    t = t.replace("✎", " ").replace("​", " ")
    t = " ".join(t.split()).strip()
    if t in ("—", "-", "–", "None", "nan"):
        return ""
    return t.casefold()


def _display(v) -> str:
    """Human-readable cell value (floats that are whole numbers shown as ints,
    datetimes as dd-Mon-yyyy HH:MM)."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d-%b-%Y %H:%M")
    if isinstance(v, date):
        return v.strftime("%d-%b-%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = " ".join(str(v).split())
    return "" if s in ("—", "-", "–") else s


def _key_value(field: str, v) -> str:
    """Normalised identity value: phones reduced to their last 10 digits, dates
    to ISO, text case/space-insensitive."""
    if isinstance(v, str) and re.match(r"^\s*\d{4}-\d{2}-\d{2}", v):
        try:                                   # same value typed as text or as a date
            v = datetime.fromisoformat(v.strip()[:19].replace("T", " "))
        except ValueError:
            pass
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d" if v.time() == time(0) else "%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = _norm(v)
    if "phone" in field.casefold():
        d = re.sub(r"\D", "", s)
        return d[-10:] if d else s
    return s


def _is_yes(v) -> bool:
    return _norm(v).lstrip("✅✔ ").startswith("yes")


def _is_no(v) -> bool:
    return _norm(v).lstrip("❌✖ ").startswith("no")


def _parse_stamp(v):
    """Follow-Up DateTime cell -> naive IST datetime, or None (blank / #REF! /
    any error text). Handles datetime objects, sheet serial numbers and the
    formatted text 'dd-Mon-yyyy HH:MM:SS'."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.replace(tzinfo=None, microsecond=0)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if 20000 < float(v) < 80000:                     # Sheets / Excel serial day
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).replace(microsecond=0)
        return None
    s = str(v).strip()
    if not s or s.startswith("#"):
        return None
    for f in ("%d-%b-%Y %H:%M:%S", "%d-%b-%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
              "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f"):
        try:
            return datetime.strptime(s, f)
        except ValueError:
            pass
    return None


def _drive_time_ist(rfc3339: str) -> datetime | None:
    """Drive createdTime/modifiedTime (UTC, RFC 3339) -> naive IST."""
    try:
        dt = datetime.strptime(rfc3339.replace("Z", "+0000"), "%Y-%m-%dT%H:%M:%S.%f%z")
    except ValueError:
        try:
            dt = datetime.strptime(rfc3339.replace("Z", "+0000"), "%Y-%m-%dT%H:%M:%S%z")
        except ValueError:
            return None
    return dt.astimezone(_IST).replace(tzinfo=None, microsecond=0)


def _fmt_hours(h) -> str:
    if h is None:
        return "—"
    m = int(round(h * 60))
    if m < 60:
        return f"{m} min"
    d, rem = divmod(m, 1440)
    hh, mm = divmod(rem, 60)
    return (f"{d}d {hh}h {mm:02d}m" if d else f"{hh}h {mm:02d}m")


def _pct(n, d):
    return (100.0 * n / d) if d else None


def _status_word(p):
    if p is None:
        return "—"
    return "On track" if p >= STATUS_ON_TRACK else ("Watch" if p >= STATUS_WATCH else "Behind")


def _status_level(word):
    return {"On track": "ok", "Watch": "medium", "Behind": "high"}.get(word, "muted")


def _call(fn, what):
    if api_retry is not None:
        return api_retry.call_with_retry(fn, what, log=log.info)
    return fn()


# =============================================================================
#  TASK-TAB PARSER  (one exported report version -> task occurrences)
# =============================================================================
def _group_for_tab(title: str):
    """Registry entry for a tab title (tolerant of the 31-character tab-name
    limit and of earlier tab names)."""
    t = _norm(title)
    for g in TASK_GROUPS:
        for alias in g["tabs"]:
            a = _norm(alias)
            if t == a or (len(t) >= 20 and a.startswith(t)) or t.startswith(a):
                return g
    return None


_SEG_SPLIT = re.compile(r"\s{2,}[·•|]\s{2,}")
_KV = re.compile(r"^([A-Za-z][A-Za-z /()&.\-]{0,30}?)\s*:\s+(.+)$")
_COUNT_SEG = re.compile(r"^\d+\s+\S")                    # "14 learners pending …"
_VOLATILE_KEYS = {"pending", "today", "max marks", "assigned"}
_CONTEXT_KEYS = {"tech name", "duration", "batch", "batch name", "technology"}


def _parse_banner(text: str) -> dict:
    """'Tech Name: X · Duration: Y · 14 pending' -> {'Tech Name': X, 'Duration': Y}.
    The leading free-text segment (a banner title) is returned as '_title'.
    Counts and other run-to-run volatile segments are dropped."""
    out = {}
    raw = unicodedata.normalize("NFKC", str(text or "")).strip()
    raw = re.sub(r"^[^\w(]+", "", raw)                  # leading emoji / symbols
    for i, seg in enumerate(_SEG_SPLIT.split(raw)):
        seg = seg.strip()
        if not seg or _COUNT_SEG.match(seg):
            continue
        m = _KV.match(seg)
        if i == 0 and (not m or m.group(1).strip().casefold() not in _CONTEXT_KEYS):
            # banner title — may itself contain a colon ("LangGraph: software …")
            out["_title"] = re.split(r"\s{2,}", seg)[0].strip()
        elif m and m.group(1).strip().casefold() not in _VOLATILE_KEYS:
            out[m.group(1).strip()] = m.group(2).strip()
    return out


def _merged_spans(ws) -> dict:
    """(row, col) of each merged range's top-left cell -> number of columns."""
    return {(m.min_row, m.min_col): m.max_col - m.min_col + 1 for m in ws.merged_cells.ranges}


def parse_task_tab(ws, sub_fill: str, section_fill: str) -> dict:
    """Parse one task tab of an exported Coordinator report.

    Returns {"has_followup": bool, "rows": [ {identity fields, follow-up values,
    what (first line of Why Flagged)} … ]}.
    Header rows are found by the 'Why Flagged' column (present in every task tab,
    old and new); follow-up columns are optional (reports before they existed).
    Full-width merged rows are section banners (context), never tasks."""
    spans = _merged_spans(ws)
    header = None                          # {normalised header: col index}
    header_names = {}                      # col index -> display header
    ctx1, ctx2 = {}, {}
    rows, has_fu = [], False
    width = ws.max_column or 1
    for r in range(1, (ws.max_row or 0) + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, width + 1)]
        normed = [_norm(v) for v in vals]
        if _norm(HEADER_ANCHOR) in normed:                    # a (section) header row
            header = {}
            header_names = {}
            for c, (nv, v) in enumerate(zip(normed, vals), 1):
                if nv and nv not in header:
                    header[nv] = c
                    header_names[c] = _display(v).replace("✎", "").strip()
            has_fu = has_fu or (_norm(FU_DONE) in header)
            ctx2 = {}
            continue
        span = spans.get((r, 1), 1)
        if span >= 3:                              # merged full-width row: banner or chrome
            fill = (ws.cell(row=r, column=1).fill.fgColor.rgb or "")[-6:].upper()
            text = "" if vals[0] is None else str(vals[0])
            # Only the report's section banners carry context; the title band,
            # guide strip, grouped super-headers and empty-state rows do not.
            if (header is not None and text.strip() and not text.strip().startswith("✅")
                    and fill not in (sub_fill.upper(), section_fill.upper())):
                # banner of an earlier (pre design-system) layout: same meaning,
                # other colours — assignment detail lines start with the 📝 marker
                fill = sub_fill.upper() if text.strip().startswith("📝") else section_fill.upper()
            if fill == sub_fill.upper() and text.strip():
                ctx2 = {("Subsection" if k == "_title" else k): v
                        for k, v in _parse_banner(text).items()}
            elif fill == section_fill.upper() and text.strip():
                ctx1 = {("Section" if k == "_title" else k): v
                        for k, v in _parse_banner(text).items()}
                ctx2 = {}
            continue
        if header is None or not any(normed):
            continue
        rec = {"_cells": {}, "_ctx": {**ctx1, **ctx2}}
        for nv, c in header.items():
            rec["_cells"][header_names[c]] = vals[c - 1] if c - 1 < len(vals) else None
        rec["_has_fu"] = _norm(FU_DONE) in header
        rows.append(rec)
    return {"has_followup": has_fu, "rows": rows}


# Earlier report layouts named the group only in the banner title (e.g.
# "AI (GenAI • Agentic AI • ML)  ·  55 pending" instead of "Tech Name: …").
_FIELD_FALLBACK = {"tech name": "Section"}


def _field(rec: dict, name: str):
    """Identity field from the row's own column, else from the banners above
    (with the documented fallback for earlier banner layouts)."""
    for k, v in rec["_cells"].items():
        if _norm(k) == _norm(name) and _display(v):
            return v
    for k, v in rec["_ctx"].items():
        if _norm(k) == _norm(name) and _display(v):
            return v
    alt = _FIELD_FALLBACK.get(_norm(name))
    if alt:
        return _field(rec, alt)
    return None


def _occurrence(group: dict, rec: dict):
    """One parsed row -> task occurrence (None when the row has no identity)."""
    fields = group.get("identity")
    if not fields:                                   # generic (unregistered) group
        fields = [k for k in rec["_cells"]
                  if _norm(k) not in _VOLATILE_HEADERS and _norm(k) not in
                  {_norm(x) for x in BC.FOLLOWUP_COLS}][:3]
        fields = list(rec["_ctx"].keys()) + fields
    who_order = [_norm(x) for x in (group.get("who") or fields[-1:])]
    ident, who, context = [], {}, []
    for f in fields:
        v = _field(rec, f)
        kv = _key_value(f, v) if v is not None else ""
        ident.append(kv)
        if kv:
            if _norm(f) in who_order:
                who[_norm(f)] = _display(v)
            else:
                context.append(_display(v))
    label = [who[w] for w in who_order if w in who]      # 'who' first, in registry order
    if not any(ident):
        return None
    cells = rec["_cells"]
    why = ""
    for k, v in cells.items():
        if _norm(k) == _norm(HEADER_ANCHOR):
            why = str(v or "")
    what = re.sub(r"^\s*(\d+\)|•)\s*", "", why.strip().split("\n")[0]).strip()

    def _fu(name):
        for k, v in cells.items():
            if _norm(k) == _norm(name):
                return v
        return None
    return {
        "key": "|".join(ident),
        "label": "  ·  ".join(dict.fromkeys(label)) or "  ·  ".join(dict.fromkeys(context)),
        "context": "  ·  ".join(dict.fromkeys(context)) if label else "",
        "what": what[:160],
        "has_fu": rec["_has_fu"],
        "done": ("yes" if _is_yes(_fu(FU_DONE)) else "no" if _is_no(_fu(FU_DONE)) else ""),
        "stamp_raw": _display(_fu(FU_DT)),
        "stamp": _parse_stamp(_fu(FU_DT)),
        "action": _display(_fu(FU_ACTION)),
        "comment": _display(_fu(FU_COMMENT)),
    }


def parse_report_workbook(wb) -> dict:
    """Exported Coordinator report workbook -> {group_key: {...}} with every task
    occurrence of every task tab. Tabs that are not task tabs are ignored."""
    out = {}
    for ws in wb.worksheets:
        group = _group_for_tab(ws.title)
        parsed = parse_task_tab(ws, BC.DS_SUB, BC.DS_SECTION)
        if group is None:
            if not parsed["has_followup"]:
                continue                                  # not a Coordinator task tab
            group = {"key": "tab:" + _norm(ws.title), "name": ws.title, "short": ws.title,
                     "identity": None, "generic": True}
            log.warning("Tab '%s' has follow-up columns but no TASK_GROUPS entry — "
                        "using a generic task identity.", ws.title)
        occ = [o for o in (_occurrence(group, r) for r in parsed["rows"]) if o]
        out[group["key"]] = {"name": group["name"], "short": group.get("short", group["name"]),
                             "has_followup": parsed["has_followup"], "tasks": occ}
    return out


# =============================================================================
#  REPORT SOURCES  (Drive: every version of every daily Coordinator report)
# =============================================================================
_VERSION = re.compile(r"-\s*Version\s*(\d+)\s*$", re.IGNORECASE)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _drive_service():
    """Impersonated Drive client — the same identity (and the same helper) that
    uploads the Coordinator reports, so every version is visible."""
    return BC._drive_client()


def _children(drive, parent_id, folders_only=False):
    return _call(lambda: BC._list_children(drive, parent_id, folders_only=folders_only),
                 "Drive: list folder")


def _in_any(day: date, ranges) -> bool:
    return any(a <= day <= b for a, b in ranges)


def discover_report_versions(drive, ranges) -> list:
    """Every native-Sheet version of the DAILY Coordinator task report whose
    report day falls inside one of `ranges` [(start, end), …].

    Where they are read from (task origin = the report day):
      * <root>/Daily Coordinator Reports/Daily DD-Mon-YYYY/   (current layout)
      * <root>/YYYY-MM-DD/                                     (legacy layout)
    Only files named like the daily task report (BC.REPORT_BASENAME) are used —
    Weekly / Monthly / Manual roll-ups have no follow-up columns and are not task
    lists, and performance reports are never read back."""
    root_children = _children(drive, BC.PARENT_FOLDER_ID, folders_only=True)
    day_folders = [f for f in root_children if CP._LEGACY_DAY_FOLDER.match(f.get("name", ""))]
    for f in root_children:
        if f.get("name") == CP.KIND_FOLDERS["Daily"]:
            day_folders += [x for x in _children(drive, f["id"], folders_only=True)
                            if CP._DAILY_FOLDER.match(x.get("name", ""))]
    versions = []
    for f in day_folders:
        day = CP.daily_folder_date(f["name"])
        if day is None or not _in_any(day, ranges):
            continue
        for x in _children(drive, f["id"]):
            if x.get("mimeType") != "application/vnd.google-apps.spreadsheet":
                continue
            if not x.get("name", "").startswith(BC.REPORT_BASENAME):
                continue
            m = _VERSION.search(x["name"])
            versions.append({"id": x["id"], "name": x["name"], "day": day,
                             "version": int(m.group(1)) if m else 1,
                             "created": _drive_time_ist(x.get("createdTime", "")),
                             "modified_raw": x.get("modifiedTime", ""),
                             "folder": f["name"]})
    versions.sort(key=lambda v: (v["day"], v["created"] or datetime.min, v["version"]))
    log.info("Found %d Coordinator report version(s) across %d report day(s).",
             len(versions), len({v["day"] for v in versions}))
    return versions


def _cache_path(v):
    safe = re.sub(r"[^0-9A-Za-z]", "", v["modified_raw"])
    return os.path.join(CACHE_DIR, f"{v['id']}_{safe}.json")


def _ser(parsed):
    def enc(o):
        o = dict(o)
        o["stamp"] = o["stamp"].isoformat() if o.get("stamp") else None
        return o
    return {k: {**g, "tasks": [enc(t) for t in g["tasks"]]} for k, g in parsed.items()}


def _deser(data):
    for g in data.values():
        for t in g["tasks"]:
            t["stamp"] = datetime.fromisoformat(t["stamp"]) if t.get("stamp") else None
    return data


def load_version(drive, v) -> dict:
    """Parsed task tabs of one report version. Exported once per file revision
    (cached by file id + modifiedTime), so re-runs only fetch changed files."""
    path = _cache_path(v)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                return _deser(json.load(fh))
        except Exception:                                        # corrupt cache -> refetch
            pass
    data = _call(lambda: drive.files().export(fileId=v["id"], mimeType=XLSX_MIME).execute(),
                 f"Drive: export {v['name']}")
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    parsed = parse_report_workbook(wb)
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(_ser(parsed), fh)
    except Exception as exc:                                     # cache is best-effort
        log.warning("Could not write cache %s (%s).", path, exc)
    return parsed


# =============================================================================
#  TASK LEDGER  (versions of each report day merged into de-duplicated tasks)
# =============================================================================
ST_ON_TIME = "Completed on time"
ST_LATE    = "Completed late"
ST_UNKNOWN = "Completed (time not recorded)"
ST_OPEN    = "Open (due today)"
ST_MISSED  = "Missed"
COMPLETED_STATES = (ST_ON_TIME, ST_LATE, ST_UNKNOWN)
PENDING_STATES = (ST_OPEN, ST_MISSED)


def _touched(o) -> bool:
    return bool(o["done"] or o["action"] or o["comment"])


def build_ledger(versions: list, loader, now: datetime) -> dict:
    """Merge every version of every report day into one task per (day, group,
    task identity). Returns {"tasks": [...], "coverage": [...], "groups": {...},
    "tracked_days": {group_key: [days]}}."""
    by_day = OrderedDict()
    for v in versions:
        by_day.setdefault(v["day"], []).append(v)

    tasks, coverage, groups = [], [], OrderedDict()
    for g in TASK_GROUPS:                                  # registry order first
        groups[g["key"]] = {"name": g["name"], "short": g["short"]}
    tracked_days = defaultdict(list)

    for day, vs in by_day.items():
        parsed = []
        for v in vs:
            try:
                parsed.append((v, loader(v)))
            except Exception as exc:                       # one bad file never stops the report
                log.warning("Skipping %s (%s).", v["name"], exc)
        cov = {"day": day, "versions": len(parsed),
               "versions_fu": sum(1 for _v, p in parsed if any(t["has_followup"] for t in p.values())),
               "groups_tracked": 0, "groups_untracked": 0, "tasks": 0, "untracked_tasks": 0,
               "dropped": 0, "multi_version": 0, "yes_no_time": 0, "conflicts": 0}
        day_start = datetime.combine(day, time(0, 0))
        due = datetime.combine(day, time(23, 59, 59))
        gkeys = OrderedDict()
        for _v, p in parsed:
            for gk, gd in p.items():
                gkeys.setdefault(gk, gd)
        for gk, gmeta in gkeys.items():
            groups.setdefault(gk, {"name": gmeta["name"], "short": gmeta["short"]})
            tabs = [(v, p[gk]) for v, p in parsed if gk in p]
            final_v, final_tab = tabs[-1]                  # latest version holding this tab
            first_seen = {}
            for v, tab in tabs:
                for o in tab["tasks"]:
                    if v["created"] and (o["key"] not in first_seen or v["created"] < first_seen[o["key"]]):
                        first_seen[o["key"]] = v["created"]
            final_keys = OrderedDict((o["key"], o) for o in final_tab["tasks"])
            if not final_tab["has_followup"]:
                cov["groups_untracked"] += 1
                cov["untracked_tasks"] += len(final_keys)
                continue
            cov["groups_tracked"] += 1
            tracked_days[gk].append(day)
            occ = defaultdict(list)                        # key -> [(version, occurrence)]
            for v, tab in tabs:
                if tab["has_followup"]:
                    for o in tab["tasks"]:
                        occ[o["key"]].append((v, o))
            universe = list(final_keys)
            for k, lst in occ.items():
                if k not in final_keys and any(_touched(o) for _v, o in lst):
                    universe.append(k)                     # actioned, then dropped off a later run
            cov["dropped"] += sum(1 for k in occ if k not in universe)

            for k in universe:
                lst = occ.get(k, [])
                ref = final_keys.get(k) or lst[-1][1]
                yes = [(v, o) for v, o in lst if o["done"] == "yes"]
                valid = sorted(o["stamp"] for _v, o in yes
                               if o["stamp"] and o["stamp"] >= day_start)
                comp = valid[0] if valid else None
                recorded = [v for v, o in lst if _touched(o)]
                if len(recorded) > 1:
                    cov["multi_version"] += 1
                if yes and any(o["done"] == "no" for _v, o in lst):
                    cov["conflicts"] += 1
                if yes and comp is None:
                    cov["yes_no_time"] += 1
                src = None                                  # entry whose action/comment we show
                if yes:
                    src = min(yes, key=lambda vo: (vo[1]["stamp"] is None or vo[1]["stamp"] < day_start,
                                                   vo[1]["stamp"] or datetime.max))[1]
                else:
                    touched = [o for _v, o in lst if _touched(o)]
                    src = touched[-1] if touched else None
                if yes:
                    status = (ST_UNKNOWN if comp is None else ST_ON_TIME if comp <= due else ST_LATE)
                else:
                    status = ST_OPEN if now <= due else ST_MISSED
                fs = first_seen.get(k)
                ttc = None
                if comp and fs:
                    ttc = max(0.0, (comp - fs).total_seconds() / 3600.0)
                tasks.append({
                    "day": day, "group": gk, "key": k, "label": ref["label"],
                    "context": ref.get("context", ""), "what": ref["what"],
                    "first_seen": fs, "due": due, "status": status,
                    "completed": bool(yes), "completed_at": comp, "ttc_h": ttc,
                    "attempted": (not yes) and any(_touched(o) for _v, o in lst),
                    "action": (src or {}).get("action", ""), "comment": (src or {}).get("comment", ""),
                    "no_action_recorded": bool(yes) and not any(o["action"] for _v, o in yes),
                    "versions_seen": sorted({v["version"] for v, tab in tabs
                                             if any(o["key"] == k for o in tab["tasks"])}),
                    "versions_recorded": sorted({v["version"] for v in recorded}),
                    "in_final": k in final_keys,
                })
                cov["tasks"] += 1
        coverage.append(cov)

    # Days on list: consecutive tracked report days (of the task's group) ending
    # at this task's day on which the item was listed and NOT completed.
    idx = {(t["group"], t["key"], t["day"]): t for t in tasks}
    for t in tasks:
        if t["completed"]:
            t["days_on_list"] = 0
            continue
        days = tracked_days[t["group"]]
        i = days.index(t["day"])
        n = 0
        while i >= 0:
            prev = idx.get((t["group"], t["key"], days[i]))
            if prev is None or prev["completed"]:
                break
            n += 1
            i -= 1
        t["days_on_list"] = n
    return {"tasks": tasks, "coverage": coverage, "groups": groups,
            "tracked_days": dict(tracked_days)}


def scope_ledger(ledger: dict, start: date, end: date) -> dict:
    """The ledger restricted to ONE reporting period: only tasks whose report day
    (task origin) is within [start, end], and "Days on List" recounted over the
    period's own report days — so no earlier day's pending/overdue workload ever
    reaches the period's figures."""
    import copy
    tasks = [copy.copy(t) for t in ledger["tasks"] if start <= t["day"] <= end]
    tracked = {g: [d for d in days if start <= d <= end]
               for g, days in ledger["tracked_days"].items()}
    idx = {(t["group"], t["key"], t["day"]): t for t in tasks}
    for t in tasks:
        if t["completed"]:
            t["days_on_list"] = 0
            continue
        days = tracked.get(t["group"], [])
        i = days.index(t["day"]) if t["day"] in days else -1
        n = 0
        while i >= 0:
            prev = idx.get((t["group"], t["key"], days[i]))
            if prev is None or prev["completed"]:
                break
            n += 1
            i -= 1
        t["days_on_list"] = n
    return {"tasks": tasks, "groups": ledger["groups"], "tracked_days": tracked,
            "coverage": [c for c in ledger["coverage"] if start <= c["day"] <= end]}


# =============================================================================
#  METRICS
# =============================================================================
def summarise(tasks: list) -> dict:
    """KPI block for any set of tasks (overall, one group, one day)."""
    n = len(tasks)
    c = defaultdict(int)
    for t in tasks:
        c[t["status"]] += 1
    completed = c[ST_ON_TIME] + c[ST_LATE] + c[ST_UNKNOWN]
    pending = c[ST_OPEN] + c[ST_MISSED]
    # every task whose timing can be judged: open tasks count as "not (yet) on
    # time", so an evening report never shows 100% timely while work is still open
    timed_den = n - c[ST_UNKNOWN]
    ttcs = [t["ttc_h"] for t in tasks if t["ttc_h"] is not None and t["status"] == ST_ON_TIME]
    comp_pct = _pct(completed, n)
    timely_pct = _pct(c[ST_ON_TIME], timed_den)
    words = [w for w in (_status_word(comp_pct), _status_word(timely_pct)) if w != "—"]
    order = {"Behind": 0, "Watch": 1, "On track": 2}
    overall = min(words, key=lambda w: order[w]) if words else "—"
    return {
        "tasks": n, "completed": completed, "pending": pending,
        "on_time": c[ST_ON_TIME], "late": c[ST_LATE], "unknown": c[ST_UNKNOWN],
        "open": c[ST_OPEN], "missed": c[ST_MISSED],
        "attempted": sum(1 for t in tasks if t["attempted"]),
        "no_action": sum(1 for t in tasks if t["no_action_recorded"]),
        "completion_pct": comp_pct, "pending_pct": _pct(pending, n),
        "timely_pct": timely_pct, "timed_den": timed_den,
        "median_ttc": median(ttcs) if ttcs else None,
        "unique_items": len({(t["group"], t["key"]) for t in tasks}),
        "carried": sum(1 for t in tasks if not t["completed"] and t.get("days_on_list", 0) >= 2),
        "status": overall,
    }


def hourly_progress(tasks: list, days, now: datetime):
    """Within-day progress: cumulative tasks generated (by the time each task
    first appeared ON its report day), cumulative completed (by its completion
    stamp on its report day) and open = generated - completed, per hour (IST).
    `days` = one report day (Daily) or the period's report days (Weekly /
    Monthly / Manual): several days are combined by hour of day, i.e. the sum
    of each day's own hour-by-hour figures — for a single day the result is
    exactly that day's progress. Completions without a valid stamp are not
    placed on the hour axis (reported separately)."""
    days = {days} if isinstance(days, date) else set(days)
    ts = [t for t in tasks if t["day"] in days]
    if not ts:
        return []
    gen_h = [t["first_seen"].hour for t in ts if t["first_seen"] and t["first_seen"].date() == t["day"]]
    done_h = [t["completed_at"].hour for t in ts
              if t["completed_at"] and t["completed_at"].date() == t["day"]]
    marks = gen_h + done_h + ([now.hour] if now.date() in days else [])
    if not marks:
        return []
    first, last = min(marks), max(marks)
    rows, d_cum = [], 0
    g_cum = sum(1 for t in ts if not t["first_seen"] or t["first_seen"].date() < t["day"])
    for h in range(first, last + 1):
        g_cum += sum(1 for x in gen_h if x == h)
        d_new = sum(1 for x in done_h if x == h)
        d_cum += d_new
        rows.append({"label": datetime(2000, 1, 1, h).strftime("%I %p").lstrip("0"),
                     "generated": g_cum, "completed_in_hour": d_new,
                     "completed": d_cum, "open": g_cum - d_cum})
    return rows


def daywise_progress(tasks: list, days: list):
    rows = []
    for d in days:
        ts = [t for t in tasks if t["day"] == d]
        s = summarise(ts)
        rows.append({"day": d, **s})
    return rows


# =============================================================================
#  WORKBOOK  (IntelliBI Coordinator design system, reused from the task report)
# =============================================================================
PCT_FMT = '0.0"%"'
CHART_COLORS = {"on_time": "2E7D32", "late": "F9A825", "unknown": "90A4AE",
                "missed": "C62828", "open": "1565C0", "generated": "1A2E5A",
                "completed": "2E7D32", "remaining": "ED7D31"}

# Dashboard scorecard ROW tint (presentation only), by group status: a lighter
# shade of the Status chip colour — On track = green, Watch = amber, Behind = red,
# No tasks = grey.
ROW_TINT = {"ok": "EDF7F0", "medium": "FFF4E3", "info": "EAF2FC", "high": "FDEBEB",
            "muted": "F3F3F3"}


def _row_tint(level, fallback):
    """Background for a data row of the given status level (fallback = the
    usual zebra shade when the status has no level)."""
    return ROW_TINT.get(level, fallback)


# ── Task-level rows (Task Register): Green / Orange / Red by severity ──────────
# Presentation only — derived from the task's EXISTING status and Days on List:
#   completed  (green)  Completed on time / Completed late / Completed (time not recorded)
#   attention  (orange) Open (due today) — still within its report day
#   urgent     (red)    Missed (report day over, not done), or still open while the
#                       same item has been pending CARRIED_DAYS+ consecutive report
#                       days (the existing "Pending 2+ Days" rule, see summarise())
CARRIED_DAYS = 2                    # = the Pending 2+ Days / Days on List highlight rule
TASK_ROW_COLORS = {                 # row background, chip background, chip / accent text
    "completed": ("E3F2E6", "C8E6C9", "1B5E20"),
    "attention": ("FFEBD2", "FFD49E", "8A4500"),
    "urgent":    ("FBDADA", "F4B4B4", "9B1C1C"),
}


def _task_severity(t) -> str:
    """'completed' | 'attention' | 'urgent' for row colouring (no logic change:
    reads the status and Days on List the ledger already computed)."""
    if t["completed"]:
        return "completed"
    if t["status"] == ST_MISSED or t.get("days_on_list", 0) >= CARRIED_DAYS:
        return "urgent"
    return "attention"


def _hdr(ws, row, headers, col0=1, height=30):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=col0 + i)
        c.value = h
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.fill = AR._fill(BC.DS_NAV2)
        c.alignment = AR._align("center", "center", wrap=True)
        c.border = BC.ds_border()
    ws.row_dimensions[row].height = height
    return row + 1


def _pct_cell(ws, row, col, p, bg=None, bold=False, level_by_status=True):
    if p is None:
        return BC.ds_cell(ws, row, col, "—", bg=bg, h_align="center", fg=BC.DS_MUTED)
    if level_by_status:
        return BC.ds_pill(ws, row, col, round(p, 1), _status_level(_status_word(p)),
                          bold=bold, number_fmt=PCT_FMT)
    return BC.ds_cell(ws, row, col, round(p, 1), bg=bg, h_align="center", bold=bold,
                      number_fmt=PCT_FMT)


def _note(ws, row, ncols, text, italic=True, size=9, height=None, color=None, bg=None):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.font = AR._font(size=size, italic=italic, color=color or BC.DS_MUTED)
    c.alignment = AR._align("left", "center", wrap=True)
    if bg:
        c.fill = AR._fill(bg)
    if height:
        ws.row_dimensions[row].height = height
    else:
        ws.row_dimensions[row].height = max(16, 15 * (1 + len(str(text)) // 170))
    return row + 1


def _kpi_tiles(ws, row, tiles, span=2):
    """A strip of KPI tiles (label / big value / small note), each `span` columns
    wide. tiles: [(label, value, note, level, number_fmt)]."""
    for i, (label, value, note, level, fmt) in enumerate(tiles):
        c0 = 1 + i * span
        c1 = c0 + span - 1
        bg, fg = BC.ds_level_colors(level)
        for r, (val, size, bold, color, h) in enumerate(
                [(label, 9, True, BC.DS_MUTED, 26), (value, 20, True, fg, 34),
                 (note, 8, False, BC.DS_MUTED, 26)]):
            if c1 > c0:
                ws.merge_cells(start_row=row + r, start_column=c0, end_row=row + r, end_column=c1)
            cell = ws.cell(row=row + r, column=c0)
            cell.value = val
            cell.font = AR._font(bold=bold, size=size, color=color)
            cell.alignment = AR._align("center", "center", wrap=True)
            if r == 1 and fmt:
                cell.number_format = fmt
            for cc in range(c0, c1 + 1):
                x = ws.cell(row=row + r, column=cc)
                x.fill = AR._fill(bg if r == 1 else ("FFFFFF" if level == "none" else BC.DS_GUIDE))
                x.border = BC.ds_border()
            ws.row_dimensions[row + r].height = h
    return row + 3


def _bar_chart(ws, title, cats, series, anchor, stacked=True, height=7.5, width=16,
               y_title=None, horizontal=False):
    """series: [(Reference, colour)] with titles from the header cell."""
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "stacked" if stacked else "clustered"
    if stacked:
        ch.overlap = 100
    ch.title = title
    ch.height, ch.width = height, width
    ch.y_axis.title = y_title
    ch.legend.position = "b"
    for ref, color in series:
        ch.add_data(ref, titles_from_data=True)
        s = ch.series[-1]
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = color
    ch.set_categories(cats)
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblSkip = 1
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"      # first group at the top
    ws.add_chart(ch, anchor)
    return ch


def _line_chart(ws, title, cats, series, anchor, height=7.5, width=16, y_title=None):
    ch = LineChart()
    ch.title = title
    ch.height, ch.width = height, width
    ch.y_axis.title = y_title
    ch.legend.position = "b"
    for ref, color in series:
        ch.add_data(ref, titles_from_data=True)
        s = ch.series[-1]
        s.graphicalProperties.line.solidFill = color
        s.graphicalProperties.line.width = 28000
        s.marker.symbol = "circle"
        s.marker.size = 6
        s.marker.graphicalProperties.solidFill = color
        s.marker.graphicalProperties.line.solidFill = color
        s.smooth = False
    ch.set_categories(cats)
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ws.add_chart(ch, anchor)
    return ch


def _group_rows(ledger, tasks):
    """Scorecard rows: every registered group (so an idle group still shows) plus
    any other group that produced tasks."""
    out = []
    for gk, meta in ledger["groups"].items():
        ts = [t for t in tasks if t["group"] == gk]
        if not ts and gk.startswith("tab:"):
            continue
        out.append((gk, meta, summarise(ts)))
    return out


# Dashboard grid: one column per scorecard column; the KPI tiles sit one per
# column on the same grid, so tiles, scorecard and attention list share edges.
DASH_COLS = ["Task Group", "Tasks", "Completed", "Pending", "Completion %", "Pending %", "Status"]
DASH_WIDTHS = [38, 19, 19, 19, 19, 19, 19]


def build_dashboard(ws, ledger, tasks, period_label, is_daily, now):
    NC = len(DASH_COLS)
    total = summarise(tasks)
    BC.ds_title(ws, NC, "Task Performance Dashboard", period_label,
                "Coordinator task completion and timeliness, overall and per task group. "
                "On time = completed on the task's report day (IST). Timely % = On time ÷ "
                "every task with a measurable time (open tasks count as not yet on time). "
                "Status bands: On track ≥ 75%, Watch ≥ 45%, Behind < 45% "
                f"(worse of Completion % and Timely %).   As of {now.strftime('%d-%b-%Y %I:%M %p')} IST.")
    for i, w in enumerate(DASH_WIDTHS, 1):              # widths first: the guide fit uses them
        ws.column_dimensions[_gcl(i)].width = w
    row = 4

    # ── Overall coordinator performance (7 tiles, one per grid column) ──────
    st = total["status"]
    row = BC.ds_section(ws, row, NC, f"OVERALL COORDINATOR PERFORMANCE   ·   {st}", level=1)
    lvl_c = _status_level(_status_word(total["completion_pct"]))
    lvl_t = _status_level(_status_word(total["timely_pct"]))
    tiles = [
        ("Tasks Generated", total["tasks"], f"{total['unique_items']} unique item(s)", "none", None),
        ("Completed", total["completed"],
         f"{total['on_time']} on time · {total['late']} late · {total['unknown']} no time", "ok", None),
        ("Pending", total["pending"], "not completed yet",
         "high" if total["pending"] else "ok", None),
        ("Completion %", round(total["completion_pct"], 1) if total["completion_pct"] is not None else "—",
         f"Pending {total['pending_pct']:.1f}%" if total["pending_pct"] is not None else "",
         lvl_c, PCT_FMT),
        ("Timely Completion %", round(total["timely_pct"], 1) if total["timely_pct"] is not None else "—",
         f"{total['on_time']} of {total['timed_den']} done on their day" if total["timed_den"] else "no timed tasks",
         lvl_t, PCT_FMT),
        ("Median Time to Complete", _fmt_hours(total["median_ttc"]),
         "generation → Done (on-time tasks)", "info", None),
        (("Attempted, Not Done", total["attempted"], "Done? = No / note, not completed",
          "medium" if total["attempted"] else "ok", None) if is_daily else
         ("Pending 2+ Days", total["carried"], "same item, consecutive report days in period",
          "medium" if total["carried"] else "ok", None)),
    ]
    row = _kpi_tiles(ws, row, tiles, span=1) + 1

    # ── Task group scorecard ───────────────────────────────────────────────
    grows = _group_rows(ledger, tasks)
    row = BC.ds_section(ws, row, NC, "TASK GROUP SCORECARD", level=1)
    hdr_row = row
    row = _hdr(ws, row, DASH_COLS, height=30)
    first = row
    for i, (gk, meta, s) in enumerate(grows):
        bg = _row_tint(_status_level(s["status"]) if s["tasks"] else "muted", BC.ds_zebra(i))
        BC.ds_priority(ws, row, 1, meta["name"], _status_level(s["status"]) if s["tasks"] else "muted",
                       h_align="left")
        for col, v in ((2, s["tasks"]), (3, s["completed"]), (4, s["pending"])):
            BC.ds_cell(ws, row, col, v, bg=bg, h_align="center", bold=(col == 4 and v > 0),
                       fg=(BC.DS_HIGH_FG if col == 4 and v > 0 else BC.DS_TEXT))
        _pct_cell(ws, row, 5, s["completion_pct"], bg=bg, bold=True)   # bg only used for "—"
        _pct_cell(ws, row, 6, s["pending_pct"], bg=bg, level_by_status=False)
        BC.ds_pill(ws, row, 7, s["status"] if s["tasks"] else "No tasks",
                   _status_level(s["status"]) if s["tasks"] else "muted")
        ws.row_dimensions[row].height = 20
        row += 1
    last = row - 1
    tb = BC.DS_SUB                                       # total row
    BC.ds_cell(ws, row, 1, "All task groups", bg=tb, bold=True)
    for col, v in ((2, total["tasks"]), (3, total["completed"]), (4, total["pending"])):
        BC.ds_cell(ws, row, col, v, bg=tb, bold=True, h_align="center")
    _pct_cell(ws, row, 5, total["completion_pct"], bold=True)
    _pct_cell(ws, row, 6, total["pending_pct"], bg=tb, bold=True, level_by_status=False)
    BC.ds_pill(ws, row, 7, total["status"], _status_level(total["status"]))
    ws.row_dimensions[row].height = 22
    row += 2                                             # one spacer row

    # ── Chart: completed vs pending per task group (scorecard columns C:D) ──
    if grows and total["tasks"]:
        row = BC.ds_section(ws, row, NC, "COMPLETED VS PENDING BY TASK GROUP", level=2)
        cats = Reference(ws, min_col=1, min_row=first, max_row=last)
        series = [(Reference(ws, min_col=3, min_row=hdr_row, max_row=last), CHART_COLORS["completed"]),
                  (Reference(ws, min_col=4, min_row=hdr_row, max_row=last), CHART_COLORS["remaining"])]
        anchor, row = _chart_below(ws, row, 7.5)          # thin spacer, then the chart
        _bar_chart(ws, "Completed vs Pending — per task group", cats, series, anchor,
                   stacked=True, horizontal=True, height=7.5,
                   width=round(_grid_width_cm(DASH_WIDTHS) - 0.4, 1), y_title="Tasks")
    BC.ds_fit_guide(ws, NC)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_NAV
    try:
        ws.print_area = f"A1:{_gcl(NC)}{row}"
    except Exception:                                    # pragma: no cover
        pass
    return ws


# Progress Trend grid: 10 columns (the day-wise table uses all of them); every
# chart sits directly under its table, within this grid.
TREND_WIDTHS = [26, 14, 13, 14, 12, 11, 11, 13, 11, 12]
TREND_ROW_PT = 15          # fixed height of the rows a chart below a table occupies


def _grid_width_cm(widths):
    """Width of a run of columns in cm (Excel: ~7 px per character + 5 px padding)."""
    return sum(w * 7 + 5 for w in widths) / 37.8


def _chart_below(ws, row, height_cm):
    """Reserve the rows under a table for a chart of `height_cm` (a thin spacer
    row first) and return (anchor cell, next free row after the chart)."""
    ws.row_dimensions[row].height = 6
    n = -(-int(height_cm * 28.35) // TREND_ROW_PT) + 1       # cm → pt → rows, +1 margin
    for r in range(row + 1, row + 1 + n):
        ws.row_dimensions[r].height = TREND_ROW_PT
    return f"A{row + 1}", row + 1 + n


def build_trend(ws, ledger, tasks, period_label, days, is_daily, now, period_text=None):
    """Progress Trend tab.
    1. WITHIN-DAY PROGRESS — <report day / period>: Generated vs Completed vs
       Open hour by hour (table, chart directly underneath). Daily = that day;
       Weekly / Monthly / Manual = the period's report days combined by hour.
    2. (periods only) DAY-WISE PROGRESS and COMPLETION % BY TASK GROUP AND DAY."""
    NC = 10
    BC.ds_title(ws, NC, "Task Progress Trend", period_label,
                ("Within-day pace: tasks generated vs completed (cumulative) and still open, hour by "
                 "hour (IST). " if is_daily else
                 "Within-day pace for the period's report days combined (hour of day, IST), then day "
                 "by day: how each report day's tasks ended — on time, late, missed or open — and "
                 "the completion / timely % trend. ")
                + "Completions without a valid Follow-Up DateTime are counted as completed but "
                  "cannot be placed on a time axis.")
    row = 4
    for i, w in enumerate(TREND_WIDTHS, 1):
        ws.column_dimensions[_gcl(i)].width = w

    # ── 1. within-day progress: hour-by-hour table + chart directly below ──
    when = period_text or (days[-1].strftime("%d-%b-%Y") if days else "")
    hp = hourly_progress(tasks, days, now) if days else []
    row = BC.ds_section(ws, row, NC, f"WITHIN-DAY PROGRESS — {when}" + (
        "" if is_daily else "  (all report days combined, by hour of day)"), level=1)
    if not hp:
        row = BC.ds_empty(ws, row, NC, "No trackable tasks on this report day." if is_daily else
                          "No trackable tasks in this period.", level="muted") + 1
    else:
        h0 = row
        row = _hdr(ws, row, ["Hour (IST)", "Tasks Generated (cumulative)", "Completed in Hour",
                             "Completed (cumulative)", "Open", "Completion %"])
        f0 = row
        for i, r in enumerate(hp):
            bg = BC.ds_zebra(i)
            BC.ds_cell(ws, row, 1, r["label"], bg=bg, h_align="center", bold=True)
            for c, v in ((2, r["generated"]), (3, r["completed_in_hour"]), (4, r["completed"]),
                         (5, r["open"])):
                BC.ds_cell(ws, row, c, v, bg=bg, h_align="center")
            _pct_cell(ws, row, 6, _pct(r["completed"], r["generated"]), bg=bg,
                      level_by_status=False)
            row += 1
        cats = Reference(ws, min_col=1, min_row=f0, max_row=row - 1)
        height = 7.5
        anchor, row = _chart_below(ws, row, height)
        _line_chart(ws, "Generated vs Completed vs Open — hour by hour", cats,
                    [(Reference(ws, min_col=2, min_row=h0, max_row=f0 + len(hp) - 1), CHART_COLORS["generated"]),
                     (Reference(ws, min_col=4, min_row=h0, max_row=f0 + len(hp) - 1), CHART_COLORS["completed"]),
                     (Reference(ws, min_col=5, min_row=h0, max_row=f0 + len(hp) - 1), CHART_COLORS["remaining"])],
                    anchor, y_title="Tasks", height=height,
                    width=round(_grid_width_cm(TREND_WIDTHS) - 0.4, 1))

    # ── 2. periods: day-wise outcome + task-group completion by day ────────
    if not is_daily:
        dp = daywise_progress(tasks, days)
        row = BC.ds_section(ws, row, NC, "DAY-WISE PROGRESS", level=1)
        if not dp:
            row = BC.ds_empty(ws, row, NC, "No trackable report days in this period.", level="muted") + 1
        else:
            h0 = row
            row = _hdr(ws, row, ["Report Day", "Tasks", "On Time", "Late", "Done, No Time",
                                 "Missed", "Open", "Completion %", "Timely %", "Pending 2+ Days"])
            f0 = row
            for i, r in enumerate(dp):
                bg = BC.ds_zebra(i)
                BC.ds_cell(ws, row, 1, r["day"].strftime("%d-%b (%a)"), bg=bg, h_align="center", bold=True)
                for c, v in ((2, r["tasks"]), (3, r["on_time"]), (4, r["late"]), (5, r["unknown"]),
                             (6, r["missed"]), (7, r["open"])):
                    BC.ds_cell(ws, row, c, v, bg=bg, h_align="center")
                _pct_cell(ws, row, 8, r["completion_pct"])
                _pct_cell(ws, row, 9, r["timely_pct"])
                BC.ds_cell(ws, row, 10, r["carried"], bg=bg, h_align="center")
                row += 1
            last = row - 1
            cats = Reference(ws, min_col=1, min_row=f0, max_row=last)
            # both charts directly under the table, side by side on the same grid:
            # outcome bars across columns A–E, the % trend across F–J
            height = 7.5
            anchor, row = _chart_below(ws, row, height)
            split = 5
            _bar_chart(ws, "Each report day's tasks by outcome", cats,
                       [(Reference(ws, min_col=c, min_row=h0, max_row=last), CHART_COLORS[k])
                        for c, k in ((3, "on_time"), (4, "late"), (5, "unknown"), (6, "missed"), (7, "open"))],
                       anchor, stacked=True, y_title="Tasks", height=height,
                       width=round(_grid_width_cm(TREND_WIDTHS[:split]) - 0.3, 1))
            _line_chart(ws, "Completion % and Timely % by report day", cats,
                        [(Reference(ws, min_col=8, min_row=h0, max_row=last), CHART_COLORS["completed"]),
                         (Reference(ws, min_col=9, min_row=h0, max_row=last), CHART_COLORS["open"])],
                        f"{_gcl(split + 1)}{anchor[1:]}", y_title="%", height=height,
                        width=round(_grid_width_cm(TREND_WIDTHS[split:]) - 0.4, 1))

            # task-group completion % by day (which category is slipping)
            row = BC.ds_section(ws, row, NC, "COMPLETION % BY TASK GROUP AND DAY", level=1)
            gks = [gk for gk in ledger["groups"] if any(t["group"] == gk for t in tasks)]
            show_days = days[-9:]
            row = _hdr(ws, row, ["Task Group"] + [d.strftime("%d-%b") for d in show_days])
            for i, gk in enumerate(gks):
                BC.ds_cell(ws, row, 1, ledger["groups"][gk]["short"], bg=BC.ds_zebra(i), bold=True)
                for j, d in enumerate(show_days, 2):
                    s = summarise([t for t in tasks if t["group"] == gk and t["day"] == d])
                    _pct_cell(ws, row, j, s["completion_pct"])
                row += 1
            if len(days) > 9:
                row = _note(ws, row, NC, f"Showing the latest 9 of {len(days)} report days.")

    BC.ds_fit_guide(ws, NC)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_SECTION
    return ws


REG_COLS = ["Report Day", "Task Group", "Item", "Context", "What Was Flagged", "Generated At", "Status",
            "Completed At", "Time to Complete", "Action Taken", "Follow-Up Comment",
            "Days on List", "Seen in Versions", "Recorded in Versions"]


def _task_row(ws, row, t, cols, groups):
    # whole row coloured by severity (green / orange / red); the Status chip is a
    # stronger shade of the same colour and column 1 carries a coloured left edge
    bg, chip_bg, accent = TASK_ROW_COLORS[_task_severity(t)]
    for c, name in enumerate(cols, 1):
        if name == "Status":
            BC.ds_cell(ws, row, c, t["status"], bg=chip_bg, fg=accent, bold=True, h_align="center",
                       left_accent=accent if c == 1 else None)
            continue
        v = {
            "Report Day": t["day"].strftime("%d-%b-%Y"),
            "Task Group": groups[t["group"]]["short"],
            "Item": t["label"], "Context": t.get("context") or "—", "What Was Flagged": t["what"],
            "Generated At": t["first_seen"].strftime("%d-%b %I:%M %p") if t["first_seen"] else "—",
            "Completed At": t["completed_at"].strftime("%d-%b %I:%M %p") if t["completed_at"] else
                            ("time not recorded" if t["completed"] else "—"),
            "Time to Complete": _fmt_hours(t["ttc_h"]),
            "Action Taken": t["action"] or "—", "Follow-Up Comment": t["comment"] or "—",
            "Days on List": t.get("days_on_list", 0) or "—",
            "Seen in Versions": ", ".join(f"V{x}" for x in t["versions_seen"]) or "—",
            "Recorded in Versions": ", ".join(f"V{x}" for x in t["versions_recorded"]) or "—",
        }.get(name, "")
        wrap = name in ("Item", "Context", "What Was Flagged", "Follow-Up Comment")
        bold = name == "Item"
        fg = BC.DS_TEXT
        if name == "Days on List" and isinstance(v, int) and v >= 2:
            fg = BC.DS_HIGH_FG
            bold = True
        BC.ds_cell(ws, row, c, v, bg=bg, wrap=wrap, bold=bold, fg=fg,
                   h_align="left" if wrap or name in ("Task Group",) else "center",
                   left_accent=accent if c == 1 else None)


def build_register(ws, ledger, tasks, period_label):
    NC = len(REG_COLS)
    BC.ds_title(ws, NC, "Task Register (audit trail)", period_label,
                "One row per task (report day × task group × item), merged across every version of "
                "that day's Coordinator report. 'Seen in' = versions listing the task; 'Recorded in' "
                "= versions where the Coordinator entered a follow-up. Pending tasks: filter Status = "
                "Missed / Open (due today); sort Days on List to see items carried across report days; "
                "a pending task with 'Recorded in' filled was attempted but not done.   Row colour:  "
                "■ Green = completed   ■ Orange = attention required (open, due today)   "
                f"■ Red = overdue / urgent (missed, or pending {CARRIED_DAYS}+ consecutive report days).")
    BC.ds_guide_count(ws, "Tasks", len(tasks))
    row = _hdr(ws, 3, REG_COLS)
    order = list(ledger["groups"])
    for t in sorted(tasks, key=lambda t: (t["day"], order.index(t["group"]), t["label"])):
        _task_row(ws, row, t, REG_COLS, ledger["groups"])
        ws.row_dimensions[row].height = 32
        row += 1
    if not tasks:
        BC.ds_empty(ws, row, NC, "No trackable tasks in this period.", level="muted")
    for c, w in enumerate([12, 16, 28, 40, 44, 15, 22, 16, 13, 16, 30, 9, 11, 11], 1):
        ws.column_dimensions[_gcl(c)].width = w
    BC.ds_fit_guide(ws, NC)
    ws.auto_filter.ref = f"A3:{_gcl(NC)}{max(row - 1, 3)}"
    ws.freeze_panes = "D4"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_SECTION
    return ws


RULES = [
    ("Source", "The Batch Coordinator daily reports (every version in <coordinator folder>/YYYY-MM-DD/) "
               "— the only place the Coordinator records Action Taken / Comment / Done? / DateTime."),
    ("Task", "One flagged row on one report day's action list, identified by the task group's stable "
             "fields (learner / assignment / session / record). Rank, row order and figures are ignored."),
    ("Versions", "All versions of a report day are merged into one list: the latest version's tasks, plus "
                 "any task actioned in an earlier version that dropped off later. A task is counted once "
                 "per day, however many versions show it."),
    ("Completed", "Follow-Up Done? = Yes in any version of that day (entries are not carried between "
                  "versions). Completion time = earliest valid Follow-Up DateTime among the Yes entries."),
    ("Pending", "Not completed. 'Open (due today)' while the report day is running; 'Missed' once it is over. "
                "Done? = No, or an Action/Comment without Done? = Yes, is shown as Attempted but stays pending."),
    ("On time / Late", "Each tab is the action list for its report day, so a task is due by 11:59 PM IST "
                       "that day. On time = completed that day; Late = completed on a later day."),
    ("Completion %", "Completed ÷ Tasks.   Pending % = Pending ÷ Tasks."),
    ("Timely %", "On time ÷ every task whose timing can be judged. Open tasks count as not (yet) on "
                 "time; completions without a valid DateTime are excluded (not measurable)."),
    ("Time to Complete", "Follow-Up DateTime − the time the task first appeared in that day's report "
                         "(file creation time). Median over on-time completions."),
    ("Reporting period", "A task belongs to the report day whose Coordinator list it appeared on. Daily = "
                         "that day; Weekly = Mon–Sun; Monthly = the calendar month; Manual = the chosen "
                         "dates. Tasks from before the period are never included, even if still pending."),
    ("Next day", "An item still listed the next day is that day's new task (fresh follow-up columns). "
                 "Days on List counts consecutive report days within the period an item stayed pending."),
    ("Status bands", f"On track ≥ {STATUS_ON_TRACK:.0f}%, Watch ≥ {STATUS_WATCH:.0f}%, Behind below — "
                     "the same bands as the Counsellor Follow-Up Trend; a group's status is the worse of "
                     "its Completion % and Timely %."),
    ("Not trackable", "Reports created before the follow-up columns existed have no place to record "
                      "completion; their tasks are listed here and never counted as pending."),
    ("DateTime caveat", "Follow-Up DateTime stamps the FIRST time Done? is set (Yes or No). A task first set "
                        "to No and later to Yes keeps the earlier time; clearing Done? clears the stamp. "
                        "#REF! stamps (sheets opened before iterative calculation) have no usable time."),
]


def build_coverage(ws, cov_rows, period_label):
    cols = ["Report Day", "Report Versions", "Versions with Follow-Up Columns", "Task Groups Tracked",
            "Task Groups Not Trackable", "Tasks Tracked", "Tasks Not Trackable",
            "Resolved Before Final Run", "Recorded in >1 Version", "Done Without Valid Time",
            "Yes / No Differ Across Versions"]
    NC = len(cols)
    BC.ds_title(ws, NC, "Data Coverage & Rules", period_label,
                "What the figures are built from, and exactly how each measure is defined.")
    row = BC.ds_section(ws, 3, NC, "REPORT DAYS READ", level=1)
    row = _hdr(ws, row, cols, height=40)
    for i, c in enumerate(cov_rows):
        bg = BC.ds_zebra(i)
        vals = [c["day"].strftime("%d-%b-%Y (%a)"), c["versions"], c["versions_fu"],
                c["groups_tracked"], c["groups_untracked"], c["tasks"], c["untracked_tasks"],
                c["dropped"], c["multi_version"], c["yes_no_time"], c["conflicts"]]
        for j, v in enumerate(vals, 1):
            warn = (j in (7, 10, 11) and v) or (j == 3 and not v)
            BC.ds_cell(ws, row, j, v, bg=(BC.DS_MED_BG if warn else bg), h_align="center",
                       bold=(j == 1))
        row += 1
    if not cov_rows:
        row = BC.ds_empty(ws, row, NC, "No Coordinator report days found for this period.", level="muted")
    row += 1
    row = BC.ds_section(ws, row, NC, "RULES & DEFINITIONS", level=1)
    for i, (k, v) in enumerate(RULES):
        BC.ds_cell(ws, row, 1, k, bg=BC.DS_SUB, bold=True, fg=BC.DS_NAV)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=NC)
        BC.ds_cell(ws, row, 2, v, wrap=True, bg=BC.ds_zebra(i))
        for cc in range(3, NC + 1):
            ws.cell(row=row, column=cc).border = BC.ds_border()
        ws.row_dimensions[row].height = 32
        row += 1
    for c, w in enumerate([20, 11, 14, 12, 13, 11, 12, 13, 12, 13, 14], 1):
        ws.column_dimensions[_gcl(c)].width = w
    BC.ds_fit_guide(ws, NC)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_MUTED
    return ws


def build_workbook(ledger, start: date, end: date, period_label: str, is_daily: bool, now: datetime):
    ledger = scope_ledger(ledger, start, end)          # this period's tasks only
    tasks = ledger["tasks"]
    days = sorted({t["day"] for t in tasks})
    cov = ledger["coverage"]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_dashboard(wb.create_sheet("Dashboard"), ledger, tasks, period_label, is_daily, now)
    build_trend(wb.create_sheet("Progress Trend"), ledger, tasks, period_label, days, is_daily, now,
                period_text=(start.strftime("%d-%b-%Y") if start == end else
                             f"{start.strftime('%d-%b-%Y')} to {end.strftime('%d-%b-%Y')}"))
    build_register(wb.create_sheet("Task Register"), ledger, tasks, period_label)
    build_coverage(wb.create_sheet("Data Coverage & Rules"), cov, period_label)
    for ws in wb.worksheets:                          # print: landscape, one page wide
        try:
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
        except Exception:                             # pragma: no cover
            pass
    return wb, tasks, summarise(tasks)


# =============================================================================
#  PLANNING / ORCHESTRATION
# =============================================================================
def _plan_jobs(today: date):
    """(jobs, errors) from the GENERATE_* configuration — shared logic in
    coordinator_periods.plan_jobs (the Batch Coordinator report uses the same
    periods and folders)."""
    return CP.plan_jobs(
        today, auto=GENERATE_AUTO, daily=GENERATE_DAILY, weekly=GENERATE_WEEKLY,
        monthly=GENERATE_MONTHLY, manual=GENERATE_MANUAL, daily_date=DAILY_DATE,
        weekly_reference_date=WEEKLY_REFERENCE_DATE, monthly_month=MONTHLY_MONTH,
        monthly_year=MONTHLY_YEAR, manual_start=MANUAL_START_DATE, manual_end=MANUAL_END_DATE)


def _file_name(job) -> str:
    safe = re.sub(r"[^0-9A-Za-z]+", "_", job["label"]).strip("_")
    return f"{REPORT_BASENAME}_{job['kind']}_{safe}"


def run_jobs(jobs, versions, loader, now, upload=None):
    """Build (and deliver) one workbook per job. The ledger is built once from
    the discovered versions; every job is then scoped to its own period
    (scope_ledger), so jobs never borrow each other's tasks. A failing job is
    recorded and the remaining jobs still run. Each workbook exists only in
    memory and is uploaded from there — nothing is written to local disk."""
    ledger = build_ledger(versions, loader, now)
    log.info("Ledger: %d trackable task(s) across %d report day(s).",
             len(ledger["tasks"]), len({t["day"] for t in ledger["tasks"]}))
    results = []
    for job in jobs:
        name = _file_name(job)
        folders = CP.folder_path(job["kind"], job["start"], job["end"])
        out = {"job": job, "name": name, "folder": "/".join(folders)}
        try:
            wb, tasks, s = build_workbook(ledger, job["start"], job["end"],
                                          f"{job['kind']}  ·  {job['label']}",
                                          job["kind"] == "Daily", now)
            out.update(summary=s, tasks=len(tasks))
            _sc = scope_ledger(ledger, job["start"], job["end"])       # for the e-mail breakdown
            out["groups"] = [(meta["name"], gs) for _gk, meta, gs in _group_rows(_sc, _sc["tasks"])]
            buf = io.BytesIO()                  # in memory only — never written to disk
            wb.save(buf)
            if upload:
                out["link"] = upload(folders, name, buf)
            buf.close()
        except Exception as exc:                                   # noqa: BLE001
            log.exception("Performance report %s %s failed: %s", job["kind"], job["label"], exc)
            out["failed"] = True
            out["error"] = str(exc)
            print(f"[FAILED] Coordinator Task Performance — {job['kind']} {job['label']}: {exc}")
            results.append(out)
            continue
        pct = lambda p: "—" if p is None else f"{p:.1f}%"
        print(f"\n{'=' * 70}\n  Coordinator Task Performance — {job['kind']} {job['label']}\n"
              f"  Period tasks: {s['tasks']} | Completed: {s['completed']} | Pending: {s['pending']} "
              f"(open {s['open']}, missed {s['missed']}) | Completion {pct(s['completion_pct'])} | "
              f"Timely {pct(s['timely_pct'])} | Status {s['status']}\n"
              f"  Folder: {out['folder']}\n"
              + (f"  [Drive] Uploaded: {out['link']}\n" if out.get("link") else
                 "  [Drive] upload OFF — built in memory only, nothing saved\n") + "=" * 70)
        results.append(out)
    return results


def cleanup_legacy_local_copies(root=None) -> int:
    """Remove report copies that earlier versions of this script saved under
    output/reports/coordinator_performance/. Deletes ONLY this report's own
    files (REPORT_BASENAME*.xlsx), then any folder left empty (including the
    root itself); any other file — and the folders holding it — is kept.
    Best-effort: a file that cannot be removed is logged and never fails the run.
    Returns the number of files removed."""
    root = root or LEGACY_OUTPUT_DIR
    if not os.path.isdir(root):
        return 0
    removed = 0
    for cur, _dirs, files in os.walk(root, topdown=False):
        for f in files:
            if f.startswith(REPORT_BASENAME) and f.lower().endswith(".xlsx"):
                try:
                    os.remove(os.path.join(cur, f))
                    removed += 1
                except OSError as exc:
                    log.warning("Could not remove old local report copy %s (%s).",
                                os.path.join(cur, f), exc)
        try:
            if not os.listdir(cur):
                os.rmdir(cur)
        except OSError as exc:
            log.warning("Could not remove empty folder %s (%s).", cur, exc)
    if removed:
        log.info("Removed %d old local report cop%s from %s (reports are kept in "
                 "Google Drive only).", removed, "y" if removed == 1 else "ies", root)
        print(f"[Cleanup] removed {removed} old local report file(s) — reports live in Google Drive only.")
    return removed


def generate():
    """Plan → discover the period's daily task reports → build → deliver.
    Returns (results, errors)."""
    cleanup_legacy_local_copies()       # first, so it happens even if a later step fails
    now = _now_ist()
    jobs, errors = _plan_jobs(now.date())
    for e in errors:
        log.error("Configuration: %s", e)
        print(f"[CONFIG ERROR] {e}")
    if not jobs:
        log.warning("No report selected (check the GENERATE_* flags).")
        return [], errors
    log.info("Planned %d report(s): %s", len(jobs),
             ", ".join(f"{j['kind']} {j['label']}" for j in jobs))
    drive = _drive_service()
    versions = discover_report_versions(drive, [(j["start"], j["end"]) for j in jobs])

    def _upload(folders, name, buf):
        # Same versioned, never-overwrite upload as the Batch Coordinator report,
        # into the same <Type>/<reporting period> folder.
        return BC.upload_report(folders, name + ".xlsx", buf, name)

    results = run_jobs(jobs, versions, lambda v: load_version(drive, v), now,
                       upload=_upload if UPLOAD_TO_DRIVE else None)
    email_results(results)
    return results, errors


def email_results(results) -> bool | None:
    """One e-mail per delivered report (same convention as the Sales lead
    reports). Controlled ONLY by SEND_EMAIL — generation and the Drive upload
    never depend on it. The e-mail carries the Google Sheet link, no attachment. Returns True (all sent), False (one or more
    failed) or None (nothing to send / disabled)."""
    done = [r for r in results if not r.get("failed")]
    if not SEND_EMAIL:
        print("[Email] SEND_EMAIL = False — report(s) generated, no e-mail sent.")
        return None
    if not done:
        print("[Email] no report delivered in this run — nothing to e-mail.")
        return None
    ok = True
    for r in done:
        kind, label, s = r["job"]["kind"], r["job"]["label"], r["summary"]
        subject = f"{kind} Coordinator Task Performance Report - {label}"
        body = CE.performance_html(kind, label, s, r.get("groups", []), r.get("link"),
                                   _fmt_hours(s["median_ttc"]), STATUS_ON_TRACK, STATUS_WATCH,
                                   benchmark=COMPLETION_BENCHMARK,
                                   group_labels=EMAIL_GROUP_LABELS)
        ok = CE.send(subject, body, EMAIL_RECIPIENTS, sender=EMAIL_SENDER,
                     star=STAR_EMAIL_IN_GMAIL) and ok
    return ok


def main(argv=None):
    """Command line (all optional — the flags above are the normal control):
        --no-upload          dry run: build in memory only (nothing written to Drive or disk)
        --no-email           do not send the e-mail (same as SEND_EMAIL = False)
        --date YYYY-MM-DD    Daily report for that date instead of the planned jobs
        --from / --to        a custom period (Manual report)"""
    import argparse
    global UPLOAD_TO_DRIVE, SEND_EMAIL, GENERATE_AUTO, GENERATE_DAILY, GENERATE_WEEKLY, GENERATE_MONTHLY
    global GENERATE_MANUAL, DAILY_DATE, MANUAL_START_DATE, MANUAL_END_DATE
    ap = argparse.ArgumentParser(description="Coordinator Task Performance & Progress Report")
    ap.add_argument("--no-upload", action="store_true")
    ap.add_argument("--no-email", action="store_true")
    ap.add_argument("--date")
    ap.add_argument("--from", dest="date_from")
    ap.add_argument("--to", dest="date_to")
    a = ap.parse_args(argv)
    if a.no_upload:
        UPLOAD_TO_DRIVE = False
    if a.no_email:
        SEND_EMAIL = False
    if a.date or (a.date_from and a.date_to):
        GENERATE_AUTO = GENERATE_WEEKLY = GENERATE_MONTHLY = False
        GENERATE_DAILY = bool(a.date)
        DAILY_DATE = a.date
        GENERATE_MANUAL = bool(a.date_from and a.date_to)
        MANUAL_START_DATE, MANUAL_END_DATE = a.date_from, a.date_to
    logging.basicConfig(level=logging.INFO if VERBOSE else logging.WARNING,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    results, errors = generate()
    failed = [r for r in results if r.get("failed")]
    print(f"\nPerformance reports delivered: {len(results) - len(failed)} | failed: {len(failed)}"
          f" | configuration errors: {len(errors)}")
    return 1 if (failed or errors) else 0


if __name__ == "__main__":
    sys.exit(main())
