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
    AND, next to that effort, the ACTUAL RESULT each group is meant to produce
    (Effort -> Completion -> Outcome; coordinator_outcomes.py): Overall Att %,
    Total Submission %, admission / Wise issues actually resolved at the evening
    re-check, sessions without escalation, Overall Interview Attendance % —
    each from its own report's calculation, with targets, Effort -> Outcome
    quadrants, the change vs the previous period and the areas that need
    management attention.

  SOURCE OF TRUTH  (read-only — nothing in the existing system is changed)
    The Batch Coordinator daily reports produced by
    pyCoordinatorTaskListReport.py are the ONLY place Coordinator
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
    Tabs: Dashboard | Effort vs Outcome Trend | Progress Trend | Task Register |
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
import pyCoordinatorTaskListReport as BC
import coordinator_periods as CP              # shared periods + Drive layout
import coordinator_email as CE                # shared e-mail (Operations Gmail account)
import coordinator_outcomes as CO             # actual performance (outcome) measures

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
EMAIL_RECIPIENTS = ["info@intellibiinnovationstechnologies.in"]#,
                   # "intellibihropsb2ch@gmail.com"]
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
    "Wise & Interview Feedback Validation":  "Wise & Interview Feedback",
    "Instructor Follow-Ups":                 "Instructor Instructions",
    "Learner Instructor Interview Reminder": "Interview Reminder",
}

# =============================================================================
#  ACTUAL PERFORMANCE (OUTCOME)  — coordinator_outcomes.py
#  Effort (Task Completion %) is shown next to the RESULT it was meant to
#  produce. Each area's Actual Performance % reuses its own report's figure
#  (see coordinator_outcomes.py and the report's "Data Coverage & Rules" tab).
# =============================================================================
CHECK_OUTCOMES = True        # False = effort only (no source reads, no evening re-check)

# Target per area — Actual Performance % at/above it = "On target". Attendance
# uses the Attendance report's own green band (Overall Att % >= 75).
OUTCOME_TARGETS = {
    "attendance": 75.0,      # Overall Att %
    "assignment": 80.0,      # Total Submission %
    "admission":  75.0,      # admission issues resolved by the evening check
    "wise":       75.0,      # Wise / interview-feedback issues resolved
    "instructor": 90.0,      # held sessions without an escalation
    "interview":  80.0,      # Overall Interview Attendance %
}
OUTCOME_NEAR_BAND = 15.0     # within this many points below target = "Near target"
EFFORT_HIGH_PCT   = 75.0     # Task Completion % at/above = "High effort" (= On track band)
TREND_DAYS_DAILY  = 7        # Daily report: the Effort vs Outcome trend shows the last N days

# Previous vs current period (Dashboard scorecard, chart and e-mail) for the
# groups whose outcome window starts BEFORE the day's list is worked, so the
# result that follows a day's follow-ups is the NEXT day's figure (Daily:
# Yesterday → Today; Weekly / Monthly / Manual: previous period → report period).
# Every figure is the one the report already computes for each period.
COMPARE_GROUPS      = ["attendance", "assignment", "instructor"]
COMPARE_STEADY_BAND = 1.0    # |outcome change| below this many points = "Steady"

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
GENERATE_AUTO    = True

GENERATE_DAILY   = True
GENERATE_WEEKLY  = False
GENERATE_MONTHLY = False
GENERATE_MANUAL  = False

DAILY_DATE            = "2026-10-02" # "YYYY-MM-DD" None
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
# Evening re-check store (Admission / Wise): the evening state of a day cannot be
# re-created later, so each delivered Daily report saves its checked items here
# (one <YYYY-MM-DD>.json per report day) and Weekly / Monthly reports add them up.
# Not shown in any report tab. Keep this folder — deleting it makes earlier days
# "not checked" in later Weekly / Monthly reports.
OUTCOME_CHECKS_DIR = os.path.join(_PROJECT, "cache", "coordinator_outcome_checks")

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
        # the row as the Coordinator saw it (display text, follow-up columns
        # excluded) + its section banner: the morning baseline of the outcome
        # re-check (coordinator_outcomes)
        "cells": {k: _display(v) for k, v in cells.items()
                  if _norm(k) not in {_norm(x) for x in BC.FOLLOWUP_COLS}},
        "section": _display(rec["_ctx"].get("Section", "")),
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


def _daily_folder_files(drive, ranges):
    """(report day, folder name, file) for every native Sheet in the Daily report-
    day folders whose day falls inside one of `ranges`:
      * <root>/Daily Coordinator Reports/Daily DD-Mon-YYYY/   (current layout)
      * <root>/YYYY-MM-DD/                                     (legacy layout)"""
    root_children = _children(drive, BC.PARENT_FOLDER_ID, folders_only=True)
    day_folders = [f for f in root_children if CP._LEGACY_DAY_FOLDER.match(f.get("name", ""))]
    for f in root_children:
        if f.get("name") == CP.KIND_FOLDERS["Daily"]:
            day_folders += [x for x in _children(drive, f["id"], folders_only=True)
                            if CP._DAILY_FOLDER.match(x.get("name", ""))]
    for f in day_folders:
        day = CP.daily_folder_date(f["name"])
        if day is None or not _in_any(day, ranges):
            continue
        for x in _children(drive, f["id"]):
            if x.get("mimeType") == "application/vnd.google-apps.spreadsheet":
                yield day, f["name"], x


def discover_report_versions(drive, ranges) -> list:
    """Every native-Sheet version of the DAILY Coordinator task report whose
    report day falls inside one of `ranges` [(start, end), …] (task origin = the
    report day; folders: _daily_folder_files).
    Only files named like the daily task report are used — the current
    "IntelliBI_Coordinator_Task_List_Report_Daily…" and the earlier
    "IntelliBI_Batch_Coordinator_Daily_Attendance_Report…" (BC.TASK_LIST_BASENAMES) —
    Weekly / Monthly / Manual roll-ups have no follow-up columns and are not task
    lists, and performance reports are never read as task lists."""
    versions = []
    for day, folder, x in _daily_folder_files(drive, ranges):
        if not x.get("name", "").startswith(tuple(getattr(BC, "TASK_LIST_BASENAMES",
                                                         (BC.REPORT_BASENAME,)))):
            continue
        m = _VERSION.search(x["name"])
        versions.append({"id": x["id"], "name": x["name"], "day": day,
                         "version": int(m.group(1)) if m else 1,
                         "created": _drive_time_ist(x.get("createdTime", "")),
                         "modified_raw": x.get("modifiedTime", ""),
                         "folder": folder})
    versions.sort(key=lambda v: (v["day"], v["created"] or datetime.min, v["version"]))
    log.info("Found %d Coordinator report version(s) across %d report day(s).",
             len(versions), len({v["day"] for v in versions}))
    return versions


def _cache_path(v):
    safe = re.sub(r"[^0-9A-Za-z]", "", v["modified_raw"])
    # _v2: occurrences carry the row cells + section (outcome re-check baseline)
    return os.path.join(CACHE_DIR, f"{v['id']}_{safe}_v2.json")


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
                    "cells": ref.get("cells", {}), "section": ref.get("section", ""),
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


def within_day_gaps(tasks: list, days):
    """What of the report day(s)' task activity is NOT on the within-day hour
    axis, and why — so the Progress Trend explains a short or empty timeline
    instead of showing a blank table:
      generated_after   tasks first listed only AFTER their report day ended
                        (the day's task list was produced later, e.g. a re-run
                        for a past date) — there was nothing to act on during
                        the day, so they cannot be placed on its timeline
      completed_after   completions stamped after the report day (late)
      completed_no_time completed without a valid Follow-Up DateTime"""
    days = {days} if isinstance(days, date) else set(days)
    ts = [t for t in tasks if t["day"] in days]
    after = [t for t in ts if t["first_seen"] and t["first_seen"].date() > t["day"]]
    return {"tasks": len(ts),
            "generated_after": len(after),
            "first_generated_after": min((t["first_seen"] for t in after), default=None),
            "completed_after": sum(1 for t in ts
                                   if t["completed_at"] and t["completed_at"].date() > t["day"]),
            "completed_no_time": sum(1 for t in ts if t["completed"] and not t["completed_at"])}


def _gap_text(g, is_daily, when, empty):
    """Plain-language explanation of within_day_gaps() for the Progress Trend."""
    first = g["first_generated_after"]
    first_txt = f" (first listed {first.strftime('%d-%b-%Y %I:%M %p')} IST)" if first else ""
    parts = []
    if empty:
        if is_daily:
            parts.append(f"No within-day progress can be shown for {when}: the task list for this day "
                         f"was generated only after the day had ended{first_txt}, so none of its "
                         f"{g['tasks']} task(s) was available to the Coordinator during {when}.")
        else:
            parts.append(f"No within-day progress can be shown: all {g['tasks']} task(s) of this period "
                         f"were generated only after their report day had ended{first_txt}.")
        parts.append("The tasks are still counted in the Dashboard and the Task Register "
                     "(due at the end of the report day, so they show as Missed or Completed late).")
    elif g["generated_after"]:
        parts.append(f"{g['generated_after']} task(s) were generated only after their report day "
                     f"had ended{first_txt} and are not part of the hour-by-hour figures above.")
    if g["completed_after"] and not empty:
        parts.append(f"{g['completed_after']} completion(s) were recorded after the report day "
                     f"(counted as late, not shown on the day's hours).")
    if g["completed_no_time"]:
        parts.append(f"{g['completed_no_time']} completion(s) have no valid Follow-Up DateTime "
                     f"(counted as completed, not placed on the hours).")
    return " ".join(parts)


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


KPI_LABEL_ROW_PT, KPI_VALUE_ROW_PT, KPI_GAP_ROW_PT = 24, 44, 8


def _kpi_cards(ws, row, cards, groups, per_row=4):
    """Uniform KPI cards: a label strip over a large value, each card spanning one
    column group of equal width (`groups` = [(first col, last col)], one per card
    position in a row), `per_row` cards per row with a thin gap row between rows.
    cards: [(label, value, level, number_fmt)]. No footer / description row.
    Returns the next free row."""
    from openpyxl.styles import Border, Side
    edge = Side(style="thin", color="C9D3E0")
    for k in range(0, len(cards), per_row):
        for (label, value, level, fmt), (c0, c1) in zip(cards[k:k + per_row], groups):
            bg, fg = BC.ds_level_colors(level)
            value_bg = "FFFFFF" if level == "none" else bg
            for r, (val, size, bold, color, fill) in enumerate(
                    [(label, 9, True, BC.DS_MUTED, BC.DS_GUIDE), (value, 22, True, fg, value_bg)]):
                rr = row + r
                if c1 > c0:
                    ws.merge_cells(start_row=rr, start_column=c0, end_row=rr, end_column=c1)
                cell = ws.cell(row=rr, column=c0)
                cell.value = val
                cell.font = AR._font(bold=bold, size=size, color=color)
                cell.alignment = AR._align("center", "center", wrap=False)
                if r == 1 and fmt:
                    cell.number_format = fmt
                for cc in range(c0, c1 + 1):          # one outlined card: label + value
                    x = ws.cell(row=rr, column=cc)
                    x.fill = AR._fill(fill)
                    x.border = Border(left=edge if cc == c0 else None,
                                      right=edge if cc == c1 else None,
                                      top=edge if r == 0 else None,
                                      bottom=edge if r == 1 else Side(style="hair", color="DCE3EC"))
            ws.row_dimensions[row].height = KPI_LABEL_ROW_PT
            ws.row_dimensions[row + 1].height = KPI_VALUE_ROW_PT
        row += 2
        if k + per_row < len(cards):
            ws.row_dimensions[row].height = KPI_GAP_ROW_PT   # breathing space between card rows
            row += 1
    return row


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


# =============================================================================
#  EFFORT -> OUTCOME  (periods, outcome view, scorecard, headline, attention)
# =============================================================================
def previous_period(kind: str, start: date, end: date):
    """The period just before (Daily: the day before; Weekly: the week before;
    Monthly: the month before; Manual: the same number of days before)."""
    if kind == "Daily":
        d = start - timedelta(days=1)
        return d, d
    if kind == "Weekly":
        return start - timedelta(days=7), end - timedelta(days=7)
    if kind == "Monthly":
        pe = start - timedelta(days=1)
        return pe.replace(day=1), pe
    n = (end - start).days + 1
    return start - timedelta(days=n), start - timedelta(days=1)


def trend_days(kind: str, start: date, end: date, today: date) -> list:
    """Daily: the last TREND_DAYS_DAILY days ending on the report day; Weekly /
    Monthly / Manual: every day of the period up to today."""
    if kind == "Daily":
        return [end - timedelta(days=i) for i in range(TREND_DAYS_DAILY - 1, -1, -1)]
    last = min(end, today)
    return [start + timedelta(days=i) for i in range((last - start).days + 1)]


def job_ranges(jobs, today: date) -> list:
    """Report days each job needs: its period, the previous period (Δ) and the
    trend days."""
    out = []
    for j in jobs:
        ps, _pe = previous_period(j["kind"], j["start"], j["end"])
        td = trend_days(j["kind"], j["start"], j["end"], today)
        out.append((min([ps, j["start"]] + td[:1]), j["end"]))
    return out


def _effort_by_group(sc) -> dict:
    out = {gk: summarise([t for t in sc["tasks"] if t["group"] == gk]) for gk in sc["groups"]}
    out["_all"] = summarise(sc["tasks"])
    return out


def compute_outcome_view(engine, ledger, job, now) -> dict:
    """Everything the Effort -> Outcome sheets and e-mail need for one job:
    cur / prev outcomes per group, previous-period effort, the day-by-day trend
    and (Daily) the evening-check detail rows. engine=None -> effort only."""
    kind, start, end = job["kind"], job["start"], job["end"]
    ps, pe = previous_period(kind, start, end)
    view = {"enabled": engine is not None, "kind": kind, "cur": {}, "prev": {},
            "prev_label": (ps.strftime("%d-%b-%Y") if ps == pe else
                           f"{ps.strftime('%d-%b-%Y')} – {pe.strftime('%d-%b-%Y')}"),
            "prev_eff": {}, "trend": [], "details": [],
            "compare": compare_labels(kind, start, end, ps, pe, now), "now": now}
    psc = scope_ledger(ledger, ps, pe)
    view["prev_eff"] = _effort_by_group(psc)
    days_with_list = lambda sc: sorted({c["day"] for c in sc["coverage"] if c["versions"]})
    if engine is not None:
        sc = scope_ledger(ledger, start, end)
        view["cur"] = engine.period_outcomes(kind, start, end, sc["tasks"], days_with_list(sc))
        view["prev"] = engine.period_outcomes(kind, ps, pe, psc["tasks"], days_with_list(psc))
        if kind == "Daily":
            try:
                view["details"] = engine.detail_rows(end, sc["tasks"])
                view["details_live"] = (getattr(engine, "_day_details", {}).get(end) or {}
                                        ).get("source") == "live"
            except Exception as exc:                                  # noqa: BLE001
                log.exception("Outcome detail %s failed: %s", end, exc)
    for d in trend_days(kind, start, end, now.date()):
        dsc = scope_ledger(ledger, d, d)
        point = {"day": d, "eff": _effort_by_group(dsc), "act": {}}
        if engine is not None:
            point["act"] = engine.period_outcomes("Daily", d, d, dsc["tasks"], days_with_list(dsc))
        view["trend"].append(point)
    return view


def _delta(a, b):
    return None if a is None or b is None else round(a - b, 1)


# ── Previous vs current period (COMPARE_GROUPS) ──────────────────────────────
def _range_text(a: date, b: date) -> str:
    return (a.strftime("%a %d-%b-%Y") if a == b else
            f"{a.strftime('%d-%b')} – {b.strftime('%d-%b-%Y')}")


def compare_labels(kind, start, end, ps, pe, now) -> dict:
    """Names and dates of the two compared periods, by their ACTUAL reporting
    dates (never the generation time): Daily = 'Yesterday' / 'Today' when the
    report day is the run day (else 'Previous Day' / 'Report Day'); Weekly /
    Monthly / Manual = previous period / report period. 'observation' says how
    the outcome relates to the effort: next_day (Daily — the newer outcome window
    starts after the older day's list was raised) or same_period."""
    today = now.date()
    if kind == "Daily":
        tags = (("Yesterday", "Today") if end == today and ps == today - timedelta(days=1)
                else ("Previous Day", "Report Day"))
    elif kind == "Weekly":
        tags = ("Previous Week", "Report Week")
    elif kind == "Monthly":
        tags = ("Previous Month", "Report Month")
    else:
        tags = ("Previous Period", "Report Period")
    if kind == "Monthly":
        prev_dates, cur_dates = ps.strftime("%b-%Y"), start.strftime("%b-%Y")
    else:
        prev_dates, cur_dates = _range_text(ps, pe), _range_text(start, end)
    return {"prev_tag": tags[0], "cur_tag": tags[1], "prev_dates": prev_dates,
            "cur_dates": cur_dates, "prev_start": ps, "prev_end": pe,
            "cur_start": start, "cur_end": end,
            "as_of": now.strftime("%I:%M %p") if end >= today else "",
            "observation": "next_day" if kind == "Daily" else "same_period"}


def outcome_availability(o, kind, end, today) -> tuple:
    """(status word, level, reason) of one period's outcome figure, so a figure
    that is not (yet) final is never shown as if it were."""
    if o is None:
        return "Not measured", "muted", ""
    if o["state"] == CO.ST_NOT_CHECKED:
        return "Not available", "high", o.get("note") or "source could not be read"
    if o["state"] == CO.ST_NO_DATA:
        return "Nothing to measure", "muted", o.get("note") or ""
    if o.get("source") == "provisional":
        return "Provisional", "medium", o.get("note") or ""
    if kind != "Daily" and end >= today:
        return "In progress", "info", f"period not finished — measured up to {today:%d-%b}"
    return "Final", "ok", ""


def compare_row(gk, s, o, prev_s, prev_o, view, now) -> dict:
    """Previous vs current period for one COMPARE_GROUPS group. Reuses the
    scorecard's own figures: effort = summarise() of each period, outcome =
    the period outcomes of compute_outcome_view (view['prev'] / view['cur'])."""
    cl = view.get("compare") or {}
    kind = view.get("kind")
    today = now.date()
    prev_comp = prev_s["completion_pct"] if prev_s and prev_s["tasks"] else None
    cur_comp = s["completion_pct"] if s and s["tasks"] else None
    prev_act = prev_o["pct"] if prev_o else None
    cur_act = o["pct"] if o else None
    change = _delta(cur_act, prev_act)
    # the effort that PRECEDED the newer outcome: Daily = the previous day's
    # follow-ups (next-day observation); period reports = the period's own
    effort_ref = prev_comp if cl.get("observation") == "next_day" else cur_comp
    verdict = CO.followup_result(effort_ref, change, COMPARE_STEADY_BAND, EFFORT_HIGH_PCT)
    prev_av = outcome_availability(prev_o, kind, cl.get("prev_end", today), today)
    cur_av = outcome_availability(o, kind, cl.get("cur_end", today), today)
    why = ""
    if change is None:
        side = cl.get("cur_tag", "current") if cur_act is None else cl.get("prev_tag", "previous")
        av = cur_av if cur_act is None else prev_av
        why = f"{side}'s outcome: {av[0]}" + (f" — {av[2]}" if av[2] else "")
    return {"prev_s": prev_s, "prev_o": prev_o, "prev_comp": prev_comp, "cur_comp": cur_comp,
            "prev_actual": prev_act, "cur_actual": cur_act, "change": change,
            "d_completion": _delta(cur_comp, prev_comp), "effort_ref": effort_ref,
            "verdict": verdict, "verdict_level": CO.FR_LEVEL.get(verdict, "muted"), "why": why,
            "prev_av": prev_av, "cur_av": cur_av,
            "prev_window": ((prev_o or {}).get("window") or {}).get("text", ""),
            "cur_window": ((o or {}).get("window") or {}).get("text", "")}


def scorecard_rows(ledger, tasks, view=None) -> list:
    """One row per task group: effort (summarise) + outcome + verdicts."""
    view = view or {}
    rows = []
    for gk, meta, s in _group_rows(ledger, tasks):
        o = (view.get("cur") or {}).get(gk)
        po = (view.get("prev") or {}).get(gk)
        pe = (view.get("prev_eff") or {}).get(gk)
        target = OUTCOME_TARGETS.get(gk)
        comp = s["completion_pct"] if s["tasks"] else None
        act = o["pct"] if o else None
        rows.append({
            "gk": gk, "name": meta["name"], "short": meta.get("short", meta["name"]), "s": s,
            "o": o, "target": target, "completion": comp, "actual": act,
            "measure": o["measure"] if o else (CO.MEASURES[gk][0] if gk in CO.MEASURES else "—"),
            "basis": CO.basis_text(o) if o else ("not measured" if gk in CO.MEASURES else "—"),
            "impact": (CO.outcome_status(act, target, OUTCOME_NEAR_BAND) if target is not None
                       else "Not measured"),
            "quadrant": (CO.quadrant(comp, act, target, EFFORT_HIGH_PCT) if target is not None
                         else CO.QUAD_NA),
            "d_actual": _delta(act, po["pct"] if po else None),
            "d_completion": _delta(comp, pe["completion_pct"] if pe and pe["tasks"] else None),
            # previous vs current period (COMPARE_GROUPS, outcome view only)
            "compare": (compare_row(gk, s, o, pe, po, view, view.get("now") or _now_ist())
                        if view.get("enabled") and gk in COMPARE_GROUPS else None),
        })
    return rows


IMPACT_LEVEL = {"On target": "ok", "Near target": "medium", "Below target": "high",
                "Not measured": "muted"}


def _avg(vals):
    vals = [v for v in vals if v is not None]
    return round(sum(vals) / len(vals), 1) if vals else None


def outcome_totals(rows) -> dict:
    measured = [r for r in rows if r["actual"] is not None]
    return {"measured": len(measured),
            "on_target": sum(1 for r in measured if r["impact"] == "On target"),
            "avg_actual": _avg(r["actual"] for r in measured),
            "attention": attention_items(rows)}


def headline(total, rows) -> str:
    """'Coordinator completed X% of required actions — what actual result?' answered."""
    if total["tasks"]:
        eff = (f"Coordinator completed {total['completion_pct']:.1f}% of required actions "
               f"({total['completed']} of {total['tasks']})")
    else:
        eff = "No Coordinator actions were required in this period"
    ot = outcome_totals(rows)
    if not ot["measured"]:
        return eff + " — actual result not measured."
    worst = min((r for r in rows if r["actual"] is not None),
                key=lambda r: (r["actual"] - (r["target"] or 0)))
    res = (f"actual result: {ot['on_target']} of {ot['measured']} areas on target "
           f"(average {ot['avg_actual']:.1f}%)")
    if worst["impact"] != "On target":
        res += (f"; weakest: {worst['short']} {worst['actual']:.1f}% "
                f"vs {worst['target']:g}% target")
    return f"{eff} — {res}."


def attention_items(rows) -> list:
    """[(group short name, level, why, action)] — the areas a manager should act on."""
    out = []
    for r in rows:
        q, s, a, t = r["quadrant"], r["s"], r["actual"], r["target"]
        comp = r["completion"]
        if q == CO.QUAD_ATTENTION:
            out.append((r["short"], "high",
                        f"Only {comp:.1f}% of {s['tasks']} task(s) completed and the result is "
                        f"{a:.1f}% against a {t:g}% target.",
                        "Make sure the list is worked in full today; review why tasks are left open."))
        elif q == CO.QUAD_NOT_CONVERTING:
            out.append((r["short"], "medium",
                        f"{comp:.1f}% of tasks completed, but the result is only {a:.1f}% "
                        f"(target {t:g}%) — {r['basis']}.",
                        "Follow-ups are not changing the outcome: review the approach or escalate."))
        elif q == CO.QUAD_NO_TASKS_LOW:
            out.append((r["short"], "medium",
                        f"Result {a:.1f}% is below the {t:g}% target, but no Coordinator task was raised.",
                        "Check whether the task list should flag these cases."))
        elif q == CO.QUAD_CHECK_TASKS:
            out.append((r["short"], "info",
                        f"Result on target ({a:.1f}%) with only {comp:.1f}% of tasks completed.",
                        "Check whether these tasks are needed or can be simplified."))
        if r["d_actual"] is not None and r["d_actual"] <= -10 and q not in (CO.QUAD_ATTENTION,):
            out.append((r["short"], "medium",
                        f"Result fell {abs(r['d_actual']):.1f} points vs the previous period.",
                        "Look at what changed this period."))
        o = r["o"]
        if o and o["state"] == CO.ST_NOT_CHECKED:
            out.append((r["short"], "muted", f"Actual result not measured: {o['note']}.",
                        "Check the source / the evening refresh (see Data Coverage & Rules)."))
    # one entry per area: most severe level first, its reasons and actions combined
    order = {"high": 0, "medium": 1, "info": 2, "muted": 3}
    merged = OrderedDict()
    for area, lev, why, act in out:
        m = merged.setdefault(area, [lev, [], []])
        if order[lev] < order[m[0]]:
            m[0] = lev
        m[1].append(why)
        if act not in m[2]:
            m[2].append(act)
    return sorted(((a, m[0], "  ".join(m[1]), "  ".join(m[2])) for a, m in merged.items()),
                  key=lambda x: order[x[1]])


def _fmt_p(p):
    return "—" if p is None else f"{p:.1f}%"


# Dashboard grid: the scorecard's 11 columns — the effort scorecard (Task Group …
# Status) followed by the actual-outcome columns. The 8 KPI cards sit on the same
# grid as 2 rows of 4 equal-width cards (DASH_CARD_GROUPS).
DASH_COLS = ["Task Group", "Tasks", "Completed", "Pending", "Completion %", "Pending %", "Status",
             "Actual Performance %", "Target", "Performance / Impact", "Effort → Outcome"]
DASH_WIDTHS = [32, 15, 15, 16, 16, 15, 15, 17, 11, 17, 19]
# KPI cards: 2 rows x 4 cards, each card one of these equal-width column groups
# (A-B, C-E, F-H, I-K = 47 characters each).
DASH_CARD_GROUPS = [(1, 2), (3, 5), (6, 8), (9, 11)]
DASH_EFFORT_COLS = (2, 7)          # super-header "EFFORT" over these columns
DASH_OUTCOME_COLS = (8, 11)        # super-header "ACTUAL OUTCOME" over these columns
CHART_EFFORT_HEX, CHART_OUTCOME_HEX = "2F5597", "ED7D31"   # blue / orange (colour-blind safe)


def _effort_outcome_chart(ws, anchor, first, last, n_groups, width_cm):
    """Clustered horizontal bars per task group: Effort (Task Completion %) vs
    Actual Performance %, read straight from the scorecard's own columns (no
    separate data table). 0–100% axis, value labels, legend on top."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.title = "Effort vs Actual Outcome by Task Group"
    ch.style = 10
    ch.height = max(7.5, 2.2 + 1.35 * n_groups)
    ch.width = width_cm
    ch.gapWidth = 55
    ch.overlap = -8
    for col, label, colour in ((5, "Effort · Task Completion %", CHART_EFFORT_HEX),
                               (8, "Actual Performance %", CHART_OUTCOME_HEX)):
        ch.add_data(Reference(ws, min_col=col, min_row=first, max_row=last), titles_from_data=False)
        s = ch.series[-1]
        s.tx = SeriesLabel(v=label)
        s.graphicalProperties.solidFill = colour
        s.graphicalProperties.line.solidFill = colour
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        s.dLbls.numFmt = '0.0"%"'
        s.dLbls.position = "outEnd"
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ch.x_axis.scaling.orientation = "maxMin"           # first task group at the top
    ch.x_axis.delete = False
    ch.x_axis.tickLblSkip = 1
    ch.y_axis.delete = False
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 110                        # room for the value labels
    ch.y_axis.majorUnit = 25
    ch.y_axis.number_format = '0"%"'
    ch.y_axis.title = "% (0–100)"
    ch.y_axis.crosses = "max"                          # value axis at the bottom (reversed categories)
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E8EF"))
    ch.legend.position = "t"
    ws.add_chart(ch, anchor)
    return ch


# ── Previous vs current period block (Dashboard scorecard) ────────────────────
CMP_COLS = ["Task Group  ·  Outcome Measure", "Tasks Done", "Effort · Completion %",
            "Outcome · Actual %", "Tasks Done", "Effort · Completion %", "Outcome · Actual %",
            "Outcome Change", "Effort Change", "Target", "Follow-up → Result"]
CMP_PREV_HEX, CMP_CHANGE_HEX = "5B6B86", "7A4A12"     # super-header fills (previous / change)
CHART_PREV_EFFORT_HEX, CHART_PREV_OUTCOME_HEX = "A9C0E4", "F6C69E"   # lighter tints = previous
CHANGE_FMT = '"▲ "0.0" pts";"▼ "0.0" pts";"● 0.0 pts"'
AVAIL_FMT = {"Provisional": '0.0"% · prov."', "In progress": '0.0"% · to date"'}


def _cmp_outcome_cell(ws, row, col, act, av, target):
    """An outcome figure of the comparison; a figure that is not available is
    shown as its status (never as a number), provisional / in-progress figures
    carry that tag in the cell."""
    from openpyxl.comments import Comment
    word, lvl, why = av
    if act is None:
        c = BC.ds_pill(ws, row, col, word, lvl if lvl != "ok" else "muted")
        c.alignment = AR._align("center", "center", wrap=True)
    else:
        lvl2 = (IMPACT_LEVEL[CO.outcome_status(act, target, OUTCOME_NEAR_BAND)]
                if target is not None else "info")
        c = BC.ds_pill(ws, row, col, round(act, 1), lvl2, number_fmt=AVAIL_FMT.get(word, PCT_FMT))
    if why:
        c.comment = Comment(f"{word}: {why}", "IntelliBI")
    return c


def _cmp_change_cell(ws, row, col, v, bg):
    if v is None:
        return BC.ds_cell(ws, row, col, "—", bg=bg, h_align="center", fg=BC.DS_MUTED)
    lvl = "ok" if v >= COMPARE_STEADY_BAND else ("high" if v <= -COMPARE_STEADY_BAND else "info")
    return BC.ds_pill(ws, row, col, v, lvl, number_fmt=CHANGE_FMT)


def compare_reading_note(cl) -> str:
    p, c = cl["prev_tag"], cl["cur_tag"]
    if cl.get("observation") == "next_day":
        return (f"How to read: {p}'s effort is followed by {c}'s outcome. {c}'s outcome window starts "
                f"after {p.lower()}'s task list was raised, so it is the first result observed after "
                f"{p.lower()}'s follow-ups (next-day observation). {c}'s own tasks are raised FROM "
                f"{c.lower()}'s outcome — their effect shows in the next Daily report. Outcome Change = "
                f"{c} − {p} (points; within ±{COMPARE_STEADY_BAND:g} = Steady). Follow-up → Result judges "
                f"{p.lower()}'s Completion % (high ≥ {EFFORT_HIGH_PCT:g}%) against that change. A change "
                f"is an observation, not proof that the follow-ups caused it.")
    return (f"How to read: same-period comparison — each period's outcome is measured over that "
            f"period's own sessions / deadlines, next to that period's follow-ups. Outcome Change = "
            f"{c} − {p} (points; within ±{COMPARE_STEADY_BAND:g} = Steady). Follow-up → Result judges "
            f"the {c.lower()}'s Completion % (high ≥ {EFFORT_HIGH_PCT:g}%) against that change. A change "
            f"is an observation, not proof that the follow-ups caused it.")


def build_compare_block(ws, row, NC, rows, view):
    """Previous vs current period for COMPARE_GROUPS, on the scorecard's grid:
    Task Group | PREVIOUS (tasks done, effort %, outcome %) | CURRENT (same) |
    Outcome Change | Effort Change | Target | Follow-up → Result, then one line
    per group with the window each outcome measures and a reading note.
    Returns (next free row, first data row, last data row) — (row, None, None)
    when there is nothing to compare."""
    crs = [r for r in rows if r.get("compare")]
    cl = (view or {}).get("compare")
    if not crs or not cl:
        return row, None, None
    row = BC.ds_section(ws, row, NC, f"{cl['prev_tag'].upper()}  →  {cl['cur_tag'].upper()}   ·   "
                        "FOLLOW-UP GROUPS: COORDINATOR EFFORT, THEN THE RESULT THAT FOLLOWED", level=2)
    cur_head = f"{cl['cur_tag'].upper()}  ·  {cl['cur_dates']}" + (
        f"  ·  as of {cl['as_of']}" if cl.get("as_of") else "")
    for (c0, c1), text, fill in (((1, 1), "", BC.DS_NAV2),
                                 ((2, 4), f"{cl['prev_tag'].upper()}  ·  {cl['prev_dates']}", CMP_PREV_HEX),
                                 ((5, 7), cur_head, BC.DS_NAV),
                                 ((8, 11), "CHANGE  ·  FOLLOW-UP → RESULT", CMP_CHANGE_HEX)):
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        for cc in range(c0, c1 + 1):
            x = ws.cell(row=row, column=cc)
            x.fill = AR._fill(fill)
            x.border = BC.ds_border()
        c = ws.cell(row=row, column=c0)
        c.value = text
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.alignment = AR._align("center", "center", wrap=True)
    ws.row_dimensions[row].height = 20
    row += 1
    row = _hdr(ws, row, CMP_COLS, height=30)
    first = row
    from openpyxl.comments import Comment
    for i, r in enumerate(crs):
        c = r["compare"]
        bg = BC.ds_zebra(i)
        BC.ds_cell(ws, row, 1, f"{r['name']}\n{r['measure']}", bg=bg, bold=True, wrap=True, size=9)
        for col0, s, comp in ((2, c["prev_s"], c["prev_comp"]), (5, r["s"], c["cur_comp"])):
            if s and s["tasks"]:
                BC.ds_cell(ws, row, col0, f"{s['completed']} / {s['tasks']}", bg=bg, h_align="center",
                           bold=True)
            else:
                BC.ds_cell(ws, row, col0, "No tasks", bg=bg, h_align="center", fg=BC.DS_MUTED,
                           italic=True)
            _pct_cell(ws, row, col0 + 1, comp, bg=bg, bold=True)
        _cmp_outcome_cell(ws, row, 4, c["prev_actual"], c["prev_av"], r["target"])
        _cmp_outcome_cell(ws, row, 7, c["cur_actual"], c["cur_av"], r["target"])
        _cmp_change_cell(ws, row, 8, c["change"], bg)
        if c["d_completion"] is None:
            BC.ds_cell(ws, row, 9, "—", bg=bg, h_align="center", fg=BC.DS_MUTED)
        else:
            BC.ds_cell(ws, row, 9, c["d_completion"], bg=bg, h_align="center", fg=BC.DS_MUTED,
                       number_fmt=CHANGE_FMT)
        BC.ds_cell(ws, row, 10, r["target"] if r["target"] is not None else "—", bg=bg,
                   h_align="center", number_fmt='0"%"', fg=BC.DS_MUTED)
        v = BC.ds_pill(ws, row, 11, c["verdict"], c["verdict_level"])
        v.alignment = AR._align("center", "center", wrap=True)
        if c["why"]:
            v.comment = Comment(c["why"], "IntelliBI")
        ws.row_dimensions[row].height = 34
        row += 1
    last = row - 1
    for r in crs:                                        # what each outcome measures
        c = r["compare"]
        text = (f"{r['short']} — {r['measure']} measured over:   {cl['prev_tag']}: "
                f"{c['prev_window'] or '—'}   →   {cl['cur_tag']}: {c['cur_window'] or '—'}"
                + (f"   ·   {c['why']}" if c["why"] else ""))
        row = _note(ws, row, NC, text, italic=False, size=9, height=17)
    row = _note(ws, row, NC, compare_reading_note(cl), italic=True, size=9,
                height=15 * max(2, -(-len(compare_reading_note(cl)) // 190)))
    return row, first, last


def _compare_chart(ws, anchor, first, last, cl, width_cm, height_cm):
    """Previous vs current period, per follow-up group: four bars in reading
    order — previous effort, previous outcome, current effort, current outcome.
    Previous = lighter tints of the same blue (effort) / orange (outcome).
    Data = the Dashboard's own Yesterday → Today block (columns C, D, F, G; a
    status text such as "Not available" in place of a figure draws no bar)."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.title = f"Follow-up groups · {cl['prev_tag']} vs {cl['cur_tag']}"
    ch.style = 10
    ch.height, ch.width = height_cm, width_cm
    ch.gapWidth = 60
    ch.overlap = -5
    for col, label, colour in (
            (3, f"{cl['prev_tag']} · Effort", CHART_PREV_EFFORT_HEX),
            (4, f"{cl['prev_tag']} · Outcome", CHART_PREV_OUTCOME_HEX),
            (6, f"{cl['cur_tag']} · Effort", CHART_EFFORT_HEX),
            (7, f"{cl['cur_tag']} · Outcome", CHART_OUTCOME_HEX)):
        ch.add_data(Reference(ws, min_col=col, min_row=first, max_row=last), titles_from_data=False)
        s = ch.series[-1]
        s.tx = SeriesLabel(v=label)
        s.graphicalProperties.solidFill = colour
        s.graphicalProperties.line.solidFill = colour
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        s.dLbls.numFmt = '0.0"%"'
        s.dLbls.position = "outEnd"
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = False
    ch.x_axis.tickLblSkip = 1
    ch.y_axis.delete = False
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 110
    ch.y_axis.majorUnit = 25
    ch.y_axis.number_format = '0"%"'
    ch.y_axis.crosses = "max"
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E8EF"))
    ch.legend.position = "t"
    ws.add_chart(ch, anchor)
    return ch


def build_dashboard(ws, ledger, tasks, period_label, is_daily, now, view=None):
    NC = len(DASH_COLS)
    total = summarise(tasks)
    rows = scorecard_rows(ledger, tasks, view)
    ot = outcome_totals(rows)
    BC.ds_title(ws, NC, "Task Performance Dashboard", period_label,
                "Coordinator task completion and timeliness (effort), next to the actual result in each "
                "area (outcome). On time = completed on the task's report day (IST). Timely % = On time ÷ "
                "every task with a measurable time (open tasks count as not yet on time). Status bands: "
                "On track ≥ 75%, Watch ≥ 45%, Behind < 45% (worse of Completion % and Timely %). "
                f"Effort → Outcome: high effort = Completion % ≥ {EFFORT_HIGH_PCT:g}%, good outcome = "
                f"Actual ≥ target.   As of {now.strftime('%d-%b-%Y %I:%M %p')} IST.")
    for i, w in enumerate(DASH_WIDTHS, 1):              # widths first: the guide fit uses them
        ws.column_dimensions[_gcl(i)].width = w
    row = 4
    # one-line headline: completed X% of required actions — what actual result?
    row = _note(ws, row, NC, headline(total, rows), italic=False, size=11, height=24,
                color=BC.DS_NAV, bg=BC.DS_SUB) + 1

    # ── Overall coordinator performance (7 effort tiles + Average Actual %) ──
    st = total["status"]
    row = BC.ds_section(ws, row, NC, f"OVERALL COORDINATOR PERFORMANCE   ·   {st}", level=1)
    lvl_c = _status_level(_status_word(total["completion_pct"]))
    lvl_t = _status_level(_status_word(total["timely_pct"]))
    avg = ot["avg_actual"]
    cards = [
        ("Tasks Generated", total["tasks"], "none", None),
        ("Completed", total["completed"], "ok", None),
        ("Pending", total["pending"], "high" if total["pending"] else "ok", None),
        ("Completion %", round(total["completion_pct"], 1) if total["completion_pct"] is not None else "—",
         lvl_c, PCT_FMT),
        ("Timely Completion %", round(total["timely_pct"], 1) if total["timely_pct"] is not None else "—",
         lvl_t, PCT_FMT),
        ("Median Time to Complete", _fmt_hours(total["median_ttc"]), "info", None),
        (("Attempted, Not Done", total["attempted"], "medium" if total["attempted"] else "ok", None)
         if is_daily else
         ("Pending 2+ Days", total["carried"], "medium" if total["carried"] else "ok", None)),
        ("Average Actual %", avg if avg is not None else "—",
         _status_level(_status_word(avg)) if avg is not None else "muted", PCT_FMT),
    ]
    row = _kpi_cards(ws, row, cards, DASH_CARD_GROUPS) + 1

    # ── Task group scorecard: effort columns, then the actual-outcome columns ──
    row = BC.ds_section(ws, row, NC, "TASK GROUP SCORECARD", level=1)
    for (c0, c1), text, fill in ((DASH_EFFORT_COLS, "EFFORT  ·  Coordinator task completion", BC.DS_NAV),
                                 (DASH_OUTCOME_COLS, "ACTUAL OUTCOME  ·  result in the area", "7A4A12")):
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        for cc in range(c0, c1 + 1):
            x = ws.cell(row=row, column=cc)
            x.fill = AR._fill(fill)
            x.border = BC.ds_border()
        c = ws.cell(row=row, column=c0)
        c.value = text
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.alignment = AR._align("center", "center")
    ws.row_dimensions[row].height = 18
    row += 1
    row = _hdr(ws, row, DASH_COLS, height=32)
    # Groups shown in the Yesterday → Today block below are not repeated here
    # (presentation only: totals, KPI cards, headline and e-mail still use every group).
    in_compare = bool((view or {}).get("compare")) and any(r.get("compare") for r in rows)
    card_rows = [r for r in rows if not (in_compare and r.get("compare"))]
    first = row
    for i, r in enumerate(card_rows):
        s, o = r["s"], r["o"]
        lvl = _status_level(s["status"]) if s["tasks"] else "muted"
        bg = _row_tint(lvl, BC.ds_zebra(i))
        BC.ds_priority(ws, row, 1, r["name"], lvl, h_align="left")
        for col, v in ((2, s["tasks"]), (3, s["completed"]), (4, s["pending"])):
            BC.ds_cell(ws, row, col, v, bg=bg, h_align="center", bold=(col == 4 and v > 0),
                       fg=(BC.DS_HIGH_FG if col == 4 and v > 0 else BC.DS_TEXT))
        _pct_cell(ws, row, 5, s["completion_pct"], bg=bg, bold=True)   # bg only used for "—"
        _pct_cell(ws, row, 6, s["pending_pct"], bg=bg, level_by_status=False)
        BC.ds_pill(ws, row, 7, s["status"] if s["tasks"] else "No tasks", lvl)
        # actual outcome
        if r["actual"] is None:
            c = BC.ds_cell(ws, row, 8, "—", bg=bg, h_align="center", fg=BC.DS_MUTED)
        else:
            c = BC.ds_pill(ws, row, 8, round(r["actual"], 1), IMPACT_LEVEL[r["impact"]],
                           number_fmt=PCT_FMT)
        basis = r["basis"] + (f" · {o['note']}" if o and o.get("note") else "")
        if basis:                                        # measure + basis as a cell note
            from openpyxl.comments import Comment
            c.comment = Comment(f"{r['measure']}: {basis}", "IntelliBI")
        BC.ds_cell(ws, row, 9, r["target"] if r["target"] is not None else "—", bg=bg,
                   h_align="center", number_fmt='0"%"', fg=BC.DS_MUTED)
        BC.ds_pill(ws, row, 10, r["impact"], IMPACT_LEVEL[r["impact"]])
        BC.ds_pill(ws, row, 11, r["quadrant"], CO.QUAD_LEVEL.get(r["quadrant"], "muted")
                   ).alignment = AR._align("center", "center", wrap=True)
        ws.row_dimensions[row].height = 28          # room for a two-line Effort → Outcome chip
        row += 1
    last = row - 1
    tb = BC.DS_SUB                                       # total row
    BC.ds_cell(ws, row, 1, "All task groups", bg=tb, bold=True)
    for col, v in ((2, total["tasks"]), (3, total["completed"]), (4, total["pending"])):
        BC.ds_cell(ws, row, col, v, bg=tb, bold=True, h_align="center")
    _pct_cell(ws, row, 5, total["completion_pct"], bold=True)
    _pct_cell(ws, row, 6, total["pending_pct"], bg=tb, bold=True, level_by_status=False)
    BC.ds_pill(ws, row, 7, total["status"], _status_level(total["status"]))
    if avg is None:
        BC.ds_cell(ws, row, 8, "—", bg=tb, h_align="center", fg=BC.DS_MUTED)
    else:
        _pct_cell(ws, row, 8, avg, bold=True)
    BC.ds_cell(ws, row, 9, "", bg=tb)
    BC.ds_cell(ws, row, 10, f"{ot['on_target']} of {ot['measured']} on target" if ot["measured"]
               else "not measured", bg=tb, bold=True, size=9, h_align="center")
    n_cmp = len(rows) - len(card_rows)
    BC.ds_cell(ws, row, 11, (f"Incl. the {n_cmp} follow-up groups below · Actual % = simple average"
                             if n_cmp else "Actual % = simple average"),
               bg=tb, size=8, h_align="center", fg=BC.DS_MUTED, italic=True, wrap=True)
    ws.row_dimensions[row].height = 30 if n_cmp else 22
    row += 2                                             # one spacer row

    # ── Previous vs current period for the follow-up groups (same grid) ───────
    row, _cf, _cl = build_compare_block(ws, row, NC, rows, view)
    if _cf is not None:
        row += 1

    # ── Charts: Effort vs Actual Outcome per task group; beside it, the
    #    follow-up groups previous vs current (data: Effort vs Outcome Trend tab)
    if rows and (total["tasks"] or ot["measured"]):
        row = BC.ds_section(ws, row, NC, "EFFORT VS OUTCOME BY TASK GROUP", level=2)
        side = _cf is not None                           # the Yesterday → Today block was drawn
        n_card = len(card_rows)                          # the scorecard's own groups (charted left)
        h = max(7.5, 2.2 + 1.35 * n_card) if n_card else 7.5
        if side:
            h = max(h, 2.6 + 2.4 * (_cl - _cf + 1))
        anchor, row = _chart_below(ws, row, h)
        split = 5                                        # A–E | F–K: two equal halves
        if n_card:
            _effort_outcome_chart(ws, anchor, first, last, n_card,
                                  round(_grid_width_cm(DASH_WIDTHS[:split] if side else DASH_WIDTHS)
                                        - 0.4, 1))
            if side:
                ws._charts[-1].height = h
                ws._charts[-1].title = (f"Other task groups · {view['compare']['cur_tag']}" if n_cmp
                                        else f"All task groups · {view['compare']['cur_tag']}")
        if side:
            _compare_chart(ws, f"{_gcl(split + 1)}{anchor[1:]}" if n_card else anchor,
                           _cf, _cl, view["compare"],
                           round(_grid_width_cm(DASH_WIDTHS[split:] if n_card else DASH_WIDTHS)
                                 - 0.4, 1), h)
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


# =============================================================================
#  EFFORT VS OUTCOME TREND  (Is performance improving?)
# =============================================================================
OT_COLS = ["Day", "Tasks Generated", "Tasks Completed", "Task Completion %", "Areas Measured",
           "Average Actual %"]
OT_WIDTHS = [22, 16, 16, 18, 16, 18]


def _trend_grid(ws, row, title, headers, data_rows, fmt_cols, ncols):
    row = BC.ds_section(ws, row, ncols, title, level=1)
    hdr = row
    row = _hdr(ws, row, headers, height=30)
    first = row
    for i, vals in enumerate(data_rows):
        bg = BC.ds_zebra(i)
        for j, v in enumerate(vals, 1):
            BC.ds_cell(ws, row, j, v if v is not None else None, bg=bg,
                       h_align="left" if j == 1 else "center", bold=(j == 1),
                       number_fmt=(PCT_FMT if j in fmt_cols else None))
        row += 1
    return hdr, first, row - 1, row


def _effort_outcome_trend_chart(ws, anchor, first, last, width_cm, height_cm):
    """Two lines over the days: Task Completion % (effort) vs Average Actual %
    (outcome) — the same blue / orange as the Dashboard chart, markers and value
    labels, fixed 0–100% scale, legend on top, gaps where a day has no value."""
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import SeriesLabel
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    ch = LineChart()
    ch.title = "Task Completion % vs Average Actual % — Day by Day"
    ch.style = 12
    ch.height, ch.width = height_cm, width_cm
    ch.display_blanks = "gap"
    for col, label, colour, pos in ((4, "Effort · Task Completion %", CHART_EFFORT_HEX, "t"),
                                    (6, "Outcome · Average Actual %", CHART_OUTCOME_HEX, "b")):
        ch.add_data(Reference(ws, min_col=col, min_row=first, max_row=last), titles_from_data=False)
        s = ch.series[-1]
        s.tx = SeriesLabel(v=label)
        s.smooth = False
        s.graphicalProperties.line.solidFill = colour
        s.graphicalProperties.line.width = 32000
        s.marker.symbol = "circle"
        s.marker.size = 7
        s.marker.graphicalProperties.solidFill = colour
        s.marker.graphicalProperties.line.solidFill = "FFFFFF"
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.showSerName = s.dLbls.showCatName = s.dLbls.showLegendKey = False
        s.dLbls.numFmt = '0.0"%"'
        s.dLbls.position = pos                     # effort labels above, outcome below the point
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ch.x_axis.delete = False
    ch.x_axis.tickLblSkip = 1
    ch.x_axis.title = "Report day"
    ch.y_axis.delete = False
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 110                    # headroom for the labels above 100%
    ch.y_axis.majorUnit = 25
    ch.y_axis.number_format = '0"%"'
    ch.y_axis.title = "%"
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E8EF"))
    ch.legend.position = "t"
    ws.add_chart(ch, anchor)
    return ch


def build_outcome_trend(ws, view, period_label, groups=None, rows=None):
    """OVERALL — DAY BY DAY: per report day, Coordinator effort (tasks generated /
    completed, Task Completion %) next to the outcome (areas measured, Average
    Actual %), and one line chart comparing the two over time. Each day uses that
    day's Daily definitions (view["trend"], compute_outcome_view)."""
    NC = len(OT_COLS)
    BC.ds_title(ws, NC, "Effort vs Outcome Trend", period_label,
                "Coordinator effort vs actual outcome, day by day. Each day uses that day's Daily "
                "definition. Average Actual % = simple average of the areas measured that day. "
                "Blank = nothing measured that day.")
    for i, w in enumerate(OT_WIDTHS, 1):
        ws.column_dimensions[_gcl(i)].width = w
    row = 4
    trend = (view or {}).get("trend") or []
    if not trend:
        row = BC.ds_empty(ws, row, NC, "No days to show.", level="muted")
    else:
        data = []
        for p in trend:
            e = p["eff"]["_all"]
            acts = [o["pct"] for o in p["act"].values() if o and o["pct"] is not None]
            data.append([p["day"].strftime("%d-%b (%a)"), e["tasks"], e["completed"],
                         round(e["completion_pct"], 1) if e["completion_pct"] is not None else None,
                         len(acts), _avg(acts)])
        _hdr_row, first, last, row = _trend_grid(ws, row, "OVERALL — DAY BY DAY", OT_COLS, data,
                                                 {4, 6}, NC)
        if len(trend) > 1:
            row = BC.ds_section(ws, row + 1, NC,
                                "TASK COMPLETION % VS AVERAGE ACTUAL % — DAY BY DAY", level=2)
            height = 9.5
            anchor, row = _chart_below(ws, row, height)
            _effort_outcome_trend_chart(ws, anchor, first, last,
                                        round(_grid_width_cm(OT_WIDTHS) - 0.4, 1), height)
    BC.ds_fit_guide(ws, NC)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_SECTION
    try:
        ws.print_area = f"A1:{_gcl(NC)}{row}"
    except Exception:                                    # pragma: no cover
        pass
    return ws


# =============================================================================
#  OUTCOME DETAIL  (Daily: every item re-checked in the evening — stored here,
#  read back by the Weekly / Monthly reports)
# =============================================================================
def save_outcome_checks(day, details, folder=None):
    """Store one report day's evening checks (Admission / Wise / Instructor items)
    as <folder>/<YYYY-MM-DD>.json — read back by later Weekly / Monthly reports.
    Written atomically; a failure is logged and never fails the report."""
    folder = folder or OUTCOME_CHECKS_DIR
    path = os.path.join(folder, f"{day.isoformat()}.json")
    try:
        os.makedirs(folder, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump([CO.detail_row_values(d) for d in details], fh, ensure_ascii=False)
        os.replace(tmp, path)
        return path
    except Exception as exc:                                     # noqa: BLE001
        log.warning("Could not store the evening checks of %s (%s).", day, exc)
        return None


def load_outcome_checks(day, folder=None):
    """The stored evening checks of a report day ([details]), or None."""
    path = os.path.join(folder or OUTCOME_CHECKS_DIR, f"{day.isoformat()}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return [d for d in (CO.detail_from_values(v) for v in json.load(fh)) if d]
    except Exception as exc:                                     # noqa: BLE001
        log.warning("Stored evening checks of %s unreadable (%s).", day, exc)
        return None


def parse_outcome_detail(wb):
    """LEGACY: the "Outcome Detail" tab that Daily reports carried until
    07-Oct-2026 -> [details] (None when the report has no such tab). Only read
    for days whose checks are not in OUTCOME_CHECKS_DIR."""
    if CO.DETAIL_TAB not in wb.sheetnames:
        return None
    ws = wb[CO.DETAIL_TAB]
    hdr_row = None
    for r in range(1, min(ws.max_row or 0, 10) + 1):
        if str(ws.cell(row=r, column=1).value or "").strip() == CO.DETAIL_COLS[0]:
            hdr_row = r
            break
    if hdr_row is None:
        return None
    out = []
    for vals in ws.iter_rows(min_row=hdr_row + 1, max_col=len(CO.DETAIL_COLS), values_only=True):
        if not vals or vals[0] is None:
            continue
        v0 = vals[0]
        if isinstance(v0, datetime):
            vals = (v0.strftime("%Y-%m-%d"),) + tuple(vals[1:])
        d = CO.detail_from_values(["" if v is None else v for v in vals])
        if d:
            out.append(d)
    return out


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


def _trend_note(ws, row, ncols, text, level="muted"):
    """A full-width, wrapped explanatory line (Progress Trend)."""
    bg, fg = BC.ds_level_colors(level)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    BC.ds_cell(ws, row, 1, text, bg=bg, fg=fg, wrap=True, italic=True)
    for cc in range(2, ncols + 1):
        ws.cell(row=row, column=cc).border = BC.ds_border()
    ws.row_dimensions[row].height = 15 * max(2, -(-len(text) // 150))
    return row + 1


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
    gaps = within_day_gaps(tasks, days) if days else {"tasks": 0}
    if not hp and not gaps["tasks"]:
        row = BC.ds_empty(ws, row, NC, "No trackable tasks on this report day." if is_daily else
                          "No trackable tasks in this period.", level="muted") + 1
    elif not hp:                               # tasks exist, but none on the day's timeline
        row = _trend_note(ws, row, NC, _gap_text(gaps, is_daily, when, empty=True), level="medium") + 1
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
        _gt = _gap_text(gaps, is_daily, when, empty=False)
        if _gt:
            row = _trend_note(ws, row, NC, _gt) + 1

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
    ("Source", "The Daily Coordinator Task List reports (formerly Batch Coordinator; every version in "
               "<coordinator folder>/Daily Coordinator Reports/Daily DD-Mon-YYYY/) "
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


OUTCOME_RULES = [
    ("Effort → Outcome", "Effort = Task Completion % (the Coordinator's own follow-ups). Outcome = Actual "
                         "Performance % — the result the work is meant to produce, from each area's own "
                         "report. The two are shown side by side; neither is derived from the other."),
    ("Attendance", "Overall Att % of the Attendance report's Session Summary (the same function). Daily = "
                   "yesterday 12:00 PM → the time of the day's Coordinator task list (daily formula); "
                   "Weekly / Monthly = sessions dated in the period (de-duplicated period formula). "
                   "Suspended learners and attendance-not-required learners excluded, as in the report."),
    ("Assignment", "Total Submission % READ from the corresponding generated Assignment Submission "
                   "Performance report (its Summary tab). Daily report of day D = the Assignment Daily report "
                   "generated that morning (deadlines of D-1); Weekly / Monthly / Manual = the Assignment "
                   "report of the same period. If that report has not been generated yet (e.g. the Monthly "
                   "Assignment report is made on the 1st of the next month) or cannot be read, the figure is "
                   "calculated now with the Assignment report's own functions for the same period and marked "
                   "PROVISIONAL in the Basis column."),
    ("Admission", "Resolution % = learners on the day's Admission Formalities list whose admission form is "
                  "SIGNED at the evening re-check ÷ learners still applicable. Signatures are refreshed "
                  "in the evening (pyZohoSignatureStatusRefresh). Follow-Up Done? = Yes does not count."),
    ("Wise & IV Feedback", "Resolution % = validation issues on the day's list fixed at the evening re-check ÷ "
                           "issues still applicable — issue level (each failing field / failed check / "
                           "missing interview feedback is one issue). Wise data refreshed in the evening "
                           "with the same validation rules; interview feedback re-checked in the "
                           "Interview Consolidate Sheet."),
    ("Instructor", "Sessions Without Escalation % = held sessions of the day's window (yesterday 12:00 PM → "
                   "task-list time; scheduled-only sessions excluded) that are NOT on the Instructor "
                   "Follow-Ups list ÷ held sessions. A session counts once however many reasons flagged "
                   "it; the reasons are counted separately."),
    ("Interview", "Overall Interview Attendance % of the IntelliBI Interview Consolidated Report (interviewed ÷ "
                  "scheduled; attended = not absent/skipped AND scored or published). Daily = yesterday → "
                  "today; Weekly = Monday → today; Monthly = the 1st → today (or the period end)."),
    ("Not applicable / not checked", "A learner / record no longer active or listed is 'No longer applicable' "
                                     "and leaves the denominator. An item whose source could not be read "
                                     "is 'Not checked' — excluded and reported, never guessed."),
    ("Stored checks", "The evening state cannot be re-created later, so each delivered Daily report stores its "
                      "checks on the reporting PC (cache/coordinator_outcome_checks); Weekly / Monthly reports "
                      "add up the stored checks of their days "
                      "(a day without a stored check is reported as not checked)."),
    ("Targets", "  ·  ".join(f"{CO.MEASURES[k][0]} {v:g}%" for k, v in OUTCOME_TARGETS.items())
                + f".  Near target = within {OUTCOME_NEAR_BAND:g} points."),
    ("Quadrants", f"High effort = Task Completion % ≥ {EFFORT_HIGH_PCT:g}%; good outcome = Actual ≥ target.  "
                  f"High·Good = {CO.QUAD_PAYING}; High·Low = {CO.QUAD_NOT_CONVERTING}; "
                  f"Low·Good = {CO.QUAD_CHECK_TASKS}; Low·Low = {CO.QUAD_ATTENTION}."),
    ("Previous vs current", "Attendance, Assignment and Instructor: the outcome window starts BEFORE the "
                            "day's list is worked (sessions from yesterday 12:00 PM; assignments due the "
                            "day before) and the day's tasks are raised FROM it, so the result that follows "
                            "a day's follow-ups is the NEXT day's figure. Daily reports therefore show "
                            "Yesterday's effort and outcome next to Today's, each labelled with its own "
                            "dates and measurement window (next-day observation). Weekly / Monthly / Manual "
                            "show the previous period next to the report period (same-period comparison). "
                            "Outcome Change = current − previous Actual % (the same Δ the report already "
                            f"uses); within ±{COMPARE_STEADY_BAND:g} point = Steady. Follow-up → Result: "
                            f"Completion % of the follow-ups before the newer outcome (Daily: yesterday's; "
                            f"period reports: the period's) high (≥ {EFFORT_HIGH_PCT:g}%) or not, against "
                            f"that change — '{CO.FR_IMPROVED_AFTER}', '{CO.FR_DECLINED_DESPITE}', "
                            f"'{CO.FR_DECLINED_INCOMPLETE}', '{CO.FR_STEADY}', …  An outcome that is not "
                            "available yet is shown as its status (Not available / Provisional / In "
                            "progress), never as a number. A change is an observation, not proof that "
                            "the follow-ups caused it."),
]


# ── Data Coverage & Rules: layout ─────────────────────────────────────────────
# The rules above, grouped for reading (texts unchanged). A rule key not listed
# here still appears, under "Other rules".
RULE_SECTIONS = [
    ("1", "DATA SOURCE & WHAT COUNTS AS A TASK", "info",
     ["Source", "Task", "Versions", "Not trackable"]),
    ("2", "REPORTING PERIODS", "info",
     ["Reporting period", "Next day"]),
    ("3", "EFFORT — TASK COMPLETION RULES", "info",
     ["Completed", "Pending", "On time / Late", "Completion %", "Timely %", "Time to Complete",
      "Status bands"]),
    ("4", "ACTUAL PERFORMANCE — HOW EACH AREA IS MEASURED", "info",
     ["Effort → Outcome", "Attendance", "Assignment", "Admission", "Wise & IV Feedback",
      "Instructor", "Interview"]),
    ("5", "TARGETS & EFFORT → OUTCOME VERDICTS", "info",
     ["Targets", "Quadrants", "Previous vs current"]),
    ("6", "IMPORTANT LIMITATIONS & CONDITIONS", "medium",
     ["Not applicable / not checked", "Stored checks", "DateTime caveat"]),
]
COV_WIDTHS = [24, 13, 13, 12, 12, 13, 13, 13, 13, 15, 16]
# report-day table: column groups (super-header) and short headers
COV_GROUPS = [((1, 1), "REPORT DAY"), ((2, 3), "VERSIONS READ"), ((4, 7), "TASK GROUPS & TASKS"),
              ((8, 11), "DATA-QUALITY CHECKS")]
COV_COLS = ["Report Day", "Report Versions", "Versions with Follow-Up Columns", "Task Groups Tracked",
            "Task Groups Not Trackable", "Tasks Tracked", "Tasks Not Trackable",
            "Resolved Before Final Run", "Recorded in >1 Version", "Done Without Valid Time",
            "Yes / No Differ Across Versions"]


def _merged_text(ws, row, c0, c1, value, bg="FFFFFF", bold=False, fg=None, size=10,
                 h_align="left", italic=False):
    """A wrapped text cell merged over c0..c1, every cell bordered."""
    if c1 > c0:
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
    c = BC.ds_cell(ws, row, c0, value, bg=bg, bold=bold, fg=fg or BC.DS_TEXT, size=size,
                   wrap=True, h_align=h_align, v_align="center", italic=italic)
    for cc in range(c0 + 1, c1 + 1):
        x = ws.cell(row=row, column=cc)
        x.border = BC.ds_border()
        x.fill = AR._fill(bg)
    return c


def _text_height(text, width_chars, size=10, min_pt=20):
    """Row height (pt) for wrapped text in a merged cell `width_chars` wide."""
    per_line = max(10, int(width_chars * (1.3 if size <= 9 else 1.2)))
    lines = sum(max(1, -(-len(part) // per_line)) for part in str(text or "").split("\n"))
    return max(min_pt, 14.5 * lines + 7)


def _coverage_cards(ws, row, cov_rows, NC):
    """At-a-glance strip: four equal cards summarising the report days read
    (totals of the table below — no new figures)."""
    days = len(cov_rows)
    versions = sum(c["versions"] for c in cov_rows)
    tracked = sum(c["tasks"] for c in cov_rows)
    notes = sum(c["untracked_tasks"] + c["yes_no_time"] + c["conflicts"] for c in cov_rows)
    cards = [("Report Days Read", days, "none", None),
             ("Report Versions Read", versions, "none", None),
             ("Tasks Tracked", tracked, "ok" if tracked else "muted", None),
             ("Data Notes (amber cells)", notes, "medium" if notes else "ok", None)]
    groups = [(1, 2), (3, 5), (6, 8), (9, 11)]
    return _kpi_cards(ws, row, cards, groups)


def _outcome_sources_table(ws, row, NC, rows, refresh_note):
    """Where each Actual Performance % of this period comes from."""
    # Area | Measure | Actual % | Target | Basis | Source | Note   (merged over the grid)
    spans = [(1, 1), (2, 3), (4, 4), (5, 5), (6, 7), (8, 9), (10, 11)]
    heads = ["Area", "Measure", "Actual %", "Target", "Basis", "Source", "Note"]
    for (c0, c1), h in zip(spans, heads):
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        for cc in range(c0, c1 + 1):
            x = ws.cell(row=row, column=cc)
            x.fill = AR._fill(BC.DS_NAV2)
            x.border = BC.ds_border()
        c = ws.cell(row=row, column=c0)
        c.value = h
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.alignment = AR._align("center", "center", wrap=True)
    ws.row_dimensions[row].height = 24
    row += 1
    width = lambda c0, c1: sum(COV_WIDTHS[c0 - 1:c1])
    for i, r in enumerate(r for r in rows if r["gk"] in CO.MEASURES):
        o = r["o"]
        bg = BC.ds_zebra(i)
        note = (o or {}).get("note", "") if o else "outcomes not checked in this run"
        if o and o.get("breakdown"):
            note = (note + "  ·  " if note else "") + ", ".join(f"{k}: {v}" for k, v in o["breakdown"])
        vals = [r["short"], r["measure"],
                round(r["actual"], 1) if r["actual"] is not None else "—",
                r["target"], r["basis"], CO.MEASURES[r["gk"]][2], note]
        h = 22
        for (c0, c1), v, k in zip(spans, vals, range(7)):
            cell = _merged_text(ws, row, c0, c1, v, bg=bg, bold=(k == 0), size=9,
                                h_align="center" if k in (2, 3) else "left",
                                fg=BC.DS_NAV if k == 0 else None)
            if k == 2 and v != "—":
                cell.number_format = PCT_FMT
            if k == 3:
                cell.number_format = '0"%"'
            if k not in (2, 3):
                h = max(h, _text_height(v, width(c0, c1), size=9, min_pt=22))
        ws.row_dimensions[row].height = h
        row += 1
    if refresh_note:
        row = _note(ws, row, NC, f"Evening source refresh: {refresh_note}")
    return row


def build_coverage(ws, cov_rows, period_label, rows=None, refresh_note=""):
    """Data Coverage & Rules, read top to bottom: at a glance → what data is
    covered → where the actual performance comes from → the rules, grouped
    (sources, periods, effort, outcome, targets, limitations)."""
    NC = len(COV_COLS)
    for c, w in enumerate(COV_WIDTHS, 1):              # widths first: heights depend on them
        ws.column_dimensions[_gcl(c)].width = w
    BC.ds_title(ws, NC, "Data Coverage & Rules", period_label,
                "What the figures are built from, where each actual result comes from, and exactly how "
                "every measure is defined.   Amber cell = a data note worth a look.")
    row = 4

    # ── At a glance ────────────────────────────────────────────────────────────
    row = BC.ds_section(ws, row, NC, "AT A GLANCE", level=1)
    row = _coverage_cards(ws, row, cov_rows, NC) + 1

    # ── A. What data is covered ───────────────────────────────────────────────
    row = BC.ds_section(ws, row, NC, "A.  WHAT DATA IS COVERED — Coordinator report days read", level=1)
    for (c0, c1), text in COV_GROUPS:
        if c1 > c0:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        for cc in range(c0, c1 + 1):
            x = ws.cell(row=row, column=cc)
            x.fill = AR._fill(BC.DS_NAV)
            x.border = BC.ds_border()
        c = ws.cell(row=row, column=c0)
        c.value = text
        c.font = AR._font(bold=True, size=8, color=AR.C_WHITE)
        c.alignment = AR._align("center", "center")
    ws.row_dimensions[row].height = 16
    row += 1
    row = _hdr(ws, row, COV_COLS, height=44)
    totals = [0] * 10
    for i, c in enumerate(cov_rows):
        bg = BC.ds_zebra(i)
        vals = [c["day"].strftime("%d-%b-%Y (%a)"), c["versions"], c["versions_fu"],
                c["groups_tracked"], c["groups_untracked"], c["tasks"], c["untracked_tasks"],
                c["dropped"], c["multi_version"], c["yes_no_time"], c["conflicts"]]
        for k, v in enumerate(vals[1:]):
            totals[k] += v or 0
        for j, v in enumerate(vals, 1):
            warn = (j in (7, 10, 11) and v) or (j == 3 and not v)
            BC.ds_cell(ws, row, j, v, bg=(BC.DS_MED_BG if warn else bg), h_align="left" if j == 1 else "center",
                       bold=(j == 1) or bool(warn), fg=(BC.DS_MED_FG if warn else BC.DS_TEXT))
        ws.row_dimensions[row].height = 20
        row += 1
    if not cov_rows:
        row = BC.ds_empty(ws, row, NC, "No Coordinator report days found for this period.", level="muted")
    else:
        BC.ds_cell(ws, row, 1, f"Total · {len(cov_rows)} day(s)", bg=BC.DS_SUB, bold=True)
        for j, v in enumerate(totals, 2):
            BC.ds_cell(ws, row, j, v, bg=BC.DS_SUB, bold=True, h_align="center")
        ws.row_dimensions[row].height = 20
        row += 1
        row = _note(ws, row, NC,
                    "Amber = worth a look: no follow-up columns in any version, tasks that cannot be "
                    "tracked, Done = Yes without a valid DateTime, or Yes / No differing between versions. "
                    "None of these changes a figure silently — see section 6 below.", size=8)
    row += 1

    # ── B. Where the actual performance comes from ────────────────────────────
    if rows:
        row = BC.ds_section(ws, row, NC, "B.  WHERE THE ACTUAL PERFORMANCE COMES FROM — this period",
                            level=1)
        row = _outcome_sources_table(ws, row, NC, rows, refresh_note) + 1

    # ── C. Rules & definitions, grouped ───────────────────────────────────────
    row = BC.ds_section(ws, row, NC, "C.  RULES & DEFINITIONS", level=1)
    texts = OrderedDict(RULES + OUTCOME_RULES)
    used = set()
    sections = list(RULE_SECTIONS)
    leftover = [k for k in texts if not any(k in keys for *_x, keys in RULE_SECTIONS)]
    if leftover:
        sections.append((str(len(sections) + 1), "OTHER RULES", "info", leftover))
    text_width = sum(COV_WIDTHS[1:])
    for num, title, level, keys in sections:
        keys = [k for k in keys if k in texts and k not in used]
        if not keys:
            continue
        row = BC.ds_section(ws, row, NC, f"{num}.  {title}", level=2, height=20)
        if level == "medium":                          # limitations: amber heading
            hc = ws.cell(row=row - 1, column=1)
            hc.fill = AR._fill(BC.DS_MED_BG)
            hc.font = AR._font(bold=True, size=9, color=BC.DS_MED_FG)
        for i, k in enumerate(keys):
            used.add(k)
            key_bg = BC.DS_MED_BG if level == "medium" else BC.DS_SUB
            key_fg = BC.DS_MED_FG if level == "medium" else BC.DS_NAV
            BC.ds_cell(ws, row, 1, k, bg=key_bg, bold=True, fg=key_fg, wrap=True, v_align="center")
            _merged_text(ws, row, 2, NC, texts[k], bg=BC.ds_zebra(i))
            ws.row_dimensions[row].height = _text_height(texts[k], text_width)
            row += 1
        ws.row_dimensions[row].height = 8                # breathing space between rule groups
        row += 1

    BC.ds_fit_guide(ws, NC)
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = BC.DS_MUTED
    try:
        ws.print_area = f"A1:{_gcl(NC)}{row}"
    except Exception:                                    # pragma: no cover
        pass
    return ws


def generated_on_text(now: datetime) -> str:
    """'Generated On: 07-Oct-2026 11:30 AM' — the report's generation time (IST),
    same wording and format as pyAssignmentSubmissionPerformanceReport._title."""
    return f"Generated On: {now.strftime('%d-%b-%Y %I:%M %p')}"


def stamp_generated_on(ws, now: datetime):
    """Append '  |  Generated On: …' to the tab's existing row-1 header (the
    merged title band). The cell keeps its merge, font, fill and alignment;
    nothing is added when the header is empty or already stamped."""
    c = ws.cell(row=1, column=1)
    text = str(c.value or "").rstrip()
    if not text or "Generated On:" in text:
        return
    c.value = f"{text}  |  {generated_on_text(now)}"
    # a narrow tab (e.g. a Weekly label on the 6-column trend tab): let the title
    # wrap onto a second line instead of being cut off at the band's edge
    span = next((m for m in ws.merged_cells.ranges if m.min_row == 1 and m.min_col == 1), None)
    last = span.max_col if span else 1
    width = sum((ws.column_dimensions[_gcl(i)].width or 8.43) for i in range(1, last + 1))
    if len(c.value) > width * 0.82:                   # ~chars that fit at the 13 pt bold title
        c.alignment = AR._align(c.alignment.horizontal or "left", "center", wrap=True)
        ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 0, 52)


def build_workbook(ledger, start: date, end: date, period_label: str, is_daily: bool, now: datetime,
                   view=None, refresh_note=""):
    """view = compute_outcome_view(...) (None = effort only, outcomes shown as
    not measured)."""
    full_ledger = ledger
    ledger = scope_ledger(ledger, start, end)          # this period's tasks only
    tasks = ledger["tasks"]
    days = sorted({t["day"] for t in tasks})
    cov = ledger["coverage"]
    if view is None:
        view = compute_outcome_view(None, full_ledger,
                                    {"kind": "Daily" if is_daily else "Manual",
                                     "start": start, "end": end}, now)
    rows = scorecard_rows(ledger, tasks, view)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_dashboard(wb.create_sheet("Dashboard"), ledger, tasks, period_label, is_daily, now, view)
    build_outcome_trend(wb.create_sheet("Effort vs Outcome Trend"), view, period_label,
                        ledger["groups"], rows)
    build_trend(wb.create_sheet("Progress Trend"), ledger, tasks, period_label, days, is_daily, now,
                period_text=(start.strftime("%d-%b-%Y") if start == end else
                             f"{start.strftime('%d-%b-%Y')} to {end.strftime('%d-%b-%Y')}"))
    build_register(wb.create_sheet("Task Register"), ledger, tasks, period_label)
    build_coverage(wb.create_sheet("Data Coverage & Rules"), cov, period_label, rows=rows,
                   refresh_note=refresh_note)
    for ws in wb.worksheets:                          # every tab: header + "Generated On"
        stamp_generated_on(ws, now)
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


def run_jobs(jobs, versions, loader, now, upload=None, engine=None, refresh_note=""):
    """Build (and deliver) one workbook per job. The ledger is built once from
    the discovered versions; every job is then scoped to its own period
    (scope_ledger), so jobs never borrow each other's tasks. A failing job is
    recorded and the remaining jobs still run. Each workbook exists only in
    memory and is uploaded from there — nothing is written to local disk.
    engine = coordinator_outcomes.OutcomeEngine (None = effort only)."""
    ledger = build_ledger(versions, loader, now)
    log.info("Ledger: %d trackable task(s) across %d report day(s).",
             len(ledger["tasks"]), len({t["day"] for t in ledger["tasks"]}))
    if engine is not None:
        cutoffs = {}
        for v in versions:
            if v.get("created") and (v["day"] not in cutoffs or v["created"] > cutoffs[v["day"]]):
                cutoffs[v["day"]] = v["created"]
        engine.cutoffs.update(cutoffs)
    results = []
    for job in jobs:
        name = _file_name(job)
        folders = CP.folder_path(job["kind"], job["start"], job["end"])
        out = {"job": job, "name": name, "folder": "/".join(folders)}
        try:
            view = compute_outcome_view(engine, ledger, job, now)
            wb, tasks, s = build_workbook(ledger, job["start"], job["end"],
                                          f"{job['kind']}  ·  {job['label']}",
                                          job["kind"] == "Daily", now, view=view,
                                          refresh_note=refresh_note)
            out.update(summary=s, tasks=len(tasks))
            _sc = scope_ledger(ledger, job["start"], job["end"])       # for the e-mail breakdown
            out["groups"] = [(meta["name"], gs) for _gk, meta, gs in _group_rows(_sc, _sc["tasks"])]
            if view.get("enabled"):
                rows = scorecard_rows(_sc, _sc["tasks"], view)
                out["outcome_rows"] = rows
                out["headline"] = headline(s, rows)
                out["attention"] = outcome_totals(rows)["attention"]
                out["compare"] = view.get("compare")
            buf = io.BytesIO()                  # in memory only — never written to disk
            wb.save(buf)
            if upload:
                out["link"] = upload(folders, name, buf)
                # a delivered Daily report keeps its LIVE evening checks for the
                # Weekly / Monthly reports (never a copy / "not checked" placeholders)
                if job["kind"] == "Daily" and view.get("details_live") and view.get("details"):
                    save_outcome_checks(job["end"], view["details"])
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
              + (f"  {out['headline']}\n" if out.get("headline") else "")
              + f"  Folder: {out['folder']}\n"
              + (f"  [Drive] Uploaded: {out['link']}\n" if out.get("link") else
                 "  [Drive] upload OFF — built in memory only, nothing saved\n") + "=" * 70)
        results.append(out)
    return results


# ── stored evening checks (Outcome Detail of earlier Daily performance reports) ──
def discover_outcome_snapshots(drive, ranges) -> dict:
    """{report day: latest Daily performance report file} for the days in
    `ranges` — the files whose Outcome Detail tab holds that day's evening check."""
    out = {}
    for day, folder, x in _daily_folder_files(drive, ranges):
        if not x.get("name", "").startswith(f"{REPORT_BASENAME}_Daily"):
            continue
        meta = {"id": x["id"], "name": x["name"], "day": day,
                "created": _drive_time_ist(x.get("createdTime", "")),
                "modified_raw": x.get("modifiedTime", "")}
        cur = out.get(day)
        if cur is None or (meta["created"] or datetime.min) > (cur["created"] or datetime.min):
            out[day] = meta
    return out


def load_snapshot(drive, meta):
    """Stored evening-check details of one Daily performance report (cached per
    file revision). None when the report has no Outcome Detail tab."""
    safe = re.sub(r"[^0-9A-Za-z]", "", meta["modified_raw"])
    path = os.path.join(CACHE_DIR, f"outcome_{meta['id']}_{safe}.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return None if data is None else [CO.detail_from_values(v) for v in data]
        except Exception:                                        # corrupt cache -> refetch
            pass
    data = _call(lambda: drive.files().export(fileId=meta["id"], mimeType=XLSX_MIME).execute(),
                 f"Drive: export {meta['name']}")
    det = parse_outcome_detail(openpyxl.load_workbook(io.BytesIO(data), data_only=True))
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(None if det is None else [CO.detail_row_values(d) for d in det], fh)
    except Exception as exc:                                     # cache is best-effort
        log.warning("Could not write cache %s (%s).", path, exc)
    return det


def make_outcome_engine(drive, ranges, now):
    """The OutcomeEngine for this run (Google sources + stored evening checks)."""
    snaps = {}
    try:
        snaps = discover_outcome_snapshots(drive, ranges)
    except Exception as exc:                                     # noqa: BLE001
        log.warning("Could not list earlier Daily performance reports (%s) — earlier days' "
                    "evening checks will show as not checked.", exc)

    def _snapshot(day):
        local = load_outcome_checks(day)
        if local is not None:
            return local
        meta = snaps.get(day)                  # legacy: a Daily report's Outcome Detail tab
        if meta is None:
            return None
        try:
            return load_snapshot(drive, meta)
        except Exception as exc:                                 # noqa: BLE001
            log.warning("Stored evening check of %s unreadable (%s).", day, exc)
            return None

    import pyAssignmentSubmissionPerformanceReport as ASR
    import pyAdmissionFormalitiesReport as ADM
    try:
        import pyWiseDataValidationReport as WISE
    except Exception:                                            # pragma: no cover
        WISE = None

    def _keyfn(r):
        g = next(x for x in TASK_GROUPS if x["key"] == "instructor")
        vals = {"Tech Name": r["sess"].get("course_name", ""),
                "Duration": r["sess"].get("course_title", ""),
                "Instructor": r.get("instr", ""),
                "Session Date": (r["actual_start_dt"].strftime("%Y-%m-%d %H:%M:%S")
                                 if r.get("actual_start_dt") is not None
                                 else str(r["sess"].get("start_time_ist", ""))[:19])}
        return "|".join(_key_value(f, vals.get(f, "")) for f in g["identity"])

    return CO.OutcomeEngine(BC, CO.GoogleOutcomeSources(BC, cache_dir=CACHE_DIR), now, snapshot=_snapshot,
                            keyfn=_keyfn, ASR=ASR, ADM=ADM, WISE=WISE)


def _refresh_note() -> str:
    """The evening batch's source-refresh status (scripts/run_evening_reports.py)."""
    raw = os.environ.get("INTELLIBI_EVENING_REFRESH", "").strip()
    if not raw:
        return ""
    parts = [p.split("=", 1) for p in raw.split(";") if "=" in p]
    bad = [n for n, st in parts if st.strip().upper() != "SUCCESS"]
    if not bad:
        return "all sources refreshed before the re-check."
    return ("NOT refreshed: " + ", ".join(bad) + " — the re-check used the data from their last "
            "successful run.")


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
    # each job's period + the previous period (Δ) + the trend days
    ranges = job_ranges(jobs, now.date())
    versions = discover_report_versions(drive, ranges)
    engine = None
    if CHECK_OUTCOMES:
        try:
            engine = make_outcome_engine(drive, ranges, now)
        except (Exception, SystemExit) as exc:                     # noqa: BLE001 (a source module may sys.exit)
            log.exception("Actual-performance checks unavailable (%s) — effort only.", exc)

    def _upload(folders, name, buf):
        # Same versioned, never-overwrite upload as the Batch Coordinator report,
        # into the same <Type>/<reporting period> folder.
        return BC.upload_report(folders, name + ".xlsx", buf, name)

    results = run_jobs(jobs, versions, lambda v: load_version(drive, v), now,
                       upload=_upload if UPLOAD_TO_DRIVE else None, engine=engine,
                       refresh_note=_refresh_note())
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
                                   group_labels=EMAIL_GROUP_LABELS,
                                   outcome_rows=r.get("outcome_rows"),
                                   headline=r.get("headline"), attention=r.get("attention"),
                                   compare=r.get("compare"), compare_band=COMPARE_STEADY_BAND)
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
