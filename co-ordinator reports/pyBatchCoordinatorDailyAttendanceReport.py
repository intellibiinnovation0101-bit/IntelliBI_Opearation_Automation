"""
================================================================================
  IntelliBI Operations Automation
  BATCH COORDINATOR — Daily Attendance Task Report
  (co-ordinator reports / pyBatchCoordinatorDailyAttendanceReport.py)
  ------------------------------------------------------------------------------
  WHAT THIS IS
    A DAILY, technology-wise, student-level attendance report that EXTENDS the
    existing Daily Attendance Report (pyAttendaceFeedbackReport.py) with each
    student's TILL-DATE aggregate attendance + feedback performance for the same
    Tech Name + Duration. It is meant to look and behave like the existing daily
    report, plus an "Overall / Till-Date Aggregate" block.

  REUSE (no logic duplicated / no existing calculation changed)
    * pyAttendaceFeedbackReport.py  → load_all_data(), every styling/colour
      helper (fonts, fills, borders, banners, band colours, _remarks, _att_bg,
      _rating_bg, auto_col_width …) and all colour constants. The DAILY columns
      use the identical values/logic as that report's Student Detail.
    * pyCoordinatorAttendanceTaskReport.py → the aggregate flagging logic
      (score_record) for the "Why Flagged" reasons, plus the same definitions for
      Last Present / Absent Streak / Feedbacks / Feedback Part. % / Avg Rating.

  OUTPUT
    One native Google Sheet per report, written into the shared Coordinator
    layout (coordinator_periods.py) so history is never overwritten and the
    Task Performance report for the same period sits in the same folder:
        <parent>/Daily Coordinator Reports/Daily 12-Sep-2026/
            IntelliBI_Batch_Coordinator_Daily_Attendance_Report_
            12-Sep-2026_10.00_AM_-_12-Sep-2026_03.00_PM
        <parent>/Weekly Coordinator Reports/Weekly 07-Sep-2026 to 13-Sep-2026/ …
        <parent>/Monthly Coordinator Reports/Monthly Sep-2026/ …
        <parent>/Manual Coordinator Reports/Manual 01-Sep-2026 to 15-Sep-2026/ …
    Re-running the same report saves the next "- Version N"; earlier versions
    are never modified.
    The workbook carries ALL the daily-report tabs, in the same order:
        Session Summary | Student Detail | Feedback Rating | Teacher_No_Feedback
    Session Summary, Feedback Rating and Teacher_No_Feedback are produced
    VERBATIM by pyAttendaceFeedbackReport.py's daily builders; Student Detail is
    the extended version (identity + daily + till-date aggregate block).

  COLUMNS
    Student:  Tech Name | Duration | Student Name | Phone
    Daily:    Status | Attendance % | Duration (min) | Joined At | Left At |
              Remarks | Feedback Given?
    Aggregate (till the report date, per Tech Name + Duration + Student):
              Sessions (P/T) | Agg Attendance % | Agg Duration Att % |
              Last Present | Absent Streak | Feedbacks | Feedback Part. % |
              Avg Rating (/10) | Overall Remarks | Why Flagged

    "Agg Attendance %"    = Present sessions / Total applicable sessions × 100.
    "Agg Duration Att %"  = total ATTENDED minutes / total applicable SESSION
                            minutes × 100 (the till-date analogue of the daily
                            duration-based Attendance %). Named distinctly so the
                            two aggregates are never confused.
    "Overall Remarks"     = the daily _remarks() bands applied to the aggregate
                            (same words as the daily Remark, computed on the
                            aggregate). Named "Overall Remarks" to avoid a second
                            column literally called "Remarks".

  RUN
    python "co-ordinator reports/pyBatchCoordinatorDailyAttendanceReport.py"
    Which reports run is set by the GENERATE_* flags in RUN CONFIGURATION below
    (GENERATE_AUTO picks Daily/Weekly/Monthly by the run date). Layer-1
    collectors must have run first.
================================================================================
"""
from __future__ import annotations

# ── bootstrap: make the project's common/ AND ops_reports_action/ importable ──
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT = os.path.dirname(_HERE)
for _p in (os.path.join(_PROJECT, "common"),
           os.path.join(_PROJECT, "ops_reports_action")):
    if os.path.isdir(_p) and _p not in sys.path:
        sys.path.insert(0, _p)

try:
    import _bootstrap  # noqa: F401  (sys.path + env defaults + config.yaml)
    from paths import CREDENTIALS_DIR
except Exception:
    CREDENTIALS_DIR = os.path.join(_PROJECT, "credentials")

import io
import logging
import calendar
from datetime import datetime, date, timedelta, time, timezone

import pandas as pd
import openpyxl

# Reuse the existing daily report wholesale (helpers, styling, colours, loader).
import pyAttendaceFeedbackReport as AR

# Shared Coordinator reporting periods + Drive layout (Daily / Weekly / Monthly /
# Manual → reporting-period folder) — the SAME module the Task Performance report
# uses, so both reports for one period always land in one folder.
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
import coordinator_periods as CP

# Reuse the coordinator flagging logic for "Why Flagged" (best-effort).
try:
    import pyCoordinatorAttendanceTaskReport as COORD
    _COORD_OK = True
except Exception as _e:                      # pragma: no cover
    COORD = None
    _COORD_OK = False

log = logging.getLogger("BatchCoordinatorDailyAttendance")

# =============================================================================
#  RUN CONFIGURATION
# =============================================================================
PARENT_FOLDER_ID  = "1BEokUc7Np7iBVSMwIyMAgZUa0mrecT-h"   # coordinator Drive folder
IMPERSONATE_USER  = "info@intellibiinnovationstechnologies.in"
# File / native-sheet base name. The report period ("duration" — the daily
# time-range label, same style as pyAttendaceFeedbackReport.py) is appended, e.g.
#   IntelliBI_Batch_Coordinator_Daily_Attendance_Report_12-Sep-2026_10.00_AM_-_12-Sep-2026_03.00_PM
REPORT_BASENAME   = "IntelliBI_Batch_Coordinator_Daily_Attendance_Report"
VERBOSE           = True

# =============================================================================
#  REPORT GENERATION CONTROL  (flag-based; mirrors pyLeadFollowUpAnalysisReport.py)
# =============================================================================
#  GENERATE_AUTO = True  → scheduler-friendly automatic selection by run date:
#      • Daily   — every run, for the current day.
#      • Weekly  — every Monday, for the previous completed week (Mon → Sun).
#      • Monthly — on the last calendar day of the month, for that whole month.
#    (The manual GENERATE_* / date flags below are IGNORED while AUTO is True.)
#
#  GENERATE_AUTO = False → use the GENERATE_DAILY / WEEKLY / MONTHLY / MANUAL
#    flags independently; several can be True and are all produced in one run.
#    The optional date variables pin a specific historical period; when None the
#    existing/default period is used (today / current week / current month).
GENERATE_AUTO    = True

GENERATE_DAILY   = True
GENERATE_WEEKLY  = True
GENERATE_MONTHLY = True
GENERATE_MANUAL  = False          # Manual = a custom start/end date range

DAILY_DATE            = None      # "YYYY-MM-DD"  (None = today)
WEEKLY_REFERENCE_DATE = None      # any date within the wanted week (None = this week)
MONTHLY_MONTH         = None      # 1-12          (None = current month)
MONTHLY_YEAR          = None      # e.g. 2026     (None = current year)
MANUAL_START_DATE     = None      # "YYYY-MM-DD"  (required when GENERATE_MANUAL)
MANUAL_END_DATE       = None      # "YYYY-MM-DD"  (required when GENERATE_MANUAL)

# =============================================================================
#  COORDINATOR FOLLOW-UP TRACKING COLUMNS  (appended to every daily action tab)
#  ---------------------------------------------------------------------------
#  Four columns the coordinator fills in on the generated report itself:
#    Action Taken        dropdown, values per tab (FOLLOWUP_ACTIONS)
#    Follow-Up Comment   free text
#    Follow-Up Done?     dropdown Yes / No
#    Follow-Up DateTime  stamps itself the moment Follow-Up Done? is set
#                        (self-referencing formula, kept stable by iterative
#                        calculation — enabled in the workbook AND on the uploaded
#                        Google Sheet by _enable_followup_timestamps()).
#  They are report-side input fields only: nothing in the data selection,
#  follow-up identification or Why Flagged logic reads or depends on them.
# =============================================================================
FOLLOWUP_COLS = ["Action Taken", "Follow-Up Comment", "Follow-Up Done?", "Follow-Up DateTime"]
FOLLOWUP_DONE_OPTIONS = ["Yes", "No"]
FOLLOWUP_ACTIONS = {
    "attendance": ["Call", "WhatsApp", "Call & WhatsApp", "Email", "No response", "Other"],
    "assignment": ["Call", "WhatsApp", "WhatsApp Group", "Call & WhatsApp", "Email", "No response", "Other"],
    "admission":  ["Form Send", "Form Signed", "Call", "WhatsApp", "Call & WhatsApp", "Email", "Other",
                   "Not Applicable"],
    "wise":       ["Corrected", "Invalid", "Call", "WhatsApp", "Call & WhatsApp", "Email", "Other"],
    "instructor": ["Call", "WhatsApp", "Call & WhatsApp", "Email", "Other"],
    "interview":  ["Call", "WhatsApp", "Call & WhatsApp", "Email", "Other"],
}
FOLLOWUP_DT_FORMAT = "dd-mmm-yyyy hh:mm:ss"      # dd-MMM-yyyy HH24:MM:SS
FOLLOWUP_TIMEZONE  = "Asia/Kolkata"              # NOW() on the uploaded sheet = IST

# =============================================================================
#  COLUMN LAYOUT
# =============================================================================
# Tech Name / Duration are NOT detail columns: they are the technology banner
# that groups the rows (shown once per group, with the pending-learner count).
STUDENT_COLS = ["Rank", "Student Name", "Phone"]
DAILY_COLS   = ["Status", "Attendance %", "Duration (min)", "Joined At",
                "Left At", "Remarks", "Feedback Given?", "Feedback Rating (/10)"]
AGG_COLS     = ["Sessions (P/T)", "Agg Attendance %", "Agg Duration Att %",
                "Last Present", "Absent Streak", "Feedbacks", "Feedback Part. %",
                "Avg Rating (/10)", "Overall Remarks", "Why Flagged"]
HEADERS = STUDENT_COLS + DAILY_COLS + AGG_COLS + FOLLOWUP_COLS
N = len(HEADERS)
_C = {h: i + 1 for i, h in enumerate(HEADERS)}   # 1-based column index by header
_AGG_START = len(STUDENT_COLS) + len(DAILY_COLS) + 1   # first aggregate column (1-based)

# ── Learner Follow-Ups thresholds ────────────────────────────────────────────
FEEDBACK_PART_THRESHOLD = 25.0   # Overall Feedback Part. % below this = feedback defaulter
LOW_RATING_THRESHOLD    = 7.0    # Daily feedback rating <= this = low-rating follow-up
CRITICAL_ATT_PCT        = 60.0   # "Critical" attendance band (matches AR._remarks)

# ── Instructor Follow-Ups ────────────────────────────────────────────────────
INSTRUCTOR_TAB_SHEET_ID = "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA"  # IntelliBIStudentInfo
INSTRUCTOR_TAB          = "Instructor"
INSTRUCTOR_NAME_SLOTS   = 5      # instructor_name_1..N -> alternative_contact_number_1..N
INSTR_SHORT_MIN         = 15.0   # (period roll-up) session ran >= this many min short = flag
INSTR_LOW_ATT_PCT       = 60.0   # (period roll-up) session Att % below this = flag
INSTR_LOW_RATING        = 7.0    # (period roll-up) session avg rating <= this = flag
# ── Instructor Follow-Ups (daily) — session follow-up thresholds ─────────────
INSTR_EARLY_MIN         = 5.0    # instructor expected to start >= this many min BEFORE scheduled
INSTR_UNDERRUN_MIN      = 30.0   # Diff Mins <= -this (ran this many min short) = Session Underrun
INSTR_LOW_FB_RATE_PCT   = 25.0   # student Feedback Rate % <= this = Low Feedback Rate follow-up

INSTR_COLS = [
    "Tech Name", "Duration", "Instructor", "Phone", "Session Date",
    "Duration (Mins)", "Scheduled (Mins)", "Diff Mins",
    "Total Enrolled", "Att. N/A Count", "Present", "Absent", "Att %",
    "Avg Time in Session %", "No. of Feedbacks", "Feedback Rate %",
    "⭐ Avg Rating(/10)", "Min Rating", "Max Rating", "Feedback Given", "Why Flagged",
] + FOLLOWUP_COLS
IN = len(INSTR_COLS)
_IC = {h: i + 1 for i, h in enumerate(INSTR_COLS)}   # 1-based column index by header

# ── Weekly/Monthly period roll-up column layouts (one row per learner / per
#    instructor — NOT per day, so daily records are never duplicated) ──────────
LP_COLS = [
    "Rank", "Tech Name", "Duration", "Student Name", "Phone",
    "Sessions (P/T) in Period", "Attendance Flag Days", "Feedback Flag Days",
    "Low-Rating Days", "Total Flag Days",
    "Agg Attendance %", "Agg Duration Att %", "Feedback Part. %", "Avg Rating (/10)",
    "Overall Remarks", "Why Flagged",
]
LPN = len(LP_COLS)
_LP = {h: i + 1 for i, h in enumerate(LP_COLS)}

IP_COLS = [
    "Tech Name", "Duration", "Instructor", "Phone", "Sessions",
    "Sessions Missing Feedback", "Feedback Given %", "Avg Att %", "Avg Rating (/10)",
    "Flagged Sessions", "Avg Diff Mins", "Why Flagged",
]
IPN = len(IP_COLS)
_IP = {h: i + 1 for i, h in enumerate(IP_COLS)}

# ── Learner Assignment Follow-Ups (assignment-submission defaulters) ──────────
#   Reuses pyAssignmentSubmissionEmailReminder's defaulter/reminder logic wholesale
#   (find_pending_reminders → consolidate_by_student → _group_rows_by_course_assignment).
AF_COLS = ["#", "Student Name", "Email", "Phone", "Deadline", "Reminder", "Why Flagged"] + FOLLOWUP_COLS
AFN = len(AF_COLS)
_AF = {h: i + 1 for i, h in enumerate(AF_COLS)}
# reminder stage → coordinator action (drives Why Flagged). Keyed by the reminder
# LEVEL produced by the reminder script, never hard-coded per technology/date.
_ASSIGN_ACTION = {
    "1st":    "Share the assignment reminder report in the WhatsApp Group",
    "2nd":    "Send report on WhatsApp Group + one-to-one reminder message to the learner",
    "final":  "Send report on WhatsApp Group + Call the learner for assignment submission — "
              "this is the final submission day",
    "missed": "Send report on WhatsApp Group + Call the learner urgently to submit — "
              "the deadline has already passed",
}

# ── Learner Admission Formalities (admission-form / e-signature status) ────────
#   Reuses pyAdmissionFormalitiesReport's EXACT student↔signer matching and
#   status/identification logic (read_records → build_signer_indexes →
#   build_rows / row_colour). No admission business logic is (re)defined here.
try:
    import pyAdmissionFormalitiesReport as ADM
    _ADM_OK = True
except Exception as _e:                      # pragma: no cover
    ADM = None
    _ADM_OK = False

# Column layout = the reference report's EXACT columns + one action column
# ("Why Flagged") immediately AFTER "Expiry Date". Falls back to a literal copy
# of the reference layout if the module can't be imported at load time.
_ADM_BASE_COLS = (list(ADM.REPORT_COLUMNS) if _ADM_OK else [
    "Student Name", "Email ID", "Phone Number", "Batch Name", "Joined On",
    "Request Form Name", "Recipient Status", "Request Status",
    "Sent Date", "Signed Date", "Expiry Date"])
ADM_COLS = _ADM_BASE_COLS + ["Why Flagged"] + FOLLOWUP_COLS
ADMN = len(ADM_COLS)
_ADMC = {h: i + 1 for i, h in enumerate(ADM_COLS)}   # 1-based column index by header


# =============================================================================
#  REPORT DESIGN SYSTEM  (presentation only — one visual language for every tab)
#  ---------------------------------------------------------------------------
#  Every tab of the Batch Coordinator report is laid out the same way so the
#  coordinator can read any row left → right as:
#        WHO / WHAT is affected  →  the numbers behind the flag  →  WHY FLAGGED
#  and, inside "Why Flagged", each reason as   issue: figures   →  ACTION.
#
#    row 1   title band          navy, white bold        (tab purpose | period)
#    row 2   guide strip         pale, small text        (how to read + colour key
#                                                         + record count)
#    row 3   column headers      dark navy, white bold   (optional grouped
#                                                         super-header above it)
#    body    white / zebra rows, thin light borders.  Colour is used ONLY where
#            it carries meaning:
#              • the first cell of a row = PRIORITY chip (High = red,
#                Medium = amber, Info = blue, OK = green) with a coloured left edge
#              • status / measure cells (Absent, No feedback, Missing, Invalid,
#                Att %, rating …) in pale semantic tints with dark coloured text
#              • "Why Flagged" always in the same pale-yellow action column:
#                numbered reasons, key figures bold red, the ACTION on its own
#                line in bold navy ("→ Call Learner")
#    section banner   blue band with white text   (technology / batch / group)
#    sub-banner       pale-blue band, navy text    (assignment / interview detail)
#    empty state      pale-green band, dark-green text
#  None of this changes what is listed or the values shown — only how they look.
# =============================================================================
from openpyxl.styles import Border as _DSBorder, Side as _DSSide
from openpyxl.utils import get_column_letter as _gcl

DS_NAV        = AR.C_NAV          # title band / super-header
DS_NAV2       = AR.C_NAV2         # column headers
DS_SECTION    = AR.C_BLUE_MID     # section banner (group)
DS_SUB        = "DCE9F7"          # sub-banner (detail line under a group)
DS_GUIDE      = "EEF3F8"          # guide strip
DS_ZEBRA      = "F6F8FB"          # alternate body row
DS_WHY        = "FFFDE7"          # "Why Flagged" action column
DS_LINE       = "D6DCE4"          # cell borders
DS_TEXT       = "222222"
DS_MUTED      = "6B7280"
# priority / semantic tints  (pale fill, dark text)
DS_HIGH_BG,   DS_HIGH_FG   = "FDE2E2", AR.C_RED_DARK
DS_MED_BG,    DS_MED_FG    = "FFF1DB", AR.C_AMBER_DARK
DS_INFO_BG,   DS_INFO_FG   = "E3EEFB", "0D47A1"
DS_OK_BG,     DS_OK_FG     = "E6F4EA", AR.C_GREEN_DARK
DS_MUTE_BG,   DS_MUTE_FG   = "EFEFEF", "757575"
_DS_LEVEL = {
    "high":   (DS_HIGH_BG, DS_HIGH_FG),
    "medium": (DS_MED_BG,  DS_MED_FG),
    "info":   (DS_INFO_BG, DS_INFO_FG),
    "ok":     (DS_OK_BG,   DS_OK_FG),
    "muted":  (DS_MUTE_BG, DS_MUTE_FG),
    "none":   (AR.C_WHITE, DS_TEXT),
}
_DS_SIDE = _DSSide(style="thin", color=DS_LINE)


def ds_border(left_accent=None):
    """Thin light border; optional medium coloured LEFT edge (priority accent)."""
    l = _DSSide(style="medium", color=left_accent) if left_accent else _DS_SIDE
    return _DSBorder(left=l, right=_DS_SIDE, top=_DS_SIDE, bottom=_DS_SIDE)


def ds_level_colors(level):
    return _DS_LEVEL.get(level or "none", _DS_LEVEL["none"])


def ds_cell(ws, row, col, value, bg=None, fg=DS_TEXT, bold=False, italic=False,
            h_align="left", v_align="center", wrap=False, number_fmt=None, size=10,
            left_accent=None):
    """One body cell in the common style. Returns the cell."""
    c = ws.cell(row=row, column=col)
    c.value = value
    c.font = AR._font(bold=bold, size=size, color=fg, italic=italic)
    c.fill = AR._fill(bg or AR.C_WHITE)
    c.alignment = AR._align(h_align, v_align, wrap=wrap)
    c.border = ds_border(left_accent)
    if number_fmt:
        c.number_format = number_fmt
    return c


def ds_pill(ws, row, col, text, level, bold=True, h_align="center", number_fmt=None):
    """A status / priority chip: pale semantic fill + dark coloured bold text."""
    bg, fg = ds_level_colors(level)
    return ds_cell(ws, row, col, text, bg=bg, fg=fg, bold=bold, h_align=h_align,
                   number_fmt=number_fmt)


def ds_priority(ws, row, col, text, level, h_align="center", wrap=True):
    """The FIRST cell of a flagged row: chip + a medium coloured left edge, so the
    priority is visible at a glance even when the row is scrolled."""
    bg, fg = ds_level_colors(level)
    return ds_cell(ws, row, col, text, bg=bg, fg=fg, bold=True, h_align=h_align,
                   left_accent=fg, wrap=wrap)


def ds_att_bg(pct, status=""):
    """Attendance-% heat tint — AR._att_bg's exact thresholds, in the report's
    pale tint family (Absent / <60 = red, 60–75 = amber, 75–95 = pale green,
    ≥95 = green)."""
    if status == "Absent" or pct < 60:
        return DS_HIGH_BG
    if pct < 75:
        return DS_MED_BG
    if pct < 95:
        return "F1F8F1"
    return DS_OK_BG


def ds_rating_bg(avg):
    """Rating (/10) heat tint — AR._rating_bg's exact thresholds, pale family."""
    if avg >= 8:
        return DS_OK_BG
    if avg >= 6:
        return "F1F8F1"
    if avg >= 4:
        return DS_MED_BG
    return DS_HIGH_BG


def ds_zebra(i):
    return AR.C_WHITE if i % 2 == 0 else DS_ZEBRA


def ds_title(ws, ncols, tab_purpose, period_label, guide_text):
    """Rows 1–2: title band + guide strip. Returns the next free row (3)."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1)
    c.value = f"  IntelliBI  |  Batch Coordinator — {tab_purpose}  |  {period_label}"
    c.font = AR._font(bold=True, size=13, color=AR.C_WHITE)
    c.fill = AR._fill(DS_NAV)
    c.alignment = AR._align("left", "center")
    ws.row_dimensions[1].height = 34
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    g = ws.cell(row=2, column=1)
    g.value = "  " + guide_text
    g.font = AR._font(size=9, color=DS_MUTED, italic=True)
    g.fill = AR._fill(DS_GUIDE)
    g.alignment = AR._align("left", "center", wrap=False)
    g.border = ds_border()
    ws.row_dimensions[2].height = 20
    return 3


DS_GUIDE_DEFAULT = ("Read each row left → right: who is affected → the figures behind the flag → "
                    "Why Flagged (issue, then the ACTION in bold).   Colour key:  "
                    "■ High = red   ■ Medium = amber   ■ Info = blue   ■ OK = green.")


def ds_guide_count(ws, label, n):
    """Append 'label: n' to the guide strip (row 2) once a tab is fully rendered."""
    g = ws.cell(row=2, column=1)
    g.value = f"{g.value}      {label}: {n}"


def ds_headers(ws, row, headers, groups=None, height=32):
    """Column-header row (+ optional grouped super-header row above it).
    groups: list of (first_col, last_col, text). Returns the first body row."""
    if groups:
        for c0, c1, text in groups:
            if c1 > c0:
                ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
            c = ws.cell(row=row, column=c0)
            c.value = text
            c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
            c.fill = AR._fill(DS_NAV)
            c.alignment = AR._align("center", "center")
            for cc in range(c0, c1 + 1):
                ws.cell(row=row, column=cc).border = ds_border()
                ws.cell(row=row, column=cc).fill = AR._fill(DS_NAV)
        ws.row_dimensions[row].height = 18
        row += 1
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col)
        c.value = h
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.fill = AR._fill(DS_NAV2)
        c.alignment = AR._align("center", "center", wrap=True)
        c.border = ds_border()
    ws.row_dimensions[row].height = height
    return row + 1


def ds_section(ws, row, ncols, text, level=1, height=None):
    """Full-width banner. level 1 = group (blue, white text); level 2 = detail
    line under a group (pale blue, navy text)."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1)
    c.value = "  " + str(text).strip()
    if level == 1:
        c.font = AR._font(bold=True, size=10, color=AR.C_WHITE)
        c.fill = AR._fill(DS_SECTION)
    else:
        c.font = AR._font(bold=True, size=9, color=DS_NAV)
        c.fill = AR._fill(DS_SUB)
    c.alignment = AR._align("left", "center")
    for cc in range(1, ncols + 1):
        ws.cell(row=row, column=cc).border = ds_border()
    ws.row_dimensions[row].height = height or (22 if level == 1 else 19)
    return row + 1


def ds_empty(ws, row, ncols, text, level="ok"):
    """'Nothing to do' band: pale green (or grey when data is unavailable)."""
    bg, fg = ds_level_colors(level)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1)
    c.value = text
    c.font = AR._font(bold=True, size=10, color=fg)
    c.fill = AR._fill(bg)
    c.alignment = AR._align("center", "center")
    for cc in range(1, ncols + 1):
        ws.cell(row=row, column=cc).border = ds_border()
    ws.row_dimensions[row].height = 24
    return row + 1


# ---- "Why Flagged" ---------------------------------------------------------------
_DS_ACTION_VERBS = ("call ", "message ", "send ", "share ", "understand ", "ask ", "review ",
                    "correct ", "fill ", "add ", "fix ", "assign ", "confirm ", "identify ",
                    "look ", "follow ", "coordinator ", "please ")


def _ds_is_action(text):
    t = str(text or "").strip().lower()
    return any(t.startswith(v) for v in _DS_ACTION_VERBS)


def _ds_split_action(runs):
    """Split one reason's runs into (issue_runs, action_runs). The action is the
    text after the LAST ' — ' / ' → ' / sentence break that starts with an action
    verb (e.g. '— Call Learner', '→ fill in this field in Wise', '. Message
    Instructor.'). If the reason itself STARTS with an action verb ('Call the
    learner to complete …'), the part before the first ' — ' is the action and the
    rest is the issue. Returns (runs, []) when no action is found. Purely
    presentational: the words are unchanged, only where they are shown."""
    flat = "".join(t for t, _ in runs)
    cands = []                                   # (issue_end, action_start)
    for sep, keep in ((" — ", 0), (" → ", 0), (". ", 1)):
        i = flat.rfind(sep)
        while i >= 0:
            if _ds_is_action(flat[i + len(sep):]):
                cands.append((i + keep, i + len(sep)))
                break
            i = flat.rfind(sep, 0, i)
    action_first = False
    if cands:
        issue_end, action_start = max(cands, key=lambda c: c[1])
    elif _ds_is_action(flat):
        firsts = [(flat.find(sep), len(sep)) for sep in (" — ", " → ") if flat.find(sep) > 0]
        if not firsts:
            return runs, []
        i, n = min(firsts)
        issue_end, action_start, action_first = i, i + n, True
    else:
        return runs, []
    a_runs, b_runs, pos = [], [], 0
    for text, red in runs:
        start, end = pos, pos + len(text)
        pos = end
        left = text[:max(0, min(len(text), issue_end - start))] if start < issue_end else ""
        right = text[max(0, action_start - start):] if end > action_start else ""
        if left:
            a_runs.append((left, red))
        if right:
            b_runs.append((right, red))
    if action_first:
        return b_runs, a_runs
    return a_runs, b_runs


def ds_why_text(ws, row, col, reasons, col_width=70, min_height=22):
    """Write the 'Why Flagged' cell for a row:
         1) issue: figures            (label + key figures bold red)
            → ACTION                  (bold navy, own line)
         2) …
    (a single reason uses '•' instead of a number). Sets the row height from the
    wrapped line count so nothing is ever cut off. Returns the estimated lines."""
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
    except Exception:                                   # very old openpyxl
        CellRichText = None
    reasons = [r for r in (reasons or []) if r]
    cpl = max(20, int(col_width * 1.25))                # ~chars per wrapped line (Arial 10)
    lines = 0
    blocks, plain = [], []
    if CellRichText is not None:
        f_base = InlineFont(rFont="Arial", sz=10, color=DS_TEXT)
        f_key = InlineFont(rFont="Arial", sz=10, b=True, color=AR.C_RED_DARK)
        f_lbl = InlineFont(rFont="Arial", sz=10, b=True, color=DS_HIGH_FG)
        f_act = InlineFont(rFont="Arial", sz=10, b=True, color=DS_NAV)
        f_num = InlineFont(rFont="Arial", sz=10, b=True, color=DS_MUTED)
    for i, runs in enumerate(reasons):
        issue, action = _ds_split_action(runs)
        if not issue:
            issue, action = runs, []
        # first letter of the issue upper-case (e.g. 'status: …' → 'Status: …')
        t0, r0 = issue[0]
        t0s = t0.lstrip()
        if t0s and t0s[0].islower():
            issue = [(t0[:len(t0) - len(t0s)] + t0s[0].upper() + t0s[1:], r0)] + list(issue[1:])
        prefix = f"{i + 1})  " if len(reasons) > 1 else "•  "
        issue_txt = "".join(t for t, _ in issue).strip()
        action_txt = "".join(t for t, _ in action).strip().rstrip(".")
        if i:
            if CellRichText is not None: blocks.append(TextBlock(f_base, "\n"))
            plain.append("\n")
        if CellRichText is not None:
            blocks.append(TextBlock(f_num, prefix))
            for k, (text, red) in enumerate(issue):
                blocks.append(TextBlock(f_lbl if (k == 0 and red) else (f_key if red else f_base), text))
            if action_txt:
                blocks.append(TextBlock(f_base, "\n      → "))
                blocks.append(TextBlock(f_act, action_txt))
        plain.append(prefix + issue_txt + (f"\n      → {action_txt}" if action_txt else ""))
        lines += max(1, -(-len(prefix + issue_txt) // cpl))
        if action_txt:
            lines += max(1, -(-(len(action_txt) + 8) // cpl))
    c = ws.cell(row=row, column=col)
    c.value = CellRichText(blocks) if (CellRichText is not None and blocks) else "".join(plain)
    c.fill = AR._fill(DS_WHY)
    c.font = AR._font(size=10, color=DS_TEXT)
    c.alignment = AR._align("left", "center", wrap=True)
    c.border = ds_border()
    need = max(min_height, 13 * lines + 8)
    cur = ws.row_dimensions[row].height or 0
    ws.row_dimensions[row].height = max(cur, need)
    return lines


def ds_fit_guide(ws, ncols):
    """Wrap the guide strip (row 2) onto two/three lines when the tab is narrower
    than the text, so nothing is cut off."""
    g = ws.cell(row=2, column=1)
    total = sum((ws.column_dimensions[_gcl(c)].width or 10) for c in range(1, ncols + 1))
    n = len(str(g.value or ""))
    lines = max(1, -(-n // max(40, int(total * 1.3))))
    if lines > 1:
        g.alignment = AR._align("left", "center", wrap=True)
        ws.row_dimensions[2].height = 14 * lines + 6


# ---- Coordinator follow-up block (Action Taken / Comment / Done? / DateTime) ----
DS_INPUT_HDR = AR.C_TEAL          # header fill of the four input columns
DS_INPUT_BG  = "F1FAF8"           # input cell background (pale teal = "yours to fill")
_DS_FU_ROWS: dict = {}            # id(ws) -> [data rows that carry follow-up cells]
FOLLOWUP_WIDTHS = {"Action Taken": 18, "Follow-Up Comment": 34,
                   "Follow-Up Done?": 13, "Follow-Up DateTime": 20}
DS_GUIDE_FOLLOWUP = ("   Teal columns are yours: pick the Action Taken, add a comment, set "
                     "Follow-Up Done? — the DateTime stamps itself.")


def ds_followup_headers(ws, row, first_col):
    """Restyle the four follow-up header cells (teal) so the input area is obvious."""
    for k, h in enumerate(FOLLOWUP_COLS):
        c = ws.cell(row=row, column=first_col + k)
        c.value = "✎ " + h
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.fill = AR._fill(DS_INPUT_HDR)
        c.alignment = AR._align("center", "center", wrap=True)
        c.border = ds_border()


def ds_followup_cells(ws, row, first_col):
    """Write the four input cells for one flagged row and register the row for
    the dropdowns. Follow-Up DateTime holds a self-referencing formula that
    freezes NOW() the first time Follow-Up Done? is set (and clears again when
    it is cleared); iterative calculation makes it stable. The `=0` guard covers
    engines whose first pass of a self-reference starts at 0 (Excel / LibreOffice)
    as well as Google Sheets' empty start, and the ISERROR guard recovers a cell
    that showed #REF! (circular dependency) before iterative calculation was
    switched on — verified: stamps once, then holds."""
    done = _gcl(first_col + 2)
    dt = _gcl(first_col + 3)
    ds_cell(ws, row, first_col, None, bg=DS_INPUT_BG, h_align="center")
    ds_cell(ws, row, first_col + 1, None, bg=DS_INPUT_BG, h_align="left", wrap=True)
    ds_cell(ws, row, first_col + 2, None, bg=DS_INPUT_BG, h_align="center", bold=True)
    c = ds_cell(ws, row, first_col + 3,
                f'=IF({done}{row}="","",IF(ISERROR({dt}{row}),NOW(),'
                f'IF(OR({dt}{row}="",{dt}{row}=0),NOW(),{dt}{row})))',
                bg=DS_INPUT_BG, h_align="center", number_fmt=FOLLOWUP_DT_FORMAT)
    _DS_FU_ROWS.setdefault(id(ws), []).append(row)
    return c


def _ds_ranges(col_letter, rows):
    """'X4:X9 X12:X15' — contiguous row runs of one column for a DataValidation."""
    out, rows = [], sorted(set(rows))
    i = 0
    while i < len(rows):
        j = i
        while j + 1 < len(rows) and rows[j + 1] == rows[j] + 1:
            j += 1
        out.append(f"{col_letter}{rows[i]}:{col_letter}{rows[j]}" if j > i
                   else f"{col_letter}{rows[i]}")
        i = j + 1
    return " ".join(out)


def ds_followup_apply(ws, first_col, actions):
    """Attach the dropdowns (Action Taken = tab-specific list, Follow-Up Done? =
    Yes/No) to every registered row of this sheet."""
    from openpyxl.worksheet.datavalidation import DataValidation
    rows = _DS_FU_ROWS.pop(id(ws), [])
    if not rows:
        return
    dv_a = DataValidation(type="list", formula1='"' + ",".join(actions) + '"',
                          allow_blank=True, showErrorMessage=True,
                          errorTitle="Action Taken", error="Please pick one of the listed actions.",
                          promptTitle="Action Taken", prompt="Choose the action you took.")
    dv_a.sqref = _ds_ranges(_gcl(first_col), rows)
    dv_d = DataValidation(type="list", formula1='"' + ",".join(FOLLOWUP_DONE_OPTIONS) + '"',
                          allow_blank=True, showErrorMessage=True,
                          errorTitle="Follow-Up Done?", error="Please choose Yes or No.",
                          promptTitle="Follow-Up Done?",
                          prompt="Yes / No — the DateTime column stamps itself.")
    dv_d.sqref = _ds_ranges(_gcl(first_col + 2), rows)
    ws.add_data_validation(dv_a)
    ws.add_data_validation(dv_d)


def ds_enable_iterative_calc(wb):
    """Workbook-level iterative calculation so the self-referencing DateTime
    formula is stable in Excel too (Google Sheets gets the same setting from
    _enable_followup_timestamps() after upload)."""
    try:
        wb.calculation.iterate = True
        wb.calculation.iterateCount = 1
        wb.calculation.iterateDelta = 0.001
        wb.calculation.fullCalcOnLoad = True
    except Exception as exc:                                  # pragma: no cover
        log.warning("Could not set iterative calculation on the workbook: %s", exc)


def ds_finish(ws, header_row, ncols, widths=None, why_col=None, why_width=70,
              default_width=14, filter_from=None, tab_color=None, followup=None,
              freeze_after_col=None):
    """Freeze below the headers, autofilter, column widths, print setup, no
    gridlines. `widths` maps header text → width; unspecified columns get
    `default_width`. The Why-Flagged column gets `why_width`."""
    hdr = {str(ws.cell(row=header_row, column=c).value or "").replace("✎ ", ""): c
           for c in range(1, ncols + 1)}
    for c in range(1, ncols + 1):
        ws.column_dimensions[_gcl(c)].width = default_width
    for name, w in {**FOLLOWUP_WIDTHS, **(widths or {})}.items():
        if name in hdr:
            ws.column_dimensions[_gcl(hdr[name])].width = w
    if followup is not None:
        ds_followup_apply(ws, followup[0], followup[1])
    if why_col:
        ws.column_dimensions[_gcl(why_col)].width = why_width
    ds_fit_guide(ws, ncols)
    last = max(ws.max_row, header_row)
    ws.auto_filter.ref = f"A{filter_from or header_row}:{_gcl(ncols)}{last}"
    # rows above the body always stay; freeze_after_col (1-based) additionally keeps
    # the identity columns up to that column visible while scrolling sideways
    ws.freeze_panes = f"{_gcl(freeze_after_col + 1) if freeze_after_col else 'A'}{header_row + 1}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_title_rows = f"1:{header_row}"
    except Exception:
        pass
    if tab_color:
        ws.sheet_properties.tabColor = tab_color


def _norm_name(s) -> str:
    """Whitespace-collapsed, case-folded name key for safe instructor matching."""
    return " ".join(str(s or "").split()).casefold()


# =============================================================================
#  SMALL PARSERS (aggregate side)
# =============================================================================
def _row_session_min(r):
    """Session length in minutes from session_start_ist → session_end_ist."""
    try:
        s = pd.to_datetime(str(r.get("session_start_ist", "")), errors="coerce")
        e = pd.to_datetime(str(r.get("session_end_ist", "")), errors="coerce")
        if pd.isna(s) or pd.isna(e):
            return None
        m = (e - s).total_seconds() / 60.0
        return round(m, 1) if m > 0 else None
    except Exception:
        return None


# =============================================================================
#  AGGREGATE COMPUTATION  (per student × Tech Name × Duration, till report date)
# =============================================================================
def build_aggregates(att_agg: pd.DataFrame, fb_agg: pd.DataFrame, cfg) -> dict:
    """Return {(student_id, course_name, course_title): {...aggregate metrics...}}.

    att_agg is the full-history attendance frame up to the report date, already
    suspension-excluded by load_all_data. Only attendance-applicable rows
    (is_attendance_required == Y) count — via the existing _applicable() filter.
    """
    out = {}
    if att_agg is None or att_agg.empty:
        return out

    appl = AR._applicable(att_agg)
    if appl.empty:
        return out

    win = int((cfg or {}).get("scoring", {}).get("trend_window_sessions", 5)) if cfg else 5

    # ── feedback grouped by (student, tech, duration) ─────────────────────────
    fb_by_key = {}
    if fb_agg is not None and not fb_agg.empty and "student_id" in fb_agg.columns:
        for _, fr in fb_agg.iterrows():
            k = (str(fr.get("student_id", "")), str(fr.get("course_name", "")),
                 str(fr.get("course_title", "")))
            sess = str(fr.get("session_id", "")) or f"_fb{len(fb_by_key)}"
            rating = fr.get("_rating")
            fb_by_key.setdefault(k, {})[sess] = rating

    group_cols = ["student_id", "course_name", "course_title"]
    for key, grp in appl.groupby(group_cols, sort=False):
        sid, cn, ct = (str(k) for k in key)

        # order this student's sessions in this tech+duration by date
        g = grp.copy()
        g["_d"] = g["_date"]
        g = g.sort_values("_d", na_position="last")

        # distinct sessions (one attendance row per session expected)
        if "session_id" in g.columns:
            g = g.drop_duplicates(subset=["session_id"], keep="last")

        total = len(g)
        present = int((g["status"] == "Present").sum())
        agg_att_pct = round(present / total * 100, 1) if total else 0.0

        # duration-based aggregate: sum(attended min) / sum(session min)
        att_min = sess_min = 0.0
        for _, rr in g.iterrows():
            dm = float(rr.get("_dur_min", 0) or 0)
            sm = _row_session_min(rr)
            if not sm or sm <= 0:
                p = float(rr.get("_pct_num", 0) or 0)      # derive from % if needed
                sm = (dm / (p / 100.0)) if (p > 0 and dm > 0) else None
            if sm and sm > 0:
                att_min += dm
                sess_min += sm
        agg_dur_pct = round(att_min / sess_min * 100, 1) if sess_min > 0 else agg_att_pct

        # latest applicable session
        _dates = [d for d in g["_date"].tolist() if d is not None and not pd.isna(d)]
        latest_date = max(_dates) if _dates else None
        latest_status = ""
        if latest_date is not None:
            _last_rows = g[g["_date"] == latest_date]
            if not _last_rows.empty:
                latest_status = str(_last_rows.iloc[-1].get("status", ""))

        # consecutive absence streak (from most recent backwards)
        streak = 0
        for st in reversed(g["status"].tolist()):
            if st == "Absent":
                streak += 1
            else:
                break

        # last present date
        _pres_dates = [d for d, st in zip(g["_date"].tolist(), g["status"].tolist())
                       if st == "Present" and d is not None and not pd.isna(d)]
        last_present = max(_pres_dates) if _pres_dates else None

        # recent-window trend (last N sessions)
        window = g.tail(win)
        w_present = int((window["status"] == "Present").sum())
        trend_pct = round(w_present / len(window) * 100, 1) if len(window) else agg_att_pct

        # feedback
        fb_map = fb_by_key.get((sid, cn, ct), {})
        n_fb = len(fb_map)
        ratings = [float(rt) for rt in fb_map.values()
                   if rt is not None and not (isinstance(rt, float) and pd.isna(rt))]
        avg_rating = round(sum(ratings) / len(ratings), 1) if ratings else None
        participation = round(min(n_fb / present * 100, 100), 1) if present else 0.0

        rec = {
            "student_id": sid, "technology": cn, "batch_timing": "",
            "student_name": str(grp.iloc[0].get("student_name", "")),
            "phone": str(grp.iloc[0].get("phone", "")), "email": "",
            "latest_date": latest_date, "latest_status": latest_status,
            "attendance_pct": agg_att_pct, "cum_pct": None,
            "streak": streak, "present": present, "total": total,
            "last_present": last_present, "trend_pct": trend_pct,
            "n_feedbacks": n_fb, "participation_pct": participation,
            "avg_rating": avg_rating, "n_ratings": len(ratings),
            # extra display field
            "agg_dur_pct": agg_dur_pct,
        }

        # Why Flagged (+ band) via the coordinator report's validated logic.
        why = ""
        if _COORD_OK and cfg is not None:
            try:
                COORD.score_record(rec, cfg)
                why = rec.get("why", "")
            except Exception as _e:               # pragma: no cover
                why = ""
        rec["why"] = why
        out[(sid, cn, ct)] = rec

    return out


# =============================================================================
#  WORKBOOK BUILDER  (mirrors pyAttendaceFeedbackReport daily Student Detail)
# =============================================================================
def _num_or_blank(v):
    return "" if v is None else v


# =============================================================================
#  LEARNER FOLLOW-UP decision (A attendance / B feedback / C low rating)
# =============================================================================
def _learner_followup_runs(r, agg, fb_pairs, fb_ratings):
    """Rich, action-oriented follow-up reasons for a learner TODAY. Returns a list
    of reasons, each a list of (text, is_red) runs (mirrors the Instructor
    Follow-Ups Why-Flagged style): the issue name, the actual values/numbers and
    the coordinator action are highlighted red; connective words stay neutral.
    Empty list ⇒ the learner is performing fine and is not listed.

      A. Attendance Defaulter — Overall Remarks = Critical AND today Absent/Critical → Call Learner
      B. Feedback Defaulter    — Overall Feedback Part. % < 25% AND no feedback today → Message Learner
      C. Low Feedback Rating   — today's rating <= 7                                  → Understand Rating Reason
    """
    status = str(r.get("status", ""))
    daily_pct = float(r.get("_pct_num", 0) or 0)
    sid = str(r.get("student_id", ""))
    key_rt = (str(r.get("session_id", "")), sid)
    runs = []

    # A ── Attendance Defaulter → Call Learner
    if agg:
        agg_dur = agg["agg_dur_pct"]
        overall_critical = AR._remarks(agg_dur, "").strip().endswith("Critical")
        today_bad = (status == "Absent") or (status == "Present" and daily_pct < CRITICAL_ATT_PCT)
        if overall_critical and today_bad:
            today_run = ("Absent", True) if status == "Absent" else (f"Only {daily_pct:.0f}% Attended", True)
            runs.append([("Critical Attendance", True), (": Overall Attendance ", False),
                         (f"{agg_dur:.0f}%", True),
                         (f" ({agg['present']}/{agg['total']} sessions), Today ", False),
                         today_run, (" — ", False), ("Call Learner", True)])

    # B ── Feedback Defaulter → Message Learner
    if agg:
        part = agg["participation_pct"]
        daily_no_fb = (status != "Absent") and (key_rt not in fb_pairs)
        if part < FEEDBACK_PART_THRESHOLD and daily_no_fb:
            runs.append([("Feedback Missing", True), (": Overall Participation ", False),
                         (f"{part:.0f}%", True),
                         (f" ({agg['n_feedbacks']}/{agg['present']} sessions), Today ", False),
                         ("No Feedback", True), (" — ", False), ("Message Learner", True)])

    # C ── Low Feedback Rating (today) → Understand Rating Reason
    if key_rt in fb_ratings:
        try:
            dr = float(fb_ratings[key_rt])
        except (TypeError, ValueError):
            dr = None
        if dr is not None and dr <= LOW_RATING_THRESHOLD:
            reason = [("Low Rating", True), (": Today ", False), (f"{dr:g}/10", True)]
            if agg and agg.get("avg_rating") is not None:
                reason += [(", Overall Avg ", False), (f"{agg['avg_rating']:g}/10", True)]
            reason += [(" — ", False), ("Understand Rating Reason", True)]
            runs.append(reason)
    return runs


def _learner_followup_reasons(r, agg, fb_pairs, fb_ratings):
    """Plain-string reasons (one per applicable follow-up) — the inclusion/rank
    signal. Derived from _learner_followup_runs so wording stays in one place."""
    return ["".join(t for t, _ in runs)
            for runs in _learner_followup_runs(r, agg, fb_pairs, fb_ratings)]


# Ranking weights — attendance is the highest-priority follow-up, then feedback
# participation, then low ratings. Each condition contributes a base weight plus
# a "how bad" component from the underlying value, and stacking multiple
# conditions adds an extra multi-issue emphasis. Higher score = more severe =
# ranked closer to 1. Nothing here is hard-coded to a student/tech/date.
_SEV_BASE_A = 1000.0   # Attendance Defaulter
_SEV_BASE_B = 500.0    # Feedback Defaulter
_SEV_BASE_C = 300.0    # Low Feedback Rating
_SEV_MULTI  = 200.0    # per additional co-occurring condition


def _learner_followup_severity(r, agg, fb_pairs, fb_ratings):
    """Severity score for one learner-follow-up row, reusing the SAME A/B/C
    conditions as _learner_followup_reasons. Returns (score, n_conditions).
    Higher score = worse performer (→ better/lower Rank number)."""
    status = str(r.get("status", ""))
    daily_pct = float(r.get("_pct_num", 0) or 0)
    sid = str(r.get("student_id", ""))
    key_rt = (str(r.get("session_id", "")), sid)
    score = 0.0
    n = 0

    # A ── Attendance Defaulter (overall Critical + today Absent/Critical)
    if agg:
        agg_dur = agg["agg_dur_pct"]
        if AR._remarks(agg_dur, "").strip().endswith("Critical") and \
                ((status == "Absent") or (status == "Present" and daily_pct < CRITICAL_ATT_PCT)):
            n += 1
            # lower overall attendance → worse; absent today worse than partial
            score += _SEV_BASE_A + max(0.0, CRITICAL_ATT_PCT - agg_dur)
            score += 40.0 if status == "Absent" else max(0.0, CRITICAL_ATT_PCT - daily_pct) / 2.0
            score += min(int(agg.get("streak", 0) or 0), 12) * 5.0   # long absent streak → worse

    # B ── Feedback Defaulter (overall participation < 25% + no feedback today)
    if agg:
        part = agg["participation_pct"]
        if part < FEEDBACK_PART_THRESHOLD and status != "Absent" and key_rt not in fb_pairs:
            n += 1
            score += _SEV_BASE_B + max(0.0, FEEDBACK_PART_THRESHOLD - part)

    # C ── Low Feedback Rating (today's rating <= 7)
    if key_rt in fb_ratings:
        try:
            dr = float(fb_ratings[key_rt])
        except (TypeError, ValueError):
            dr = None
        if dr is not None and dr <= LOW_RATING_THRESHOLD:
            n += 1
            score += _SEV_BASE_C + max(0.0, LOW_RATING_THRESHOLD - dr) * 20.0

    if n > 1:
        score += (n - 1) * _SEV_MULTI          # stacking issues → more severe
    return score, n


def _compute_learner_ranks(att_daily, agg_map, fb_pairs, fb_ratings):
    """Technology-wise severity ranking of learner-follow-up rows. Ranking is
    computed SEPARATELY for each Technology (course_name × course_title), so the
    worst learner in each technology is Rank 1 and the numbering restarts from 1
    when the next technology begins. Returns
    {(student_id, session_id, course_name, course_title): rank}. One rank per
    learner-follow-up record (deduped by that key), so no learner is double-counted."""
    scored = {}
    if att_daily is None or att_daily.empty:
        return {}
    for _, r in att_daily.iterrows():
        if "_attn_applicable" in r.index and not bool(r.get("_attn_applicable", True)):
            continue
        sid = str(r.get("student_id", ""))
        cn, ct = str(r.get("course_name", "")), str(r.get("course_title", ""))
        agg = agg_map.get((sid, cn, ct))
        if not _learner_followup_reasons(r, agg, fb_pairs, fb_ratings):
            continue
        key = (sid, str(r.get("session_id", "")), cn, ct)
        sc, n = _learner_followup_severity(r, agg, fb_pairs, fb_ratings)
        agg_dur = agg["agg_dur_pct"] if agg else 100.0
        streak = int(agg.get("streak", 0) or 0) if agg else 0
        # keep the worst instance if the same key somehow appears twice
        cand = (sc, n, agg_dur, streak, str(r.get("student_name", "")))
        if key not in scored or cand[:2] > scored[key][:2]:
            scored[key] = cand
    # group by Technology (course_name, course_title) and rank WITHIN each group,
    # so every technology restarts at Rank 1 (worst performer first).
    from collections import defaultdict
    by_tech = defaultdict(list)
    for key, v in scored.items():
        by_tech[(key[2], key[3])].append((key, v))
    rank_map = {}
    for _tech, items in by_tech.items():
        items.sort(key=lambda kv: (-kv[1][0], kv[1][2], -kv[1][3], kv[1][4], kv[0]))
        for i, (key, _v) in enumerate(items, 1):
            rank_map[key] = i
    return rank_map


# =============================================================================
#  INSTRUCTOR FOLLOW-UPS  (per-session; reuses AR Session-Summary / Feedback /
#  Teacher-No-Feedback calculations exactly)
# =============================================================================
_NULLish = {"", "nan", "nat", "none", "null", "-", "—"}


def _parse_dt(v):
    """Parse a Sessions-tab timestamp cell to a datetime, or None. Handles the
    'YYYY-MM-DD HH:MM:SS' the collector writes plus ISO-T / with-Z variants, and
    treats blank / null-token cells as missing (never raises)."""
    s = str(v).strip() if v is not None else ""
    if s.lower() in _NULLish:
        return None
    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        # last-ditch explicit formats
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M"):
            try:
                return datetime.strptime(s[:19], fmt)
            except (ValueError, TypeError):
                continue
        return None
    return dt.to_pydatetime() if hasattr(dt, "to_pydatetime") else dt


def _scheduled_min(sess):
    """Scheduled session length (min) from THIS session's own Sessions-tab
    scheduled columns (Session Scheduled Start / End). Returns None when the
    slot is missing or invalid (blank end, or end <= start) so the report shows
    a clean '—' rather than a wrong number. The value is read from the same
    session row that supplies session_id/duration, so the mapping is exact."""
    try:
        s = _parse_dt(sess.get("Session Scheduled Start", ""))
        e = _parse_dt(sess.get("Session Scheduled End", ""))
        if s is None or e is None:
            return None
        m = (e - s).total_seconds() / 60.0
        return round(m, 1) if m > 0 else None
    except Exception:
        return None


def _first_valid_dt(sa, col):
    """First non-blank, parseable datetime in an attendance column, or None."""
    if sa is None or sa.empty or col not in sa.columns:
        return None
    for v in sa[col].tolist():
        d = _parse_dt(v)
        if d is not None:
            return d
    return None


def _actual_session_times(sess, sa):
    """True ACTUAL conducted start/end for a session, plus its duration (min).

    ONE SOURCE OF TRUTH: the actual session start/end is `session_start_ist` /
    `session_end_ist` — the same value the portal shows and that the collector
    also writes into the Attendance / Student_Feedback / Teacher_Feedback rows.
    The Sessions tab's start_time_ist can lag (it keeps the first same-day sync's
    value, which is often the SCHEDULED start, because the start-time watermark
    won't re-admit a later, earlier actual start), so we PREFER the attendance
    `session_start_ist` / `session_end_ist` and fall back to the Sessions row's
    start_time_ist / end_time_ist only when attendance has none. This never uses
    the scheduled slot and is not keyed to any technology/date. Returns
    (start_dt, end_dt, dur_min)."""
    feed_s = _parse_dt(sess.get("start_time_ist", ""))
    feed_e = _parse_dt(sess.get("end_time_ist", ""))
    # authoritative actual from attendance (session_start_ist/session_end_ist)
    obs_s = _first_valid_dt(sa, "session_start_ist")
    obs_e = _first_valid_dt(sa, "session_end_ist")

    a_start = obs_s if obs_s is not None else feed_s
    a_end = obs_e if obs_e is not None else feed_e

    dur = None
    if a_start is not None and a_end is not None:
        m = (a_end - a_start).total_seconds() / 60.0
        dur = round(m, 1) if m > 0 else None
    # last-resort duration: longest single attendance duration (no timestamps)
    if (dur is None or dur == 0) and sa is not None and not sa.empty and "_dur_min" in sa.columns:
        _m = sa["_dur_min"].max()
        if _m and _m > 0:
            dur = round(float(_m), 1)
    return a_start, a_end, dur


def load_instructor_phones(service):
    """Map normalized instructor name -> phone, from the Instructor master tab.
    Only Is_Active == Y rows; instructor_name_i -> alternative_contact_number_i."""
    out = {}
    try:
        df = AR.read_sheet_df(service, INSTRUCTOR_TAB_SHEET_ID, INSTRUCTOR_TAB)
    except Exception as e:
        log.warning("Could not read Instructor tab (%s) — instructor phones blank.", e)
        return out
    if df is None or df.empty:
        return out
    if "Is_Active" in df.columns:
        df = df[df["Is_Active"].astype(str).str.strip().str.upper() == "Y"]
    for _, r in df.iterrows():
        for i in range(1, INSTRUCTOR_NAME_SLOTS + 1):
            nm = _norm_name(r.get(f"instructor_name_{i}", ""))
            ph = str(r.get(f"alternative_contact_number_{i}", "") or "").strip()
            if nm and ph:
                out.setdefault(nm, ph)
    return out


def _finish_instr(ws, row_num):
    ds_finish(ws, 4, IN, widths={
        "Tech Name": 22, "Duration": 26, "Instructor": 18, "Phone": 15, "Session Date": 19,
        "Duration (Mins)": 10, "Scheduled (Mins)": 10, "Diff Mins": 9, "Total Enrolled": 9,
        "Att. N/A Count": 9, "Present": 8, "Absent": 8, "Att %": 8, "Avg Time in Session %": 11,
        "No. of Feedbacks": 10, "Feedback Rate %": 10, "⭐ Avg Rating(/10)": 10,
        "Min Rating": 8, "Max Rating": 8, "Feedback Given": 10,
    }, why_col=_IC["Why Flagged"], why_width=70, default_width=10, tab_color=DS_SECTION,
       followup=(_IC["Action Taken"], FOLLOWUP_ACTIONS["instructor"]),
       freeze_after_col=_IC["Phone"])


def _fmt_clock(dt):
    """datetime -> '7:05 PM' (no leading zero on the hour). '' when missing."""
    if dt is None:
        return ""
    try:
        return dt.strftime("%I:%M %p").lstrip("0")
    except Exception:
        return ""


def _why_flagged_richtext(reasons):
    """Build a Why-Flagged cell from a list of reasons, each reason a list of
    (text, is_red) runs. Key figures/times (is_red=True) render bold red; the
    rest render in a neutral dark colour. Reasons are shown as wrapped bullets.
    Falls back to plain red text if rich text is unavailable."""
    try:
        from openpyxl.cell.rich_text import CellRichText, TextBlock
        from openpyxl.cell.text import InlineFont
        red = InlineFont(rFont="Calibri", sz=10, b=True, color=AR.C_RED_DARK)
        base = InlineFont(rFont="Calibri", sz=10, color="333333")
        blocks = []
        for i, runs in enumerate(reasons):
            if i:
                blocks.append(TextBlock(base, "\n"))
            blocks.append(TextBlock(base, "• "))
            for text, is_red in runs:
                blocks.append(TextBlock(red if is_red else base, text))
        return CellRichText(blocks)
    except Exception:
        # graceful fallback: plain text (whole-cell red handled by caller)
        return "\n".join("• " + "".join(t for t, _ in runs) for runs in reasons)


def build_instructor_followups(ws, sess_daily, att_daily, fb_daily, tf_daily,
                               instr_phones, report_date, period_label=None):
    """Per-SESSION instructor follow-up action list. A session is listed only when
    it trips at least one of the five follow-up conditions:
      1. Not started >= 5 min before scheduled start   → Coordinator Message
      2. Started after the scheduled start (late)        → Coordinator Call
      3. Cancelled (reuses AR._classify_session_type)    → Coordinator attention
      4. Underrun: Diff Mins <= -30                      → Coordinator Call
      5. Instructor feedback missing (Feedback Given=No) → Coordinator Message
    Multiple conditions collapse into ONE record; Why Flagged combines every
    applicable reason with dynamic times/numbers (key figures highlighted red).
    All metrics/columns reuse the existing Session-Summary / Feedback-Rating /
    Teacher-No-Feedback logic exactly.
    Presentation: common report design system (title band, guide strip, grouped
    headers, priority chip on Tech Name, semantic tints only on status/measure
    cells, numbered Why-Flagged reasons with the action on its own line)."""
    period = period_label or report_date.strftime('%d-%b-%Y')
    ds_title(ws, IN, "Instructor Follow-Ups", period,
             "One row per session that needs a coordinator touch. Priority chip on Tech Name: "
             "High = call the instructor, Medium = message / remind.   "
             "Colour key: ■ High = red  ■ Medium = amber  ■ Info = blue  ■ OK = green."
             + DS_GUIDE_FOLLOWUP)
    HDR_ROW = 4
    groups = [(_IC["Tech Name"], _IC["Session Date"], "SESSION"),
              (_IC["Duration (Mins)"], _IC["Diff Mins"], "TIMING (actual vs scheduled)"),
              (_IC["Total Enrolled"], _IC["Avg Time in Session %"], "ATTENDANCE"),
              (_IC["No. of Feedbacks"], _IC["Feedback Given"], "FEEDBACK"),
              (_IC["Why Flagged"], _IC["Why Flagged"], "ACTION"),
              (_IC["Action Taken"], _IC["Follow-Up DateTime"], "COORDINATOR FOLLOW-UP (fill in)")]
    row_num = ds_headers(ws, 3, INSTR_COLS, groups=groups)
    ds_followup_headers(ws, 4, _IC["Action Taken"])

    if sess_daily is None or sess_daily.empty:
        ds_empty(ws, row_num, IN, "No sessions in this period.", level="muted")
        _finish_instr(ws, row_num + 1)
        return ws

    sess_f = AR._prefer_instructor_name(sess_daily)
    # Feedback Given = a Teacher_Feedback row WITH content (same rule as the
    # AR Teacher_No_Feedback tab); a completion-only placeholder row does not count.
    tf_ids = AR.teacher_feedback_session_ids(tf_daily)

    def _dedup_fb(sub):
        if sub is None or sub.empty:
            return sub
        cols = [c for c in ["session_id", "student_id"] if c in sub.columns]
        return sub.drop_duplicates(subset=cols) if cols else sub

    any_rendered = False
    n_rows = 0
    _order = sess_f.sort_values(["course_name", "course_title", "start_time_ist"],
                                na_position="last")
    for _, sess in _order.iterrows():
        sid = sess.get("session_id", "")
        sa = (att_daily[att_daily["session_id"] == sid]
              if (att_daily is not None and not att_daily.empty) else pd.DataFrame())
        sa_app = AR._applicable(sa) if not sa.empty else sa
        na_cnt = AR._na_count(sa) if not sa.empty else 0
        present = sa_app[sa_app["status"] == "Present"] if not sa_app.empty else sa_app
        n_present = len(present)
        total_enrolled = len(sa)
        absent_n = len(sa_app) - n_present
        att_pct = round(n_present / len(sa_app) * 100, 1) if len(sa_app) else 0.0

        # TRUE actual conducted start/end/duration (never the scheduled slot)
        actual_start_dt, actual_end_dt, dur_min = _actual_session_times(sess, sa)
        dur_min = dur_min or 0.0
        avg_time = (round(present["_dur_min"].mean() / dur_min * 100, 1)
                    if (dur_min > 0 and n_present > 0) else 0.0)

        sched_min = _scheduled_min(sess)
        diff_min = round(dur_min - sched_min, 1) if (sched_min is not None) else None

        sf = (_dedup_fb(fb_daily[fb_daily["session_id"] == sid])
              if (fb_daily is not None and not fb_daily.empty) else pd.DataFrame())
        n_fb = len(sf) if sf is not None else 0
        fb_rt = round(n_fb / n_present * 100, 1) if n_present else 0.0
        ratings = (sf["_rating"].dropna().tolist()
                   if (n_fb and sf is not None and "_rating" in sf.columns) else [])
        avg_r = round(sum(ratings) / len(ratings), 2) if ratings else None
        min_r = int(min(ratings)) if ratings else None
        max_r = int(max(ratings)) if ratings else None
        fb_given = sid in tf_ids
        instr = str(sess.get("tutor_name", ""))
        phone = instr_phones.get(_norm_name(instr), "")

        # ── instructor follow-up conditions ──────────────────────────────────
        # Session type reuses AR's exact cancelled/scheduled classifier.
        _end_raw = str(sess.get("end_time_ist", "")).strip()
        _end_blank = _end_raw in ("", "nan", "NaT", "None", "NAN")
        _sess_type = AR._classify_session_type(
            len(sa) > 0, dur_min > 0, _end_blank, str(sess.get("start_time_ist", "")).strip())

        sched_start_dt = _parse_dt(sess.get("Session Scheduled Start", ""))
        # actual_start_dt already resolved to the TRUE actual by _actual_session_times

        # reasons: list of reason-run-lists → [(text, is_red), …]; is_red = key figure
        reasons = []
        if _sess_type == "cancelled":
            # 3. Session Cancelled — coordinator may be unaware
            reasons.append([("Session Cancelled", True),
                            (" — Coordinator may be unaware; confirm with the instructor.", False)])
        elif _sess_type == "scheduled":
            pass                                   # not yet conducted → not a follow-up
        else:
            # 1 & 2. Start-time checks (need both scheduled & actual start)
            if sched_start_dt is not None and actual_start_dt is not None:
                late_min = round((actual_start_dt - sched_start_dt).total_seconds() / 60.0)
                sched_txt, actual_txt = _fmt_clock(sched_start_dt), _fmt_clock(actual_start_dt)
                if late_min > 0:
                    # 2. Session Not Started On Time → Coordinator Call
                    reasons.append([("Started Late: Scheduled ", False), (sched_txt, True),
                                    (", Actual ", False), (actual_txt, True),
                                    (" (", False), (f"{late_min} min late", True),
                                    (") — Call Instructor", False)])
                elif late_min > -INSTR_EARLY_MIN:
                    # 1. Session Not Started 5 Mins Prior → Coordinator Message
                    exp_txt = _fmt_clock(sched_start_dt - timedelta(minutes=INSTR_EARLY_MIN))
                    var_min = int(round(INSTR_EARLY_MIN + late_min))   # min after expected early start
                    reasons.append([("Not Started 5 Min Early: Scheduled ", False), (sched_txt, True),
                                    (f", expected by {exp_txt}", True),
                                    (", Actual ", False), (actual_txt, True),
                                    (" (", False), (f"{var_min} min late vs expected", True),
                                    (") — Message Instructor", False)])
            # 4. Session Underrun (Diff Mins <= -30) → Coordinator Call
            if diff_min is not None and diff_min <= -INSTR_UNDERRUN_MIN:
                under = abs(int(round(diff_min)))
                reasons.append([("Session Underrun: Scheduled ", False), (f"{sched_min:.0f} min", True),
                                (", Actual ", False), (f"{dur_min:.0f} min", True),
                                (" (", False), (f"{under} min short", True),
                                (") — Call Instructor", False)])
            # 5. Instructor Feedback Missing → Coordinator Message
            if not fb_given:
                reasons.append([("Feedback Missing", True),
                                (" — instructor feedback for this session is pending. Message Instructor.", False)])
            # 6. Low Feedback Rate (student feedback participation <= 25%) →
            #    Coordinator asks students to submit their feedback. Combined into
            #    this session's existing reasons (never a duplicate record).
            if n_present > 0 and fb_rt <= INSTR_LOW_FB_RATE_PCT:
                reasons.append([("Low Feedback Rate", True), (": Only ", False),
                                (f"{n_fb} of {n_present}", True),
                                (" students submitted feedback (", False),
                                (f"{fb_rt:g}%", True), (") — ", False),
                                ("Ask students to submit their feedback.", True)])
        if not reasons:
            continue

        # ── render ───────────────────────────────────────────────────────────
        # Priority = the escalation the reasons themselves ask for: a "Call"
        # reason (late start, underrun) or a cancelled session = High; message-
        # level reasons (not 5 min early, feedback missing, low fb rate) = Medium.
        _flat = " ".join("".join(t for t, _ in r) for r in reasons).lower()
        level = ("high" if ("call instructor" in _flat or "session cancelled" in _flat)
                 else "medium")
        bg = ds_zebra(n_rows)
        ds_priority(ws, row_num, _IC["Tech Name"], sess.get("course_name", ""), level,
                    h_align="left")
        ds_cell(ws, row_num, _IC["Duration"], sess.get("course_title", ""), bg=bg)
        ds_cell(ws, row_num, _IC["Instructor"], instr, bg=bg, bold=True)
        ds_cell(ws, row_num, _IC["Phone"], phone or "—", bg=bg, h_align="center")
        ds_cell(ws, row_num, _IC["Session Date"],
                (actual_start_dt.strftime("%Y-%m-%d %H:%M:%S") if actual_start_dt is not None
                 else str(sess.get("start_time_ist", ""))[:19]), bg=bg, h_align="center")
        ds_cell(ws, row_num, _IC["Duration (Mins)"], dur_min if dur_min else "", bg=bg,
                h_align="center", number_fmt="0.0")
        ds_cell(ws, row_num, _IC["Scheduled (Mins)"],
                (round(sched_min, 1) if sched_min is not None else "—"), bg=bg,
                h_align="center", number_fmt="0.0")
        # Diff Mins — short / long / on-time tint (same thresholds as before)
        if diff_min is not None:
            if diff_min <= -INSTR_SHORT_MIN:
                dlev = "high"
            elif diff_min >= INSTR_SHORT_MIN:
                dlev = "medium"
            else:
                dlev = "ok"
            ds_pill(ws, row_num, _IC["Diff Mins"], f"{diff_min:+.0f}", dlev)
        else:
            ds_cell(ws, row_num, _IC["Diff Mins"], "—", bg=bg, h_align="center")
        for hname, val in (("Total Enrolled", total_enrolled), ("Att. N/A Count", na_cnt),
                           ("Present", n_present), ("Absent", absent_n)):
            ds_cell(ws, row_num, _IC[hname], val, bg=bg, h_align="center")
        ds_cell(ws, row_num, _IC["Att %"], att_pct, bg=ds_att_bg(att_pct, ""),
                h_align="center", number_fmt='0.0"%"')
        ds_cell(ws, row_num, _IC["Avg Time in Session %"], avg_time,
                bg=ds_att_bg(avg_time, ""), h_align="center", number_fmt='0.0"%"')
        ds_cell(ws, row_num, _IC["No. of Feedbacks"], n_fb, bg=bg, h_align="center")
        ds_pill(ws, row_num, _IC["Feedback Rate %"], fb_rt,
                ("ok" if fb_rt >= 50 else "medium" if fb_rt > 0 else "high"),
                bold=False, number_fmt='0.0"%"')
        if avg_r is not None:
            ds_cell(ws, row_num, _IC["⭐ Avg Rating(/10)"], avg_r, bg=ds_rating_bg(avg_r),
                    h_align="center", number_fmt='0.0')
        else:
            ds_cell(ws, row_num, _IC["⭐ Avg Rating(/10)"], "—", bg=bg, h_align="center")
        ds_cell(ws, row_num, _IC["Min Rating"], (min_r if min_r is not None else "—"),
                bg=bg, h_align="center")
        ds_cell(ws, row_num, _IC["Max Rating"], (max_r if max_r is not None else "—"),
                bg=bg, h_align="center")
        ds_pill(ws, row_num, _IC["Feedback Given"], "Yes" if fb_given else "No",
                "ok" if fb_given else "high")
        ds_why_text(ws, row_num, _IC["Why Flagged"], reasons, col_width=70)
        ds_followup_cells(ws, row_num, _IC["Action Taken"])
        any_rendered = True
        n_rows += 1
        row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, IN, "✅  No instructor follow-ups required for this period.")
        row_num += 1
    ds_guide_count(ws, "Sessions needing follow-up", n_rows)
    _finish_instr(ws, row_num)
    return ws


# =============================================================================
#  LEARNER ASSIGNMENT FOLLOW-UPS  (assignment-submission defaulters)
#  Reuses pyAssignmentSubmissionEmailReminder's EXACT defaulter/reminder logic —
#  no separate definition of "defaulter" is created here.
# =============================================================================
def load_assignment_followups(service, report_date):
    """Return (assignment-details, tech→[assignment groups]) for the assignment
    defaulters ON report_date, using the reminder script's own logic:
    find_pending_reminders → _dedup_records → consolidate_by_student →
    _group_rows_by_course_assignment. Returns {} on any failure so the tab simply
    shows an empty banner and the rest of the report is unaffected."""
    try:
        import pyAssignmentSubmissionEmailReminder as ASG   # bootstrap already on sys.path
    except Exception as e:                                  # pragma: no cover
        log.warning("Assignment reminder module unavailable (%s) — "
                    "Learner Assignment Follow-Ups will be empty.", e)
        return {}
    try:
        subs = ASG.read_sheet_df(service, ASG.SUBMISSION_SHEET_ID, ASG.SUBMISSIONS_TAB)
        recs = ASG.find_pending_reminders(subs, report_date)   # same defaulter rule
        recs = ASG._dedup_records(recs)
        consolidated = ASG.consolidate_by_student(recs)
        groups = ASG._group_rows_by_course_assignment(consolidated)
        # regroup by Technology (class_name) → list of (title, aid, rows)
        from collections import defaultdict as _dd
        by_tech = _dd(list)
        for (cn, title, aid), rows in groups.items():
            by_tech[cn].append((title, aid, rows))
        log.info("Assignment defaulters: %d assignment group(s) across %d technology(ies).",
                 len(groups), len(by_tech))
        return dict(by_tech)
    except Exception as e:                                  # pragma: no cover
        log.warning("Could not compute assignment defaulters (%s).", e)
        return {}


def _assignment_why_runs(row):
    """Action-oriented Why-Flagged runs for one assignment-defaulter row, driven
    by the reminder STAGE (never hard-coded). Highlights the reminder stage, the
    deadline and the action in red (same rich-text style as the other tabs)."""
    lvl = str(row.get("reminder_level", ""))
    label = str(row.get("reminder_label", "") or lvl or "Reminder")
    ddl = str(row.get("deadline_str", "") or "")
    action = _ASSIGN_ACTION.get(lvl, "Follow up with the learner for assignment submission")
    runs = [(label, True), (": ", False)]
    if lvl == "missed":
        runs += [("Deadline passed", True)]
        if ddl:
            runs += [(f" (was due {ddl})", True)]
    elif lvl == "final":
        runs += [("Final submission day", True)]
        if ddl:
            runs += [(f" (due {ddl})", True)]
    else:
        runs += [("Assignment due ", False), (ddl or "soon", True)]
    runs += [(" — ", False), (action, True), (".", False)]
    return [[r for r in runs if r[0]]]


def _assign_banner(ws, row_num, ncols, text, bg, txt_color="FFFFFF", h_align="left"):
    """Full-width merged banner row (technology / assignment-detail header)."""
    ws.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=ncols)
    c = ws.cell(row=row_num, column=1)
    c.value = text
    c.fill = AR._fill(bg)
    c.font = AR._font(bold=True, size=10, color=txt_color)
    c.alignment = AR._align(h_align, "center")
    c.border = AR._border()
    ws.row_dimensions[row_num].height = 20


def build_assignment_followups(ws, by_tech, report_date, period_label=None):
    """Learner Assignment Follow-Ups tab. Grouped Technology-wise (like the
    Learner Attendance Follow-Ups tab); each Technology shows the relevant
    Assignment Details header, then every pending-submission learner with an
    action-oriented Why Flagged. All defaulter identification/data is reused from
    pyAssignmentSubmissionEmailReminder.
    Presentation: technology = section banner, assignment details = pale
    sub-banner, '#' cell = priority chip (Missed / Final = High, 2nd = Medium,
    1st = Info), Reminder stage as a chip, neutral zebra rows."""
    period = period_label or report_date.strftime('%d-%b-%Y')
    ds_title(ws, AFN, "Learner Assignment Follow-Ups", period,
             "Learners with a pending assignment submission, grouped by technology (blue banner: "
             "Tech Name · Duration · pending count) and assignment (pale banner). "
             "Priority chip on '#': High = deadline missed / final day, Medium = 2nd reminder, "
             "Info = 1st reminder.   Colour key: ■ High = red  ■ Medium = amber  ■ Info = blue."
             + DS_GUIDE_FOLLOWUP)
    HDR_ROW = 3
    row_num = ds_headers(ws, HDR_ROW, AF_COLS)
    ds_followup_headers(ws, HDR_ROW, _AF["Action Taken"])

    any_rendered = False
    n_rows = 0
    try:
        import pyAssignmentSubmissionEmailReminder as ASG
        _lp = ASG.LEVEL_PRIORITY
    except Exception:
        _lp = {"1st": 1, "2nd": 2, "final": 3, "missed": 4}

    for cn in sorted((by_tech or {}).keys(), key=lambda x: str(x).lower()):
        assignments = by_tech[cn]
        n_tech = sum(len(rows) for _t, _a, rows in assignments if rows)
        # Duration(s) of this technology's pending assignments (normally one batch)
        durations = []
        for _t, _a, rows in assignments:
            if rows:
                d = str(rows[0].get("class_subject", "") or "").strip()
                if d and d not in durations:
                    durations.append(d)
        one_duration = (len(durations) == 1)
        head = f"Tech Name: {cn or '(Unknown Technology)'}"
        if durations:
            head += f"   ·   Duration: {' / '.join(durations)}"
        head += f"   ·   {n_tech} pending"
        row_num = ds_section(ws, row_num, AFN, head, level=1)

        for title_txt, aid, rows in sorted(assignments, key=lambda t: str(t[0]).lower()):
            if not rows:
                continue
            r0 = rows[0]
            # Assignment Details header (reuses the same meta shown in the PDF)
            _assigned = r0.get("assigned_date_str", "") or "—"
            details = (f"📝  {title_txt or '(Untitled Assignment)'}"
                       + ("" if one_duration else
                          f"      •  Duration: {r0.get('class_subject','') or '—'}")
                       + f"      •  Max Marks: {r0.get('maximum_marks','') or '—'}"
                       f"      •  Assigned: {_assigned}"
                       f"      •  Deadline: {r0.get('deadline_str','') or '—'}"
                       f"      •  Pending: {len(rows)}")
            row_num = ds_section(ws, row_num, AFN, details, level=2)

            # learners, most-urgent reminder first (same sort as the PDF)
            rows_sorted = sorted(
                rows, key=lambda r: (-_lp.get(r.get("reminder_level", ""), 0),
                                     str(r.get("student_name", "")).lower()))
            for i, r in enumerate(rows_sorted, 1):
                lvl = str(r.get("reminder_level", ""))
                level = ("high" if lvl in ("final", "missed") else "medium" if lvl == "2nd" else "info")
                bg = ds_zebra(i)
                ds_priority(ws, row_num, _AF["#"], i, level)
                ds_cell(ws, row_num, _AF["Student Name"], r.get("student_name", "") or "—",
                        bg=bg, bold=True)
                ds_cell(ws, row_num, _AF["Email"], r.get("student_email", "") or "—", bg=bg)
                ds_cell(ws, row_num, _AF["Phone"], r.get("student_phone", "") or "—", bg=bg,
                        h_align="center")
                ds_cell(ws, row_num, _AF["Deadline"], r.get("deadline_str", "") or "—", bg=bg,
                        h_align="center", bold=(level == "high"),
                        fg=(AR.C_RED_DARK if level == "high" else DS_TEXT))
                ds_pill(ws, row_num, _AF["Reminder"], r.get("reminder_label", "") or "—", level)
                ds_why_text(ws, row_num, _AF["Why Flagged"], _assignment_why_runs(r), col_width=64)
                ds_followup_cells(ws, row_num, _AF["Action Taken"])
                any_rendered = True
                n_rows += 1
                row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, AFN, "✅  No assignment submission follow-ups for this day.")
        row_num += 1
    ds_guide_count(ws, "Pending submissions", n_rows)

    ds_finish(ws, HDR_ROW, AFN, widths={
        "#": 6, "Student Name": 24, "Email": 32, "Phone": 16, "Deadline": 13, "Reminder": 16,
    }, why_col=_AF["Why Flagged"], why_width=64, default_width=14, tab_color=DS_SECTION,
       followup=(_AF["Action Taken"], FOLLOWUP_ACTIONS["assignment"]),
       freeze_after_col=_AF["Phone"])
    return ws


# =============================================================================
#  LEARNER ADMISSION FORMALITIES  (admission-form / e-signature status)
#  Reuses pyAdmissionFormalitiesReport's EXACT matching + status logic — the
#  "pending admission formalities" set and every value are taken as-is; only an
#  action-oriented "Why Flagged" column is added for the coordinator.
# =============================================================================
def load_admission_formalities(service):
    """Return the still-pending admission-formality rows, using
    pyAdmissionFormalitiesReport's own logic end-to-end (read_records →
    build_signer_indexes → build_rows with current_only=True, so already-signed
    students are excluded exactly as the source 'Current' report does). Every
    field/value is taken as-is. Returns [] on any failure so the tab shows an
    empty banner and the rest of the report is unaffected."""
    try:
        import pyAdmissionFormalitiesReport as ADM   # bootstrap already on sys.path
    except Exception as e:                            # pragma: no cover
        log.warning("Admission-formalities module unavailable (%s) — "
                    "Learner Admission Formalities will be empty.", e)
        return []
    try:
        students = ADM.read_records(service, ADM.STUDENT_SHEET_ID, ADM.STUDENTS_TAB)
        signers  = ADM.read_records(service, ADM.SIGNERS_SHEET_ID, ADM.SIGNERS_TAB)
        by_email, by_name = ADM.build_signer_indexes(signers)
        start_date = ADM.to_ist_date(ADM.CURRENT_REPORT_START_DATE) or date(2026, 7, 1)
        rows, _stats = ADM.build_rows(students, by_email, by_name,
                                      current_only=True, start_date=start_date)
        # newest joiners first (same ordering the source 'Current' report uses)
        rows.sort(key=lambda r: r.get("Joined On", ""), reverse=True)
        log.info("Admission formalities: %d pending student(s).", len(rows))
        return rows
    except Exception as e:                            # pragma: no cover
        log.warning("Could not compute admission formalities (%s).", e)
        return []


def _adm_norm(v) -> str:
    return " ".join(str(v or "").split()).strip().lower()


def _adm_is_not_sent(recipient_status) -> bool:
    """True when no signing request has effectively been sent — covers the
    reference sentinel 'Form Not Sent' and a plain 'Not Sent'."""
    n = _adm_norm(recipient_status)
    return (n == "" or "not sent" in n)


def _admission_why_runs(row):
    """Action-oriented Why-Flagged runs for one admission-formality row, driven
    only by Recipient Status (never hard-coded per student). Key status, dates/
    expiry and the required action render bold red (same rich-text style as the
    other follow-up tabs).

      • Not Sent → Send the form + share the guideline email for signing +
        call the learner to get the formalities completed.
      • Any other pending/incomplete status → Call the learner to complete the
        pending formalities, with the relevant status/details shown dynamically.
    """
    rs_raw = str(row.get("Recipient Status", "") or "").strip()
    qs_raw = str(row.get("Request Status", "") or "").strip()
    sent   = str(row.get("Sent Date", "") or "").strip()
    expiry = str(row.get("Expiry Date", "") or "").strip()
    status_label = rs_raw or "Not Sent"

    if _adm_is_not_sent(rs_raw):
        runs = [
            (status_label, True),
            (" — Send the admission form, share the guideline email for signing, and ", False),
            ("call the learner", True),
            (" to get the formalities completed.", False),
        ]
        return [runs]

    runs = [
        ("Call the learner", True),
        (" to complete the pending admission formalities", False),
        (" — status: ", False), (status_label, True),
    ]
    if qs_raw and _adm_norm(qs_raw) != _adm_norm(status_label):
        runs += [(", request ", False), (qs_raw, True)]
    if sent:
        runs += [(", sent ", False), (sent, False)]
    if expiry:
        runs += [(", ", False), ("expires " + expiry, True)]
    runs += [(".", False)]
    return [[r for r in runs if r[0]]]


def _adm_bg(recipient_status, request_status):
    """Row background for one admission row, reusing ADM.row_colour's EXACT
    status classification and mapping it onto this report's openpyxl palette."""
    if not _ADM_OK:
        return AR.C_RED_LITE if _adm_is_not_sent(recipient_status) else AR.C_WHITE
    try:
        c = ADM.row_colour(recipient_status, request_status)
    except Exception:
        c = None
    if c == ADM.COL_RED:
        return AR.C_RED_LITE
    if c == ADM.COL_GREEN:
        return AR.C_GREEN
    if c == ADM.COL_LIGHT_GREEN:
        return AR.C_GREEN_PALE
    if c == ADM.COL_AMBER:
        return AR.C_AMBER
    return AR.C_WHITE


def _adm_severity_key(row):
    """Sort key: most-urgent first — red (not-sent / expired / declined) → amber
    (sent, unopened) → others. Reuses ADM.row_colour's classification (no new
    severity rule); stable sort preserves the newest-joined order within a band."""
    bg = _adm_bg(row.get("Recipient Status", ""), row.get("Request Status", ""))
    return -{AR.C_RED_LITE: 3, AR.C_AMBER: 2}.get(bg, 1)


def build_admission_formalities(ws, rows, report_date, period_label=None):
    """Learner Admission Formalities tab. Lists learners with PENDING admission
    formalities (identification, status and every value reused as-is from
    pyAdmissionFormalitiesReport), plus one action-oriented 'Why Flagged' column
    immediately after 'Expiry Date'.
    Presentation: Student Name = priority chip (red = not sent / expired /
    declined, amber = sent but unopened, blue = other pending — the reference
    report's own classification), Recipient Status as a chip, expiry date bold red
    when live, neutral zebra rows, numbered Why Flagged with the action line."""
    period = period_label or report_date.strftime('%d-%b-%Y')
    ds_title(ws, ADMN, "Learner Admission Formalities", period,
             "Learners whose admission form / e-signature is still pending, most urgent first. "
             "Priority chip on Student Name: High = form not sent, expired or declined; "
             "Medium = sent but not opened; Info = other pending.   "
             "Colour key: ■ High = red  ■ Medium = amber  ■ Info = blue." + DS_GUIDE_FOLLOWUP)
    HDR_ROW = 3
    row_num = ds_headers(ws, HDR_ROW, ADM_COLS)
    ds_followup_headers(ws, HDR_ROW, _ADMC["Action Taken"])

    _text_cols = {"Student Name", "Email ID", "Batch Name", "Request Form Name",
                  "Recipient Status", "Request Status"}
    any_rendered = False
    n_rows = 0
    # most-urgent first (red → amber → rest); stable, so newest-joined order stays
    ordered = sorted(rows or [], key=_adm_severity_key)
    for r in ordered:
        rs = r.get("Recipient Status", "")
        qs = r.get("Request Status", "")
        bg_ref = _adm_bg(rs, qs)                 # reference classification (unchanged)
        level = ("high" if bg_ref == AR.C_RED_LITE else "medium" if bg_ref == AR.C_AMBER
                 else "ok" if bg_ref in (AR.C_GREEN, AR.C_GREEN_PALE) else "info")
        bg = ds_zebra(n_rows)
        for c in _ADM_BASE_COLS:
            val = r.get(c, "")
            ds_cell(ws, row_num, _ADMC[c], (val if str(val).strip() else "—"), bg=bg,
                    h_align="left" if c in _text_cols else "center")
        # Student Name — the priority chip (who is affected + how urgent)
        ds_priority(ws, row_num, _ADMC["Student Name"],
                    (r.get("Student Name", "") if str(r.get("Student Name", "")).strip() else "—"),
                    level, h_align="left")
        # Recipient Status — highlight the key status as a chip.
        ds_pill(ws, row_num, _ADMC["Recipient Status"],
                (rs if str(rs).strip() else "—"), level, h_align="left")
        # Expiry Date — a live deadline; flag red when present on an attention row.
        if str(r.get("Expiry Date", "") or "").strip() and level in ("high", "medium"):
            ds_cell(ws, row_num, _ADMC["Expiry Date"], r.get("Expiry Date", ""), bg=bg,
                    h_align="center", bold=True, fg=AR.C_RED_DARK)
        ds_why_text(ws, row_num, _ADMC["Why Flagged"], _admission_why_runs(r), col_width=66)
        ds_followup_cells(ws, row_num, _ADMC["Action Taken"])
        any_rendered = True
        n_rows += 1
        row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, ADMN, "✅  No pending admission formalities for this day.")
        row_num += 1
    ds_guide_count(ws, "Learners with pending formalities", n_rows)

    ds_finish(ws, HDR_ROW, ADMN, widths={
        "Student Name": 24, "Email ID": 30, "Phone Number": 16, "Batch Name": 14,
        "Joined On": 12, "Request Form Name": 30, "Recipient Status": 20,
        "Request Status": 30, "Sent Date": 20, "Signed Date": 20, "Expiry Date": 20,
    }, why_col=_ADMC["Why Flagged"], why_width=66, default_width=14, tab_color=DS_SECTION,
       followup=(_ADMC["Action Taken"], FOLLOWUP_ACTIONS["admission"]),
       freeze_after_col=_ADMC["Phone Number"])
    return ws


# =============================================================================
#  LEARNER WISE VALIDATION  (combined Student + Course + Instructor data checks)
#  Reuses pyWiseDataValidationReport's EXACT validation logic (build_student_rows /
#  build_course_rows / build_instructor_rows). Only the *failed / attention-required*
#  records it already returns are re-presented in ONE coordinator tab, each section
#  separated by a header banner and ending in an action-oriented, red-highlighted
#  "Why Flagged" column. No validation business rule is (re)defined here.
# =============================================================================
_WISE_SEV = {"Invalid": 3, "Missing": 2, "Warning": 1}    # record severity ranking
_WISE_ACTION = {
    "Invalid": "correct this value in Wise",
    "Missing": "fill in this field in Wise",
    "Warning": "review and confirm",
}
_WISE_STU_FIELDS = [   # (label, status key, description key or None, synthetic desc)
    ("Student Name",    "Student Name Status",    "Student Name Validation Description", ""),
    ("Email ID",        "Email ID Status",        "Email ID Validation Description", ""),
    ("Phone Number",    "Phone Number Status",    "Phone Number Validation Description", ""),
    ("Tag Name",        "Tag Name Status",        "Tag Name Validation Description", ""),
    ("Private Note",    "Private Note Status",    None, "Private note not captured for this student."),
    ("Profile Picture", "Profile Picture Status", None, "Profile picture not uploaded for this student."),
]
_WISE_CRS_FIELDS = [
    ("Course Title",    "Course Title Status",    "Course Title Validation Description"),
    ("Course Subtitle", "Course Subtitle Status", "Course Subtitle Validation Description"),
    ("Course Tag Name", "Course Tag Name Status", "Course Tag Name Validation Description"),
]


def load_wise_validation():
    """Return {'student':[...], 'course':[...], 'instructor':[...]} — the failed /
    attention-required validation records, produced by pyWiseDataValidationReport's
    OWN logic end-to-end (identical to its standalone run). Returns None on any
    failure so the tab shows an 'unavailable' banner and the rest of the report is
    unaffected. No validation rule is changed here."""
    try:
        import pyWiseDataValidationReport as WISE   # bootstrap already on sys.path
    except Exception as e:                          # pragma: no cover
        log.warning("Wise validation module unavailable (%s) — "
                    "Learner Wise Validation will be empty.", e)
        return None
    try:
        ref = WISE.load_reference_data()
        sheets, _drive = WISE.google_services()
        student_src    = WISE.read_tab(sheets, WISE.SOURCE_SHEET_ID, WISE.STUDENTS_TAB)
        combined_src   = WISE.read_tab(sheets, WISE.SOURCE_SHEET_ID, WISE.COMBINED_TAB)
        instructor_src = WISE.read_tab(sheets, WISE.SOURCE_SHEET_ID, WISE.INSTRUCTOR_TAB)

        faculty_names, _lc = WISE.load_faculty_names(sheets)
        faculty_initials = {WISE.to_first_initial(n) for n in faculty_names if WISE.to_first_initial(n)}
        faculty_initials |= {WISE.to_first_initial(n) for n in ref["old_instructor_names"]
                             if WISE.to_first_initial(n)}
        onb_pool = WISE.build_onboarding_pool(WISE.load_onboarding_records(sheets))
        future_class_ids = WISE.load_future_session_class_ids()
        enrolled_ids = {str(r.get("student_id", "")).strip() for r in combined_src
                        if str(r.get("student_id", "")).strip()}

        student_rows, _st    = WISE.build_student_rows(student_src, ref, enrolled_ids)
        course_rows, _ct     = WISE.build_course_rows(combined_src, faculty_initials, ref, future_class_ids)
        instructor_rows, _it = WISE.build_instructor_rows(instructor_src, onb_pool)
        log.info("Wise validation: %d student, %d course, %d instructor flagged record(s).",
                 len(student_rows), len(course_rows), len(instructor_rows))
        return {"student": student_rows, "course": course_rows, "instructor": instructor_rows}
    except Exception as e:                          # pragma: no cover
        log.warning("Could not compute Wise validation (%s).", e)
        return None


def _wise_status_style(status):
    """(bg, font) for a validation-status cell, matching the report palette."""
    s = str(status or "").strip()
    if s == "Invalid":
        return AR.C_RED_LITE, AR.C_RED_DARK
    if s == "Missing":
        return AR.C_AMBER, AR.C_AMBER_DARK
    if s == "Warning":
        return AR.C_AMBER_PALE, AR.C_AMBER_DARK
    if s == "Valid":
        return AR.C_GREEN_PALE, AR.C_GREEN_DARK
    return AR.C_WHITE, "333333"


def _wise_sev_bg(sev):
    return {3: AR.C_RED_LITE, 2: AR.C_AMBER, 1: AR.C_AMBER_PALE}.get(sev, AR.C_WHITE)


def _wise_field_reason(label, status, desc):
    """One action-oriented Why-Flagged bullet (list of (text, is_red) runs) for a
    flagged field: the field, the Invalid/Missing/Warning status and the required
    action render bold red; the existing validation description shows as-is."""
    runs = [(label, True), (": ", False), (status, True)]
    d = str(desc or "").strip()
    if d:
        runs += [(" — ", False), (d, False)]
    act = _WISE_ACTION.get(str(status or "").strip())
    if act:
        runs += [(" → ", False), (act, True)]
    return [r for r in runs if r[0]]


def _wise_instr_action(vtype):
    t = str(vtype or "").lower()
    if "missing" in t:        return "add the missing contact detail"
    if "not found" in t:      return "add the instructor to Faculty Onboarding"
    if "wrong sequence" in t: return "fix the contact-to-name mapping"
    if "mismatch" in t:       return "correct the contact detail"
    if "assignment" in t:     return "assign the instructor name(s)"
    return "correct the instructor record"


def _wise_instr_reason(vtype, msg):
    runs = [(str(vtype or "Issue"), True), (": ", False)]
    m = str(msg or "").strip()
    if m:
        runs += [(m, False)]
    runs += [(" → ", False), (_wise_instr_action(vtype), True)]
    return [r for r in runs if r[0]]


def _wise_joined_on_ist(value):
    """Display-only: render a Wise 'Joined On' timestamp in IST.
    Wise returns it as UTC ISO-8601 (e.g. '2026-09-29T11:04:12.729Z'); the same
    instant is shown as '2026-09-29 16:34:12 IST' — the format the Course
    section's 'Created On' already uses. A value that is not a parseable
    timestamp is shown exactly as received; the record itself is never changed."""
    raw = str(value or "").strip()
    if not raw:
        return raw
    txt = raw[:-1] + "+00:00" if raw.endswith(("Z", "z")) else raw
    try:
        dt = datetime.fromisoformat(txt)
    except ValueError:
        try:                                           # e.g. '2026-09-29 11:04:12'
            dt = datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return raw
    if dt.tzinfo is None:                              # naive → treat as UTC (Wise export)
        dt = dt.replace(tzinfo=timezone.utc)
    ist = dt.astimezone(timezone(timedelta(hours=5, minutes=30)))
    return ist.strftime("%Y-%m-%d %H:%M:%S IST")


def _wise_student_display(records):
    disp = []
    for i, rec in enumerate(records or [], 1):
        reasons, sev, statuses = [], 0, {}
        for k, (label, skey, dkey, syn) in enumerate(_WISE_STU_FIELDS):
            st = str(rec.get(skey, "")).strip()
            statuses[3 + k] = st                      # status cells at data index 3..8
            rank = _WISE_SEV.get(st, 0)
            if rank > 0:
                reasons.append(_wise_field_reason(label, st, rec.get(dkey, "") if dkey else syn))
                sev = max(sev, rank)
        cells = [i, rec.get("Student Name", "") or "—", rec.get("Batch Name", "") or "—",
                 "", "", "", "", "", "", _wise_joined_on_ist(rec.get("Joined On", "")) or "—"]
        disp.append({"cells": cells, "status_cells": statuses,
                     "severity": sev, "why_reasons": reasons})
    return disp


def _wise_course_display(records):
    disp = []
    for i, rec in enumerate(records or [], 1):
        reasons, sev, statuses = [], 0, {}
        for k, (label, skey, dkey) in enumerate(_WISE_CRS_FIELDS):
            st = str(rec.get(skey, "")).strip()
            statuses[4 + k] = st                      # status cells at data index 4..6
            rank = _WISE_SEV.get(st, 0)
            if rank > 0:
                reasons.append(_wise_field_reason(label, st, rec.get(dkey, "")))
                sev = max(sev, rank)
        cells = [i, rec.get("Course Title", "") or "—", rec.get("Course Subtitle", "") or "—",
                 rec.get("Instructor_Name", "") or "—", "", "", "", rec.get("Created On", "") or "—"]
        disp.append({"cells": cells, "status_cells": statuses,
                     "severity": sev, "why_reasons": reasons})
    return disp


def _wise_instructor_display(records):
    """Group the reference's per-check instructor rows by (Instructor ID, Name) so a
    single instructor with several failed checks is ONE record with a combined Why
    Flagged (no duplicate records)."""
    from collections import OrderedDict
    groups = OrderedDict()
    for rec in records or []:
        key = (str(rec.get("Instructor ID", "")).strip(),
               str(rec.get("Instructor Name", "")).strip())
        groups.setdefault(key, []).append(rec)
    disp = []
    for i, ((iid, iname), checks) in enumerate(groups.items(), 1):
        reasons, hard = [], False
        for chk in checks:
            vtype = chk.get("Validation Type", "")
            reasons.append(_wise_instr_reason(vtype, chk.get("Detailed Validation Message", "")))
            if "missing" not in str(vtype).lower():
                hard = True
        cells = [i, iid or "—", iname or "—", len(checks)]
        disp.append({"cells": cells, "status_cells": {},
                     "severity": (3 if hard else 2), "why_reasons": reasons})
    return disp


def _wise_render_section(ws, start_row, ncols_max, banner_text, data_headers,
                         display_rows, text_col_idxs):
    """Render one validation section: a section banner, a header row (the data
    headers, then 'Why Flagged' spanning every remaining column so it is wide and
    aligned across sections), then the failed records — first cell = priority
    chip (Invalid = High, Missing = Medium, Warning = Info), Valid / Missing /
    Invalid statuses as chips, neutral zebra rows. Returns the next free row
    (with a trailing blank)."""
    r = start_row
    total_cols = ncols_max + len(FOLLOWUP_COLS)           # why span + follow-up block
    fu_c0 = ncols_max + 1
    r = ds_section(ws, r, total_cols, banner_text, level=1)
    ndata = len(data_headers)
    why_c0 = ndata + 1
    # header row: data headers + one merged "Why Flagged" header
    for col, h in enumerate(data_headers, 1):
        c = ws.cell(row=r, column=col)
        c.value = h
        c.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
        c.fill = AR._fill(DS_NAV2)
        c.alignment = AR._align("center", "center", wrap=True)
        c.border = ds_border()
    if ncols_max > why_c0:
        ws.merge_cells(start_row=r, start_column=why_c0, end_row=r, end_column=ncols_max)
    for col in range(why_c0, ncols_max + 1):
        c = ws.cell(row=r, column=col)
        c.fill = AR._fill(DS_NAV2); c.border = ds_border()
    wc = ws.cell(row=r, column=why_c0)
    wc.value = "Why Flagged"
    wc.font = AR._font(bold=True, size=9, color=AR.C_WHITE)
    wc.alignment = AR._align("center", "center", wrap=True)
    ds_followup_headers(ws, r, fu_c0)
    ws.row_dimensions[r].height = 28
    r += 1
    if not display_rows:
        r = ds_empty(ws, r, total_cols, "✅  No attention-required records in this section.")
        return r + 1
    # width available to the merged Why-Flagged cell (for the row-height estimate)
    why_width = sum((ws.column_dimensions[_gcl(c)].width or 14) for c in range(why_c0, ncols_max + 1))
    for i, dr in enumerate(display_rows):
        sev = dr["severity"]
        level = {3: "high", 2: "medium", 1: "info"}.get(sev, "info")
        bg = ds_zebra(i)
        cells = dr["cells"]
        for ci in range(ndata):
            st = dr["status_cells"].get(ci)
            if st is not None:
                s = str(st or "").strip()
                lev = {"Invalid": "high", "Missing": "medium", "Warning": "info",
                       "Valid": "ok"}.get(s, "none")
                ds_pill(ws, r, ci + 1, s or "—", lev)
            else:
                val = cells[ci] if ci < len(cells) else ""
                ds_cell(ws, r, ci + 1, (val if str(val).strip() != "" else "—"), bg=bg,
                        h_align="left" if ci in text_col_idxs else "center", wrap=True,
                        bold=(ci in text_col_idxs and ci == min(text_col_idxs, default=-1)))
        # first cell = priority chip
        first_val = ws.cell(row=r, column=1).value
        ds_priority(ws, r, 1, first_val, level,
                    h_align="left" if 0 in text_col_idxs else "center")
        if ncols_max > why_c0:
            ws.merge_cells(start_row=r, start_column=why_c0, end_row=r, end_column=ncols_max)
            for col in range(why_c0 + 1, ncols_max + 1):
                ws.cell(row=r, column=col).border = ds_border()
                ws.cell(row=r, column=col).fill = AR._fill(DS_WHY)
        ds_why_text(ws, r, why_c0, dr["why_reasons"], col_width=why_width)
        ds_followup_cells(ws, r, fu_c0)
        r += 1
    return r + 1


def build_wise_validation(ws, data, report_date, period_label=None, interview_rows=None):
    """Wise & Interview Feedback Validation tab — combines pyWiseDataValidationReport's
    Student, Course and Instructor validation FAILURES, PLUS an 'Interview Feedback Not
    Completed' section, into one coordinator tab. Each section is separated by a header
    banner and ends in an action-oriented 'Why Flagged' column. The existing Wise
    validation logic is reused unchanged; the interview-feedback section is purely
    additive. Presentation follows the common report design system."""
    NMAX = 11
    # Interview Feedback Not Completed — appended as the LAST section of this tab.
    IFV_DATA = ["#", "Interview Start Date", "Interviewer Name", "Tech Stack",
                "Batch Name", "Batch Title / Duration"]
    # the interview rows carry [date, interviewer, tech, batch, title]; number them
    # like every other section so the '#' priority chip lines up (display only)
    ifv_rows = [dict(r, cells=[i] + list(r.get("cells", [])))
                for i, r in enumerate(interview_rows or [], 1)]
    IFV_BANNER = ("INTERVIEW FEEDBACK NOT COMPLETED  —  Feedback missing in the "
                  "Interview Consolidate Sheet")
    period = period_label or report_date.strftime('%d-%b-%Y')
    NTOT = NMAX + len(FOLLOWUP_COLS)
    ds_title(ws, NTOT, "Wise & Interview Feedback Validation", period,
             "Data-quality checks in Wise (student / course / instructor records) plus interviews "
             "without feedback. Priority chip on the first cell: High = Invalid value, "
             "Medium = Missing value, Info = Warning.   Colour key: ■ High = red  ■ Medium = amber  "
             "■ Info = blue  ■ Valid = green." + DS_GUIDE_FOLLOWUP)
    # column widths first — the merged Why-Flagged width drives the row heights
    _widths = [6, 26, 24, 18, 12, 12, 12, 20, 10, 24, 40] + [FOLLOWUP_WIDTHS[h] for h in FOLLOWUP_COLS]
    for i, w in enumerate(_widths, 1):
        ws.column_dimensions[_gcl(i)].width = w
    row = 3

    if not data:
        row = ds_empty(ws, row, NTOT, "Wise validation data is unavailable for this run.",
                       level="muted")
        # Still show the Interview Feedback validation section (independent of Wise data).
        _wise_render_section(ws, row + 1, NMAX, IFV_BANNER, IFV_DATA, ifv_rows, {2, 3, 4, 5})
        ds_followup_apply(ws, NMAX + 1, FOLLOWUP_ACTIONS["wise"])
        ws.freeze_panes = "E3"          # title + guide rows, and the first 4 columns
        ws.sheet_view.showGridLines = False
        ws.sheet_properties.tabColor = DS_SECTION
        return ws

    STU_DATA = ["#", "Student Name", "Batch Name", "Name", "Email", "Phone",
                "Tag", "Note", "Picture", "Joined On"]
    CRS_DATA = ["#", "Course Title", "Course Subtitle", "Instructor (Tag)",
                "Title", "Subtitle", "Tag", "Created On"]
    INS_DATA = ["#", "Instructor ID", "Instructor Name", "Failed Checks"]

    n_total = sum(len(v or []) for v in (data.get("student"), data.get("course"),
                                          data.get("instructor"))) + len(interview_rows or [])
    row = _wise_render_section(
        ws, row, NMAX, "STUDENT VALIDATION  —  Failed / Attention-Required Records",
        STU_DATA, _wise_student_display(data.get("student", [])), {1, 2})
    row = _wise_render_section(
        ws, row, NMAX, "COURSE VALIDATION  —  Failed / Attention-Required Records",
        CRS_DATA, _wise_course_display(data.get("course", [])), {1, 2, 3})
    row = _wise_render_section(
        ws, row, NMAX, "INSTRUCTOR VALIDATION  —  Failed / Attention-Required Records",
        INS_DATA, _wise_instructor_display(data.get("instructor", [])), {2})
    # NEW — Interview Feedback Not Completed (last section, additive).
    row = _wise_render_section(
        ws, row, NMAX, IFV_BANNER, IFV_DATA, ifv_rows, {2, 3, 4, 5})
    ds_guide_count(ws, "Records needing attention", n_total)
    ds_followup_apply(ws, NMAX + 1, FOLLOWUP_ACTIONS["wise"])

    ds_fit_guide(ws, NTOT)
    ws.freeze_panes = "E3"              # title + guide rows, and the first 4 columns
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = DS_SECTION
    try:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    except Exception:
        pass
    return ws


def build_student_detail_ext(ws, att_daily: pd.DataFrame, susp_daily: pd.DataFrame,
                             fb_daily: pd.DataFrame, agg_map: dict, report_date: date,
                             yest_date=None, period_label=None):
    """Extended daily Student Detail written into the provided worksheet `ws`.
    Mirrors pyAttendaceFeedbackReport's daily Student Detail (identity + daily
    columns, identical logic) and appends the Overall / Till-Date Aggregate block.
    Presentation: common report design system — grouped headers (Student / Daily /
    Overall), the Rank cell as the priority chip, neutral zebra rows, semantic
    tints only on status & measure cells, and a numbered Why Flagged with the
    action on its own line."""
    period = period_label or report_date.strftime('%d-%b-%Y')
    ds_title(ws, N, "Learner Attendance Follow-Ups", period,
             "One row per learner who needs a follow-up today, grouped by technology (the blue "
             "banner names the Tech / Duration and how many learners are pending); Rank 1 = "
             "most urgent within the technology (chip: High = rank 1–3, Medium = 4–6, Info = others). "
             "Peach rows = yesterday's session.   Colour key: ■ High = red  ■ Medium = amber  "
             "■ Info = blue  ■ OK = green." + DS_GUIDE_FOLLOWUP)
    groups = [(1, 1, ""),
              (2, len(STUDENT_COLS), "STUDENT"),
              (len(STUDENT_COLS) + 1, len(STUDENT_COLS) + len(DAILY_COLS),
               f"DAILY — {report_date.strftime('%d-%b-%Y')}"),
              (_AGG_START, _C["Why Flagged"] - 1, "OVERALL / TILL-DATE AGGREGATE"),
              (_C["Why Flagged"], _C["Why Flagged"], "ACTION"),
              (_C["Action Taken"], _C["Follow-Up DateTime"], "COORDINATOR FOLLOW-UP (fill in)")]
    HDR_ROW = 4
    row_num = ds_headers(ws, 3, HEADERS, groups=groups)
    ds_followup_headers(ws, HDR_ROW, _C["Action Taken"])

    # feedback given pairs (session_id, student_id) for today
    fb_pairs = set()
    fb_ratings = {}          # (session_id, student_id) -> rating (/10) for that session
    if fb_daily is not None and not fb_daily.empty \
            and "session_id" in fb_daily.columns and "student_id" in fb_daily.columns:
        for _, fr in fb_daily.iterrows():
            _k = (str(fr.get("session_id", "")), str(fr.get("student_id", "")))
            fb_pairs.add(_k)
            _rt = fr.get("_rating")
            if _rt is not None and not (isinstance(_rt, float) and pd.isna(_rt)):
                fb_ratings[_k] = _rt

    def _agg_cells(row_num, key, row_bg):
        """Write the aggregate columns for a given (sid, cn, ct)."""
        a = agg_map.get(key)
        if not a:
            for col in range(_AGG_START, _C["Why Flagged"] + 1):
                ds_cell(ws, row_num, col, "—", bg=row_bg, h_align="center")
            return
        present, total = a["present"], a["total"]
        agg_att = a["attendance_pct"]
        agg_dur = a["agg_dur_pct"]
        # Sessions (P/T)
        ds_cell(ws, row_num, _C["Sessions (P/T)"], f"{present}/{total}", bg=row_bg,
                bold=True, h_align="center")
        # Agg Attendance %  /  Agg Duration Att %  (heat tint = meaning)
        for hname, v in (("Agg Attendance %", agg_att), ("Agg Duration Att %", agg_dur)):
            ds_cell(ws, row_num, _C[hname], v, bg=ds_att_bg(v, ""), h_align="center",
                    number_fmt='0.0"%"', bold=(v < 75),
                    fg=(AR.C_RED_DARK if v < 75 else AR.C_GREEN_DARK if v >= 95 else DS_TEXT))
        # Last Present
        lp = a["last_present"].strftime("%d-%b-%Y") if a["last_present"] else "—"
        ds_cell(ws, row_num, _C["Last Present"], lp, bg=row_bg, h_align="center")
        # Absent Streak
        streak = a["streak"]
        ds_pill(ws, row_num, _C["Absent Streak"], streak,
                ("high" if streak >= 3 else "medium" if streak >= 1 else "ok"),
                bold=(streak >= 1))
        # Feedbacks
        ds_cell(ws, row_num, _C["Feedbacks"], a["n_feedbacks"], bg=row_bg, h_align="center")
        # Feedback Part. %
        part = a["participation_pct"]
        ds_pill(ws, row_num, _C["Feedback Part. %"], part,
                ("ok" if part >= 50 else "medium" if part > 0 else "high"),
                bold=False, number_fmt='0.0"%"')
        # Avg Rating (/10)
        if a["avg_rating"] is not None:
            ds_cell(ws, row_num, _C["Avg Rating (/10)"], a["avg_rating"],
                    bg=ds_rating_bg(a["avg_rating"]), h_align="center", number_fmt='0.0')
        else:
            ds_cell(ws, row_num, _C["Avg Rating (/10)"], "—", bg=row_bg, h_align="center")
        # Overall Remarks (daily _remarks bands, on the aggregate duration %)
        orem = AR._remarks(agg_dur, "")
        ds_cell(ws, row_num, _C["Overall Remarks"], orem, bg=ds_att_bg(agg_dur, ""),
                bold=(agg_dur < 75),
                fg=(AR.C_RED_DARK if agg_dur < 75 else AR.C_GREEN_DARK if agg_dur >= 95 else DS_TEXT))

    # global severity ranking (rank 1 = worst) across all follow-up rows
    rank_map = _compute_learner_ranks(att_daily, agg_map, fb_pairs, fb_ratings)

    # ── FOLLOW-UP rows only, grouped by (course_name, course_title) ──────────
    #   A student is kept ONLY when they qualify for at least one follow-up type
    #   (Attendance defaulter / Feedback defaulter / Low rating). Why Flagged
    #   combines every applicable reason so the coordinator can cover them all in
    #   one conversation. Not-applicable and suspended students are never
    #   follow-ups and are omitted from this action list.
    any_rendered = False
    n_rows = 0
    if att_daily is not None and not att_daily.empty:
        _tmp = att_daily.copy()
        # order each technology's follow-up rows by their (technology-wise) rank,
        # so the Rank column reads 1, 2, 3, 4 … down each technology block.
        def _rk_of(r):
            return rank_map.get((str(r.get("student_id", "")), str(r.get("session_id", "")),
                                 str(r.get("course_name", "")), str(r.get("course_title", ""))), 10**9)
        _tmp["_rank_sort"] = _tmp.apply(_rk_of, axis=1)
        sorted_att = _tmp.sort_values(
            ["course_name", "course_title", "_rank_sort"], na_position="last")

        # Pass 1 — the follow-up rows, in their final order (same qualification
        # rule as before: applicable AND at least one reason). Collected first so
        # every technology banner can show how many learners it holds.
        flagged = []
        for _, r in sorted_att.iterrows():
            # not-applicable students can't be follow-ups here
            if "_attn_applicable" in r.index and not bool(r.get("_attn_applicable", True)):
                continue
            cn = r.get("course_name", "")
            ct = r.get("course_title", "")
            sid = str(r.get("student_id", ""))
            agg_key = (sid, str(cn), str(ct))
            agg = agg_map.get(agg_key)
            reason_runs = _learner_followup_runs(r, agg, fb_pairs, fb_ratings)
            if not reason_runs:
                continue                        # performing fine -> not on the list
            flagged.append((r, cn, ct, sid, agg_key, reason_runs))
        group_n = {}
        for _r, cn, ct, _s, _k, _rr in flagged:
            group_n[(cn, ct)] = group_n.get((cn, ct), 0) + 1

        # Pass 2 — render
        banner_group = None
        alt_i = 0
        for r, cn, ct, sid, agg_key, reason_runs in flagged:
            if (cn, ct) != banner_group:        # technology banner (once per group)
                banner_group = (cn, ct)
                alt_i = 0
                n_grp = group_n[(cn, ct)]
                banner = (f"Tech Name: {cn}   ·   Duration: {ct}" if ct else f"Tech Name: {cn}")
                banner += f"   ·   {n_grp} learner{'s' if n_grp != 1 else ''} pending follow-up"
                row_num = ds_section(ws, row_num, N, banner, level=1)

            status = str(r.get("status", ""))
            pct = r.get("_pct_num", 0.0)
            row_bg = ds_zebra(alt_i)
            if yest_date is not None and r.get("_date") == yest_date:
                row_bg = "FFF6EA"                # yesterday's session — soft peach
            alt_i += 1
            icon = AR._att_icon(pct, status)
            status_disp = "❌  Absent" if status == "Absent" else "✅  Present"
            att_disp = f"{icon}  Absent" if status == "Absent" else f"{icon}  {pct:.1f}%"

            # Rank (1 = worst) — reflects combined A/B/C follow-up severity
            _rk = rank_map.get((sid, str(r.get("session_id", "")), str(cn), str(ct)))
            level = ("high" if (_rk and _rk <= 3) else "medium" if (_rk and _rk <= 6) else "info")
            ds_priority(ws, row_num, _C["Rank"], _rk if _rk else "—", level)

            ds_cell(ws, row_num, _C["Student Name"], r.get("student_name", ""), bg=row_bg, bold=True)
            ds_cell(ws, row_num, _C["Phone"], r.get("phone", ""), bg=row_bg)

            ds_pill(ws, row_num, _C["Status"], status_disp,
                    "high" if status == "Absent" else "ok")
            ds_cell(ws, row_num, _C["Attendance %"], att_disp, bg=ds_att_bg(pct, status),
                    h_align="center", bold=(status == "Absent" or pct < 75),
                    fg=(AR.C_RED_DARK if (status == "Absent" or pct < 75) else AR.C_GREEN_DARK))
            ds_cell(ws, row_num, _C["Duration (min)"],
                    r.get("_dur_min", 0) if r.get("_dur_min", 0) > 0 else "",
                    bg=row_bg, h_align="center")
            ds_cell(ws, row_num, _C["Joined At"], str(r.get("first_join_ist", ""))[:19],
                    bg=row_bg, h_align="center")
            ds_cell(ws, row_num, _C["Left At"], str(r.get("last_leave_ist", ""))[:19],
                    bg=row_bg, h_align="center")
            ds_cell(ws, row_num, _C["Remarks"], AR._remarks(pct, status), bg=row_bg,
                    fg=(AR.C_RED_DARK if (status == "Absent" or pct < 75)
                        else AR.C_GREEN_DARK if pct >= 95 else DS_TEXT))

            if status == "Absent":
                fb_val, fb_lev = "Absent", "muted"
            elif (str(r.get("session_id", "")), sid) in fb_pairs:
                fb_val, fb_lev = "✅  Yes", "ok"
            else:
                fb_val, fb_lev = "❌  No", "high"
            ds_pill(ws, row_num, _C["Feedback Given?"], fb_val, fb_lev)

            _rt_key = (str(r.get("session_id", "")), sid)
            if status != "Absent" and _rt_key in fb_ratings:
                _rt_val = fb_ratings[_rt_key]
                ds_cell(ws, row_num, _C["Feedback Rating (/10)"], _rt_val,
                        bg=ds_rating_bg(_rt_val), h_align="center", number_fmt="0.#")
            else:
                ds_cell(ws, row_num, _C["Feedback Rating (/10)"], "—", bg=row_bg, h_align="center")

            # aggregate block, then the combined, action-oriented Why Flagged
            _agg_cells(row_num, agg_key, row_bg)
            ds_why_text(ws, row_num, _C["Why Flagged"], reason_runs, col_width=62)
            ds_followup_cells(ws, row_num, _C["Action Taken"])
            any_rendered = True
            n_rows += 1
            row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, N, "✅  No learner follow-ups required for this day.")
        row_num += 1
    ds_guide_count(ws, "Learners needing follow-up", n_rows)

    ds_finish(ws, HDR_ROW, N, widths={
        "Rank": 7, "Student Name": 22, "Phone": 15,
        "Status": 12, "Attendance %": 12, "Duration (min)": 10, "Joined At": 19, "Left At": 19,
        "Remarks": 12, "Feedback Given?": 11, "Feedback Rating (/10)": 10,
        "Sessions (P/T)": 10, "Agg Attendance %": 11, "Agg Duration Att %": 11,
        "Last Present": 12, "Absent Streak": 9, "Feedbacks": 9, "Feedback Part. %": 10,
        "Avg Rating (/10)": 9, "Overall Remarks": 14,
    }, why_col=_C["Why Flagged"], why_width=62, default_width=11, tab_color=DS_SECTION,
       followup=(_C["Action Taken"], FOLLOWUP_ACTIONS["attendance"]),
       freeze_after_col=_C["Phone"])
    return ws


# =============================================================================
#  DRIVE OUTPUT — date-wise folder + native Google Sheet (history preserved)
# =============================================================================
def _find_or_create_folder(drive, parent_id, name):
    safe = name.replace("'", "\\'")
    q = (f"'{parent_id}' in parents and name='{safe}' and "
         f"mimeType='application/vnd.google-apps.folder' and trashed=false")
    res = drive.files().list(q=q, fields="files(id,name)", supportsAllDrives=True,
                             includeItemsFromAllDrives=True).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    f = drive.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    log.info("Created date folder '%s'", name)
    return f["id"]


def _drive_client():
    """Impersonated Drive v3 client (drive scope) — the identity that owns the
    Coordinator report folders."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    creds = service_account.Credentials.from_service_account_file(
        AR.SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive"]
    ).with_subject(IMPERSONATE_USER)
    return gbuild("drive", "v3", credentials=creds, cache_discovery=False)


def _list_children(drive, parent_id, folders_only=False):
    """All (non-trashed) children of a Drive folder: [{id, name, mimeType,
    createdTime, modifiedTime}], following pagination."""
    q = f"'{parent_id}' in parents and trashed=false"
    if folders_only:
        q += " and mimeType='application/vnd.google-apps.folder'"
    out, token = [], None
    while True:
        res = drive.files().list(q=q, fields="nextPageToken, files(id,name,mimeType,createdTime,modifiedTime)",
                                 pageSize=1000, pageToken=token, supportsAllDrives=True,
                                 includeItemsFromAllDrives=True).execute()
        out += res.get("files", [])
        token = res.get("nextPageToken")
        if not token:
            return out


def upload_report(folder_name, filename: str, buf: io.BytesIO, base_prefix: str) -> str:
    """Create <parent>/<folder_name>/ and drop the workbook there as a native
    Google Sheet (converted on upload). `folder_name` is one folder name or a
    sequence of nested names — the Coordinator layout passes
    (report-type folder, reporting-period folder) from coordinator_periods.
    Existing reports for this report/date are
    NEVER overwritten or modified — the Coordinator may have added manual follow-up
    comments to them. If a report with this base_prefix already exists in the
    folder, the new run is saved as the next version instead:
        first run      -> the plain file name
        already exists -> "<name> - Version 2"
        Version 2 too  -> "<name> - Version 3"  (increments dynamically)
    Every previous version is kept unchanged. base_prefix scopes this per report
    type, so each type versions independently and other reports are untouched."""
    import re
    from googleapiclient.http import MediaIoBaseUpload

    drive = _drive_client()
    names = [folder_name] if isinstance(folder_name, str) else list(folder_name)
    folder_id = CP.resolve_folder(drive, PARENT_FOLDER_ID, names, _find_or_create_folder)
    folder_name = "/".join(names)
    drive_name = filename[:-5] if filename.lower().endswith(".xlsx") else filename
    base = base_prefix.replace("'", "\\'")
    q = f"'{folder_id}' in parents and name contains '{base}' and trashed=false"
    existing = drive.files().list(q=q, fields="files(id,name)", supportsAllDrives=True,
                                  includeItemsFromAllDrives=True).execute().get("files", [])

    # Do NOT delete/replace any existing report — preserve every version. If one
    # already exists for this report type in this folder, name the new run as the
    # next available version (the un-versioned file counts as Version 1).
    if existing:
        _ver_re = re.compile(r"-\s*Version\s*(\d+)\s*$", re.IGNORECASE)
        max_ver = 1
        for _f in existing:
            _m = _ver_re.search(_f.get("name", ""))
            if _m:
                max_ver = max(max_ver, int(_m.group(1)))
        target_name = f"{drive_name} - Version {max_ver + 1}"
    else:
        target_name = drive_name

    buf.seek(0)
    media = MediaIoBaseUpload(
        buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=False)
    meta = {"name": target_name, "parents": [folder_id],
            "mimeType": "application/vnd.google-apps.spreadsheet"}
    up = drive.files().create(body=meta, media_body=media, fields="id,webViewLink",
                              supportsAllDrives=True).execute()
    link = up.get("webViewLink") or f"https://docs.google.com/spreadsheets/d/{up.get('id','')}/edit"
    log.info("Uploaded native Google Sheet → %s / %s", folder_name, target_name)
    _enable_followup_timestamps(up.get("id", ""))
    return link


def _enable_followup_timestamps(spreadsheet_id: str) -> None:
    """Make the coordinator follow-up columns work on the uploaded Google Sheet:
    turn on iterative calculation (so the self-referencing Follow-Up DateTime
    formula keeps the time at which Follow-Up Done? was set instead of ticking
    with NOW()) and pin the sheet's time zone to IST so the stamp is local time.
    Best-effort: a failure is logged with the manual fix and never blocks the run."""
    if not spreadsheet_id:
        return
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    body = {"requests": [{"updateSpreadsheetProperties": {
        "properties": {"timeZone": FOLLOWUP_TIMEZONE,
                       "iterativeCalculationSettings": {"maxIterations": 1,
                                                        "convergenceThreshold": 0.0}},
        "fields": "timeZone,iterativeCalculationSettings"}}]}
    # The Sheets API accepts the Drive scope, so the FIRST attempt uses exactly the
    # credentials the upload itself used (impersonated, drive scope) — no extra
    # domain-wide-delegation scope is needed. Fallback: the plain service account
    # (works when the coordinator folder is shared with it).
    attempts = [
        ("impersonated " + IMPERSONATE_USER,
         lambda: service_account.Credentials.from_service_account_file(
             AR.SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive"]
         ).with_subject(IMPERSONATE_USER)),
        ("service account",
         lambda: service_account.Credentials.from_service_account_file(
             AR.SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])),
    ]
    errors = []
    for who, make_creds in attempts:
        try:
            sheets = gbuild("sheets", "v4", credentials=make_creds(), cache_discovery=False)
            sheets.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            log.info("Follow-up DateTime enabled on the sheet as %s (iterative calculation "
                     "on, tz=%s).", who, FOLLOWUP_TIMEZONE)
            return
        except Exception as exc:                               # noqa: BLE001
            errors.append(f"{who}: {exc}")
    bar = "!" * 70
    log.warning("%s\nFOLLOW-UP DATETIME NOT ENABLED on the uploaded sheet — its cells will "
                "show #REF! (circular dependency) until iterative calculation is switched "
                "on by hand: open the sheet ▸ File ▸ Settings ▸ Calculation ▸ Iterative "
                "calculation = On (Max iterations 1), then clear and re-select Follow-Up "
                "Done? on any row already marked.\nAttempts: %s\n%s",
                bar, " | ".join(errors), bar)


def upload_datewise(report_date: date, filename: str, buf: io.BytesIO) -> str:
    """Daily report → <parent>/Daily Coordinator Reports/Daily DD-Mon-YYYY/
    (the shared Coordinator layout; thin wrapper over upload_report)."""
    return upload_report(CP.folder_path("Daily", report_date, report_date),
                         filename, buf, REPORT_BASENAME)


# =============================================================================
#  ORCHESTRATION
# =============================================================================
# =============================================================================
#  WEEKLY / MONTHLY PERIOD ROLL-UPS  (one row per learner / per instructor over
#  the whole period — preserves daily history, avoids per-day duplication)
# =============================================================================
def _session_metrics(sess, att_period, fb_period, tf_ids):
    """Per-session instructor metrics (same rules as the daily Instructor tab)."""
    sid = sess.get("session_id", "")
    sa = (att_period[att_period["session_id"] == sid]
          if (att_period is not None and not att_period.empty) else pd.DataFrame())
    sa_app = AR._applicable(sa) if not sa.empty else sa
    present = sa_app[sa_app["status"] == "Present"] if not sa_app.empty else sa_app
    n_present, n_app = len(present), len(sa_app)
    att_pct = round(n_present / n_app * 100, 1) if n_app else 0.0
    # TRUE actual conducted duration (never the scheduled slot)
    _as, _ae, dur_min = _actual_session_times(sess, sa)
    dur_min = dur_min or 0.0
    sched_min = _scheduled_min(sess)
    diff_min = round(dur_min - sched_min, 1) if sched_min is not None else None
    sf = (fb_period[fb_period["session_id"] == sid]
          if (fb_period is not None and not fb_period.empty) else pd.DataFrame())
    if sf is not None and not sf.empty:
        _cols = [c for c in ["session_id", "student_id"] if c in sf.columns]
        sf = sf.drop_duplicates(subset=_cols) if _cols else sf
    ratings = (sf["_rating"].dropna().tolist()
               if (sf is not None and not sf.empty and "_rating" in sf.columns) else [])
    avg_r = round(sum(ratings) / len(ratings), 2) if ratings else None
    fb_given = sid in tf_ids
    flagged = ((not fb_given)
               or (diff_min is not None and diff_min <= -INSTR_SHORT_MIN)
               or (n_app > 0 and att_pct < INSTR_LOW_ATT_PCT)
               or (avg_r is not None and avg_r <= INSTR_LOW_RATING))
    return dict(att_pct=att_pct, avg_r=avg_r, diff_min=diff_min,
                fb_given=fb_given, flagged=flagged)


def build_learner_followups_period(ws, att_period, fb_period, agg_map, start, end, label):
    """One row per (student, tech) flagged on >=1 day in the period, with the
    count of Attendance / Feedback / Low-Rating days and till-date aggregate."""
    ds_title(ws, LPN, "Learner Attendance Follow-Ups (period roll-up)", label,
             "One row per learner flagged on at least one day in the period; Rank 1 = most urgent "
             "within the technology (chip: High = rank 1–3, Medium = 4–6, Info = others).   "
             "Colour key: ■ High = red  ■ Medium = amber  ■ Info = blue  ■ OK = green.")
    row_num = ds_headers(ws, 3, LP_COLS)

    fb_pairs = set()
    fb_ratings = {}
    if fb_period is not None and not fb_period.empty:
        for _, fr in fb_period.iterrows():
            k = (str(fr.get("session_id", "")), str(fr.get("student_id", "")))
            fb_pairs.add(k)
            rt = fr.get("_rating")
            if rt is not None and not (isinstance(rt, float) and pd.isna(rt)):
                fb_ratings[k] = rt

    # ── Phase 1: collect one payload per flagged (student, tech) over the period ──
    payloads = []
    if att_period is not None and not att_period.empty:
        appl = AR._applicable(att_period)
        for (sid, cn, ct), grp in appl.groupby(["student_id", "course_name", "course_title"], sort=False):
            agg = agg_map.get((str(sid), str(cn), str(ct)))
            if not agg:
                continue
            overall_critical = AR._remarks(agg["agg_dur_pct"], "").strip().endswith("Critical")
            part = agg["participation_pct"]
            a_days = b_days = c_days = 0
            present = total = 0
            for _, r in grp.iterrows():
                total += 1
                status = str(r.get("status", ""))
                dp = float(r.get("_pct_num", 0) or 0)
                key = (str(r.get("session_id", "")), str(sid))
                if status == "Present":
                    present += 1
                if overall_critical and ((status == "Absent") or (status == "Present" and dp < CRITICAL_ATT_PCT)):
                    a_days += 1
                if part < FEEDBACK_PART_THRESHOLD and status == "Present" and key not in fb_pairs:
                    b_days += 1
                if key in fb_ratings:
                    try:
                        if float(fb_ratings[key]) <= LOW_RATING_THRESHOLD:
                            c_days += 1
                    except (TypeError, ValueError):
                        pass
            total_flag = a_days + b_days + c_days
            if total_flag == 0:
                continue
            # action-oriented reasons as (text, is_red) runs — same red-highlight
            # style as the daily Learner / Instructor Follow-Ups.
            reason_runs = []
            if a_days:
                reason_runs.append([("Critical Attendance", True), (": Overall Attendance ", False),
                                    (f"{agg['agg_dur_pct']:.0f}%", True), (", Flagged ", False),
                                    (f"{a_days} day(s)", True), (" — ", False), ("Call Learner", True)])
            if b_days:
                reason_runs.append([("Feedback Missing", True), (": Overall Participation ", False),
                                    (f"{part:.0f}%", True), (", Missing ", False),
                                    (f"{b_days} day(s)", True), (" — ", False), ("Message Learner", True)])
            if c_days:
                cr = [("Low Rating", True), (": Flagged ", False), (f"{c_days} day(s)", True)]
                if agg.get("avg_rating") is not None:
                    cr += [(", Overall Avg ", False), (f"{agg['avg_rating']:g}/10", True)]
                cr += [(" — ", False), ("Understand Rating Reason", True)]
                reason_runs.append(cr)
            # severity: weighted flag-days by condition priority + attendance gap
            sev = (a_days * _SEV_BASE_A + b_days * _SEV_BASE_B + c_days * _SEV_BASE_C
                   + max(0.0, CRITICAL_ATT_PCT - agg["agg_dur_pct"]))
            payloads.append(dict(
                sev=sev, total_flag=total_flag, agg_dur=agg["agg_dur_pct"],
                a_days=a_days, b_days=b_days, c_days=c_days, present=present, total=total,
                part=part, agg=agg, reason_runs=reason_runs,
                nm=str(grp.iloc[0].get("student_name", "")), ph=str(grp.iloc[0].get("phone", "")),
                cn=cn, ct=ct))

    # Technology-wise ranking: rank restarts at 1 for each Technology
    # (course_name × course_title), worst performer first. Rows are grouped by
    # technology and ordered by that rank.
    from collections import defaultdict as _dd
    _by_tech = _dd(list)
    for p in payloads:
        _by_tech[(p["cn"], p["ct"])].append(p)
    ordered = []
    for _tech in sorted(_by_tech.keys()):
        grp = sorted(_by_tech[_tech],
                     key=lambda p: (-p["sev"], -p["total_flag"], p["agg_dur"], p["nm"]))
        for rk, p in enumerate(grp, 1):
            p["_rank"] = rk
            ordered.append(p)
    any_rendered = False
    n_rows = 0
    _last_tech = None
    for p in ordered:
        rk = p["_rank"]
        agg = p["agg"]
        if (p["cn"], p["ct"]) != _last_tech:            # technology banner
            _last_tech = (p["cn"], p["ct"])
            row_num = ds_section(ws, row_num, LPN,
                                 f"{p['cn']}  —  {p['ct']}" if p["ct"] else f"{p['cn']}", level=1)
        bg = ds_zebra(rk)
        vals = {
            "Rank": rk, "Tech Name": p["cn"], "Duration": p["ct"],
            "Student Name": p["nm"], "Phone": p["ph"],
            "Sessions (P/T) in Period": f"{p['present']}/{p['total']}",
            "Attendance Flag Days": p["a_days"], "Feedback Flag Days": p["b_days"],
            "Low-Rating Days": p["c_days"], "Total Flag Days": p["total_flag"],
            "Overall Remarks": AR._remarks(p["agg_dur"], ""),
        }
        for c in LP_COLS:
            if c in vals:
                ds_cell(ws, row_num, _LP[c], vals[c], bg=bg,
                        bold=(c == "Student Name"),
                        h_align="left" if c in ("Tech Name", "Duration", "Student Name")
                        else "center")
        ds_priority(ws, row_num, _LP["Rank"], rk,
                    "high" if rk <= 3 else "medium" if rk <= 6 else "info")
        for hname, v in (("Attendance Flag Days", p["a_days"]), ("Feedback Flag Days", p["b_days"]),
                         ("Low-Rating Days", p["c_days"]), ("Total Flag Days", p["total_flag"])):
            ds_pill(ws, row_num, _LP[hname], v, ("high" if v >= 3 else "medium" if v >= 1 else "none"),
                    bold=(v >= 1))
        ds_cell(ws, row_num, _LP["Agg Attendance %"], agg["attendance_pct"],
                bg=ds_att_bg(agg["attendance_pct"], ""), h_align="center", number_fmt='0.0"%"')
        ds_cell(ws, row_num, _LP["Agg Duration Att %"], p["agg_dur"],
                bg=ds_att_bg(p["agg_dur"], ""), h_align="center", number_fmt='0.0"%"')
        ds_pill(ws, row_num, _LP["Feedback Part. %"], p["part"],
                ("ok" if p["part"] >= 50 else "medium" if p["part"] > 0 else "high"),
                bold=False, number_fmt='0.0"%"')
        if agg.get("avg_rating") is not None:
            ds_cell(ws, row_num, _LP["Avg Rating (/10)"], agg["avg_rating"],
                    bg=ds_rating_bg(agg["avg_rating"]), h_align="center", number_fmt='0.0')
        else:
            ds_cell(ws, row_num, _LP["Avg Rating (/10)"], "—", bg=bg, h_align="center")
        ds_cell(ws, row_num, _LP["Overall Remarks"], AR._remarks(p["agg_dur"], ""),
                bg=ds_att_bg(p["agg_dur"], ""), bold=(p["agg_dur"] < 75),
                fg=(AR.C_RED_DARK if p["agg_dur"] < 75 else AR.C_GREEN_DARK if p["agg_dur"] >= 95 else DS_TEXT))
        ds_why_text(ws, row_num, _LP["Why Flagged"], p["reason_runs"], col_width=60)
        any_rendered = True
        n_rows += 1
        row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, LPN, "✅  No learner follow-ups in this period.")
        row_num += 1
    ds_guide_count(ws, "Learners needing follow-up", n_rows)
    ds_finish(ws, 3, LPN, widths={
        "Rank": 7, "Tech Name": 22, "Duration": 26, "Student Name": 22, "Phone": 15,
        "Sessions (P/T) in Period": 12, "Attendance Flag Days": 11, "Feedback Flag Days": 11,
        "Low-Rating Days": 11, "Total Flag Days": 10, "Agg Attendance %": 11,
        "Agg Duration Att %": 11, "Feedback Part. %": 10, "Avg Rating (/10)": 9, "Overall Remarks": 14,
    }, why_col=_LP["Why Flagged"], why_width=60, default_width=11, tab_color=DS_SECTION)
    return ws


def build_instructor_followups_period(ws, sess_period, att_period, fb_period, tf_period,
                                      instr_phones, label):
    """One row per (instructor, tech) over the period: session counts, missing
    feedback, avg attendance/rating, flagged-session count, avg schedule diff."""
    ds_title(ws, IPN, "Instructor Follow-Ups (period roll-up)", label,
             "One row per instructor and technology with at least one missing feedback or flagged "
             "session in the period. Priority chip on Tech Name: High = feedback missing in more "
             "than half the sessions or sessions running short, Medium = otherwise.   "
             "Colour key: ■ High = red  ■ Medium = amber  ■ OK = green.")
    row_num = ds_headers(ws, 3, IP_COLS)
    _ip_widths = {"Tech Name": 22, "Duration": 26, "Instructor": 20, "Phone": 15, "Sessions": 9,
                  "Sessions Missing Feedback": 12, "Feedback Given %": 11, "Avg Att %": 10,
                  "Avg Rating (/10)": 10, "Flagged Sessions": 10, "Avg Diff Mins": 10}

    if sess_period is None or sess_period.empty:
        ds_empty(ws, row_num, IPN, "No sessions in this period.", level="muted")
        ds_finish(ws, 3, IPN, widths=_ip_widths, why_col=_IP["Why Flagged"], why_width=60,
                  default_width=11, tab_color=DS_SECTION)
        return ws

    sess_f = AR._prefer_instructor_name(sess_period)
    tf_ids = AR.teacher_feedback_session_ids(tf_period)   # content-bearing rows only

    any_rendered = False
    n_rows = 0
    for (cn, ct, instr), grp in sess_f.groupby(["course_name", "course_title", "tutor_name"], sort=True):
        n = len(grp)
        miss = flagged = 0
        att_vals, rating_vals, diff_vals = [], [], []
        for _, sess in grp.iterrows():
            m = _session_metrics(sess, att_period, fb_period, tf_ids)
            if not m["fb_given"]:
                miss += 1
            if m["flagged"]:
                flagged += 1
            att_vals.append(m["att_pct"])
            if m["avg_r"] is not None:
                rating_vals.append(m["avg_r"])
            if m["diff_min"] is not None:
                diff_vals.append(m["diff_min"])
        if miss == 0 and flagged == 0:
            continue
        avg_att = round(sum(att_vals) / len(att_vals), 1) if att_vals else 0.0
        avg_rat = round(sum(rating_vals) / len(rating_vals), 1) if rating_vals else None
        avg_diff = round(sum(diff_vals) / len(diff_vals), 1) if diff_vals else None
        given_pct = round((n - miss) / n * 100, 1) if n else 0.0
        phone = instr_phones.get(_norm_name(instr), "")

        reasons = []
        if miss:
            reasons.append(f"Feedback missing in {miss}/{n} session(s).")
        if flagged:
            reasons.append(f"{flagged}/{n} session(s) flagged for follow-up.")
        if avg_rat is not None and avg_rat <= INSTR_LOW_RATING:
            reasons.append(f"Low avg student rating {avg_rat:.1f}/10.")
        if avg_diff is not None and avg_diff <= -INSTR_SHORT_MIN:
            reasons.append(f"Sessions run ~{abs(avg_diff):.0f} min short on average.")

        vals = {
            "Tech Name": cn, "Duration": ct, "Instructor": instr, "Phone": phone or "—",
            "Sessions": n, "Sessions Missing Feedback": miss,
            "Flagged Sessions": flagged,
            "Avg Diff Mins": (f"{avg_diff:+.0f}" if avg_diff is not None else "—"),
        }
        bg = ds_zebra(n_rows)
        for c in IP_COLS:
            if c in vals:
                ds_cell(ws, row_num, _IP[c], vals[c], bg=bg, bold=(c == "Instructor"),
                        h_align="left" if c in ("Tech Name", "Duration", "Instructor")
                        else "center")
        # priority = the period's own severity signals: feedback missing in more
        # than half the sessions, or sessions running short on average
        _short = (avg_diff is not None and avg_diff <= -INSTR_SHORT_MIN)
        ds_priority(ws, row_num, _IP["Tech Name"], cn,
                    "high" if (_short or (n and miss / n > 0.5)) else "medium", h_align="left")
        ds_pill(ws, row_num, _IP["Sessions Missing Feedback"], miss,
                ("high" if miss else "ok"), bold=bool(miss))
        ds_pill(ws, row_num, _IP["Feedback Given %"], given_pct,
                ("ok" if given_pct >= 80 else "medium" if given_pct > 0 else "high"),
                bold=False, number_fmt='0.0"%"')
        ds_cell(ws, row_num, _IP["Avg Att %"], avg_att, bg=ds_att_bg(avg_att, ""),
                h_align="center", number_fmt='0.0"%"')
        if avg_rat is not None:
            ds_cell(ws, row_num, _IP["Avg Rating (/10)"], avg_rat, bg=ds_rating_bg(avg_rat),
                    h_align="center", number_fmt='0.0')
        else:
            ds_cell(ws, row_num, _IP["Avg Rating (/10)"], "—", bg=bg, h_align="center")
        ds_pill(ws, row_num, _IP["Flagged Sessions"], flagged, ("medium" if flagged else "ok"),
                bold=bool(flagged))
        ds_why_text(ws, row_num, _IP["Why Flagged"], [[(t, True)] for t in reasons], col_width=60)
        any_rendered = True
        n_rows += 1
        row_num += 1

    if not any_rendered:
        ds_empty(ws, row_num, IPN, "✅  No instructor follow-ups in this period.")
        row_num += 1
    ds_guide_count(ws, "Instructors needing follow-up", n_rows)
    ds_finish(ws, 3, IPN, widths=_ip_widths, why_col=_IP["Why Flagged"], why_width=60,
              default_width=11, tab_color=DS_SECTION)
    return ws


def _generate_period(service, report_type, start, end, plabel,
                     sess_agg, att_agg, fb_agg, tf_agg, susp_agg, gen_date):
    """Period roll-up (Weekly / Monthly / Manual): one row per learner and per
    instructor over [start, end] (no daily duplication). Daily history stays in
    the day-wise sheets. `report_type` only labels/names the output; the period
    window is passed in explicitly."""
    log.info("Period (%s): %s → %s  (%s)", report_type, start, end, plabel)

    def _le(df, d):
        if df is None or df.empty or "_date" not in df.columns:
            return df.iloc[0:0].copy() if df is not None else df
        return df[df["_date"].apply(lambda x: x is not None and not pd.isna(x) and x <= d)].copy()

    def _between(df, a, b):
        if df is None or df.empty or "_date" not in df.columns:
            return df.iloc[0:0].copy() if df is not None else df
        return df[df["_date"].apply(lambda x: x is not None and not pd.isna(x) and a <= x <= b)].copy()

    # aggregates as of the period close
    cfg = COORD.load_config() if _COORD_OK else None
    agg_map = build_aggregates(_le(att_agg, end), _le(fb_agg, end), cfg)

    # period-window rows (calendar dates within [start, end])
    att_p = _between(att_agg, start, end)
    fb_p = _between(fb_agg, start, end)
    sess_p = _between(sess_agg, start, end)
    tf_p = _between(tf_agg, start, end)

    instr_phones = load_instructor_phones(service)
    log.info("Period rows: %d session(s), %d attendance; instructors=%d",
             0 if sess_p is None else len(sess_p),
             0 if att_p is None else len(att_p), len(instr_phones))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_learner_followups_period(wb.create_sheet("Learner Attendance Follow-Ups"),
                                   att_p, fb_p, agg_map, start, end, plabel)
    build_instructor_followups_period(wb.create_sheet("Instructor Follow-Ups"),
                                      sess_p, att_p, fb_p, tf_p, instr_phones, plabel)
    buf = io.BytesIO()
    wb.save(buf)

    base = f"IntelliBI_Batch_Coordinator_{report_type.title()}_Follow_Ups"
    safe = plabel.replace(" ", "_").replace("–", "to").replace("/", "-").replace(":", ".")
    filename = f"{base}_{safe}.xlsx"
    # Save Weekly / Monthly / Manual roll-ups in the shared Coordinator layout:
    # <parent>/<Type> Coordinator Reports/<reporting-period folder>/ — the same
    # folder the Task Performance report for this type + period uses. File name
    # is unchanged; the per-report base_prefix keeps versioning per report type.
    folder = CP.folder_path(report_type.title(), start, end)
    link = upload_report(folder, filename, buf, base)
    print(f"\n{'='*64}\n  Batch Coordinator {report_type.title()} Follow-Ups — {plabel}")
    print(f"  Period sessions: {0 if sess_p is None else len(sess_p)} | Agg groups: {len(agg_map)}")
    print(f"  Link: {link}\n{'='*64}\n")
    return {"link": link, "report_type": report_type, "period": plabel,
            "start": start.isoformat(), "end": end.isoformat()}


def _le_date(df, d):
    """Rows whose _date is <= d (as-of slice). Safe on empty/missing frames."""
    if df is None or df.empty or "_date" not in df.columns:
        return df.iloc[0:0].copy() if df is not None else df
    return df[df["_date"].apply(lambda x: x is not None and not pd.isna(x) and x <= d)].copy()


# =============================================================================
#  LEARNER & INSTRUCTOR INTERVIEW REMINDER  (ADDITIVE TAB — upcoming interviews)
#  -----------------------------------------------------------------------------
#  Reads the interview-schedule Google Sheets in the coordinator interview folder,
#  finds interviews whose follow-up date falls on the report date, and lists the
#  Instructor reminder (first, on a distinct row) followed by the learner
#  reminders (sorted by interview time). Follow-up rule, per interview date D:
#     Call    = last non-Sunday day strictly before D          (1 day prior)
#     Message = last non-Sunday day strictly before the Call    (2 days prior)
#  Sundays are never used and the pair shifts backward automatically. Every value
#  (dates, batches, instructors, candidates, phones, times) is data-driven — none
#  is hard-coded. This block is fully self-contained and only ADDS a tab; it never
#  changes any existing tab, calculation, formatting, Drive/email or report logic.
# =============================================================================
import re as _iv_re

INTERVIEW_FOLDER_ID    = "1PzfXzmpLk_O9vBur6g7azkcxKWMikeKP"   # coordinator interview-schedule folder
INTERVIEW_HELPER_TAB   = "Interview_Helper"
INTERVIEW_FEEDBACK_TAB = "Interview Feedback"
# Interview Consolidate Sheet — where completed interview feedback is published.
INTERVIEW_CONSOLIDATE_SHEET_ID = "16IQtgrlvYZpEpsmtzWhyaS9ZgRHYB_jzQBF--DckCPg"

IV_COLS = ["Interview Time", "Batch Name (Class)", "Batch Title / Duration",
           "Candidate Name", "Phone", "Why Flagged"] + FOLLOWUP_COLS
IVN  = len(IV_COLS)
_IVC = {h: i + 1 for i, h in enumerate(IV_COLS)}


def _iv_services():
    """Impersonated Sheets + Drive clients (same identity that owns the interview
    folder) so every schedule file in it is readable regardless of service-account
    sharing. Uses ONLY the Drive scope — exactly the scope upload_report already
    impersonates with successfully — because the service account's domain-wide
    delegation is authorised for that scope. (Requesting an extra, un-delegated
    scope such as spreadsheets makes the whole impersonated token fail, which
    silently emptied the reminder tab.) The Sheets API accepts the Drive scope for
    reads, so both clients are built from the same Drive-scoped credentials."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    creds = service_account.Credentials.from_service_account_file(
        AR.SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/drive"],
    ).with_subject(IMPERSONATE_USER)
    sheets = gbuild("sheets", "v4", credentials=creds, cache_discovery=False)
    drive  = gbuild("drive", "v3", credentials=creds, cache_discovery=False)
    return sheets, drive


def _iv_prev_non_sunday(d):
    """Latest day strictly before d that is not a Sunday (weekday 6)."""
    d = d - timedelta(days=1)
    while d.weekday() == 6:
        d = d - timedelta(days=1)
    return d


def _iv_reminder_days(interview_date):
    """(message_day, call_day) for an interview date. Call is the last non-Sunday
    before the interview; Message the last non-Sunday before the Call. Guarantees
    two distinct non-Sunday days, Call closest to the interview."""
    call_day = _iv_prev_non_sunday(interview_date)
    msg_day  = _iv_prev_non_sunday(call_day)
    return msg_day, call_day


def _iv_filename_date(name):
    """Interview start date from the file name's last '_'-token (e.g. '21-Sep-2026')."""
    tok = str(name or "").strip().split("_")[-1]
    for fmt in ("%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(tok, fmt).date()
        except ValueError:
            continue
    return None


def _iv_parse_slot(cell):
    """Parse an 'Interview Time' cell like '21 Sep 2026  08:00 PM - 08:10 PM'.
    Returns (interview_date, start_datetime, start_label, display_text) or None."""
    s = _iv_re.sub(r"\s+", " ", str(cell or "")).strip()
    if not s:
        return None
    parts = _iv_re.split(r"\s*[–—\-]\s*", s)      # en/em/hyphen dash
    head = parts[0].strip()                                 # "21 Sep 2026 08:00 PM"
    end  = parts[1].strip() if len(parts) > 1 else ""
    dt = None
    for fmt in ("%d %b %Y %I:%M %p", "%d %B %Y %I:%M %p"):
        try:
            dt = datetime.strptime(head, fmt); break
        except ValueError:
            continue
    if dt is None:
        return None
    start_label = dt.strftime("%I:%M %p").lstrip("0")
    disp = f"{dt.strftime('%d-%b-%Y')}  {start_label}" + (f" - {end}" if end else "")
    return dt.date(), dt, start_label, disp


def _iv_read_grid(sheets, fid, tab):
    """Raw cell grid (list of rows) for one tab; [] on any error."""
    try:
        resp = sheets.spreadsheets().values().get(
            spreadsheetId=fid, range=f"'{tab}'!A1:Z200").execute()
        return resp.get("values", [])
    except Exception as e:                                  # pragma: no cover
        log.warning("Interview file %s tab '%s' unreadable (%s).", fid, tab, e)
        return []


def _iv_label_value(grid, label):
    """First non-blank value to the RIGHT of the cell whose text equals/starts with
    `label` (case-insensitive, trailing ':' ignored). '' if not found."""
    lab = label.strip().lower().rstrip(":")
    for row in grid:
        for j, cell in enumerate(row):
            t = str(cell or "").strip().lower().rstrip(":")
            if t == lab or (lab and t.startswith(lab)):
                for k in range(j + 1, len(row)):
                    v = str(row[k] or "").strip()
                    if v:
                        return v
    return ""


def _iv_feedback_rows(grid):
    """Yield (time_cell, candidate_name) from the Interview Feedback grid, locating
    the header row (the one carrying both an 'Interview Time' and a 'Candidate'
    column) dynamically so exact row positions are never assumed."""
    hdr_idx = time_col = name_col = None
    for i, row in enumerate(grid[:8]):
        low = [str(c or "").strip().lower().replace("\n", " ") for c in row]
        has_time = any(("interview" in c and "time" in c) for c in low)
        has_cand = any("candidate" in c for c in low)
        if has_time and has_cand:
            hdr_idx = i
            for j, c in enumerate(low):
                if time_col is None and ("interview" in c and "time" in c):
                    time_col = j
                if name_col is None and "candidate" in c:
                    name_col = j
            break
    if hdr_idx is None or time_col is None or name_col is None:
        return
    for row in grid[hdr_idx + 1:]:
        tcell = row[time_col] if time_col < len(row) else ""
        ncell = row[name_col] if name_col < len(row) else ""
        name = str(ncell or "").strip()
        if name:
            yield str(tcell or "").strip(), name


def _iv_instructor_index(sess_agg):
    """From the report's own session data: (norm course_name, norm course_title) ->
    instructor name, and (norm course_name) -> instructor name. Latest session wins."""
    by_full, by_name = {}, {}
    if sess_agg is None or getattr(sess_agg, "empty", True):
        return by_full, by_name
    df = AR._prefer_instructor_name(sess_agg)
    if "course_name" not in df.columns or "tutor_name" not in df.columns:
        return by_full, by_name
    d = df.copy()
    if "start_time_ist" in d.columns:
        d["_iv_k"] = d["start_time_ist"].apply(_parse_dt)
        d = d.sort_values("_iv_k", na_position="first")
    for _, r in d.iterrows():
        nm = str(r.get("tutor_name", "") or "").strip()
        if not nm:
            continue
        cn = _norm_name(r.get("course_name", ""))
        ct = _norm_name(r.get("course_title", ""))
        if cn:
            by_name[cn] = nm
            if ct:
                by_full[(cn, ct)] = nm
    return by_full, by_name


def _iv_phone_index(att_agg):
    """From the report's own attendance data: (norm student_name, norm course_name)
    -> phone, and (norm student_name) -> phone (first non-blank wins)."""
    by_full, by_name = {}, {}
    if att_agg is None or getattr(att_agg, "empty", True):
        return by_full, by_name
    if "student_name" not in att_agg.columns:
        return by_full, by_name
    for _, r in att_agg.iterrows():
        nm = _norm_name(r.get("student_name", ""))
        ph = str(r.get("phone", "") or "").strip()
        if not nm or not ph:
            continue
        cn = _norm_name(r.get("course_name", ""))
        by_name.setdefault(nm, ph)
        if cn:
            by_full.setdefault((nm, cn), ph)
    return by_full, by_name


def _iv_resolve_instructor(inst_full, inst_name, batch_name, batch_title):
    cn, ct = _norm_name(batch_name), _norm_name(batch_title)
    if cn and ct and (cn, ct) in inst_full:
        return inst_full[(cn, ct)]
    return inst_name.get(cn, "")


def _iv_resolve_phone(ph_full, ph_name, candidate, batch_name):
    nm, cn = _norm_name(candidate), _norm_name(batch_name)
    if cn and (nm, cn) in ph_full:
        return ph_full[(nm, cn)]
    return ph_name.get(nm, "")


def _iv_learner_runs(action, name, time_label, date_label):
    """Person-specific, action-oriented Why-Flagged runs for a learner. The learner
    name, date, time and the Call/Message action are highlighted; wording is built
    dynamically from those values (nothing hard-coded per person)."""
    who = str(name or "").strip() or "the learner"
    if action == "message":                                    # 2 days prior
        return [[("Interview on ", False), (date_label, True), (" at ", False),
                 (time_label, True), (" — ", False),
                 ("Message ", True), (who, True),
                 (" one-to-one, share the interview details, and ask them to "
                  "prepare and join on time.", False)]]
    return [[("Interview tomorrow at ", False), (time_label, True),      # 1 day prior
             (" — ", False), ("Call ", True), (who, True),
             (", confirm attendance, remind them to prepare well, follow interview "
              "etiquette, and join on time.", False)]]


def _iv_instructor_runs(action, time_label, date_label):
    """Action-oriented Why-Flagged runs for the instructor (key info highlighted)."""
    if action == "message":
        return [[("Interview scheduled at ", False), (time_label, True),
                 (" on ", False), (date_label, True), (" — ", False),
                 ("Message Instructor", True),
                 (", confirm availability, ensure the interviewer guidelines & Q&A "
                  "are reviewed, and be ready to start on time.", False)]]
    return [[("Interview tomorrow at ", False), (time_label, True),
             (" on ", False), (date_label, True), (" — ", False),
             ("Call Instructor", True),
             (", confirm availability, ensure guidelines/Q&A are reviewed, follow "
              "interview etiquette, and start on time.", False)]]


def _iv_name_full_key(name):
    """Cleansed, case-insensitive full-name key (collapse whitespace, casefold)."""
    return " ".join(str(name or "").split()).casefold()


def _iv_name_comp_keys(name):
    """'First name + Last-name initial' compressed key(s), lowercase, no spaces —
    e.g. 'Kumar Aradhya' -> {'kumara'}. A single-token name -> just that token. A
    3+-token name returns both first+last-initial and first+second-initial so either
    naming convention matches. Generic for any interviewer name (nothing hard-coded)."""
    toks = [t for t in str(name or "").split() if t]
    if not toks:
        return set()
    first = toks[0].casefold()
    if len(toks) == 1:
        return {first}
    return {(first + toks[-1][0]).casefold(), (first + toks[1][0]).casefold()}


def load_interviewer_phone_index(service):
    """Build interviewer phone lookups from IntelliBIStudentInfo → Instructor tab.
    Returns {'active': {...}, 'inactive': {...}}, each a {'full': {key->phone},
    'comp': {key->phone}} keyed by the cleansed full name AND the compressed
    First+Last-Initial form, so an interviewer name can be matched either way.
    Active and Inactive rows are kept separate to honour the search priority."""
    empty = lambda: {"full": {}, "comp": {}}
    idx = {"active": empty(), "inactive": empty()}
    try:
        df = AR.read_sheet_df(service, INSTRUCTOR_TAB_SHEET_ID, INSTRUCTOR_TAB)
    except Exception as e:
        log.warning("Could not read Instructor tab for interviewer match (%s).", e)
        return idx
    if df is None or df.empty:
        return idx
    has_active = "Is_Active" in df.columns
    for _, r in df.iterrows():
        is_active = (str(r.get("Is_Active", "")).strip().upper() == "Y") if has_active else True
        bucket = idx["active"] if is_active else idx["inactive"]
        for i in range(1, INSTRUCTOR_NAME_SLOTS + 1):
            raw = str(r.get(f"instructor_name_{i}", "") or "").strip()
            ph  = str(r.get(f"alternative_contact_number_{i}", "") or "").strip()
            if not raw or not ph:
                continue
            fk = _iv_name_full_key(raw)
            if fk:
                bucket["full"].setdefault(fk, ph)
            for ck in _iv_name_comp_keys(raw):
                bucket["comp"].setdefault(ck, ph)
    return idx


def _iv_match_interviewer_phone(idx, interviewer_name):
    """Phone for an interviewer name, in priority order: Active exact full-name,
    Active First+Last-Initial, then the same for Inactive. Cleansed & case-
    insensitive. '' when no match (caller keeps the phone blank)."""
    if not str(interviewer_name or "").strip():
        return ""
    fk = _iv_name_full_key(interviewer_name)
    cks = _iv_name_comp_keys(interviewer_name)
    for state in ("active", "inactive"):
        bucket = idx.get(state, {"full": {}, "comp": {}})
        if fk and fk in bucket["full"]:            # exact full-name match first
            return bucket["full"][fk]
        for ck in cks:                             # then First + Last-Initial
            if ck in bucket["comp"]:
                return bucket["comp"][ck]
            if ck in bucket["full"]:               # tab stored the compressed form
                return bucket["full"][ck]
    return ""


def _iv_interviewer_runs(action, name, time_label, date_label, note=None):
    """Person-specific, action-oriented Why-Flagged runs for the INTERVIEWER. The
    interviewer name, date, time and the Call/Message action are highlighted; the
    wording is built dynamically from those values. When the interviewer name is
    unknown (older files) it reads 'the interviewer'. An optional `note` (phone/name
    not available) is appended as a second highlighted bullet."""
    who = str(name or "").strip() or "the interviewer"
    if action == "message":                                    # 2 days prior
        lead = [("Interview on ", False), (date_label, True), (" at ", False),
                (time_label, True), (" — ", False),
                ("Message ", True), (who, True),
                (" to confirm interview readiness and availability, and to review "
                 "the interviewer guidelines & Q&A before the interview.", False)]
    else:                                                      # 1 day prior
        lead = [("Interview tomorrow at ", False), (time_label, True),
                (" on ", False), (date_label, True), (" — ", False),
                ("Call ", True), (who, True),
                (" and confirm interview readiness and availability.", False)]
    runs = [lead]
    if note:
        runs.append([(note, True)])
    return runs


def _iv_norm_date(value):
    """Parse a 'dd-MMM-yyyy' (or full-month) date to a date object; None if it is
    not a date. Used for filename dates and the Consolidate 'Interview Date'."""
    s = str(value or "").strip()
    if not s:
        return None
    for fmt in ("%d-%b-%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _iv_filename_dates(name):
    """(start_date, end_date) from the interview file name's trailing date token(s).
    Newer files carry BOTH dates ('..._22-Sep-2026_23-Sep-2026'); older files carry
    one ('..._21-Sep-2026') -> start == end. Returns None when no trailing date."""
    toks = str(name or "").strip().split("_")
    dts = []
    for tok in reversed(toks):
        d = _iv_norm_date(tok)
        if d is None:
            break                       # stop at the first non-date token from the end
        dts.append(d)
    if not dts:
        return None
    return min(dts), max(dts)


def _iv_meta_value(grid, key):
    """Value of a 'KEY:value' metadata cell (e.g. 'TECH_STACKS:Python') from the
    Interview Feedback tab's machine-readable header row. '' if not present."""
    key_l = str(key).strip().lower()
    for row in grid[:6]:
        for cell in row:
            s = str(cell or "")
            if ":" in s and s.split(":", 1)[0].strip().lower() == key_l:
                return s.split(":", 1)[1].strip()
    return ""


def _iv_read_range(sheets, fid, a1):
    """Raw grid for an explicit A1 range (first sheet when no tab prefix). [] on error."""
    try:
        resp = sheets.spreadsheets().values().get(spreadsheetId=fid, range=a1).execute()
        return resp.get("values", [])
    except Exception as e:                                  # pragma: no cover
        log.warning("Sheet range %s unreadable (%s).", a1, e)
        return []


def _iv_consolidate_index(sheets):
    """Set of (norm Batch Name, norm Batch Title, Interview Date) present in the
    Interview Consolidate Sheet. Returns None when the sheet/header can't be read,
    so callers can skip flagging rather than raise false 'not completed' alerts."""
    grid = _iv_read_range(sheets, INTERVIEW_CONSOLIDATE_SHEET_ID, "A1:U50000")
    if not grid:
        return None
    hdr_i = bn_c = bt_c = dt_c = None
    for i, row in enumerate(grid[:8]):
        low = [str(c or "").strip().lower() for c in row]
        if "batch name" in low and "batch title" in low and "interview date" in low:
            hdr_i = i
            bn_c = low.index("batch name")
            bt_c = low.index("batch title")
            dt_c = low.index("interview date")
            break
    if hdr_i is None:
        return None                     # header not found -> cannot validate
    idx = set()
    for row in grid[hdr_i + 1:]:
        bn = _norm_name(row[bn_c]) if bn_c < len(row) else ""
        bt = _norm_name(row[bt_c]) if bt_c < len(row) else ""
        d  = _iv_norm_date(row[dt_c]) if dt_c < len(row) else None
        if bn and bt and d:
            idx.add((bn, bt, d))
    return idx


def _iv_feedback_missing_runs(end_date_label, interviewer, batch_name):
    """Dynamic, action-oriented Why-Flagged runs for a missing interview feedback.
    Highlights the end date, the 'feedback not available' point and the interviewer
    name; when the interviewer name is unknown the wording omits it gracefully."""
    who = str(interviewer or "").strip()
    lead = [("Interview ended on ", False), (end_date_label, True),
            (", but feedback is not available in the Interview Consolidate Sheet", True),
            (" — ", False)]
    if who:
        tail = [("Check with ", False), (who, True),
                (" whether the interview feedback has been completed and submitted.",
                 False)]
    else:
        tail = [("Please check whether the interview feedback for ", False),
                (str(batch_name or "this batch"), True),
                (" has been completed and submitted.", False)]
    return [lead + tail]


def load_interview_feedback_validation(sheets, drive, report_date):
    """Interviews whose END date was yesterday (report_date == End Date + 1) and
    whose feedback is NOT yet present in the Interview Consolidate Sheet. Returns
    display rows for _wise_render_section: cells = [Interview Start Date, Interviewer
    Name, Tech Stack, Batch Name, Batch Title / Duration]. Nothing hard-coded."""
    try:
        q = (f"'{INTERVIEW_FOLDER_ID}' in parents and "
             f"mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")
        files = drive.files().list(
            q=q, fields="files(id,name)", pageSize=1000, supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute().get("files", [])
    except Exception as e:
        log.warning("Interview folder unreadable for feedback validation (%s).", e)
        return []

    # Only files whose interview END date was yesterday need checking today.
    todo = []
    for fmeta in files:
        span = _iv_filename_dates(fmeta.get("name", ""))
        if not span:
            continue
        start_d, end_d = span
        if report_date == end_d + timedelta(days=1):
            todo.append((fmeta.get("id", ""), start_d, end_d))
    if not todo:
        return []

    consolidated = _iv_consolidate_index(sheets)
    if consolidated is None:            # can't read the Consolidate Sheet -> don't flag
        log.warning("Interview Consolidate Sheet not readable — feedback validation skipped.")
        return []

    rows, seen = [], set()
    for fid, start_d, end_d in todo:
        helper = _iv_read_grid(sheets, fid, INTERVIEW_HELPER_TAB)
        batch_name  = _iv_label_value(helper, "Batch Name (Class)")
        batch_title = _iv_label_value(helper, "Batch Title / Duration")
        interviewer = _iv_label_value(helper, "Interviewer Name")
        # Tech Stack — from the Interview Feedback metadata; fall back to Batch Name.
        tech = _iv_meta_value(_iv_read_grid(sheets, fid, INTERVIEW_FEEDBACK_TAB),
                              "TECH_STACKS") or batch_name

        dk = (_norm_name(batch_name), _norm_name(batch_title), start_d, end_d)
        if dk in seen:
            continue
        seen.add(dk)

        # Feedback present if ANY interview day in [start..end] is in the Consolidate.
        found, d = False, start_d
        while d <= end_d:
            if (_norm_name(batch_name), _norm_name(batch_title), d) in consolidated:
                found = True
                break
            d += timedelta(days=1)
        if found:
            continue

        rows.append({
            "cells": [start_d.strftime("%d-%b-%Y"), interviewer, tech,
                      batch_name, batch_title],
            "status_cells": {}, "severity": 2,
            "why_reasons": _iv_feedback_missing_runs(
                end_d.strftime("%d-%b-%Y"), interviewer, batch_name),
        })
    return rows


def load_interview_reminders(sheets, drive, report_date):
    """Interview follow-up GROUPS due on report_date. Each group is one
    (batch, interview_date, action) with its due candidates. A candidate is due
    when report_date equals its Message day or Call day (per-candidate interview
    date; Sundays skipped/shifted)."""
    try:
        q = (f"'{INTERVIEW_FOLDER_ID}' in parents and "
             f"mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")
        files = drive.files().list(
            q=q, fields="files(id,name,modifiedTime)", pageSize=1000,
            orderBy="modifiedTime desc", supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute().get("files", [])
    except Exception as e:
        log.warning("Interview folder unreadable (%s) — reminder tab will be empty.", e)
        return []

    groups, seen = {}, set()
    lo, hi = report_date - timedelta(days=5), report_date + timedelta(days=60)
    for fmeta in files:
        fid, fname = fmeta.get("id", ""), fmeta.get("name", "")
        fd = _iv_filename_date(fname)
        if fd is not None and not (lo <= fd <= hi):
            continue                                   # far-past / far-future schedule
        helper = _iv_read_grid(sheets, fid, INTERVIEW_HELPER_TAB)
        batch_name  = _iv_label_value(helper, "Batch Name (Class)")
        batch_title = _iv_label_value(helper, "Batch Title / Duration")
        # Interviewer Name — read straight from the Interview_Helper tab (present
        # only in the latest schedule files, just below "Batch Title / Duration:").
        # Older files lack it -> stays blank; we NEVER fall back to the Instructor.
        interviewer_name = _iv_label_value(helper, "Interviewer Name")
        for tcell, cname in _iv_feedback_rows(_iv_read_grid(sheets, fid, INTERVIEW_FEEDBACK_TAB)):
            slot = _iv_parse_slot(tcell)
            if slot is None:
                continue
            idate, sdt, slabel, disp = slot
            msg_day, call_day = _iv_reminder_days(idate)
            if report_date == msg_day:
                action = "message"
            elif report_date == call_day:
                action = "call"
            else:
                continue
            dk = (idate, action, _norm_name(cname), sdt.strftime("%H:%M"))
            if dk in seen:
                continue
            seen.add(dk)
            key = (_norm_name(batch_name), _norm_name(batch_title), idate, action)
            g = groups.setdefault(key, {
                "batch_name": batch_name, "batch_title": batch_title,
                "interviewer_name": interviewer_name,
                "interview_date": idate, "action": action, "candidates": []})
            g["candidates"].append({"sort": sdt, "time_disp": disp,
                                    "start_label": slabel, "name": cname})

    out = []
    for g in groups.values():
        g["candidates"].sort(key=lambda c: c["sort"])
        g["batch_time_label"] = g["candidates"][0]["start_label"] if g["candidates"] else ""
        out.append(g)
    out.sort(key=lambda g: (g["interview_date"], _norm_name(g["batch_name"]),
                            0 if g["action"] == "message" else 1))
    return out


def _iv_write_row(ws, row_num, values, why_runs, level, bold, zebra_i=0, chip_text=None):
    """One reminder row in the common style: the first cell (Interview Time) is
    the priority chip (Medium = call today, Info = message today; the interviewer
    row is a bold 'Info' row), other cells neutral, Why Flagged numbered with the
    action line."""
    bg = ds_zebra(zebra_i)
    for col_name in IV_COLS:
        cidx = _IVC[col_name]
        if col_name in FOLLOWUP_COLS:
            continue                                   # written once below
        if col_name == "Why Flagged":
            ds_why_text(ws, row_num, cidx, why_runs, col_width=66, min_height=34)
        elif col_name == "Interview Time":
            ds_priority(ws, row_num, cidx, values.get(col_name, ""), level, h_align="left")
        else:
            h_align = "center" if col_name == "Phone" else "left"
            ds_cell(ws, row_num, cidx, values.get(col_name, ""), bg=bg, bold=bold,
                    h_align=h_align, wrap=True)
    ds_followup_cells(ws, row_num, _IVC["Action Taken"])


def _iv_finish(ws, row_num):
    ds_finish(ws, 3, IVN, widths={
        "Interview Time": 32, "Batch Name (Class)": 20, "Batch Title / Duration": 28,
        "Candidate Name": 28, "Phone": 16,
    }, why_col=_IVC["Why Flagged"], why_width=66, default_width=16, tab_color=DS_SECTION,
       followup=(_IVC["Action Taken"], FOLLOWUP_ACTIONS["interview"]),
       freeze_after_col=_IVC["Phone"])


def build_interview_reminders(ws, sheets, drive, service, att_agg,
                              report_date, period_label=None):
    """Learner Instructor Interview Reminder tab. Interviewer row first (distinct
    background), then learner rows sorted by interview time; a banner separates
    each batch/interview. The interviewer comes from the schedule file's
    Interview_Helper tab (never the batch Instructor). Additive only.
    Presentation: common report design system — interview = section banner with
    today's action (MESSAGE / CALL), interviewer row = bold Info row, learners as
    neutral zebra rows, Interview Time cell = priority chip."""
    period = period_label or report_date.strftime('%d-%b-%Y')
    ds_title(ws, IVN, "Learner & Instructor Interview Reminders", period,
             "Interviews that need a reminder today: MESSAGE two days before, CALL one day before. "
             "The interviewer row comes first in each group, then the candidates by time. "
             "Priority chip on Interview Time: Medium = call today, Info = message today."
             + DS_GUIDE_FOLLOWUP)
    HDR_ROW = 3
    row_num = ds_headers(ws, HDR_ROW, IV_COLS)
    ds_followup_headers(ws, HDR_ROW, _IVC["Action Taken"])

    try:
        groups = load_interview_reminders(sheets, drive, report_date)
    except Exception as e:                              # never break the report
        log.exception("Interview reminder tab failed to load: %s", e)
        groups = []

    if not groups:
        ds_empty(ws, row_num, IVN, "✅  No interview follow-ups due today.")
        ds_guide_count(ws, "Reminders due today", 0)
        _iv_finish(ws, row_num + 1)
        return ws

    iv_phone_idx = load_interviewer_phone_index(service)   # Instructor tab -> phone
    ph_full, ph_name = _iv_phone_index(att_agg)

    n_rows = 0
    for g in groups:
        bname  = g["batch_name"] or "—"
        btitle = g["batch_title"] or "—"
        idate  = g["interview_date"]
        action = g["action"]
        date_label = idate.strftime("%d-%b-%Y")
        action_txt = "MESSAGE (2 days prior)" if action == "message" else "CALL (1 day prior)"
        level = "medium" if action == "call" else "info"
        # ── batch/interview separator banner ─────────────────────────────────
        row_num = ds_section(
            ws, row_num, IVN,
            f"Interview — {bname}   ·   {btitle}   ·   {date_label}"
            f"      Today: {action_txt}   ·   {len(g['candidates'])} candidate(s)", level=1)

        # ── Interviewer reminder — FIRST record, distinct (bold) row ──────────
        # Interviewer identity comes ONLY from the schedule file (Interview_Helper);
        # phone from the Instructor master tab (Active→Inactive, exact→initial). No
        # fallback to the batch Instructor. Name/phone stay blank when unavailable;
        # the record is always kept and Why Flagged explains any gap.
        interviewer = str(g.get("interviewer_name", "") or "").strip()
        batch_time = g["batch_time_label"] or "—"
        if interviewer:
            iv_phone = _iv_match_interviewer_phone(iv_phone_idx, interviewer)
            iv_name_cell = f"Interviewer: {interviewer}"
            iv_note = (None if iv_phone else
                       "Interviewer phone number not available — please look it up "
                       "before following up.")
        else:
            iv_phone = ""
            iv_name_cell = ""            # keep blank; never fall back to Instructor
            iv_note = ("Interviewer name not provided in the interview schedule — "
                       "please identify the interviewer before following up.")
        _iv_write_row(ws, row_num, {
            "Interview Time": f"{date_label}  from {batch_time}",
            "Batch Name (Class)": bname,
            "Batch Title / Duration": btitle,
            "Candidate Name": iv_name_cell,
            "Phone": iv_phone,           # blank when not found (no placeholder)
        }, _iv_interviewer_runs(action, interviewer, batch_time, date_label, iv_note),
           level=level, bold=True, zebra_i=0)
        # make the interviewer row visibly distinct from the candidates below
        for cidx in range(2, _IVC["Why Flagged"]):
            ws.cell(row=row_num, column=cidx).fill = AR._fill(DS_INFO_BG)
        row_num += 1
        n_rows += 1

        # ── Learner reminders — sorted by interview time ascending ───────────
        for i, c in enumerate(g["candidates"]):
            phone = _iv_resolve_phone(ph_full, ph_name, c["name"], bname)
            _iv_write_row(ws, row_num, {
                "Interview Time": c["time_disp"],
                "Batch Name (Class)": bname,
                "Batch Title / Duration": btitle,
                "Candidate Name": c["name"],
                "Phone": phone or "—",
            }, _iv_learner_runs(action, c["name"], c["start_label"], date_label),
               level=level, bold=False, zebra_i=i)
            row_num += 1
            n_rows += 1

    ds_guide_count(ws, "Reminders due today", n_rows)
    _iv_finish(ws, row_num)
    return ws


def _generate_daily(service, report_date, sess_agg, att_agg, fb_agg, tf_agg, susp_agg):
    """Daily report for `report_date` — unchanged daily logic/tabs/upload; only the
    report date is a parameter and aggregates are sliced as-of that date."""
    # ── daily window (exactly as the existing daily report: yest-noon → now) ──
    win_start = datetime.combine(report_date - timedelta(days=1), time(12, 0, 0))
    win_end = (AR._ist_now() if report_date == AR._ist_today()
               else datetime.combine(report_date, time(23, 59, 59)))
    yest_date = report_date - timedelta(days=1)

    def _in_window(df):
        if df is None or df.empty or "_dt" not in df.columns:
            return df.iloc[0:0].copy() if df is not None else df
        mask = df["_dt"].apply(lambda x: x is not None and not pd.isna(x) and win_start <= x <= win_end)
        return df[mask].copy()

    att_daily = _in_window(att_agg)
    fb_daily = _in_window(fb_agg)
    susp_daily = _in_window(susp_agg)
    sess_daily = _in_window(sess_agg)
    tf_daily = _in_window(tf_agg)
    log.info("Daily window %s → %s : %d session(s), %d attendance row(s), %d suspended.",
             win_start, win_end, 0 if sess_daily is None else len(sess_daily),
             0 if att_daily is None else len(att_daily),
             0 if susp_daily is None else len(susp_daily))

    # daily window label — same format as the existing daily report
    label = (f"{win_start.strftime('%d-%b-%Y %I:%M %p')} - "
             f"{win_end.strftime('%d-%b-%Y %I:%M %p')}")

    # ── aggregates (till report date) ─────────────────────────────────────────
    cfg = COORD.load_config() if _COORD_OK else None
    agg_map = build_aggregates(_le_date(att_agg, report_date), _le_date(fb_agg, report_date), cfg)
    log.info("Computed aggregates for %d (student × tech × duration) group(s).", len(agg_map))

    # ── instructor phone directory (Instructor master; Is_Active=Y) ───────────
    instr_phones = load_instructor_phones(service)
    log.info("Instructor phone directory: %d name(s).", len(instr_phones))

    # ── build workbook: Learner Attendance Follow-Ups + Learner Assignment
    #    Follow-Ups + Learner Admission Formalities + Learner Wise Validation +
    #    Instructor Follow-Ups ──────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_student_detail_ext(wb.create_sheet("Learner Attendance Follow-Ups"),
                             att_daily, susp_daily, fb_daily, agg_map, report_date,
                             yest_date=yest_date)
    # Learner Assignment Follow-Ups — assignment-submission defaulters, reusing
    # pyAssignmentSubmissionEmailReminder's own reminder/defaulter logic.
    build_assignment_followups(wb.create_sheet("Learner Assignment Follow-Ups"),
                               load_assignment_followups(service, report_date), report_date)
    # Learner Admission Formalities — admission-form / e-signature pending list,
    # reusing pyAdmissionFormalitiesReport's own matching/status logic wholesale.
    build_admission_formalities(wb.create_sheet("Learner Admission Formalities"),
                                load_admission_formalities(service), report_date)
    # Wise & Interview Feedback Validation — Student / Course / Instructor data-
    # validation failures (reused wholesale) PLUS an additive 'Interview Feedback
    # Not Completed' section (interviews that ended yesterday with no feedback yet
    # in the Interview Consolidate Sheet). The interview part is wrapped so it can
    # never block the rest of the report.
    _ifv_rows = []
    try:
        _ifv_sheets, _ifv_drive = _iv_services()
        _ifv_rows = load_interview_feedback_validation(_ifv_sheets, _ifv_drive, report_date)
    except Exception as _ifve:
        log.exception("Interview Feedback validation skipped (%s).", _ifve)
    build_wise_validation(wb.create_sheet("Wise & Interview Feedback Validation"),
                          load_wise_validation(), report_date, interview_rows=_ifv_rows)
    build_instructor_followups(wb.create_sheet("Instructor Follow-Ups"),
                               sess_daily, att_daily, fb_daily, tf_daily,
                               instr_phones, report_date)
    # Learner Instructor Interview Reminder — upcoming-interview follow-ups
    # (Message 2 days prior / Call 1 day prior; Sundays skipped and shifted back).
    # Additive tab only; wrapped so it can never block the rest of the report.
    try:
        _iv_sheets, _iv_drive = _iv_services()
        build_interview_reminders(
            wb.create_sheet("Learner Instructor Interview Reminder"),
            _iv_sheets, _iv_drive, service, att_agg, report_date)
    except Exception as _ive:
        log.exception("Interview Reminder tab skipped (%s).", _ive)
    ds_enable_iterative_calc(wb)            # keeps the Follow-Up DateTime formula stable
    buf = io.BytesIO()
    wb.save(buf)
    # Filename carries the report period ("duration"), same convention/transform
    # as pyAttendaceFeedbackReport.py's daily report.
    safe_label = label.replace(" ", "_").replace("–", "to").replace("/", "-").replace(":", ".")
    filename = f"{REPORT_BASENAME}_{safe_label}.xlsx"
    log.info("Workbook built with tabs: %s (%d bytes).",
             ", ".join(wb.sheetnames), len(buf.getvalue()))

    # ── upload date-wise as a native Google Sheet ─────────────────────────────
    link = upload_datewise(report_date, filename, buf)
    n_daily = 0 if att_daily is None else len(att_daily)
    print(f"\n{'='*64}\n  Batch Coordinator Daily Attendance — {report_date}")
    print(f"  Daily rows: {n_daily} | Aggregate groups: {len(agg_map)}")
    print(f"  Link: {link}\n{'='*64}\n")
    return {"link": link, "report_date": report_date.isoformat(),
            "daily_rows": n_daily, "agg_groups": len(agg_map)}


# =============================================================================
#  REPORT PLANNING  (flag-based; mirrors pyLeadFollowUpAnalysisReport.py)
# =============================================================================
def _week_label(mon, sun):
    return f"{mon.strftime('%d %b')} – {sun.strftime('%d %b %Y')}"


def _month_label(y, m):
    return date(y, m, 1).strftime("%B %Y")


def _plan_jobs(today):
    """Decide which report(s) to generate and for what period, from the control
    flags. Returns a list of job dicts:
        {"kind": "daily",  "date": <date>}
        {"kind": "period", "report_type": <str>, "start": <date>, "end": <date>,
         "label": <str>}
    Only WHICH report runs and its period are decided here — the per-report logic
    (tabs, calculations, formatting, upload) is unchanged."""
    jobs = []

    # ── AUTO: pick reports from the run date (manual flags ignored) ───────────
    if GENERATE_AUTO:
        # Daily — every run, for the current day.
        jobs.append({"kind": "daily", "date": today})
        # Weekly — every Monday, for the previous completed week (Mon → Sun).
        if today.weekday() == 0:                       # Monday
            mon = today - timedelta(days=7)
            sun = mon + timedelta(days=6)
            jobs.append({"kind": "period", "report_type": "weekly",
                         "start": mon, "end": sun, "label": _week_label(mon, sun)})
        # Monthly — on the last calendar day of the month, for that whole month.
        _last = calendar.monthrange(today.year, today.month)[1]
        if today.day == _last:
            first = date(today.year, today.month, 1)
            last = date(today.year, today.month, _last)
            jobs.append({"kind": "period", "report_type": "monthly",
                         "start": first, "end": last,
                         "label": _month_label(today.year, today.month)})
        return jobs

    # ── MANUAL flag mode: each flag independent; several may run in one pass ──
    if GENERATE_DAILY:
        d = datetime.strptime(DAILY_DATE, "%Y-%m-%d").date() if DAILY_DATE else today
        jobs.append({"kind": "daily", "date": d})
    if GENERATE_WEEKLY:
        ref = (datetime.strptime(WEEKLY_REFERENCE_DATE, "%Y-%m-%d").date()
               if WEEKLY_REFERENCE_DATE else today)
        mon = ref - timedelta(days=ref.weekday())      # Monday of ref's week
        sun = mon + timedelta(days=6)
        jobs.append({"kind": "period", "report_type": "weekly",
                     "start": mon, "end": sun, "label": _week_label(mon, sun)})
    if GENERATE_MONTHLY:
        yr = MONTHLY_YEAR or today.year
        mo = MONTHLY_MONTH or today.month
        first = date(yr, mo, 1)
        last = date(yr, mo, calendar.monthrange(yr, mo)[1])
        jobs.append({"kind": "period", "report_type": "monthly",
                     "start": first, "end": last, "label": _month_label(yr, mo)})
    if GENERATE_MANUAL:
        if not (MANUAL_START_DATE and MANUAL_END_DATE):
            log.warning("GENERATE_MANUAL is on but MANUAL_START_DATE / MANUAL_END_DATE "
                        "are not set — skipping the Manual report.")
        else:
            ms = datetime.strptime(MANUAL_START_DATE, "%Y-%m-%d").date()
            me = datetime.strptime(MANUAL_END_DATE, "%Y-%m-%d").date()
            jobs.append({"kind": "period", "report_type": "manual",
                         "start": ms, "end": me,
                         "label": f"{ms.strftime('%d-%b-%Y')} to {me.strftime('%d-%b-%Y')}"})
    return jobs


def generate():
    from utils import get_sheets_service        # plain service account (reads sources)
    service = get_sheets_service(AR.SERVICE_ACCOUNT_FILE)

    today = AR._ist_today()
    jobs = _plan_jobs(today)
    if not jobs:
        log.warning("No reports selected to generate (check the GENERATE_* flags).")
        return []
    log.info("Planned %d report job(s): %s", len(jobs),
             ", ".join(j["kind"] if j["kind"] == "daily" else j["report_type"] for j in jobs))

    # ── ONE full-history load up to the latest date any job needs ─────────────
    max_end = today
    for j in jobs:
        d = j["date"] if j["kind"] == "daily" else j["end"]
        if d > max_end:
            max_end = d
    early = date(2000, 1, 1)
    log.info("Loading attendance/feedback up to %s …", max_end)
    sess_agg, att_agg, fb_agg, tf_agg, susp_agg = AR.load_all_data(service, early, max_end)

    # ── run each planned report on the shared, in-memory data ─────────────────
    results = []
    for j in jobs:
        try:
            if j["kind"] == "daily":
                results.append(_generate_daily(
                    service, j["date"], sess_agg, att_agg, fb_agg, tf_agg, susp_agg))
            else:
                results.append(_generate_period(
                    service, j["report_type"], j["start"], j["end"], j["label"],
                    sess_agg, att_agg, fb_agg, tf_agg, susp_agg, today))
        except Exception as _je:
            log.exception("Report job %s failed: %s", j, _je)
            results.append({"failed": True, "job": str(j), "error": str(_je)})
    return results


def main():
    logging.basicConfig(level=logging.INFO if VERBOSE else logging.WARNING,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    return generate()


if __name__ == "__main__":
    # Non-zero exit when any planned report failed, so the Operations pipeline
    # (scripts/run_reports_action.py) records the failure and retries/alerts.
    _res = main()
    sys.exit(1 if any(isinstance(r, dict) and r.get("failed") for r in (_res or [])) else 0)
