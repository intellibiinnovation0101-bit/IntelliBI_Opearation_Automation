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
    One native Google Sheet per report date, written into a DATE-WISE sub-folder
    under the coordinator Drive folder, so history is never overwritten. The file
    name carries the report period ("duration"), like the daily report, e.g.:
        <parent>/YYYY-MM-DD/IntelliBI_Batch_Coordinator_Daily_Attendance_Report_
                            12-Sep-2026_10.00_AM_-_12-Sep-2026_03.00_PM
    Re-running the same day replaces that day's sheet only (matched by base name).
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
from datetime import datetime, date, timedelta, time

import pandas as pd
import openpyxl

# Reuse the existing daily report wholesale (helpers, styling, colours, loader).
import pyAttendaceFeedbackReport as AR

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
#  COLUMN LAYOUT
# =============================================================================
STUDENT_COLS = ["Rank", "Tech Name", "Duration", "Student Name", "Phone"]
DAILY_COLS   = ["Status", "Attendance %", "Duration (min)", "Joined At",
                "Left At", "Remarks", "Feedback Given?", "Feedback Rating (/10)"]
AGG_COLS     = ["Sessions (P/T)", "Agg Attendance %", "Agg Duration Att %",
                "Last Present", "Absent Streak", "Feedbacks", "Feedback Part. %",
                "Avg Rating (/10)", "Overall Remarks", "Why Flagged"]
HEADERS = STUDENT_COLS + DAILY_COLS + AGG_COLS
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
]
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
AF_COLS = ["#", "Student Name", "Email", "Phone", "Deadline", "Reminder", "Why Flagged"]
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
ADM_COLS = _ADM_BASE_COLS + ["Why Flagged"]
ADMN = len(ADM_COLS)
_ADMC = {h: i + 1 for i, h in enumerate(ADM_COLS)}   # 1-based column index by header


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
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(IN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_IC["Why Flagged"])].width = 46
    ws.column_dimensions[get_column_letter(_IC["Instructor"])].width = 20


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
    Teacher-No-Feedback logic exactly."""
    title = (f"IntelliBI  |  Batch Coordinator — Instructor Follow-Ups  |  "
             f"{period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, IN, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, INSTR_COLS)
    row_num = 3

    if sess_daily is None or sess_daily.empty:
        AR.write_section_banner(ws, row_num, IN, "  No sessions in this period.",
                                AR.C_GREY_BD, h_align="center")
        _finish_instr(ws, row_num + 1)
        return ws

    sess_f = AR._prefer_instructor_name(sess_daily)
    tf_ids = (set(tf_daily["session_id"].dropna().unique())
              if (tf_daily is not None and not tf_daily.empty and "session_id" in tf_daily.columns)
              else set())

    def _dedup_fb(sub):
        if sub is None or sub.empty:
            return sub
        cols = [c for c in ["session_id", "student_id"] if c in sub.columns]
        return sub.drop_duplicates(subset=cols) if cols else sub

    any_rendered = False
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
                                ("Ask students to submit their session feedback.", True)])
        if not reasons:
            continue

        # ── render ───────────────────────────────────────────────────────────
        bg = AR.C_WHITE
        plain = {
            "Tech Name": sess.get("course_name", ""),
            "Duration": sess.get("course_title", ""),
            "Instructor": instr,
            "Phone": phone or "—",
            "Session Date": (actual_start_dt.strftime("%Y-%m-%d %H:%M:%S")
                             if actual_start_dt is not None
                             else str(sess.get("start_time_ist", ""))[:19]),
            "Duration (Mins)": dur_min if dur_min else "",
            "Scheduled (Mins)": (round(sched_min, 1) if sched_min is not None else "—"),
            "Diff Mins": (f"{diff_min:+.0f}" if diff_min is not None else "—"),
            "Total Enrolled": total_enrolled,
            "Att. N/A Count": na_cnt,
            "Present": n_present,
            "Absent": absent_n,
            "No. of Feedbacks": n_fb,
            "Min Rating": (min_r if min_r is not None else "—"),
            "Max Rating": (max_r if max_r is not None else "—"),
        }
        for c in INSTR_COLS:
            if c in plain:
                AR.style_data_cell(ws, row_num, _IC[c], plain[c], bg=bg,
                                   h_align="left" if c in ("Tech Name", "Duration", "Instructor")
                                   else "center")
        AR.style_data_cell(ws, row_num, _IC["Att %"], att_pct, bg=AR._att_bg(att_pct, ""),
                           h_align="center", number_fmt='0.0"%"')
        AR.style_data_cell(ws, row_num, _IC["Avg Time in Session %"], avg_time,
                           bg=AR._att_bg(avg_time, ""), h_align="center", number_fmt='0.0"%"')
        AR.style_data_cell(ws, row_num, _IC["Feedback Rate %"], fb_rt,
                           bg=(AR.C_GREEN_PALE if fb_rt >= 50 else AR.C_AMBER if fb_rt > 0 else AR.C_RED_LITE),
                           h_align="center", number_fmt='0.0"%"')
        if avg_r is not None:
            AR.style_data_cell(ws, row_num, _IC["⭐ Avg Rating(/10)"], avg_r,
                               bg=AR._rating_bg(avg_r), h_align="center", number_fmt='0.0')
        else:
            AR.style_data_cell(ws, row_num, _IC["⭐ Avg Rating(/10)"], "—", bg=bg, h_align="center")

        dc = ws.cell(row=row_num, column=_IC["Diff Mins"])
        if diff_min is not None:
            if diff_min <= -INSTR_SHORT_MIN:
                dcol, dbg = AR.C_RED_DARK, AR.C_RED_LITE
            elif diff_min >= INSTR_SHORT_MIN:
                dcol, dbg = AR.C_AMBER_DARK, AR.C_AMBER
            else:
                dcol, dbg = AR.C_GREEN_DARK, AR.C_GREEN_PALE
            dc.fill = AR._fill(dbg); dc.font = AR._font(bold=True, size=10, color=dcol)
            dc.alignment = AR._align("center", "center"); dc.border = AR._border()

        fgc = ws.cell(row=row_num, column=_IC["Feedback Given"])
        fgc.value = "Yes" if fb_given else "No"
        fgc.fill = AR._fill(AR.C_GREEN if fb_given else AR.C_RED_LITE)
        fgc.font = AR._font(bold=True, size=10, color=AR.C_GREEN_DARK if fb_given else AR.C_RED_DARK)
        fgc.alignment = AR._align("center", "center"); fgc.border = AR._border()

        wc = ws.cell(row=row_num, column=_IC["Why Flagged"])
        _rt = _why_flagged_richtext(reasons)
        wc.value = _rt
        wc.fill = AR._fill(AR.C_AMBER_PALE)
        # rich text carries its own per-run colours; the base font only matters
        # for the plain-text fallback (kept red so key info still stands out).
        wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
        wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()

        ws.row_dimensions[row_num].height = max(30, 15 * len(reasons) + 8)
        any_rendered = True
        row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, IN,
                                "  ✅  No instructor follow-ups required for this period.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1
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
    pyAssignmentSubmissionEmailReminder."""
    title = (f"IntelliBI  |  Batch Coordinator — Learner Assignment Follow-Ups  |  "
             f"{period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, AFN, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, AF_COLS)
    row_num = 3

    any_rendered = False
    try:
        import pyAssignmentSubmissionEmailReminder as ASG
        _lp = ASG.LEVEL_PRIORITY
    except Exception:
        _lp = {"1st": 1, "2nd": 2, "final": 3, "missed": 4}

    for cn in sorted((by_tech or {}).keys(), key=lambda x: str(x).lower()):
        assignments = by_tech[cn]
        # technology banner (same blue as the attendance tab's tech banner)
        _assign_banner(ws, row_num, AFN, f"  {cn or '(Unknown Technology)'}",
                       AR.C_BLUE_MID, h_align="center")
        row_num += 1

        for title_txt, aid, rows in sorted(assignments, key=lambda t: str(t[0]).lower()):
            if not rows:
                continue
            r0 = rows[0]
            # Assignment Details header (reuses the same meta shown in the PDF)
            _assigned = r0.get("assigned_date_str", "") or "—"
            details = (f"  📝  {title_txt or '(Untitled Assignment)'}"
                       f"      •  Duration: {r0.get('class_subject','') or '—'}"
                       f"      •  Max Marks: {r0.get('maximum_marks','') or '—'}"
                       f"      •  Assigned: {_assigned}"
                       f"      •  Deadline: {r0.get('deadline_str','') or '—'}"
                       f"      •  Pending: {len(rows)}")
            _assign_banner(ws, row_num, AFN, details, AR.C_TEAL, h_align="left")
            row_num += 1

            # learners, most-urgent reminder first (same sort as the PDF)
            rows_sorted = sorted(
                rows, key=lambda r: (-_lp.get(r.get("reminder_level", ""), 0),
                                     str(r.get("student_name", "")).lower()))
            for i, r in enumerate(rows_sorted, 1):
                lvl = str(r.get("reminder_level", ""))
                row_bg = (AR.C_RED_LITE if lvl in ("final", "missed")
                          else AR.C_AMBER if lvl == "2nd" else AR.C_WHITE)
                AR.style_data_cell(ws, row_num, _AF["#"], i, bg=row_bg, h_align="center")
                AR.style_data_cell(ws, row_num, _AF["Student Name"],
                                   r.get("student_name", "") or "—", bg=row_bg, h_align="left")
                AR.style_data_cell(ws, row_num, _AF["Email"],
                                   r.get("student_email", "") or "—", bg=row_bg, h_align="left")
                AR.style_data_cell(ws, row_num, _AF["Phone"],
                                   r.get("student_phone", "") or "—", bg=row_bg, h_align="center")
                AR.style_data_cell(ws, row_num, _AF["Deadline"],
                                   r.get("deadline_str", "") or "—", bg=row_bg, h_align="center")
                rc = ws.cell(row=row_num, column=_AF["Reminder"])
                rc.value = r.get("reminder_label", "") or "—"
                _rcol = (AR.C_RED_DARK if lvl in ("final", "missed")
                         else AR.C_AMBER_DARK if lvl == "2nd" else AR.C_GREEN_DARK)
                rc.font = AR._font(bold=True, size=10, color=_rcol)
                rc.fill = AR._fill(row_bg); rc.alignment = AR._align("center", "center")
                rc.border = AR._border()
                wc = ws.cell(row=row_num, column=_AF["Why Flagged"])
                _rt = _why_flagged_richtext(_assignment_why_runs(r))
                wc.value = _rt
                wc.fill = AR._fill(AR.C_AMBER_PALE)
                wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
                wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()
                ws.row_dimensions[row_num].height = 28
                any_rendered = True
                row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, AFN,
                                "  ✅  No assignment submission follow-ups for this day.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1

    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(AFN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_AF["Why Flagged"])].width = 60
    ws.column_dimensions[get_column_letter(_AF["Email"])].width = 30
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
    immediately after 'Expiry Date'. Important status, dates/expiry and the
    required action are highlighted red, matching the other follow-up tabs."""
    title = (f"IntelliBI  |  Batch Coordinator — Learner Admission Formalities  |  "
             f"{period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, ADMN, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, ADM_COLS)
    row_num = 3

    _text_cols = {"Student Name", "Email ID", "Batch Name", "Request Form Name",
                  "Recipient Status", "Request Status"}
    any_rendered = False
    # most-urgent first (red → amber → rest); stable, so newest-joined order stays
    ordered = sorted(rows or [], key=_adm_severity_key)
    for r in ordered:
        rs = r.get("Recipient Status", "")
        qs = r.get("Request Status", "")
        bg = _adm_bg(rs, qs)
        for c in _ADM_BASE_COLS:
            val = r.get(c, "")
            AR.style_data_cell(ws, row_num, _ADMC[c],
                               (val if str(val).strip() else "—"),
                               bg=bg, h_align="left" if c in _text_cols else "center")
        # Recipient Status — highlight the key status (bold, status-coloured).
        sc = ws.cell(row=row_num, column=_ADMC["Recipient Status"])
        _scol = (AR.C_RED_DARK if bg == AR.C_RED_LITE
                 else AR.C_AMBER_DARK if bg == AR.C_AMBER
                 else AR.C_GREEN_DARK if bg in (AR.C_GREEN, AR.C_GREEN_PALE)
                 else "333333")
        sc.font = AR._font(bold=True, size=10, color=_scol)
        # Expiry Date — a live deadline; flag red when present on an attention row.
        if str(r.get("Expiry Date", "") or "").strip() and bg in (AR.C_RED_LITE, AR.C_AMBER):
            ec = ws.cell(row=row_num, column=_ADMC["Expiry Date"])
            ec.font = AR._font(bold=True, size=10, color=AR.C_RED_DARK)
        # Why Flagged — action-oriented, red-highlighted rich text.
        wc = ws.cell(row=row_num, column=_ADMC["Why Flagged"])
        _rt = _why_flagged_richtext(_admission_why_runs(r))
        wc.value = _rt
        wc.fill = AR._fill(AR.C_AMBER_PALE)
        wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
        wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()
        ws.row_dimensions[row_num].height = 30
        any_rendered = True
        row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, ADMN,
                                "  ✅  No pending admission formalities for this day.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1

    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(ADMN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_ADMC["Why Flagged"])].width = 64
    ws.column_dimensions[get_column_letter(_ADMC["Email ID"])].width = 30
    ws.column_dimensions[get_column_letter(_ADMC["Recipient Status"])].width = 20
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
                 "", "", "", "", "", "", rec.get("Joined On", "") or "—"]
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
    """Render one validation section: a full-width separator banner, a header row
    (with 'Why Flagged' pinned to the last column so it aligns across sections),
    then the failed records. Returns the next free row (with a trailing blank)."""
    r = start_row
    _assign_banner(ws, r, ncols_max, banner_text, AR.C_BLUE_MID, h_align="center")
    r += 1
    ndata = len(data_headers)
    full_hdr = list(data_headers) + [""] * (ncols_max - 1 - ndata) + ["Why Flagged"]
    AR.write_header_row(ws, r, full_hdr)
    r += 1
    if not display_rows:
        AR.write_section_banner(ws, r, ncols_max,
                                "  ✅  No attention-required records in this section.",
                                AR.C_GREEN_DARK, h_align="center")
        return r + 2
    for dr in display_rows:
        bg = _wise_sev_bg(dr["severity"])
        cells = dr["cells"]
        for ci in range(ndata):
            st = dr["status_cells"].get(ci)
            if st is not None:
                sbg, sfg = _wise_status_style(st)
                AR.style_data_cell(ws, r, ci + 1, st or "—", bg=sbg, h_align="center")
                ws.cell(row=r, column=ci + 1).font = AR._font(bold=True, size=10, color=sfg)
            else:
                val = cells[ci] if ci < len(cells) else ""
                AR.style_data_cell(ws, r, ci + 1, (val if str(val).strip() != "" else "—"),
                                   bg=bg, h_align="left" if ci in text_col_idxs else "center")
        for ci in range(ndata, ncols_max - 1):        # gap cells → keep the row band
            AR.style_data_cell(ws, r, ci + 1, "", bg=bg, h_align="center")
        wc = ws.cell(row=r, column=ncols_max)
        _rt = _why_flagged_richtext(dr["why_reasons"])
        wc.value = _rt
        wc.fill = AR._fill(AR.C_AMBER_PALE)
        wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
        wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()
        ws.row_dimensions[r].height = max(30, 15 * len(dr["why_reasons"]) + 8)
        r += 1
    return r + 1


def build_wise_validation(ws, data, report_date, period_label=None, interview_rows=None):
    """Wise & Interview Feedback Validation tab — combines pyWiseDataValidationReport's
    Student, Course and Instructor validation FAILURES, PLUS an 'Interview Feedback Not
    Completed' section, into one coordinator tab. Each section is separated by a header
    banner and ends in an action-oriented, red-highlighted 'Why Flagged' column. The
    existing Wise validation logic is reused unchanged; the interview-feedback section
    is purely additive."""
    from openpyxl.utils import get_column_letter
    NMAX = 11
    # Interview Feedback Not Completed — appended as the LAST section of this tab.
    IFV_DATA = ["Interview Start Date", "Interviewer Name", "Tech Stack",
                "Batch Name", "Batch Title / Duration"]
    IFV_BANNER = ("  INTERVIEW FEEDBACK NOT COMPLETED  —  Feedback Missing in the "
                  "Interview Consolidate Sheet")
    title = (f"IntelliBI  |  Batch Coordinator — Wise & Interview Feedback Validation  |  "
             f"{period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, NMAX, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    row = 2

    if not data:
        AR.write_section_banner(ws, row, NMAX,
                                "  Wise validation data is unavailable for this run.",
                                AR.C_GREY_BD, h_align="center")
        # Still show the Interview Feedback validation section (independent of Wise data).
        _wise_render_section(ws, row + 2, NMAX, IFV_BANNER, IFV_DATA,
                             interview_rows or [], {1, 2, 3, 4})
        ws.freeze_panes = "A2"
        AR.auto_col_width(ws)
        ws.column_dimensions[get_column_letter(NMAX)].width = 72
        return ws

    STU_DATA = ["#", "Student Name", "Batch Name", "Name", "Email", "Phone",
                "Tag", "Note", "Picture", "Joined On"]
    CRS_DATA = ["#", "Course Title", "Course Subtitle", "Instructor (Tag)",
                "Title", "Subtitle", "Tag", "Created On"]
    INS_DATA = ["#", "Instructor ID", "Instructor Name", "Failed Checks"]

    row = _wise_render_section(
        ws, row, NMAX, "  STUDENT VALIDATION  —  Failed / Attention-Required Records",
        STU_DATA, _wise_student_display(data.get("student", [])), {1, 2})
    row = _wise_render_section(
        ws, row, NMAX, "  COURSE VALIDATION  —  Failed / Attention-Required Records",
        CRS_DATA, _wise_course_display(data.get("course", [])), {1, 2, 3})
    row = _wise_render_section(
        ws, row, NMAX, "  INSTRUCTOR VALIDATION  —  Failed / Attention-Required Records",
        INS_DATA, _wise_instructor_display(data.get("instructor", [])), {2})
    # NEW — Interview Feedback Not Completed (last section, additive).
    row = _wise_render_section(
        ws, row, NMAX, IFV_BANNER, IFV_DATA, interview_rows or [], {1, 2, 3, 4})

    ws.freeze_panes = "A2"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(NMAX)].width = 72     # Why Flagged
    return ws


def build_student_detail_ext(ws, att_daily: pd.DataFrame, susp_daily: pd.DataFrame,
                             fb_daily: pd.DataFrame, agg_map: dict, report_date: date,
                             yest_date=None, period_label=None):
    """Extended daily Student Detail written into the provided worksheet `ws`.
    Mirrors pyAttendaceFeedbackReport's daily Student Detail (identity + daily
    columns, identical logic/styling) and appends the Overall / Till-Date
    Aggregate block."""
    title = (f"IntelliBI  |  Batch Coordinator — Learner Attendance Follow-Ups  |  "
             f"{period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, N, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")

    # ── row 2: section super-headers ─────────────────────────────────────────
    def _superhead(c_start, c_end, text, bg):
        ws.merge_cells(start_row=2, start_column=c_start, end_row=2, end_column=c_end)
        c = ws.cell(row=2, column=c_start)
        c.value = text
        c.font = AR._font(bold=True, size=10, color=AR.C_WHITE)
        c.fill = AR._fill(bg)
        c.alignment = AR._align("center", "center")
        c.border = AR._border()
    _superhead(1, 1, "Rank", AR.C_NAV2)
    _superhead(2, len(STUDENT_COLS), "Student", AR.C_NAV2)
    _superhead(len(STUDENT_COLS) + 1, len(STUDENT_COLS) + len(DAILY_COLS),
               f"Daily — {report_date.strftime('%d-%b-%Y')}", AR.C_BLUE_MID)
    _superhead(_AGG_START, N, "Overall / Till-Date Aggregate", AR.C_TEAL)
    ws.row_dimensions[2].height = 20

    # ── row 3: column headers ────────────────────────────────────────────────
    AR.write_header_row(ws, 3, HEADERS)

    row_num = 4

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
        """Write the 10 aggregate columns for a given (sid, cn, ct)."""
        a = agg_map.get(key)
        if not a:
            for col in range(_AGG_START, N + 1):
                AR.style_data_cell(ws, row_num, col, "—", bg=row_bg, h_align="center")
            return
        present, total = a["present"], a["total"]
        agg_att = a["attendance_pct"]
        agg_dur = a["agg_dur_pct"]
        # Sessions (P/T)
        AR.style_data_cell(ws, row_num, _C["Sessions (P/T)"], f"{present}/{total}",
                           bg=row_bg, bold=True, h_align="center")
        # Agg Attendance %
        c = AR.style_data_cell(ws, row_num, _C["Agg Attendance %"], agg_att,
                               bg=AR._att_bg(agg_att, ""), h_align="center",
                               number_fmt='0.0"%"')
        c.font = AR._font(bold=(agg_att < 75), size=10,
                          color=AR.C_RED_DARK if agg_att < 75 else AR.C_GREEN_DARK if agg_att >= 95 else "000000")
        # Agg Duration Att %
        c = AR.style_data_cell(ws, row_num, _C["Agg Duration Att %"], agg_dur,
                               bg=AR._att_bg(agg_dur, ""), h_align="center",
                               number_fmt='0.0"%"')
        c.font = AR._font(bold=(agg_dur < 75), size=10,
                          color=AR.C_RED_DARK if agg_dur < 75 else AR.C_GREEN_DARK if agg_dur >= 95 else "000000")
        # Last Present
        lp = a["last_present"].strftime("%d-%b-%Y") if a["last_present"] else "—"
        AR.style_data_cell(ws, row_num, _C["Last Present"], lp, bg=row_bg, h_align="center")
        # Absent Streak
        streak = a["streak"]
        if streak >= 3:
            s_bg, s_col = AR.C_RED_LITE, AR.C_RED_DARK
        elif streak >= 1:
            s_bg, s_col = AR.C_AMBER, AR.C_AMBER_DARK
        else:
            s_bg, s_col = AR.C_GREEN_PALE, AR.C_GREEN_DARK
        c = AR.style_data_cell(ws, row_num, _C["Absent Streak"], streak, bg=s_bg, h_align="center")
        c.font = AR._font(bold=(streak >= 1), size=10, color=s_col)
        # Feedbacks
        AR.style_data_cell(ws, row_num, _C["Feedbacks"], a["n_feedbacks"], bg=row_bg, h_align="center")
        # Feedback Part. %
        part = a["participation_pct"]
        if part >= 50:
            p_bg, p_col = AR.C_GREEN_PALE, AR.C_GREEN_DARK
        elif part > 0:
            p_bg, p_col = AR.C_AMBER, AR.C_AMBER_DARK
        else:
            p_bg, p_col = AR.C_RED_LITE, AR.C_RED_DARK
        c = AR.style_data_cell(ws, row_num, _C["Feedback Part. %"], part, bg=p_bg,
                               h_align="center", number_fmt='0.0"%"')
        c.font = AR._font(size=10, color=p_col)
        # Avg Rating (/10)
        if a["avg_rating"] is not None:
            c = AR.style_data_cell(ws, row_num, _C["Avg Rating (/10)"], a["avg_rating"],
                                   bg=AR._rating_bg(a["avg_rating"]), h_align="center",
                                   number_fmt='0.0')
        else:
            AR.style_data_cell(ws, row_num, _C["Avg Rating (/10)"], "—", bg=row_bg, h_align="center")
        # Overall Remarks (daily _remarks bands, on the aggregate duration %)
        orem = AR._remarks(agg_dur, "")
        c = AR.style_data_cell(ws, row_num, _C["Overall Remarks"], orem,
                               bg=AR._att_bg(agg_dur, ""), h_align="left")
        c.font = AR._font(size=10, color=AR.C_RED_DARK if agg_dur < 75
                          else AR.C_GREEN_DARK if agg_dur >= 95 else "000000")
        # Why Flagged
        why = a.get("why", "") or ""
        c = AR.style_data_cell(ws, row_num, _C["Why Flagged"], why or "—",
                               bg=(AR.C_AMBER_PALE if why else row_bg), h_align="left", wrap=True)
        c.font = AR._font(size=10, color=AR.C_AMBER_DARK if why else "808080")

    # global severity ranking (rank 1 = worst) across all follow-up rows
    rank_map = _compute_learner_ranks(att_daily, agg_map, fb_pairs, fb_ratings)

    # ── FOLLOW-UP rows only, grouped by (course_name, course_title) ──────────
    #   A student is kept ONLY when they qualify for at least one follow-up type
    #   (Attendance defaulter / Feedback defaulter / Low rating). Why Flagged
    #   combines every applicable reason so the coordinator can cover them all in
    #   one conversation. Not-applicable and suspended students are never
    #   follow-ups and are omitted from this action list.
    any_rendered = False
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

        banner_group = None
        alt_i = 0
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

            if (cn, ct) != banner_group:        # lazy technology banner
                banner_group = (cn, ct)
                alt_i = 0
                banner = f"  {cn}  —  {ct}" if ct else f"  {cn}"
                AR.write_section_banner(ws, row_num, N, banner, AR.C_BLUE_MID, h_align="center")
                row_num += 1

            status = str(r.get("status", ""))
            pct = r.get("_pct_num", 0.0)
            row_bg = AR._row_bg(pct, status, alt_i)
            if yest_date is not None and r.get("_date") == yest_date:
                row_bg = AR.C_YEST_HL
            alt_i += 1
            icon = AR._att_icon(pct, status)
            status_disp = "❌  Absent" if status == "Absent" else "✅  Present"
            att_disp = f"{icon}  Absent" if status == "Absent" else f"{icon}  {pct:.1f}%"

            for col, v in enumerate([r.get("course_name", ""), r.get("course_title", ""),
                                     r.get("student_name", ""), r.get("phone", "")], 2):
                AR.style_data_cell(ws, row_num, col, v, bg=row_bg, h_align="left")

            # Rank (1 = worst) — reflects combined A/B/C follow-up severity
            _rk = rank_map.get((sid, str(r.get("session_id", "")), str(cn), str(ct)))
            rkc = ws.cell(row=row_num, column=_C["Rank"])
            rkc.value = _rk if _rk else "—"
            if _rk and _rk <= 3:
                _rk_bg, _rk_col = AR.C_RED_LITE, AR.C_RED_DARK
            elif _rk and _rk <= 6:
                _rk_bg, _rk_col = AR.C_AMBER, AR.C_AMBER_DARK
            else:
                _rk_bg, _rk_col = row_bg, "000000"
            rkc.font = AR._font(bold=True, size=11, color=_rk_col)
            rkc.fill = AR._fill(_rk_bg)
            rkc.alignment = AR._align("center", "center"); rkc.border = AR._border()

            sc = ws.cell(row=row_num, column=_C["Status"])
            sc.value = status_disp
            sc.font = AR._font(bold=True, size=10,
                               color=AR.C_RED_DARK if status == "Absent" else AR.C_GREEN_DARK)
            sc.fill = AR._fill(AR.C_RED_LITE if status == "Absent" else AR.C_GREEN)
            sc.alignment = AR._align("center", "center"); sc.border = AR._border()

            ac = ws.cell(row=row_num, column=_C["Attendance %"])
            ac.value = att_disp
            ac.font = AR._font(bold=(status == "Absent" or pct < 75), size=10,
                               color=AR.C_RED_DARK if (status == "Absent" or pct < 75) else AR.C_GREEN_DARK)
            ac.fill = AR._fill(AR._att_bg(pct, status))
            ac.alignment = AR._align("center", "center"); ac.border = AR._border()

            AR.style_data_cell(ws, row_num, _C["Duration (min)"],
                               r.get("_dur_min", 0) if r.get("_dur_min", 0) > 0 else "",
                               bg=row_bg, h_align="center")
            AR.style_data_cell(ws, row_num, _C["Joined At"],
                               str(r.get("first_join_ist", ""))[:19], bg=row_bg, h_align="center")
            AR.style_data_cell(ws, row_num, _C["Left At"],
                               str(r.get("last_leave_ist", ""))[:19], bg=row_bg, h_align="center")

            rc = ws.cell(row=row_num, column=_C["Remarks"])
            rc.value = AR._remarks(pct, status)
            rc.font = AR._font(size=10, color=AR.C_RED_DARK if (status == "Absent" or pct < 75)
                               else AR.C_GREEN_DARK if pct >= 95 else "000000")
            rc.fill = AR._fill(row_bg); rc.alignment = AR._align("left", "center"); rc.border = AR._border()

            if status == "Absent":
                fb_val, fb_col, fb_bg = "Absent", AR.C_GREY_BD, AR.C_GREY_LITE
            elif (str(r.get("session_id", "")), sid) in fb_pairs:
                fb_val, fb_col, fb_bg = "✅  Yes", AR.C_GREEN_DARK, AR.C_GREEN
            else:
                fb_val, fb_col, fb_bg = "❌  No", AR.C_RED_DARK, AR.C_RED_LITE
            fbc = ws.cell(row=row_num, column=_C["Feedback Given?"])
            fbc.value = fb_val
            fbc.font = AR._font(bold=True, size=10, color=fb_col)
            fbc.fill = AR._fill(fb_bg); fbc.alignment = AR._align("center", "center"); fbc.border = AR._border()

            _rt_key = (str(r.get("session_id", "")), sid)
            if status != "Absent" and _rt_key in fb_ratings:
                _rt_val = fb_ratings[_rt_key]
                AR.style_data_cell(ws, row_num, _C["Feedback Rating (/10)"], _rt_val,
                                   bg=AR._rating_bg(_rt_val), h_align="center", number_fmt="0.#")
            else:
                AR.style_data_cell(ws, row_num, _C["Feedback Rating (/10)"], "—",
                                   bg=row_bg, h_align="center")

            # aggregate block, then OVERWRITE Why Flagged with the combined,
            # action-oriented reasons (key figures/actions highlighted red).
            _agg_cells(row_num, agg_key, row_bg)
            wc = ws.cell(row=row_num, column=_C["Why Flagged"])
            _rt = _why_flagged_richtext(reason_runs)
            wc.value = _rt
            wc.fill = AR._fill(AR.C_AMBER_PALE)
            wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
            wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()

            ws.row_dimensions[row_num].height = max(30, 15 * len(reason_runs) + 8)
            any_rendered = True
            row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, N,
                                "  ✅  No learner follow-ups required for this day.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1

    # ── freeze / filter / widths (same convention as the daily report) ───────
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A3:{get_column_letter(N)}{max(row_num - 1, 3)}"
    ws.freeze_panes = "A4"
    AR.auto_col_width(ws)
    # a touch wider for the wordy aggregate columns
    ws.column_dimensions[get_column_letter(_C["Why Flagged"])].width = 48
    ws.column_dimensions[get_column_letter(_C["Overall Remarks"])].width = 16
    ws.column_dimensions[get_column_letter(_C["Rank"])].width = 7
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


def upload_report(folder_name: str, filename: str, buf: io.BytesIO, base_prefix: str) -> str:
    """Create <parent>/<folder_name>/ and drop the workbook there as a native
    Google Sheet (converted on upload). Existing reports for this report/date are
    NEVER overwritten or modified — the Coordinator may have added manual follow-up
    comments to them. If a report with this base_prefix already exists in the
    folder, the new run is saved as the next version instead:
        first run      -> the plain file name
        already exists -> "<name> - Version 2"
        Version 2 too  -> "<name> - Version 3"  (increments dynamically)
    Every previous version is kept unchanged. base_prefix scopes this per report
    type, so each type versions independently and other reports are untouched."""
    import re
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    from googleapiclient.http import MediaIoBaseUpload

    creds = service_account.Credentials.from_service_account_file(
        AR.SERVICE_ACCOUNT_FILE, scopes=["https://www.googleapis.com/auth/drive"]
    ).with_subject(IMPERSONATE_USER)
    drive = gbuild("drive", "v3", credentials=creds, cache_discovery=False)

    folder_id = _find_or_create_folder(drive, PARENT_FOLDER_ID, folder_name)
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
    return link


def upload_datewise(report_date: date, filename: str, buf: io.BytesIO) -> str:
    """Daily report → <parent>/YYYY-MM-DD/ (thin wrapper over upload_report)."""
    return upload_report(report_date.strftime("%Y-%m-%d"), filename, buf, REPORT_BASENAME)


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
    AR.style_title_row(ws, 1, 1, LPN,
                       f"IntelliBI  |  Learner Attendance Follow-Ups — {label}")
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, LP_COLS)
    row_num = 3

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
    for p in ordered:
        rk = p["_rank"]
        agg = p["agg"]
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
                AR.style_data_cell(ws, row_num, _LP[c], vals[c], bg=AR.C_WHITE,
                                   h_align="left" if c in ("Tech Name", "Duration", "Student Name")
                                   else "center")
        rkc = ws.cell(row=row_num, column=_LP["Rank"])
        _rk_bg, _rk_col = ((AR.C_RED_LITE, AR.C_RED_DARK) if rk <= 3
                           else (AR.C_AMBER, AR.C_AMBER_DARK) if rk <= 6 else (AR.C_WHITE, "000000"))
        rkc.font = AR._font(bold=True, size=11, color=_rk_col)
        rkc.fill = AR._fill(_rk_bg); rkc.alignment = AR._align("center", "center"); rkc.border = AR._border()
        AR.style_data_cell(ws, row_num, _LP["Agg Attendance %"], agg["attendance_pct"],
                           bg=AR._att_bg(agg["attendance_pct"], ""), h_align="center", number_fmt='0.0"%"')
        AR.style_data_cell(ws, row_num, _LP["Agg Duration Att %"], p["agg_dur"],
                           bg=AR._att_bg(p["agg_dur"], ""), h_align="center", number_fmt='0.0"%"')
        AR.style_data_cell(ws, row_num, _LP["Feedback Part. %"], p["part"],
                           bg=(AR.C_GREEN_PALE if p["part"] >= 50 else AR.C_AMBER if p["part"] > 0 else AR.C_RED_LITE),
                           h_align="center", number_fmt='0.0"%"')
        if agg.get("avg_rating") is not None:
            AR.style_data_cell(ws, row_num, _LP["Avg Rating (/10)"], agg["avg_rating"],
                               bg=AR._rating_bg(agg["avg_rating"]), h_align="center", number_fmt='0.0')
        else:
            AR.style_data_cell(ws, row_num, _LP["Avg Rating (/10)"], "—", bg=AR.C_WHITE, h_align="center")
        wc = ws.cell(row=row_num, column=_LP["Why Flagged"])
        _rt = _why_flagged_richtext(p["reason_runs"])
        wc.value = _rt
        wc.fill = AR._fill(AR.C_AMBER_PALE)
        wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(_rt, str) else "333333")
        wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()
        ws.row_dimensions[row_num].height = max(30, 15 * len(p["reason_runs"]) + 8)
        any_rendered = True
        row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, LPN,
                                "  ✅  No learner follow-ups in this period.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(LPN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_LP["Why Flagged"])].width = 50
    ws.column_dimensions[get_column_letter(_LP["Rank"])].width = 7
    return ws


def build_instructor_followups_period(ws, sess_period, att_period, fb_period, tf_period,
                                      instr_phones, label):
    """One row per (instructor, tech) over the period: session counts, missing
    feedback, avg attendance/rating, flagged-session count, avg schedule diff."""
    AR.style_title_row(ws, 1, 1, IPN,
                       f"IntelliBI  |  Instructor Follow-Ups — {label}")
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, IP_COLS)
    row_num = 3

    if sess_period is None or sess_period.empty:
        AR.write_section_banner(ws, row_num, IPN, "  No sessions in this period.",
                                AR.C_GREY_BD, h_align="center")
        from openpyxl.utils import get_column_letter
        ws.freeze_panes = "A3"; AR.auto_col_width(ws)
        return ws

    sess_f = AR._prefer_instructor_name(sess_period)
    tf_ids = (set(tf_period["session_id"].dropna().unique())
              if (tf_period is not None and not tf_period.empty and "session_id" in tf_period.columns)
              else set())

    any_rendered = False
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
        for c in IP_COLS:
            if c in vals:
                AR.style_data_cell(ws, row_num, _IP[c], vals[c], bg=AR.C_WHITE,
                                   h_align="left" if c in ("Tech Name", "Duration", "Instructor")
                                   else "center")
        AR.style_data_cell(ws, row_num, _IP["Feedback Given %"], given_pct,
                           bg=(AR.C_GREEN_PALE if given_pct >= 80 else AR.C_AMBER if given_pct > 0 else AR.C_RED_LITE),
                           h_align="center", number_fmt='0.0"%"')
        AR.style_data_cell(ws, row_num, _IP["Avg Att %"], avg_att, bg=AR._att_bg(avg_att, ""),
                           h_align="center", number_fmt='0.0"%"')
        if avg_rat is not None:
            AR.style_data_cell(ws, row_num, _IP["Avg Rating (/10)"], avg_rat,
                               bg=AR._rating_bg(avg_rat), h_align="center", number_fmt='0.0')
        else:
            AR.style_data_cell(ws, row_num, _IP["Avg Rating (/10)"], "—", bg=AR.C_WHITE, h_align="center")
        wc = ws.cell(row=row_num, column=_IP["Why Flagged"])
        wc.value = "\n".join(f"• {t}" for t in reasons)
        wc.fill = AR._fill(AR.C_AMBER_PALE); wc.font = AR._font(size=10, color=AR.C_RED_DARK)
        wc.alignment = AR._align("left", "center", wrap=True); wc.border = AR._border()
        ws.row_dimensions[row_num].height = 30
        any_rendered = True
        row_num += 1

    if not any_rendered:
        AR.write_section_banner(ws, row_num, IPN,
                                "  ✅  No instructor follow-ups in this period.",
                                AR.C_GREEN_DARK, h_align="center")
        row_num += 1
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(IPN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_IP["Why Flagged"])].width = 50
    ws.column_dimensions[get_column_letter(_IP["Instructor"])].width = 20
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
    # Save Weekly / Monthly / Manual roll-ups INSIDE the run's date-wise folder
    # (the report-generation date, same folder the Daily report uses), as separate
    # files — no separate Weekly/Monthly folder. File name is unchanged; the
    # per-report base_prefix keeps each report type replacing only its own file.
    folder = gen_date.strftime("%Y-%m-%d")
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
           "Candidate Name", "Phone", "Why Flagged"]
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


def _iv_write_row(ws, row_num, values, why_runs, bg, why_bg, bold):
    for col_name in IV_COLS:
        cidx = _IVC[col_name]
        if col_name == "Why Flagged":
            wc = ws.cell(row=row_num, column=cidx)
            rt = _why_flagged_richtext(why_runs)
            wc.value = rt
            wc.fill = AR._fill(why_bg)
            wc.font = AR._font(size=10, color=AR.C_RED_DARK if isinstance(rt, str) else "333333")
            wc.alignment = AR._align("left", "center", wrap=True)
            wc.border = AR._border()
        else:
            h_align = "center" if col_name == "Phone" else "left"
            AR.style_data_cell(ws, row_num, cidx, values.get(col_name, ""),
                               bg=bg, bold=bold, h_align=h_align, wrap=True)
    ws.row_dimensions[row_num].height = 46


def _iv_finish(ws, row_num):
    from openpyxl.utils import get_column_letter
    ws.auto_filter.ref = f"A2:{get_column_letter(IVN)}{max(row_num - 1, 2)}"
    ws.freeze_panes = "A3"
    AR.auto_col_width(ws)
    ws.column_dimensions[get_column_letter(_IVC["Why Flagged"])].width = 66
    ws.column_dimensions[get_column_letter(_IVC["Interview Time"])].width = 26
    ws.column_dimensions[get_column_letter(_IVC["Batch Title / Duration"])].width = 24
    ws.column_dimensions[get_column_letter(_IVC["Candidate Name"])].width = 24


def build_interview_reminders(ws, sheets, drive, service, att_agg,
                              report_date, period_label=None):
    """Learner Instructor Interview Reminder tab. Interviewer row first (distinct
    background), then learner rows sorted by interview time; a coloured banner
    separates each batch/interview. The interviewer comes from the schedule file's
    Interview_Helper tab (never the batch Instructor). Additive only."""
    title = (f"IntelliBI  |  Batch Coordinator — Learner & Instructor Interview "
             f"Reminders  |  {period_label or report_date.strftime('%d-%b-%Y')}")
    AR.style_title_row(ws, 1, 1, IVN, title)
    ws.cell(row=1, column=1).alignment = AR._align("center", "center")
    AR.write_header_row(ws, 2, IV_COLS)
    row_num = 3

    try:
        groups = load_interview_reminders(sheets, drive, report_date)
    except Exception as e:                              # never break the report
        log.exception("Interview reminder tab failed to load: %s", e)
        groups = []

    if not groups:
        AR.write_section_banner(ws, row_num, IVN,
                                "  No interview follow-ups due today.",
                                AR.C_GREEN_DARK, h_align="center")
        _iv_finish(ws, row_num + 1)
        return ws

    iv_phone_idx = load_interviewer_phone_index(service)   # Instructor tab -> phone
    ph_full, ph_name = _iv_phone_index(att_agg)

    for g in groups:
        bname  = g["batch_name"] or "—"
        btitle = g["batch_title"] or "—"
        idate  = g["interview_date"]
        action = g["action"]
        date_label = idate.strftime("%d-%b-%Y")
        action_txt = "MESSAGE (2 days prior)" if action == "message" else "CALL (1 day prior)"
        # ── batch/interview separator banner ─────────────────────────────────
        AR.write_section_banner(
            ws, row_num, IVN,
            f"  Interview — {bname}   ·   {btitle}   ·   {date_label}"
            f"      Today: {action_txt}",
            AR.C_TEAL, h_align="left")
        ws.row_dimensions[row_num].height = 24
        row_num += 1

        # ── Interviewer reminder — FIRST record, distinct blue background ─────
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
           bg=AR.C_BLUE_LITE, why_bg=AR.C_BLUE_PALE, bold=True)
        row_num += 1

        # ── Learner reminders — sorted by interview time ascending ───────────
        for i, c in enumerate(g["candidates"]):
            phone = _iv_resolve_phone(ph_full, ph_name, c["name"], bname)
            alt = AR.C_WHITE if i % 2 == 0 else AR.C_ROW_ALT
            _iv_write_row(ws, row_num, {
                "Interview Time": c["time_disp"],
                "Batch Name (Class)": bname,
                "Batch Title / Duration": btitle,
                "Candidate Name": c["name"],
                "Phone": phone or "—",
            }, _iv_learner_runs(action, c["name"], c["start_label"], date_label),
               bg=alt, why_bg=AR.C_AMBER_PALE, bold=False)
            row_num += 1

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
    return results


def main():
    logging.basicConfig(level=logging.INFO if VERBOSE else logging.WARNING,
                        format="%(asctime)s [%(levelname)s] %(message)s")
    return generate()


if __name__ == "__main__":
    main()
