"""
================================================================================
  IntelliBI Operations Automation — Daily Batch Coordinator
  ATTENDANCE Task & Performance Monitoring Report
  (ops_reports_action / pyCoordinatorAttendanceTaskReport.py)
  ------------------------------------------------------------------------------
  PURPOSE
    Produce, every day, a *technology-wise action list* that tells the Batch
    Coordinator WHICH students genuinely need follow-up (not merely who was
    absent today) — fusing today's/latest performance with till-date history
    (attendance, feedback participation, number of feedbacks, ratings, recent
    trend, absence streaks) into one explainable "Attention Score". The same
    artifact is the evidence base Management uses to see whether the required
    follow-ups were completed, and the future source for the coordinator's
    daily / weekly / monthly KRA.

    Daily Monitoring -> Intelligent Follow-Up Identification -> Coordinator
    Action -> Follow-Up Tracking -> Management Performance Monitoring.

  WHAT IT PRODUCES
    A day-wise sub-folder under a configured Google Drive folder, containing a
    single NATIVE Google Sheet with four tabs:
        Summary          — management KPIs (per technology + overall), with LIVE
                           completion formulas that update as the coordinator
                           fills the action list.
        Follow-Up Tasks  — the flagged students only: protected read-only
                           diagnostics + editable coordinator entry columns
                           (dropdowns, conditional formatting). Clean, filterable,
                           sorted Technology -> Priority.
        Master Data      — every applicable student with full metrics + score +
                           band (auditable KRA / analytics feed).
        Guide            — legend, scoring explanation, how-to.

    Previous days are never overwritten (each day = its own folder + file).
    Re-running on the SAME day preserves coordinator entries already typed in.

  DATA SOURCES (produced by Layer 1, read-only here)
    IntellBIAttendance   (1TqDjq4gAyo32eRNMbuLd6uu0eCNZb7h1j5YH-q68AhU)
        Sessions, Attendance, Student_Feedback
    IntelliBIStudentInfo (1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA)
        Students   (phone, Batch_Timing, is_attendance_required, Is_Deleted)

  DESIGN NOTES
    * Approach is a transparent, weighted, RULE-BASED score (no ML) — reliable,
      explainable and operationally useful, exactly as the brief asks.
    * Every weight / threshold / band cut-off / option list is config-driven via
      config/coordinator_attendance_config.json (optional — the script self-
      defaults if the file is absent), so tuning needs no code change.
    * Attendance maths mirror the validated pyAttendaceFeedbackReport.py
      (Present / distinct sessions; suspended and is_attendance_required != Y
      excluded) so this report never conflicts with the existing calculation.
    * Nothing is hardcoded per student / technology / date / session.

  RUN
    Configure the small block under "RUN CONFIGURATION" (or rely on defaults),
    then:  python ops_reports_action/pyCoordinatorAttendanceTaskReport.py
    Self-test (no Google calls):
           python ops_reports_action/pyCoordinatorAttendanceTaskReport.py --selftest

  Layer-1 dependencies (when wired into scripts/run_reports_action.py):
    pySessionAttendanceStudentTeacherFeedbacks, pyStudentPaymentClassesStudentEnrolled
================================================================================
"""
from __future__ import annotations

# ── IntelliBI Operations Automation portability bootstrap ────────────────────
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_COMMON = os.path.join(os.path.dirname(_HERE), "common")
if os.path.isdir(_COMMON) and _COMMON not in sys.path:
    sys.path.insert(0, _COMMON)

try:                                            # project bootstrap (preferred)
    import _bootstrap  # noqa: F401  (sys.path + env defaults + config.yaml)
    from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR
except Exception:                               # standalone / test fallback
    CREDENTIALS_DIR = os.path.join(os.path.dirname(_HERE), "credentials")
    CONFIG_DIR = os.path.join(os.path.dirname(_HERE), "config")
    LOGS_DIR = os.path.join(os.path.dirname(_HERE), "logs")
# ── end bootstrap ────────────────────────────────────────────────────────────

import json
import logging
import statistics
from datetime import datetime, timezone, timedelta

IST = timezone(timedelta(hours=5, minutes=30))

# =============================================================================
#  RUN CONFIGURATION  (operator-facing; safe defaults)
# =============================================================================
SEND_EMAIL = False          # email the day's Sheet link (default OFF — nothing sends unexpectedly)
REPORT_DATE = None          # "YYYY-MM-DD" to pin the report day; None = derive from latest data
VERBOSE = True

# =============================================================================
#  DEFAULT CONFIG  (overlaid by config/coordinator_attendance_config.json)
# =============================================================================
DEFAULT_CONFIG = {
    "drive": {
        "parent_folder_id": "1BEokUc7Np7iBVSMwIyMAgZUa0mrecT-h",
        "impersonate_user": "info@intellibiinnovationstechnologies.in",
        "day_folder_format": "%Y-%m-%d",
        "sheet_title_prefix": "Coordinator Attendance Tasks",
        "share_with": [],                 # optional extra emails to grant writer access
    },
    "sources": {
        "attendance_sheet_id": "1TqDjq4gAyo32eRNMbuLd6uu0eCNZb7h1j5YH-q68AhU",
        "students_sheet_id":   "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA",
        "sessions_tab":        "Sessions",
        "attendance_tab":      "Attendance",
        "student_feedback_tab": "Student_Feedback",
        "students_tab":        "Students",
    },
    "scoring": {
        "attendance_target_pct": 75,      # "healthy" attendance line
        "attendance_critical_pct": 60,    # below this = critical
        "trend_window_sessions": 5,       # recent-momentum window
        "streak_flag": 3,                 # consecutive absences that force attention
        "rating_good": 7.0,               # rating (/10) considered healthy
        "rating_low": 5.0,                # rating (/10) considered a concern
        "min_present_for_feedback": 2,    # need >= this many attended sessions before
                                          #   we hold "no feedback" against a student
        "weights": {                      # sum = 100  ->  score is already 0..100
            "overall_attendance":     30,
            "latest_absent":          15,
            "absent_streak":          20,
            "recent_trend":           10,
            "feedback_participation": 12,
            "feedback_rating":         8,
            "feedback_volume":         5,
        },
        "band_cutoffs": {"critical": 65, "high": 45, "medium": 30},
    },
    "options": {
        "action_taken": ["Call", "WhatsApp", "Email", "In-person",
                         "Parent contacted", "No response", "Not required", "Other"],
        "outcome": ["Completed", "In Progress", "Pending", "No Response",
                    "Escalated", "Not Required"],
        "follow_up_done": ["Yes", "No"],
    },
    "email": {
        "recipients": ["info@intellibiinnovationstechnologies.in"],
    },
}

# --- tab names (plain — kept formula-safe) -----------------------------------
SUMMARY_TAB = "Summary"
TASKS_TAB = "Follow-Up Tasks"
MASTER_TAB = "Master Data"
GUIDE_TAB = "Guide"
TAB_ORDER = [SUMMARY_TAB, TASKS_TAB, MASTER_TAB, GUIDE_TAB]

# --- diagnostic (read-only) columns, shared by Tasks + Master ----------------
DIAG_COLS = [
    "Priority", "Attention Score", "Technology", "Batch Timing", "Student Name",
    "Phone", "Email", "Student ID", "Latest Session", "Latest Status",
    "Attendance %", "Attendance Band", "Absent Streak", "Sessions (P/T)",
    "Last Present", "Recent Trend %", "Feedbacks", "Feedback Part. %",
    "Avg Rating (/10)", "Why Flagged", "Recommended Action",
]
# --- coordinator entry (editable) columns, Tasks tab only --------------------
ENTRY_COLS = [
    "Follow-Up Done?", "Action Taken", "Follow-Up Comment", "Outcome",
    "Coordinator", "Follow-Up Date", "Next Action",
]
TASKS_COLS = DIAG_COLS + ENTRY_COLS
MASTER_COLS = DIAG_COLS + ["Requires Follow-Up"]

_ID_COL_IDX = DIAG_COLS.index("Student ID")          # hidden helper column
_PCT_COL_IDX = DIAG_COLS.index("Attendance %")
_TREND_COL_IDX = DIAG_COLS.index("Recent Trend %")
_PART_COL_IDX = DIAG_COLS.index("Feedback Part. %")
_RATING_COL_IDX = DIAG_COLS.index("Avg Rating (/10)")
_SCORE_COL_IDX = DIAG_COLS.index("Attention Score")
_PRIORITY_COL_IDX = DIAG_COLS.index("Priority")
_TECH_COL_IDX = DIAG_COLS.index("Technology")

# colours (Sheets API rgb 0..1)
C_NAVY = {"red": 0.106, "green": 0.208, "blue": 0.369}
C_WHITE = {"red": 1, "green": 1, "blue": 1}
C_RED = {"red": 0.98, "green": 0.80, "blue": 0.80}
C_RED_TXT = {"red": 0.69, "green": 0.11, "blue": 0.11}
C_AMBER = {"red": 1.0, "green": 0.90, "blue": 0.70}
C_AMBER_TXT = {"red": 0.54, "green": 0.43, "blue": 0.0}
C_GREEN = {"red": 0.78, "green": 0.90, "blue": 0.79}
C_GREEN_TXT = {"red": 0.10, "green": 0.48, "blue": 0.20}
C_ORANGE = {"red": 0.99, "green": 0.75, "blue": 0.55}
C_YELLOW = {"red": 1.0, "green": 0.95, "blue": 0.60}
C_HEADER_TXT = {"red": 1, "green": 1, "blue": 1}

log = logging.getLogger("CoordinatorAttendanceTask")


# =============================================================================
#  CONFIG LOADING
# =============================================================================
def _deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config() -> dict:
    """DEFAULT_CONFIG overlaid by config/coordinator_attendance_config.json (if any)."""
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))   # deep copy
    path = os.path.join(CONFIG_DIR, "coordinator_attendance_config.json")
    try:
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as fh:
                cfg = _deep_merge(cfg, json.load(fh) or {})
            log.info("Loaded overrides from %s", path)
    except Exception as e:
        log.warning("Could not read %s (%s) — using defaults.", path, e)
    return cfg


# =============================================================================
#  SMALL PARSERS
# =============================================================================
def _s(v) -> str:
    return "" if v is None else str(v).strip()


def _pdate(v):
    """'YYYY-MM-DD HH:MM:SS' (or ISO / date) -> datetime.date, else None."""
    s = _s(v)
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:19] if len(s) >= 19 and fmt == "%Y-%m-%d %H:%M:%S" else s[:10], fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00")).date()
    except Exception:
        return None


def _pfloat(v):
    s = _s(v).replace("%", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _rows_to_dicts(values):
    if not values:
        return []
    header = [_s(h) for h in values[0]]
    out = []
    for r in values[1:]:
        out.append({header[i]: (r[i] if i < len(r) else "") for i in range(len(header))})
    return out


# =============================================================================
#  GOOGLE SERVICES  (impersonated: sheets + drive)
# =============================================================================
def build_services(cfg: dict):
    """Return (sheets_read, sheets_write, drive) using exactly the credential
    pattern the rest of this project already relies on:

      * sheets_read  — the PLAIN service account (no impersonation), spreadsheets
                       scope. It is already shared on the SOURCE data sheets, so
                       it reads them directly (same as get_sheets_service()).
      * drive + sheets_write — the service account impersonating info@ with the
                       DRIVE scope ONLY. That is the single scope the domain-wide
                       delegation is authorized for (what the existing Drive
                       uploads use). The Google Sheets API accepts the drive
                       scope, so this SAME credential both creates the day's
                       native Sheet (owned by info@) and writes its values +
                       formatting. We never request the spreadsheets scope under
                       impersonation (that is not in the delegation allow-list and
                       causes 'unauthorized_client')."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    sa_file = os.path.join(CREDENTIALS_DIR, "service_account.json")

    read_creds = service_account.Credentials.from_service_account_file(
        sa_file, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sheets_read = build("sheets", "v4", credentials=read_creds, cache_discovery=False)

    imp_creds = service_account.Credentials.from_service_account_file(
        sa_file, scopes=["https://www.googleapis.com/auth/drive"]
    ).with_subject(cfg["drive"]["impersonate_user"])
    drive = build("drive", "v3", credentials=imp_creds, cache_discovery=False)
    sheets_write = build("sheets", "v4", credentials=imp_creds, cache_discovery=False)

    return sheets_read, sheets_write, drive


def _retry(fn, label="", tries=5):
    from googleapiclient.errors import HttpError
    import time
    delay = 2
    for attempt in range(1, tries + 1):
        try:
            return fn()
        except HttpError as e:
            status = getattr(e, "resp", None) and e.resp.status
            if status in (429, 500, 503) and attempt < tries:
                log.warning("  [retry] %s — %s, waiting %ss (%d/%d)", label, status, delay, attempt, tries)
                time.sleep(delay)
                delay *= 2
            else:
                raise


def read_tab(sheets, sheet_id, tab):
    res = _retry(
        lambda: sheets.spreadsheets().values()
        .get(spreadsheetId=sheet_id, range=f"'{tab}'!A:ZZ").execute(),
        label=f"read {tab}",
    )
    return _rows_to_dicts(res.get("values", []))


# =============================================================================
#  DATA MODEL  —  build one record per (student, technology)
# =============================================================================
def build_records(att_rows, fb_rows, students, cfg):
    """Return (records, meta) where records is a list of per (student,technology)
    dicts with attendance + feedback metrics; meta carries run-level info."""
    # ── student master maps ───────────────────────────────────────────────
    attn_req, phone_m, timing_m, name_m, email_m, deleted_m = {}, {}, {}, {}, {}, {}
    for s in students:
        sid = _s(s.get("student_id"))
        if not sid:
            continue
        attn_req[sid] = _s(s.get("is_attendance_required")).upper()
        phone_m[sid] = _s(s.get("phone"))
        timing_m[sid] = _s(s.get("Batch_Timing"))
        name_m[sid] = _s(s.get("student_name"))
        email_m[sid] = _s(s.get("email"))
        deleted_m[sid] = _s(s.get("Is_Deleted")).upper()

    # ── group attendance rows by (student, technology) ────────────────────
    groups = {}       # (sid, tech) -> {session_id: (date, status, cum_pct)}
    excluded_suspended = excluded_not_applicable = excluded_deleted = 0
    all_dates = []
    for r in att_rows:
        if _s(r.get("suspend_status")).lower() == "suspended":
            excluded_suspended += 1
            continue
        sid = _s(r.get("student_id"))
        tech = _s(r.get("course_name"))
        if not sid or not tech:
            continue
        if deleted_m.get(sid) == "Y":
            excluded_deleted += 1
            continue
        # mirror the validated attendance report: only is_attendance_required == Y counts
        if attn_req.get(sid, "") != "Y":
            excluded_not_applicable += 1
            continue
        d = _pdate(r.get("session_start_ist"))
        if d:
            all_dates.append(d)
        sess = _s(r.get("session_id")) or f"__{d}"
        status = "Present" if _s(r.get("status")).lower() == "present" else "Absent"
        cum = _pfloat(r.get("attendance_percent"))
        groups.setdefault((sid, tech), {})[sess] = (d, status, cum)

    # ── group feedback rows by (student, technology) ──────────────────────
    fb_groups = {}    # (sid, tech) -> {session_id: rating|None}
    for r in fb_rows:
        sid = _s(r.get("student_id"))
        tech = _s(r.get("course_name"))
        if not sid or not tech:
            continue
        sess = _s(r.get("session_id")) or f"fb_{len(fb_groups)}"
        fb_groups.setdefault((sid, tech), {})[sess] = _pfloat(r.get("rating"))

    sc = cfg["scoring"]
    win = int(sc["trend_window_sessions"])
    records = []
    for (sid, tech), sess_map in groups.items():
        # order this student's sessions in this technology by date (undated last)
        items = sorted(sess_map.values(), key=lambda t: (t[0] is None, t[0] or datetime.min.date()))
        total = len(items)
        present = sum(1 for _, st, _ in items if st == "Present")
        pct = round(present / total * 100, 1) if total else 0.0
        latest_date, latest_status = (items[-1][0], items[-1][1]) if items else (None, "")
        cum_latest = next((c for _, _, c in reversed(items) if c is not None), None)

        # consecutive absence streak from the most recent session backwards
        streak = 0
        for _, st, _ in reversed(items):
            if st == "Absent":
                streak += 1
            else:
                break
        # recent-window trend
        window = items[-win:] if total >= 1 else []
        w_present = sum(1 for _, st, _ in window if st == "Present")
        trend_pct = round(w_present / len(window) * 100, 1) if window else pct
        last_present = max((d for d, st, _ in items if st == "Present" and d), default=None)

        # feedback
        fb_map = fb_groups.get((sid, tech), {})
        n_fb = len(fb_map)
        ratings = [rt for rt in fb_map.values() if rt is not None]
        avg_rating = round(statistics.mean(ratings), 1) if ratings else None
        participation = round(min(n_fb / present * 100, 100), 1) if present else 0.0

        records.append({
            "student_id": sid,
            "technology": tech,
            "batch_timing": timing_m.get(sid, ""),
            "student_name": name_m.get(sid) or "",
            "phone": phone_m.get(sid, ""),
            "email": email_m.get(sid, ""),
            "latest_date": latest_date,
            "latest_status": latest_status,
            "attendance_pct": pct,
            "cum_pct": cum_latest,
            "streak": streak,
            "present": present,
            "total": total,
            "last_present": last_present,
            "trend_pct": trend_pct,
            "n_feedbacks": n_fb,
            "participation_pct": participation,
            "avg_rating": avg_rating,
            "n_ratings": len(ratings),
        })

    meta = {
        "report_date": (max(all_dates) if all_dates else datetime.now(IST).date()),
        "excluded_suspended": excluded_suspended,
        "excluded_not_applicable": excluded_not_applicable,
        "excluded_deleted": excluded_deleted,
        "n_records": len(records),
    }
    return records, meta


# =============================================================================
#  SCORING ENGINE  (transparent weighted rules)
# =============================================================================
def _shortfall(pct, target, critical):
    if pct is None:
        return 0.0
    if pct >= target:
        return 0.0
    if pct <= critical:
        return 1.0
    return round((target - pct) / (target - critical), 4)


def score_record(rec, cfg):
    """Return the record enriched with score, band, requires_followup, reasons,
    recommended_action, attendance_band. Pure & deterministic."""
    sc = cfg["scoring"]
    w = sc["weights"]
    target = sc["attendance_target_pct"]
    crit = sc["attendance_critical_pct"]
    streak_flag = int(sc["streak_flag"])
    rating_good = float(sc["rating_good"])
    rating_low = float(sc["rating_low"])
    min_present = int(sc["min_present_for_feedback"])

    pct = rec["attendance_pct"]
    streak = rec["streak"]
    present = rec["present"]
    latest_absent = rec["latest_status"] == "Absent"

    # ── risk components (0..1) ────────────────────────────────────────────
    r_overall = _shortfall(pct, target, crit)
    r_latest = 1.0 if latest_absent else 0.0
    r_streak = min(streak / streak_flag, 1.0) if streak_flag else 0.0
    r_trend = _shortfall(rec["trend_pct"], target, crit)
    if present >= min_present:
        r_part = round(1.0 - min(rec["participation_pct"], 100) / 100.0, 4)
        r_vol = 1.0 if rec["n_feedbacks"] == 0 else 0.0
    else:
        r_part = r_vol = 0.0
    if rec["n_ratings"] >= 1 and rec["avg_rating"] is not None:
        ar = rec["avg_rating"]
        if ar >= rating_good:
            r_rating = 0.0
        elif ar <= rating_low:
            r_rating = 1.0
        else:
            r_rating = round((rating_good - ar) / (rating_good - rating_low), 4)
    else:
        r_rating = 0.0

    score = (w["overall_attendance"] * r_overall
             + w["latest_absent"] * r_latest
             + w["absent_streak"] * r_streak
             + w["recent_trend"] * r_trend
             + w["feedback_participation"] * r_part
             + w["feedback_rating"] * r_rating
             + w["feedback_volume"] * r_vol)
    score = int(round(score))

    # ── band from score, then elevate on hard triggers ───────────────────
    bc = sc["band_cutoffs"]
    order = {"OK": 0, "Medium": 1, "High": 2, "Critical": 3}
    band = "OK"
    if score >= bc["medium"]:
        band = "Medium"
    if score >= bc["high"]:
        band = "High"
    if score >= bc["critical"]:
        band = "Critical"

    def elevate(current, floor):
        return floor if order[floor] > order[current] else current

    if pct is not None and pct < crit:
        band = elevate(band, "Critical")
    if streak >= streak_flag + 1:
        band = elevate(band, "Critical")
    if streak >= streak_flag:
        band = elevate(band, "High")
    if latest_absent and pct is not None and pct < target:
        band = elevate(band, "High")
    if (not latest_absent) and pct is not None and pct < target:
        band = elevate(band, "Medium")     # classic "at-risk": present but low history
    if present >= min_present and rec["n_feedbacks"] == 0:
        band = elevate(band, "Medium")
    if rec["n_ratings"] >= 1 and rec["avg_rating"] is not None and rec["avg_rating"] <= rating_low:
        band = elevate(band, "Medium")

    requires = band != "OK"

    # ── attendance band label (display) ──────────────────────────────────
    if pct is None:
        att_band = "—"
    elif rec["latest_status"] == "Absent" and pct >= target:
        att_band = "Absent (latest)"
    elif pct >= 95:
        att_band = "Excellent"
    elif pct >= 85:
        att_band = "Good"
    elif pct >= target:
        att_band = "Satisfactory"
    elif pct >= crit:
        att_band = "Low"
    else:
        att_band = "Critical"

    # ── dynamic reasons ──────────────────────────────────────────────────
    reasons = []
    if latest_absent:
        ld = rec["latest_date"].strftime("%d-%b") if rec["latest_date"] else "latest session"
        reasons.append(f"Absent in latest session ({ld})")
    if streak >= 2:
        reasons.append(f"{streak} consecutive absences")
    if pct is not None and pct < crit:
        reasons.append(f"Critical attendance {pct:g}% (below {crit}%)")
    elif pct is not None and pct < target:
        reasons.append(f"Low attendance {pct:g}% (below {target}%)")
    if rec["total"] >= 2 and rec["trend_pct"] < target and rec["trend_pct"] < pct:
        reasons.append(f"Declining trend — last {len(range(min(rec['total'], int(sc['trend_window_sessions']))))} sessions {rec['trend_pct']:g}%")
    if present >= min_present and rec["n_feedbacks"] == 0:
        reasons.append(f"No feedback despite {present} sessions attended")
    elif present >= min_present and rec["participation_pct"] < 50:
        reasons.append(f"Low feedback participation {rec['participation_pct']:g}% ({rec['n_feedbacks']}/{present})")
    if rec["n_ratings"] >= 1 and rec["avg_rating"] is not None and rec["avg_rating"] <= rating_low:
        reasons.append(f"Low avg rating {rec['avg_rating']:g}/10")
    why = "; ".join(reasons)

    # ── recommended action (priority-ordered, dynamic) ───────────────────
    if pct is not None and pct < crit:
        action = "Immediate counselling call — attendance critical"
    elif streak >= streak_flag:
        action = f"Call today — {streak} sessions missed in a row"
    elif latest_absent and pct is not None and pct < target:
        action = "Call/WhatsApp — absent again & history below target"
    elif latest_absent:
        action = "Quick check-in — reason for latest absence"
    elif pct is not None and pct < target:
        action = f"Counsel on attendance — trending below {target}%"
    elif present >= min_present and rec["n_feedbacks"] == 0:
        action = "Remind to submit session feedback"
    elif rec["n_ratings"] >= 1 and rec["avg_rating"] is not None and rec["avg_rating"] <= rating_low:
        action = "Discuss course experience — low satisfaction signal"
    elif present >= min_present and rec["participation_pct"] < 50:
        action = "Nudge for regular feedback participation"
    else:
        action = "Monitor"

    rec.update({
        "score": score, "band": band, "requires_followup": requires,
        "att_band": att_band, "why": why, "action": action,
        "risk": {"overall": r_overall, "latest": r_latest, "streak": r_streak,
                 "trend": r_trend, "participation": r_part, "rating": r_rating,
                 "volume": r_vol},
    })
    return rec


_BAND_RANK = {"Critical": 0, "High": 1, "Medium": 2, "OK": 3}


# =============================================================================
#  ROW / MATRIX BUILDERS
# =============================================================================
def _diag_row(rec):
    def num(v):
        return "" if v is None else v
    sess = f"{rec['present']}/{rec['total']}"
    return {
        "Priority": rec["band"],
        "Attention Score": rec["score"],
        "Technology": rec["technology"],
        "Batch Timing": rec["batch_timing"],
        "Student Name": rec["student_name"],
        "Phone": rec["phone"],
        "Email": rec["email"],
        "Student ID": rec["student_id"],
        "Latest Session": rec["latest_date"].strftime("%Y-%m-%d") if rec["latest_date"] else "",
        "Latest Status": rec["latest_status"],
        "Attendance %": num(rec["attendance_pct"]),
        "Attendance Band": rec["att_band"],
        "Absent Streak": rec["streak"],
        "Sessions (P/T)": sess,
        "Last Present": rec["last_present"].strftime("%Y-%m-%d") if rec["last_present"] else "",
        "Recent Trend %": num(rec["trend_pct"]),
        "Feedbacks": rec["n_feedbacks"],
        "Feedback Part. %": num(rec["participation_pct"]),
        "Avg Rating (/10)": num(rec["avg_rating"]),
        "Why Flagged": rec["why"],
        "Recommended Action": rec["action"],
    }


def _matrix(cols, dict_rows):
    return [cols] + [[d.get(c, "") for c in cols] for d in dict_rows]


def build_tasks_matrix(flagged, preserved):
    """flagged: sorted list of scored recs requiring follow-up.
    preserved: {(student_id, technology): {entry_col: value}} carried from a
    same-day earlier run."""
    rows = []
    for rec in flagged:
        d = _diag_row(rec)
        keep = preserved.get((rec["student_id"], rec["technology"]), {})
        for c in ENTRY_COLS:
            d[c] = keep.get(c, "")
        rows.append(d)
    return _matrix(TASKS_COLS, rows)


def build_master_matrix(all_recs):
    rows = []
    for rec in sorted(all_recs, key=lambda r: (-r["score"], r["technology"], r["student_name"])):
        d = _diag_row(rec)
        d["Requires Follow-Up"] = "Yes" if rec["requires_followup"] else "No"
        rows.append(d)
    return _matrix(MASTER_COLS, rows)


def build_summary_matrix(all_recs, flagged, meta, cfg):
    """Static per-technology counts + LIVE completion formulas referencing the
    Follow-Up Tasks tab (update as the coordinator fills entries)."""
    from collections import defaultdict

    tq = f"'{TASKS_TAB}'"
    tech_c = _col(_TECH_COL_IDX)
    pri_c = _col(_PRIORITY_COL_IDX)
    # Outcome lives in the entry block on the Tasks tab
    out_c = _col(len(DIAG_COLS) + ENTRY_COLS.index("Outcome"))

    per = defaultdict(lambda: {"appl": 0, "flag": 0, "Critical": 0, "High": 0, "Medium": 0})
    for r in all_recs:
        per[r["technology"]]["appl"] += 1
    for r in flagged:
        per[r["technology"]]["flag"] += 1
        per[r["technology"]][r["band"]] += 1

    gen = datetime.now(IST).strftime("%Y-%m-%d %H:%M IST")
    rd = meta["report_date"].strftime("%Y-%m-%d")
    m = []
    m.append([f"IntelliBI — Batch Coordinator Attendance Tasks"])
    m.append([f"Report day: {rd}    |    Generated: {gen}"])
    m.append([f"Applicable students: {len(all_recs)}    |    Requiring follow-up: {len(flagged)}    |    "
              f"Excluded (not-attendance-required / suspended / left): "
              f"{meta['excluded_not_applicable']} / {meta['excluded_suspended']} / {meta['excluded_deleted']}"])
    m.append([""])
    header = ["Technology", "Applicable", "Flagged", "Critical", "High", "Medium",
              "Completed", "Pending", "Completion %", "Critical Unattended"]
    m.append(header)

    first_data_row = len(m) + 1     # 1-based sheet row of the first technology row

    def tech_formulas(tech, flag_count):
        t = tech.replace('"', '""')
        completed = (f'=COUNTIFS({tq}!{tech_c}2:{tech_c},"{t}",{tq}!{out_c}2:{out_c},"Completed")')
        # Pending references the same-row Flagged (col C) and Completed (col G)
        return completed

    grand = {"appl": 0, "flag": 0, "Critical": 0, "High": 0, "Medium": 0}
    row_idx = first_data_row
    for tech in sorted(per.keys()):
        v = per[tech]
        for k in grand:
            grand[k] += v[k]
        t = tech.replace('"', '""')
        completed_f = f'=COUNTIFS({tq}!{tech_c}2:{tech_c},"{t}",{tq}!{out_c}2:{out_c},"Completed")'
        pending_f = f"=MAX(0,C{row_idx}-G{row_idx})"
        comp_pct_f = f'=IF(C{row_idx}=0,"—",TEXT(G{row_idx}/C{row_idx},"0%"))'
        crit_un_f = (f'=COUNTIFS({tq}!{tech_c}2:{tech_c},"{t}",{tq}!{pri_c}2:{pri_c},'
                     f'"Critical",{tq}!{out_c}2:{out_c},"<>Completed")')
        m.append([tech, v["appl"], v["flag"], v["Critical"], v["High"], v["Medium"],
                  completed_f, pending_f, comp_pct_f, crit_un_f])
        row_idx += 1

    # grand total row
    total_row = row_idx
    completed_all = f"=SUM(G{first_data_row}:G{row_idx-1})" if row_idx > first_data_row else 0
    pending_all = f"=MAX(0,C{total_row}-G{total_row})"
    comp_pct_all = f'=IF(C{total_row}=0,"—",TEXT(G{total_row}/C{total_row},"0%"))'
    crit_un_all = f"=SUM(J{first_data_row}:J{row_idx-1})" if row_idx > first_data_row else 0
    m.append(["TOTAL", grand["appl"], grand["flag"], grand["Critical"], grand["High"],
              grand["Medium"], completed_all, pending_all, comp_pct_all, crit_un_all])

    return m, {"header_row": first_data_row - 1, "first_data_row": first_data_row,
               "total_row": total_row, "n_tech": len(per)}


def build_guide_matrix(cfg):
    sc = cfg["scoring"]
    w = sc["weights"]
    rows = [
        ["IntelliBI — Batch Coordinator Attendance Task Report — Guide"],
        [""],
        ["What this is"],
        ["A daily, technology-wise action list of the students who genuinely need follow-up, "
         "judged on today's/latest performance AND till-date history. Students performing well "
         "are intentionally left off so you don't chase them unnecessarily."],
        [""],
        ["How a student gets flagged — Attention Score (0–100)"],
        ["The score is a transparent weighted sum of risk signals (higher = more attention needed):"],
        [f"  • Overall attendance shortfall   weight {w['overall_attendance']}"],
        [f"  • Absent in latest session       weight {w['latest_absent']}"],
        [f"  • Consecutive-absence streak     weight {w['absent_streak']}"],
        [f"  • Recent {sc['trend_window_sessions']}-session trend         weight {w['recent_trend']}"],
        [f"  • Feedback participation         weight {w['feedback_participation']}"],
        [f"  • Feedback rating (/10)          weight {w['feedback_rating']}"],
        [f"  • No feedback despite attending  weight {w['feedback_volume']}"],
        [""],
        ["Priority bands"],
        [f"  Critical  ≥ {sc['band_cutoffs']['critical']}   |   High ≥ {sc['band_cutoffs']['high']}   "
         f"|   Medium ≥ {sc['band_cutoffs']['medium']}   |   below = not flagged"],
        ["Hard triggers also force a flag regardless of score: attendance below "
         f"{sc['attendance_critical_pct']}%, {sc['streak_flag']}+ absences in a row, absent-again with "
         f"history below {sc['attendance_target_pct']}%, zero feedback despite attending, or a low rating."],
        [""],
        ["Why 'absent today' is NOT the same as 'needs follow-up'"],
        ["A student absent once but with excellent history and no streak scores low and is left off. "
         "A student present today but with weak history/feedback IS flagged. The score handles both."],
        [""],
        ["Your job (Coordinator) — fill these columns on the 'Follow-Up Tasks' tab"],
        ["  Follow-Up Done?  ·  Action Taken  ·  Follow-Up Comment (required)  ·  Outcome  ·  "
         "Coordinator  ·  Follow-Up Date  ·  Next Action"],
        ["The grey diagnostic columns are auto-generated — please don't edit them."],
        [""],
        ["Management view — 'Summary' tab"],
        ["Per technology and overall: applicable, flagged, by priority, and LIVE follow-ups "
         "completed / pending / completion % / critical cases still unattended — these update "
         "automatically as you fill the Outcome column."],
        [""],
        ["Tuning"],
        ["All weights, thresholds and option lists live in "
         "config/coordinator_attendance_config.json (optional). Change them there — no code edit."],
    ]
    return rows


# =============================================================================
#  SHEETS: column letters + formatting request builders
# =============================================================================
def _col(i):
    s = ""
    n = i + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _header_fmt(gid, ncols):
    return {"repeatCell": {
        "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {
            "backgroundColor": C_NAVY,
            "textFormat": {"bold": True, "foregroundColor": C_HEADER_TXT, "fontSize": 10},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP"}},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy)"}}


def _freeze(gid, rows=1, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": gid, "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"}}


def _width(gid, idx, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}


def _hide(gid, idx):
    return {"updateDimensionProperties": {
        "range": {"sheetId": gid, "dimension": "COLUMNS", "startIndex": idx, "endIndex": idx + 1},
        "properties": {"hiddenByUser": True}, "fields": "hiddenByUser"}}


def _numfmt(gid, idx, pattern, start_row=1):
    return {"repeatCell": {
        "range": {"sheetId": gid, "startRowIndex": start_row, "startColumnIndex": idx, "endColumnIndex": idx + 1},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": pattern}}},
        "fields": "userEnteredFormat.numberFormat"}}


def _validation(gid, col_idx, nrows, options):
    return {"setDataValidation": {
        "range": {"sheetId": gid, "startRowIndex": 1, "endRowIndex": 1 + nrows,
                  "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1},
        "rule": {"condition": {"type": "ONE_OF_LIST",
                               "values": [{"userEnteredValue": v} for v in options]},
                 "strict": False, "showCustomUi": True}}}


def _cf_text(gid, col_idx, nrows, text, bg, txt=None, bold=True):
    fmt = {"backgroundColor": bg}
    if txt:
        fmt["textFormat"] = {"foregroundColor": txt, "bold": bold}
    return {"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": gid, "startRowIndex": 1, "endRowIndex": 1 + nrows,
                    "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}],
        "booleanRule": {"condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": text}]},
                        "format": fmt}}}}


def _cf_numless(gid, col_idx, nrows, threshold, bg, txt=None):
    fmt = {"backgroundColor": bg}
    if txt:
        fmt["textFormat"] = {"foregroundColor": txt, "bold": True}
    return {"addConditionalFormatRule": {"index": 0, "rule": {
        "ranges": [{"sheetId": gid, "startRowIndex": 1, "endRowIndex": 1 + nrows,
                    "startColumnIndex": col_idx, "endColumnIndex": col_idx + 1}],
        "booleanRule": {"condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": str(threshold)}]},
                        "format": fmt}}}}


def _protect(gid, start_col, end_col, desc):
    return {"addProtectedRange": {"protectedRange": {
        "range": {"sheetId": gid, "startColumnIndex": start_col, "endColumnIndex": end_col},
        "warningOnly": True, "description": desc}}}


# width hints by header name (px); default 105
_WIDTH = {
    "Student Name": 170, "Email": 210, "Phone": 120, "Why Flagged": 340,
    "Recommended Action": 260, "Technology": 150, "Attendance Band": 130,
    "Follow-Up Comment": 280, "Next Action": 200, "Action Taken": 140,
    "Outcome": 130, "Coordinator": 130, "Follow-Up Date": 120, "Batch Timing": 110,
    "Sessions (P/T)": 100, "Latest Session": 110, "Last Present": 110,
}


def _diag_format_reqs(gid, cols, nrows, protect_to):
    """Formatting common to Tasks & Master diagnostic blocks."""
    reqs = [_header_fmt(gid, len(cols)), _freeze(gid, 1, 5)]
    for i, c in enumerate(cols):
        reqs.append(_width(gid, i, _WIDTH.get(c, 105)))
    reqs.append(_hide(gid, _ID_COL_IDX))
    if nrows > 0:
        reqs.append(_numfmt(gid, _PCT_COL_IDX, '0.0"%"'))
        reqs.append(_numfmt(gid, _TREND_COL_IDX, '0.0"%"'))
        reqs.append(_numfmt(gid, _PART_COL_IDX, '0.0"%"'))
        reqs.append(_numfmt(gid, _RATING_COL_IDX, '0.0'))
        reqs.append(_numfmt(gid, _SCORE_COL_IDX, '0'))
        # priority colours
        reqs.append(_cf_text(gid, _PRIORITY_COL_IDX, nrows, "Critical", C_RED, C_RED_TXT))
        reqs.append(_cf_text(gid, _PRIORITY_COL_IDX, nrows, "High", C_ORANGE, C_RED_TXT))
        reqs.append(_cf_text(gid, _PRIORITY_COL_IDX, nrows, "Medium", C_YELLOW, C_AMBER_TXT))
        # attendance % thresholds (order: <60 first, then <75)
        reqs.append(_cf_numless(gid, _PCT_COL_IDX, nrows, 60, C_RED, C_RED_TXT))
        reqs.append(_cf_numless(gid, _PCT_COL_IDX, nrows, 75, C_AMBER, C_AMBER_TXT))
        # latest status Absent -> red
        reqs.append(_cf_text(gid, DIAG_COLS.index("Latest Status"), nrows, "Absent", C_RED, C_RED_TXT))
    # protect the diagnostic block (warning-only so automation is never blocked)
    reqs.append(_protect(gid, 0, protect_to, "Auto-generated diagnostics — please do not edit"))
    return reqs


def _reset_format_reqs(current_meta, gid):
    """Delete existing conditional-format rules / protected ranges / bandings on a
    tab so a same-day re-run does not pile them up."""
    reqs = []
    props = current_meta.get(gid, {})
    for i in range(props.get("cf", 0) - 1, -1, -1):
        reqs.append({"deleteConditionalFormatRule": {"sheetId": gid, "index": i}})
    for pid in props.get("protected", []):
        reqs.append({"deleteProtectedRange": {"protectedRangeId": pid}})
    for bid in props.get("banded", []):
        reqs.append({"deleteBanding": {"bandedRangeId": bid}})
    return reqs


# =============================================================================
#  DRIVE: day folder + native sheet (find-or-create, idempotent)
# =============================================================================
def find_or_create_folder(drive, parent_id, name):
    safe = name.replace("'", "\\'")
    q = (f"'{parent_id}' in parents and name='{safe}' and "
         f"mimeType='application/vnd.google-apps.folder' and trashed=false")
    res = _retry(lambda: drive.files().list(
        q=q, fields="files(id,name)", supportsAllDrives=True,
        includeItemsFromAllDrives=True).execute(), label="find folder")
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    f = _retry(lambda: drive.files().create(body=meta, fields="id", supportsAllDrives=True).execute(),
               label="create folder")
    log.info("Created day folder '%s'", name)
    return f["id"]


def find_sheet(drive, folder_id, title):
    safe = title.replace("'", "\\'")
    q = (f"'{folder_id}' in parents and name='{safe}' and "
         f"mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")
    res = _retry(lambda: drive.files().list(
        q=q, fields="files(id,name)", supportsAllDrives=True,
        includeItemsFromAllDrives=True).execute(), label="find sheet")
    files = res.get("files", [])
    return files[0]["id"] if files else None


def create_sheet(sheets, drive, folder_id, title):
    body = {"properties": {"title": title},
            "sheets": [{"properties": {"title": t}} for t in TAB_ORDER]}
    ss = _retry(lambda: sheets.spreadsheets().create(
        body=body, fields="spreadsheetId").execute(), label="create sheet")
    sid = ss["spreadsheetId"]
    finfo = _retry(lambda: drive.files().get(fileId=sid, fields="parents",
                   supportsAllDrives=True).execute(), label="get parents")
    prev = ",".join(finfo.get("parents", []))
    _retry(lambda: drive.files().update(
        fileId=sid, addParents=folder_id, removeParents=prev,
        fields="id", supportsAllDrives=True).execute(), label="move sheet")
    log.info("Created native Google Sheet '%s'", title)
    return sid


def sheet_meta(sheets, sid):
    """Return (title->gid, gid->{cf,protected,banded}) for the workbook."""
    meta = _retry(lambda: sheets.spreadsheets().get(
        spreadsheetId=sid,
        fields="sheets(properties(sheetId,title),conditionalFormats,protectedRanges(protectedRangeId),bandedRanges(bandedRangeId))"
    ).execute(), label="sheet meta")
    gids, feats = {}, {}
    for s in meta.get("sheets", []):
        p = s["properties"]
        gids[p["title"]] = p["sheetId"]
        feats[p["sheetId"]] = {
            "cf": len(s.get("conditionalFormats", []) or []),
            "protected": [pr["protectedRangeId"] for pr in s.get("protectedRanges", []) or []],
            "banded": [b["bandedRangeId"] for b in s.get("bandedRanges", []) or []],
        }
    return gids, feats


def ensure_tabs(sheets, sid, gids):
    """Add any missing tab (keeps IDs of existing ones)."""
    reqs = [{"addSheet": {"properties": {"title": t}}} for t in TAB_ORDER if t not in gids]
    if reqs:
        _retry(lambda: sheets.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": reqs}).execute(), label="add tabs")
        return True
    return False


def read_preserved_entries(sheets, sid):
    """From an existing same-day Tasks tab, capture coordinator entries keyed by
    (student_id, technology) so a re-run never wipes them."""
    try:
        vals = read_tab(sheets, sid, TASKS_TAB)
    except Exception:
        return {}
    out = {}
    for d in vals:
        key = (_s(d.get("Student ID")), _s(d.get("Technology")))
        if not any(key):
            continue
        entry = {c: _s(d.get(c)) for c in ENTRY_COLS if _s(d.get(c))}
        if entry:
            out[key] = entry
    return out


def write_values(sheets, sid, tab, matrix, user_entered=False):
    _retry(lambda: sheets.spreadsheets().values().clear(
        spreadsheetId=sid, range=f"'{tab}'").execute(), label=f"clear {tab}")
    if not matrix:
        return
    _retry(lambda: sheets.spreadsheets().values().update(
        spreadsheetId=sid, range=f"'{tab}'!A1",
        valueInputOption="USER_ENTERED" if user_entered else "RAW",
        body={"values": matrix}).execute(), label=f"write {tab}")


# =============================================================================
#  SUMMARY / GUIDE formatting
# =============================================================================
def _summary_format_reqs(gid, info, ncols=10):
    reqs = [_freeze(gid, info["first_data_row"] - 1, 0)]
    # title
    reqs.append({"repeatCell": {
        "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 13, "foregroundColor": C_NAVY}}},
        "fields": "userEnteredFormat.textFormat"}})
    # column header row of the per-tech table (info["header_row"] is 1-based)
    hdr = info["header_row"] - 1
    reqs.append({"repeatCell": {
        "range": {"sheetId": gid, "startRowIndex": hdr, "endRowIndex": hdr + 1, "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {"backgroundColor": C_NAVY,
                 "textFormat": {"bold": True, "foregroundColor": C_HEADER_TXT},
                 "horizontalAlignment": "CENTER", "wrapStrategy": "WRAP"}},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,wrapStrategy)"}})
    # total row bold
    tr = info["total_row"] - 1
    reqs.append({"repeatCell": {
        "range": {"sheetId": gid, "startRowIndex": tr, "endRowIndex": tr + 1, "startColumnIndex": 0, "endColumnIndex": ncols},
        "cell": {"userEnteredFormat": {"backgroundColor": C_GREEN,
                 "textFormat": {"bold": True}}},
        "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    reqs.append(_width(gid, 0, 160))
    for i in range(1, ncols):
        reqs.append(_width(gid, i, 120))
    return reqs


def _guide_format_reqs(gid):
    return [
        {"repeatCell": {
            "range": {"sheetId": gid, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontSize": 13, "foregroundColor": C_NAVY}}},
            "fields": "userEnteredFormat.textFormat"}},
        _width(gid, 0, 900),
    ]


# =============================================================================
#  ORCHESTRATION
# =============================================================================
def generate(cfg, sheets_read, sheets_write, drive):
    src = cfg["sources"]
    log.info("Reading source data ...")
    att_rows = read_tab(sheets_read, src["attendance_sheet_id"], src["attendance_tab"])
    fb_rows = read_tab(sheets_read, src["attendance_sheet_id"], src["student_feedback_tab"])
    students = read_tab(sheets_read, src["students_sheet_id"], src["students_tab"])
    log.info("Loaded: attendance=%d feedback=%d students=%d", len(att_rows), len(fb_rows), len(students))
    # all target-sheet operations (create / read-back / write / format) run as
    # the impersonated info@ credential:
    sheets = sheets_write

    records, meta = build_records(att_rows, fb_rows, students, cfg)
    if REPORT_DATE:
        try:
            meta["report_date"] = datetime.strptime(REPORT_DATE, "%Y-%m-%d").date()
        except ValueError:
            log.warning("Bad REPORT_DATE %r — using derived date.", REPORT_DATE)
    for rec in records:
        score_record(rec, cfg)
    flagged = [r for r in records if r["requires_followup"]]
    flagged.sort(key=lambda r: (r["technology"], _BAND_RANK[r["band"]], -r["score"], r["student_name"]))
    log.info("Applicable=%d  Flagged=%d  (Critical=%d High=%d Medium=%d)",
             len(records), len(flagged),
             sum(1 for r in flagged if r["band"] == "Critical"),
             sum(1 for r in flagged if r["band"] == "High"),
             sum(1 for r in flagged if r["band"] == "Medium"))

    # ── Drive: day folder + sheet (find-or-create) ────────────────────────
    day_name = meta["report_date"].strftime(cfg["drive"]["day_folder_format"])
    title = f"{cfg['drive']['sheet_title_prefix']} — {day_name}"
    folder_id = find_or_create_folder(drive, cfg["drive"]["parent_folder_id"], day_name)
    sid = find_sheet(drive, folder_id, title)
    preserved = {}
    if sid:
        log.info("Existing sheet for today found — preserving coordinator entries.")
        preserved = read_preserved_entries(sheets, sid)
    else:
        sid = create_sheet(sheets, drive, folder_id, title)

    gids, feats = sheet_meta(sheets, sid)
    if ensure_tabs(sheets, sid, gids):
        gids, feats = sheet_meta(sheets, sid)

    # ── build matrices ────────────────────────────────────────────────────
    tasks_m = build_tasks_matrix(flagged, preserved)
    master_m = build_master_matrix(records)
    summary_m, sinfo = build_summary_matrix(records, flagged, meta, cfg)
    guide_m = build_guide_matrix(cfg)

    # ── write values ──────────────────────────────────────────────────────
    write_values(sheets, sid, SUMMARY_TAB, summary_m, user_entered=True)
    write_values(sheets, sid, TASKS_TAB, tasks_m)
    write_values(sheets, sid, MASTER_TAB, master_m)
    write_values(sheets, sid, GUIDE_TAB, guide_m)

    # ── formatting (reset first so re-runs don't pile up CF/protection) ───
    reqs = []
    for gid in (gids[SUMMARY_TAB], gids[TASKS_TAB], gids[MASTER_TAB], gids[GUIDE_TAB]):
        reqs += _reset_format_reqs(feats, gid)

    n_tasks = max(len(tasks_m) - 1, 0)
    n_master = max(len(master_m) - 1, 0)

    # Tasks tab
    tg = gids[TASKS_TAB]
    reqs += _diag_format_reqs(tg, TASKS_COLS, n_tasks, protect_to=len(DIAG_COLS))
    if n_tasks > 0:
        opt = cfg["options"]
        reqs.append(_validation(tg, len(DIAG_COLS) + ENTRY_COLS.index("Follow-Up Done?"), n_tasks, opt["follow_up_done"]))
        reqs.append(_validation(tg, len(DIAG_COLS) + ENTRY_COLS.index("Action Taken"), n_tasks, opt["action_taken"]))
        reqs.append(_validation(tg, len(DIAG_COLS) + ENTRY_COLS.index("Outcome"), n_tasks, opt["outcome"]))
        oc = len(DIAG_COLS) + ENTRY_COLS.index("Outcome")
        reqs.append(_cf_text(tg, oc, n_tasks, "Completed", C_GREEN, C_GREEN_TXT))
        reqs.append(_cf_text(tg, oc, n_tasks, "Pending", C_AMBER, C_AMBER_TXT))
        reqs.append(_cf_text(tg, oc, n_tasks, "No Response", C_RED, C_RED_TXT))
        reqs.append(_cf_text(tg, oc, n_tasks, "Escalated", C_RED, C_RED_TXT))
        for c in ("Follow-Up Comment", "Next Action"):
            reqs.append(_width(tg, len(DIAG_COLS) + ENTRY_COLS.index(c), _WIDTH.get(c, 200)))
        # entry header tint (light) so coordinator sees where to type
        reqs.append({"repeatCell": {
            "range": {"sheetId": tg, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": len(DIAG_COLS), "endColumnIndex": len(TASKS_COLS)},
            "cell": {"userEnteredFormat": {"backgroundColor": C_GREEN_TXT,
                     "textFormat": {"bold": True, "foregroundColor": C_HEADER_TXT}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}})

    # Master tab
    mg = gids[MASTER_TAB]
    reqs += _diag_format_reqs(mg, MASTER_COLS, n_master, protect_to=len(MASTER_COLS))
    if n_master > 0:
        rc = MASTER_COLS.index("Requires Follow-Up")
        reqs.append(_cf_text(mg, rc, n_master, "Yes", C_AMBER, C_AMBER_TXT))

    # Summary + Guide
    reqs += _summary_format_reqs(gids[SUMMARY_TAB], sinfo)
    reqs += _guide_format_reqs(gids[GUIDE_TAB])

    # chunk the batchUpdate to stay well under request limits
    for i in range(0, len(reqs), 200):
        chunk = reqs[i:i + 200]
        _retry(lambda c=chunk: sheets.spreadsheets().batchUpdate(
            spreadsheetId=sid, body={"requests": c}).execute(), label="format batch")

    link = f"https://docs.google.com/spreadsheets/d/{sid}/edit"

    # optional sharing
    for email in cfg["drive"].get("share_with", []) or []:
        try:
            _retry(lambda e=email: drive.permissions().create(
                fileId=sid, sendNotificationEmail=False,
                body={"type": "user", "role": "writer", "emailAddress": e},
                supportsAllDrives=True).execute(), label="share")
        except Exception as e:
            log.warning("Share with %s failed: %s", email, e)

    log.info("Done. %d flagged of %d applicable. Link: %s", len(flagged), len(records), link)
    return {"link": link, "spreadsheet_id": sid, "day": day_name,
            "flagged": len(flagged), "applicable": len(records), "meta": meta,
            "critical": sum(1 for r in flagged if r["band"] == "Critical")}


def send_email(cfg, result):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    try:
        from email_config import GMAIL_SENDER, GMAIL_APP_PASS
    except Exception as e:
        log.warning("email_config not available (%s) — skipping email.", e)
        return
    to = cfg["email"]["recipients"]
    msg = MIMEMultipart()
    msg["From"] = GMAIL_SENDER
    msg["To"] = ", ".join(to)
    msg["Subject"] = f"Batch Coordinator Attendance Tasks — {result['day']}"
    body = (f"<p>Daily attendance follow-up list for <b>{result['day']}</b> is ready.</p>"
            f"<p><b>{result['flagged']}</b> of {result['applicable']} applicable students need "
            f"follow-up ({result['critical']} critical).</p>"
            f'<p><a href="{result["link"]}">Open the Google Sheet</a></p>')
    msg.attach(MIMEText(body, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_SENDER, GMAIL_APP_PASS)
        server.sendmail(GMAIL_SENDER, to, msg.as_string())
    log.info("Email sent to %s", ", ".join(to))


def main():
    logging.basicConfig(
        level=logging.INFO if VERBOSE else logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s")
    cfg = load_config()
    sheets_read, sheets_write, drive = build_services(cfg)
    result = generate(cfg, sheets_read, sheets_write, drive)
    if SEND_EMAIL:
        send_email(cfg, result)
    return result


# =============================================================================
#  SELF-TEST  (no Google calls) — exercises parsing, scoring, matrices
# =============================================================================
def _selftest():
    print("Running self-test (synthetic data, no Google calls) ...")
    cfg = load_config()
    today = datetime.now(IST).date()

    def d(n):
        return (today - timedelta(days=n)).strftime("%Y-%m-%d %H:%M:%S")

    students = [
        {"student_id": "S1", "student_name": "Ravi Kumar", "email": "ravi@x.com", "phone": "900000001",
         "Batch_Timing": "Morning", "is_attendance_required": "Y", "Is_Deleted": "N"},
        {"student_id": "S2", "student_name": "Asha Rao", "email": "asha@x.com", "phone": "900000002",
         "Batch_Timing": "Evening", "is_attendance_required": "Y", "Is_Deleted": "N"},
        {"student_id": "S3", "student_name": "Neha Jain", "email": "neha@x.com", "phone": "900000003",
         "Batch_Timing": "Weekend", "is_attendance_required": "Y", "Is_Deleted": "N"},
        {"student_id": "S4", "student_name": "Zoya Ali", "email": "zoya@x.com", "phone": "900000004",
         "Batch_Timing": "Morning", "is_attendance_required": "N", "Is_Deleted": "N"},  # excluded
    ]
    att = []

    def add(sid, tech, offsets, statuses, cum):
        for off, st in zip(offsets, statuses):
            att.append({"student_id": sid, "course_name": tech, "session_id": f"{tech}-{off}",
                        "session_start_ist": d(off), "status": st, "suspend_status": "Active",
                        "attendance_percent": f"{cum:.2f}%"})

    # S1: excellent history but absent latest -> should NOT be flagged
    add("S1", "Data Analytics", [10, 8, 6, 4, 2, 0], ["Present"]*5 + ["Absent"], 96)
    # S2: present latest but weak history (~58%) -> flagged (critical)
    add("S2", "Data Analytics", [10, 8, 6, 4, 2, 0],
        ["Absent", "Present", "Absent", "Absent", "Present", "Present"], 55)
    # S3: streak of 3 absences, mid history -> flagged
    add("S3", "Full Stack", [10, 8, 6, 4, 2, 0],
        ["Present", "Present", "Present", "Absent", "Absent", "Absent"], 70)
    # S4: not attendance-required -> excluded entirely
    add("S4", "Full Stack", [4, 2, 0], ["Absent", "Absent", "Absent"], 20)

    fb = [
        {"student_id": "S1", "course_name": "Data Analytics", "session_id": "Data Analytics-10", "rating": "9"},
        {"student_id": "S1", "course_name": "Data Analytics", "session_id": "Data Analytics-8", "rating": "8"},
        {"student_id": "S2", "course_name": "Data Analytics", "session_id": "Data Analytics-2", "rating": "3"},
        # S3: no feedback at all
    ]

    records, meta = build_records(att, fb, students, cfg)
    for r in records:
        score_record(r, cfg)
    records.sort(key=lambda r: -r["score"])

    print(f"\nReport date (derived): {meta['report_date']}")
    print(f"Applicable records: {len(records)}  | excluded not-applicable: {meta['excluded_not_applicable']}")
    assert all(r["student_id"] != "S4" for r in records), "S4 should be excluded"
    print(f"\n{'ID':>3} {'Tech':<15} {'Att%':>5} {'Streak':>6} {'Latest':>8} {'Score':>5} {'Band':>8}  Why")
    for r in records:
        print(f"{r['student_id']:>3} {r['technology']:<15} {r['attendance_pct']:>5} "
              f"{r['streak']:>6} {r['latest_status']:>8} {r['score']:>5} {r['band']:>8}  {r['why']}")

    by_id = {r["student_id"]: r for r in records}
    assert by_id["S1"]["band"] == "OK", f"S1 should be OK, got {by_id['S1']['band']} — absent once but excellent"
    assert by_id["S2"]["requires_followup"], "S2 (weak history) must be flagged"
    assert by_id["S3"]["requires_followup"], "S3 (streak) must be flagged"
    assert by_id["S3"]["streak"] == 3, f"S3 streak should be 3, got {by_id['S3']['streak']}"

    flagged = [r for r in records if r["requires_followup"]]
    tasks_m = build_tasks_matrix(flagged, {})
    master_m = build_master_matrix(records)
    summary_m, sinfo = build_summary_matrix(records, flagged, meta, cfg)
    guide_m = build_guide_matrix(cfg)
    assert tasks_m[0] == TASKS_COLS
    assert master_m[0] == MASTER_COLS
    assert len(tasks_m) - 1 == len(flagged)
    # every row width matches header width
    for m, name in ((tasks_m, "tasks"), (master_m, "master")):
        w = len(m[0])
        assert all(len(row) == w for row in m), f"{name} ragged"

    # entry preservation
    preserved = {("S2", "Data Analytics"): {"Follow-Up Comment": "Called, will rejoin", "Outcome": "Completed"}}
    tasks_m2 = build_tasks_matrix(flagged, preserved)
    ci = TASKS_COLS.index("Follow-Up Comment")
    idi = TASKS_COLS.index("Student ID")
    got = [row[ci] for row in tasks_m2[1:] if row[idi] == "S2"]
    assert got and got[0] == "Called, will rejoin", "entry preservation failed"

    print(f"\nSummary rows: {len(summary_m)} | tech table starts row {sinfo['first_data_row']} "
          f"| total row {sinfo['total_row']}")
    print("Sample Summary completion formula:",
          [c for c in summary_m[sinfo['first_data_row'] - 1] if isinstance(c, str) and c.startswith("=COUNTIFS")][:1])
    print(f"Guide rows: {len(guide_m)}")
    print("\n✓ Self-test passed.")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        _selftest()
    else:
        main()
