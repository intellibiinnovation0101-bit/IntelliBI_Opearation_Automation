"""
================================================================================
  IntelliBI Student Info Pipeline
  Loads three tabs in one run:
    Pipeline A — ClassLearnerTeacherEnrolled  (GET /institutes/{id}/sessions)
                   one row per class × enrolled student, with instructor(s)
    Pipeline B — Students + Payments          (GET /institutes/v3/{id}/students)

  All tabs live in: IntelliBIStudentInfo
  SHEET ID: 1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA

  FAILURE PROTECTION:
    - If any paginated fetch fails mid-way, the entire pipeline for that
      section is aborted and NOTHING is written to the sheet.
    - Each pipeline is independent — a failure in one does not block others.

SETUP:
  1. Place service_account.json in the same folder as this script.
  2. Share 'IntelliBIStudentInfo' with the service account email (Editor access).
  3. pip install -r requirements.txt
================================================================================
"""

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import json
import os
import re
import hashlib
import requests
from datetime import datetime

from utils import (
    get_sheets_service, upsert_rows, upsert_scd2, overwrite_rows,
    sort_sheet_by_column,
    to_ist_ymd, now_ist_ymd, col_letter,
)

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

from wise_config import HEADERS   # central headers; rotate API key in config_files/config.py only
# Refresh mode (internal config; replaces the old --force-refresh CLI flag).
#   "cache"         -> use the file cache (default).
#   "force-refresh" -> bypass the cache and fetch everything fresh from the API.
REFRESH_MODE = "cache"
INSTITUTE_ID         = "672a0918ae3d6e9fadfbc622"
BASE_URL             = "https://api.wiseapp.live"
PAGE_SIZE            = 50
SERVICE_ACCOUNT_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json"
)

SHEET_ID = "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA"   # IntelliBIStudentInfo

COMBINED_TAB     = "ClassLearnerTeacherEnrolled"
STUDENTS_TAB     = "Students"
PAYMENTS_TAB     = "Payments"
INSTRUCTORS_TAB  = "Instructor"

# Instructor API endpoint (relative to BASE_URL). {institute_id} is substituted.
# Confirmed live route: /institutes/v2/{id}/teachers -> data.teachers[] each with a
# nested userId object. Query params below mirror the portal's own request.
INSTRUCTORS_PATH = "/institutes/v2/{institute_id}/teachers"
INSTRUCTORS_PARAMS = {
    "showGoogleCalendarStatus": "true",
    "showClasses":              "true",
    "showClassCount":           "false",
    "showOwner":                "true",
}

# Per-instructor identity enrichment. Adding an alternate number/email in the
# portal POSTs to /institutes/{id}/participants/{pid}/addNewIdentity, so the
# participant resource holds ALL identities (primary + alternates). We GET it per
# instructor to fill the alternative_* columns. Best-effort & non-fatal; flip the
# switch off to skip these calls. Adjust PARTICIPANT_PATH if the GET route differs.
PARTICIPANT_PATH = "/institutes/{institute_id}/participants/{participant_id}"
ENRICH_PARTICIPANT_IDENTITIES = True
TTL_PARTICIPANTS_SP = 24 * 3600   # 24 hours - identities rarely change

_CLASS_INSTR_CACHE = None
def _fetch_class_instructor_map() -> dict:
    """{class_id: "Tag1, Tag2"} built from the Classes API metadata.tags.
    The sessions API classId object carries NO metadata, so the instructor
    tag(s) must be joined from GET /institutes/{id}/classes by class id.
    Fetched once per process run."""
    global _CLASS_INSTR_CACHE
    if _CLASS_INSTR_CACHE is not None:
        return _CLASS_INSTR_CACHE
    m = {}
    try:
        resp = requests.get(f"{BASE_URL}/institutes/{INSTITUTE_ID}/classes",
                            headers=HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if isinstance(data, dict):
            classes = ((data.get("data") or {}).get("classes")
                       or data.get("data") or data.get("classes") or [])
        elif isinstance(data, list):
            classes = data
        else:
            classes = []
        for c in classes:
            if not isinstance(c, dict):
                continue
            cid = c.get("_id") or ""
            md = c.get("metadata") or {}
            tags = md.get("tags") if isinstance(md, dict) else None
            if isinstance(tags, list):
                name = ", ".join(str(t).strip() for t in tags if str(t).strip())
            else:
                name = str(tags).strip() if tags else ""
            if cid:
                m[cid] = name
        print(f"[Instructor] class->instructor map: {len(m)} classes")
    except Exception as e:
        print(f"[Instructor] WARN could not fetch class tags: {e}")
    _CLASS_INSTR_CACHE = m
    return m

COMBINED_COLUMNS = [
    "class_id", "class_name", "class_subject", "class_type",
    "teacher_ids", "teachers", "Instructor_Name",
    "student_id", "student_count",
    "start_time", "end_time", "created_at",
    "synced_at",
]
STUDENTS_COLUMNS = [
    "student_id", "student_name", "email", "phone",
    "batch_name", "Batch_Timing", "joined_on",
    "candidate_name", "highest_education", "passout_year",
    "how_heard_intellibi", "reference_name",
    "fresher_working_professional", "years_experience", "it_non_it",
    "it_domain", "non_it_industry", "follow_up_comments",
    "is_candidate_active", "is_attendance_required",
    "reg_created_at", "reg_updated_at", "reg_enabled", "reg_status",
    "Is_Deleted",
    "synced_at",
    "profile_picture",      # kept LAST so the column appears at the end of the tab
]
PAYMENTS_COLUMNS = [
    "student_id", "student_name", "class_id", "class_name",
    "total_paid", "total_paid_all_time", "total_due",
    "total_overdue", "total_remaining", "currency", "due_date", "synced_at",
]
# Instructor tab — SCD Type-2. Surrogate_Key / Is_Active / Record_Version /
# Start_Effective_Date / End_Effective_Date are managed by upsert_scd2().
# The teacher `tags` (real instructor names, comma-separated) are split into this
# many instructor_name_N columns. Increase if any instructor ever has more names.
INSTRUCTOR_NAME_SLOTS = 5
_INSTRUCTOR_NAME_COLS = [f"instructor_name_{i}" for i in range(1, INSTRUCTOR_NAME_SLOTS + 1)]

# Number of alternative_email_* columns. An instructor row can carry several
# assigned faculty members (instructor_name_1..N), so we keep the same number of
# alternate-email slots to avoid dropping any faculty email (previously capped at 2).
INSTRUCTOR_EMAIL_SLOTS = 5
_INSTRUCTOR_ALT_EMAIL_COLS = [f"alternative_email_{i}" for i in range(1, INSTRUCTOR_EMAIL_SLOTS + 1)]

# Number of alternative_contact_number_* columns — same rationale as emails, so
# no assigned faculty phone number is dropped when a row has 3+ instructors.
INSTRUCTOR_CONTACT_SLOTS = 5
_INSTRUCTOR_ALT_CONTACT_COLS = [f"alternative_contact_number_{i}" for i in range(1, INSTRUCTOR_CONTACT_SLOTS + 1)]

# Fixed number written as primary_contact_number for EVERY instructor record; the
# record's own numbers are shifted down one alternate slot (see _apply_contact_shift).
DEFAULT_PRIMARY_CONTACT_NUMBER = "+91 70206 29915"

INSTRUCTOR_COLUMNS = [
    "Surrogate_Key",
    "instructor_id",
    "instructor_details",
    "role",
    *_INSTRUCTOR_NAME_COLS,
    "primary_contact_number",
    *_INSTRUCTOR_ALT_CONTACT_COLS,
    "primary_email",
    *_INSTRUCTOR_ALT_EMAIL_COLS,
    "profile_picture",
    "google_calendar_linked",
    "joined_on",
    "class_count",
    "Is_Active",
    "Record_Version",
    "Start_Effective_Date",
    "End_Effective_Date",
    "synced_at",
]
# Business attributes compared to decide whether a new SCD-2 version is needed.
INSTRUCTOR_COMPARE_COLS = [
    "instructor_details", "role",
    *_INSTRUCTOR_NAME_COLS,
    "primary_contact_number", *_INSTRUCTOR_ALT_CONTACT_COLS,
    "primary_email", *_INSTRUCTOR_ALT_EMAIL_COLS,
]

# HEADERS is imported from config (see top of file)


# ─────────────────────────────────────────────────────────────────────────────
#  FILE-BASED API CACHE
#  Caches expensive API responses (sessions, students, enriched profiles)
#  to avoid redundant hits. Use --force-refresh to bypass cache entirely.
#  NOTE: Existing upsert logic is NOT affected — cache only wraps fetches.
# ─────────────────────────────────────────────────────────────────────────────

_SCRIPT_DIR_CACHE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR         = os.path.join(str(PROJECT_CACHE_DIR), "student_payment")

TTL_SESSIONS_SP    = 12 * 3600    # 12 hours — session/class list changes slowly
TTL_STUDENTS_SP    = 12 * 3600    # 12 hours — student roster is stable
TTL_ENRICHED       = 24 * 3600    # 24 hours — profile extras rarely change
TTL_INSTRUCTORS_SP = 12 * 3600    # 12 hours — instructor roster is stable

def _ensure_cache_dirs():
    """Create cache directory structure if missing."""
    for sub in ["", "profiles"]:
        path = os.path.join(CACHE_DIR, sub) if sub else CACHE_DIR
        os.makedirs(path, exist_ok=True)


def _cache_path(category: str, key: str = "") -> str:
    """Build the file path for a cache entry. KEYED entries get their own file in
    a per-category subdir so different keys never collide. (Previously every key
    in a non-"profiles" category overwrote one shared "{category}.json" file,
    which poisoned per-participant caches — e.g. sp_participant — making every
    instructor read the first-written participant's identities on cached runs.)"""
    if key:
        safe_key = hashlib.md5(key.encode()).hexdigest()
        return os.path.join(CACHE_DIR, category, f"{safe_key}.json")
    if category == "profiles":
        return os.path.join(CACHE_DIR, category, "data.json")
    return os.path.join(CACHE_DIR, f"{category}.json")


def _cache_get(category: str, key: str = "", ttl_seconds: int = 3600):
    """Read a cached JSON entry if it exists and is within TTL. Returns None if stale/missing."""
    path = _cache_path(category, key)
    if not os.path.isfile(path):
        return None
    try:
        mtime = os.path.getmtime(path)
        age   = datetime.now().timestamp() - mtime
        if age > ttl_seconds:
            return None  # stale
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _cache_put(category: str, key: str, data):
    """Write a JSON-serialisable object to cache."""
    path = _cache_path(category, key)
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception as e:
        print(f"  [Cache] Warning — could not write {path}: {e}")


def _cache_clear_all():
    """Remove all cached data (used with --force-refresh)."""
    import shutil
    if os.path.isdir(CACHE_DIR):
        shutil.rmtree(CACHE_DIR)
        print("[Cache] All cached data cleared.")


# ── Data-cleaning helpers ─────────────────────────────────────────────────────

def _trim(val) -> str:
    """TRIM: strip spaces, tabs and newlines from both sides."""
    if val is None:
        return ""
    return str(val).strip()


def _tc(val) -> str:
    """TRIM + InitCap: strip whitespace then title-case every word."""
    v = _trim(val)
    return v.title() if v else ""


def _lower(val) -> str:
    """TRIM + lowercase."""
    return _trim(val).lower()


# ── additionalNote parser ─────────────────────────────────────────────────────

# Map normalised key text → column name
_NOTE_KEY_MAP = {
    "candidate name":                   "candidate_name",
    "highest education":                "highest_education",
    "passout year":                     "passout_year",
    "how did you hear about intellibi": "how_heard_intellibi",
    "reference name":                   "reference_name",
    "fresher / working professional":   "fresher_working_professional",
    "fresher/working professional":     "fresher_working_professional",
    "no of years experience":           "years_experience",
    "no. of years experience":          "years_experience",
    "it / non it":                      "it_non_it",
    "it/ non it":                       "it_non_it",
    "it/non it":                        "it_non_it",
    "it domain":                        "it_domain",
    "non-it industry":                  "non_it_industry",
    "follow-up comments":               "follow_up_comments",
    "iscandidateactive":                "is_candidate_active",
    "is candidate active":              "is_candidate_active",
    "isattendancerequired":             "is_attendance_required",
    "is attendance required":           "is_attendance_required",
}

_NOTE_COLUMNS = {
    "candidate_name", "highest_education", "passout_year",
    "how_heard_intellibi", "reference_name",
    "fresher_working_professional", "years_experience", "it_non_it",
    "it_domain", "non_it_industry", "follow_up_comments",
    "is_candidate_active", "is_attendance_required",
}


def _norm_note_key(s: str) -> str:
    """Normalise a note key so matching is tolerant of real-world formatting.
    Drops zero-width / BOM characters (a stray BOM on the first line, e.g.
    'Candidate Name', was silently breaking the candidate_name match), collapses
    internal whitespace, lower-cases, and trims surrounding punctuation such as
    a trailing '?' on 'How did you hear about IntelliBI?'."""
    s = (s or "").replace("﻿", "").replace("​", "")
    s = re.sub(r"\s+", " ", s).strip().lower()
    # Also strip surrounding quotes (straight or smart) — notes are often stored
    # wrapped in quotes, which would otherwise corrupt the first key.
    return s.strip(" ?:.\t\"'“”‘’")


# Pre-normalised lookup so both the map keys and the incoming keys are compared
# on the same normalised footing.
_NOTE_KEY_MAP_NORM = {_norm_note_key(k): v for k, v in _NOTE_KEY_MAP.items()}


# ── Follow-Up Comments formatting ─────────────────────────────────────────────
# Each follow-up entry is a "<date> <comment>" pair. Coordinators type these by
# hand, so several entries often run together, the date format is NOT standardised
# and the separator after the date is unreliable — it may be ':', '-', a dash, or
# missing entirely. The *date* is therefore the only dependable signal for where a
# new entry begins, so detection is driven purely by these date patterns (no
# separator required). The date itself is never parsed or validated — the patterns
# are used solely to locate entry boundaries.
_FOLLOW_UP_DATE_PATTERNS = [
    r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",              # 31/03/2026 · 16-03-2026
    r"\d{4}[/-]\d{1,2}[/-]\d{1,2}",                 # 2026-03-31 (ISO / year-first)
    r"\d{1,2}[\s/-]+[A-Za-z]{3,9}[\s/-]+\d{2,4}",  # 16 Mar 2026 · 16-Mar-2026
    r"[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}",         # March 16, 2026
]
# A new entry begins at each date occurrence. A date immediately preceded by a
# digit, '/' or '-' is skipped (it's part of a larger number/date, not a new
# entry start). No trailing separator is required.
_FOLLOW_UP_ENTRY_RE = re.compile(
    r"(?<![\d/-])(?:" + "|".join(_FOLLOW_UP_DATE_PATTERNS) + r")"
)


def _format_follow_up_comments(text: str) -> str:
    """Place each date-wise follow-up entry on its own line.

    Locates every '<date> -' entry start and splits the text there, so multiple
    follow-ups that were run together (or wrapped arbitrarily) each end up on a
    separate line. The comment text, punctuation and special characters are left
    untouched — only whitespace/newlines *between* entries are normalised so each
    entry occupies exactly one line. If fewer than two entries are detected the
    text is returned as-is (only outer whitespace trimmed).
    """
    if not text or not text.strip():
        return (text or "").strip()

    starts = [m.start() for m in _FOLLOW_UP_ENTRY_RE.finditer(text)]
    if len(starts) <= 1:
        # 0 or 1 entry — nothing to separate; collapse internal wrapping only.
        return re.sub(r"\s*\n\s*", " ", text).strip()

    # Cut the text at the start of the 2nd..Nth entries (the 1st stays put).
    segments, prev = [], 0
    for cut in starts[1:]:
        segments.append(text[prev:cut])
        prev = cut
    segments.append(text[prev:])

    # Collapse any internal line-wrapping inside an entry to a single space so
    # each Date + Comment sits on one line; keep the entry text otherwise intact.
    cleaned = [re.sub(r"\s*\n\s*", " ", s).strip() for s in segments]
    return "\n".join(s for s in cleaned if s)

# Placeholder values that should be treated as "not provided".
_NOTE_PLACEHOLDERS = {"", "n/a", "na", "none", "null", "-", "--"}

# Admin flag keys that commonly appear AFTER 'Follow-Up Comments' in the note.
# The follow-up capture must stop when it reaches one of these so the flag is
# still parsed (otherwise it gets swallowed into the multi-line comment value).
_FOLLOW_UP_TRAILING_KEYS = {"is_candidate_active", "is_attendance_required"}


def _parse_additional_note(note_str: str) -> dict:
    """
    Parse additionalNote (newline-separated 'Key: Value' pairs) into a dict by
    splitting each line on its FIRST ':' (so a value may itself contain ':').
    Only the key-value pairs actually PRESENT are populated; any field that is
    absent is left blank — a missing field is never an error and never skips the
    record. Keys are matched after normalisation (see _norm_note_key) so stray BOM /
    zero-width characters, extra spacing, or trailing punctuation no longer cause
    a field (notably candidate_name) to silently go unmatched.
    Returns a dict with all known column names initialised to "" by default.

    'Follow-Up Comments' is a special MULTI-LINE field and is always the last
    key in the note. Coordinators enter one follow-up entry per line, each
    beginning with a free-form (unvalidated) date. A naive line-by-line parse
    only kept the first line and dropped the rest. To fix that, once the
    Follow-Up Comments key is reached the remainder of that line PLUS every
    subsequent line is captured verbatim — newlines and original ordering
    preserved, dates left exactly as typed, nothing truncated.
    """
    result = {col: "" for col in _NOTE_COLUMNS}
    if not note_str:
        return result
    # Notes are frequently stored WRAPPED in surrounding quotes, e.g.
    # '"IsCandidateActive:N"' or '"Candidate Name: ...\n...\n"'. The wrapping quote
    # corrupts the FIRST key (so it never matches) and pollutes the LAST value
    # (leaving a stray '"'). Strip surrounding quotes/whitespace before splitting so
    # every key-value pair — including the first and last — parses correctly.
    note_str = str(note_str).strip().strip("\"'“”‘’").strip()
    if not note_str:
        return result
    lines = note_str.split("\n")
    i, n = 0, len(lines)
    while i < n:
        line = lines[i].strip()
        # Split on the FIRST ':' only. A key never contains a colon, but a value
        # can — e.g. Follow-Up Comments: '25/03/2026 : Revoked access ...'. Splitting
        # on the first ':' therefore keeps the whole value (colons and all) intact,
        # and also handles the compact 'Key:Value' form (e.g. 'IsCandidateActive:N')
        # as well as the spaced 'Key : Value' form (the key is trimmed on match).
        if ":" not in line:
            i += 1
            continue
        raw_key, _, value = line.partition(":")
        col_name = _NOTE_KEY_MAP_NORM.get(_norm_note_key(raw_key))
        if col_name == "follow_up_comments":
            # Multi-line field: grab the value on this line plus subsequent lines
            # verbatim, BUT stop as soon as a line begins a trailing admin key
            # (IsCandidateActive / IsAttendanceRequired) that sometimes appears
            # AFTER Follow-Up Comments — those must still be parsed, not swallowed.
            fu_lines = [value]
            j = i + 1
            while j < n:
                nxt = lines[j]
                if ":" in nxt and _NOTE_KEY_MAP_NORM.get(
                        _norm_note_key(nxt.partition(":")[0])) in _FOLLOW_UP_TRAILING_KEYS:
                    break
                fu_lines.append(nxt)
                j += 1
            result[col_name] = _format_follow_up_comments("\n".join(fu_lines))
            i = j                     # resume parsing at the trailing admin key (if any)
            continue
        if col_name:
            result[col_name] = _trim(value)
        i += 1
    return result


# ── Batch timing helper ───────────────────────────────────────────────────────

# A student Tag Name follows <ShortCode><MMYY><OptionalSegment>, e.g. 'DA0626',
# 'DA0626E', 'DA0626W'. The optional segment after the 4-digit MMYY encodes the
# batch timing.
_BATCH_TAG_RE = re.compile(r"^[A-Za-z]+\d{4}(.*)$")


def _derive_batch_timing(batch_name: str) -> str:
    """
    Derive Batch_Timing (Morning / Evening / Weekend) from the optional suffix
    segment of the batch tag:
        suffix 'E' → Evening
        suffix 'W' → Weekend
        no suffix (or anything else, e.g. 'M') → Morning  (default)
    batch_name may hold several comma-separated tags; the first tag is used.
    """
    if not batch_name:
        return "Morning"
    first_tag = batch_name.split(",")[0].strip()
    m = _BATCH_TAG_RE.match(first_tag)
    seg = (m.group(1).strip() if m else "").upper()
    if seg == "E":
        return "Evening"
    if seg == "W":
        return "Weekend"
    return "Morning"


# ─────────────────────────────────────────────────────────────────────────────
#  PIPELINE A — ClassLearnerTeacherEnrolled  (sessions API)
# ─────────────────────────────────────────────────────────────────────────────

def fetch_all_sessions(use_cache: bool = True) -> list:
    """
    Paginate GET /institutes/{id}/sessions (all classes, no status filter).
    Uses paginateBy=COUNT; stops when page_number exceeds page_count.
    Raises RuntimeError if ANY page fails — caller must abort before writing.
    """
    # ── Check cache first ────────────────────────────────────────────────────
    if use_cache:
        cached = _cache_get("sp_sessions", "", TTL_SESSIONS_SP)
        if cached is not None:
            print(f"[Fetch Sessions] {len(cached)} session(s) loaded from cache")
            return cached

    all_sessions = []
    page_number  = 1
    total_pages  = None   # resolved from first response

    while True:
        params = {
            "paginateBy":       "COUNT",
            "page_number":      page_number,
            "page_size":        PAGE_SIZE,
            "includeCancelled": "false",
        }
        try:
            resp = requests.get(
                f"{BASE_URL}/institutes/{INSTITUTE_ID}/sessions",
                params=params, headers=HEADERS, timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            raise RuntimeError(
                f"Failed on page {page_number} of sessions fetch: {e}\n"
                f"  → {len(all_sessions)} sessions already fetched will NOT be written."
            )

        payload     = data.get("data") or {}
        new_sessions = payload.get("sessions") or []
        if not isinstance(new_sessions, list):
            new_sessions = []

        all_sessions.extend(new_sessions)

        # Resolve total page count from first response
        if total_pages is None:
            total_pages = int(payload.get("page_count") or 1)

        if not new_sessions or page_number >= total_pages:
            break
        page_number += 1

    print(f"[Fetch Sessions] {len(all_sessions)} session(s) fetched across {page_number} page(s)")

    # ── Save to cache ────────────────────────────────────────────────────────
    if use_cache:
        _cache_put("sp_sessions", "", all_sessions)

    return all_sessions


def transform_combined(all_sessions: list) -> list:
    """
    Group sessions by class_id.  For each class collect:
      - class metadata (name, subject, classType)
      - all instructor IDs + names seen across sessions
      - enrolled student IDs (taken from the session with the most students)
    Returns one row per class × student.
    """
    synced_at  = now_ist_ymd()
    class_map  = {}   # class_id → {meta, teachers{id:name}, students set, student_count}

    for session in all_sessions:
        cls_obj = session.get("classId") or {}
        if isinstance(cls_obj, str):
            class_id = cls_obj
            class_name = class_subject = class_type = ""
        else:
            class_id      = cls_obj.get("_id") or ""
            class_name    = _trim(cls_obj.get("name") or "")
            class_subject = _trim(cls_obj.get("subject") or "")
            class_type    = _trim(cls_obj.get("classType") or "")

        if not class_id:
            continue

        if class_id not in class_map:
            class_map[class_id] = {
                "class_name":    class_name,
                "class_subject": class_subject,
                "class_type":    class_type,
                "teachers":      {},    # teacher_id → teacher_name
                "students":      set(),
                "student_count": 0,
                "start_time":    "",    # earliest session start_time for this class
                "end_time":      "",    # latest session end_time for this class
                "created_at":    "",    # earliest session createdAt for this class
            }
        entry = class_map[class_id]

        # Instructor for this session
        instr = session.get("userId") or {}
        if isinstance(instr, dict):
            tid   = instr.get("_id") or ""
            tname = _tc(instr.get("name") or "")
            if tid:
                entry["teachers"][tid] = tname

        # Enrolled students — keep the largest set seen (enrollment doesn't change)
        raw_students = session.get("students") or []
        if isinstance(raw_students, list) and len(raw_students) > len(entry["students"]):
            entry["students"]      = set(s for s in raw_students if s)
            entry["student_count"] = int(session.get("studentCount") or len(raw_students))

        # Track earliest start_time across all sessions for this class
        raw_start = session.get("start_time") or ""
        if raw_start and (not entry["start_time"] or raw_start < entry["start_time"]):
            entry["start_time"] = to_ist_ymd(raw_start)

        # Track latest end_time across all sessions for this class
        raw_end = session.get("end_time") or ""
        if raw_end and (not entry["end_time"] or raw_end > entry["end_time"]):
            entry["end_time"] = to_ist_ymd(raw_end)

        # Track earliest createdAt across all sessions for this class
        raw_created = session.get("createdAt") or ""
        if raw_created and (not entry["created_at"] or raw_created < entry["created_at"]):
            entry["created_at"] = to_ist_ymd(raw_created)

    # Build one row per class × student
    rows = []
    for class_id, info in class_map.items():
        teacher_ids_str = ", ".join(info["teachers"].keys())
        teachers_str    = ", ".join(info["teachers"].values())
        student_count   = info["student_count"] or len(info["students"])

        for student_id in sorted(info["students"]):
            rows.append({
                "class_id":      class_id,
                "class_name":    info["class_name"],
                "class_subject": info["class_subject"],
                "class_type":    info["class_type"],
                "teacher_ids":   teacher_ids_str,
                "teachers":      teachers_str,
                "Instructor_Name": _fetch_class_instructor_map().get(class_id, ""),
                "student_id":    student_id,
                "student_count": str(student_count),
                "start_time":    info["start_time"],
                "end_time":      info["end_time"],
                "created_at":    info["created_at"],
                "synced_at":     synced_at,
            })

    classes_seen = len(class_map)
    print(f"[Transform A] Classes: {classes_seen} | Combined rows: {len(rows)}")
    return rows


def run_pipeline_a(service, use_cache: bool = True) -> bool:
    """
    ClassLearnerTeacherEnrolled pipeline — uses the sessions API to combine
    class info, enrolled students, and instructor(s) into a single tab.
    Returns True on success, False on fetch failure (no writes performed).
    """
    print("\n" + "-" * 64)
    print("  Pipeline A  —  ClassLearnerTeacherEnrolled")
    print("-" * 64)

    try:
        all_sessions = fetch_all_sessions(use_cache=use_cache)
    except RuntimeError as e:
        print(f"\n[Pipeline A] ABORTED — {e}")
        print("[Pipeline A] Sheet not modified. Fix the error and re-run.")
        return False

    combined_rows = transform_combined(all_sessions)

    # ClassLearnerTeacherEnrolled is a full snapshot rebuilt from scratch every run,
    # so overwrite the tab completely. This guarantees every field lands under its
    # correct header (no column drift) and clears any stale mis-aligned rows left
    # over from an earlier layout change (e.g. when Instructor_Name was inserted).
    overwrite_rows(service, SHEET_ID, COMBINED_TAB, COMBINED_COLUMNS, combined_rows)

    print(f"  Combined rows written : {len(combined_rows)}")
    return True


# ─────────────────────────────────────────────────────────────────────────────
#  PIPELINE B — Students & Payments
# ─────────────────────────────────────────────────────────────────────────────

PROFILE_REQUEST_DELAY = 0.2   # seconds between per-student profile calls


# Fields (in priority order, per container) that may hold a student's private
# note. `additionalNote` is the primary one used by ~97% of students; the rest are
# fallbacks so a note stored under an alternate key is still recovered instead of
# showing blank. Alternates are consulted ONLY when the primary is empty, so a
# student who already has an additionalNote is never affected.
_NOTE_FIELD_CANDIDATES = (
    "additionalNote", "additional_note", "note", "notes",
    "privateNote", "private_note", "remark", "remarks",
)


def _first_note(*containers) -> str:
    """Return the first non-empty note-like string found across the given dicts,
    trying each note field name in priority order. Never raises."""
    for src in containers:
        if not isinstance(src, dict):
            continue
        for key in _NOTE_FIELD_CANDIDATES:
            v = src.get(key)
            if isinstance(v, str) and v.strip():
                return v
    return ""


def _fetch_profile_extras(student_id: str) -> dict:
    """
    Fetch tags, additionalNote and registrationData fields for one student
    from the profile API.
    Returns a dict with all fields, or empty defaults on error. The `_ok` flag
    reports whether the API call actually succeeded, so callers can tell a genuine
    "no data" from a failed fetch and avoid caching a failure as if it were empty.
    """
    defaults = {
        "tags": [], "additionalNote": "",
        "reg_enabled": "", "reg_status": "", "_ok": False,
    }
    try:
        resp = requests.get(
            f"{BASE_URL}/public/institutes/{INSTITUTE_ID}/studentReports/{student_id}",
            params={
                "showContractData":     "false",
                "showRegistrationData": "true",
                "showParent":           "false",
            },
            headers=HEADERS, timeout=30,
        )
        resp.raise_for_status()
        report = (resp.json().get("data") or {}).get("studentReport") or {}
        user   = report.get("user") or {}
        reg    = report.get("registrationData") or {}
        return {
            # Tags may live on the user or (defensively) on the report itself.
            "tags":           user.get("tags") or report.get("tags") or [],
            # Note may be stored under an alternate key or on the report/registration
            # object for some students — search them all rather than only user.additionalNote.
            "additionalNote": _first_note(user, report, reg),
            "reg_enabled":    str(reg.get("enabled")) if reg.get("enabled") is not None else "",
            "reg_status":     _trim(reg.get("status") or ""),
            "_ok":            True,
        }
    except Exception as e:
        print(f"  [Profile extras] Skipped {student_id}: {e}")
        return defaults


def enrich_students_with_profile(all_students: list, use_cache: bool = True) -> list:
    """
    For each student in all_students, call the profile API and merge
    'tags' and 'additionalNote' into the student dict.
    Returns the enriched list (same objects, mutated in-place).

    Caching strategy:
      - First checks for a complete enriched snapshot (all students combined).
      - If found and covers all current student IDs, applies it directly.
      - Otherwise falls back to per-student API calls.
    """
    import time
    total = len(all_students)

    def _apply_extras(student: dict, extras: dict) -> None:
        """Merge profile extras INTO the student WITHOUT ever letting an empty
        fetched value wipe good data the student already carries.

        Root-cause guard: the profile-extras API intermittently returns an empty
        `tags`/`additionalNote` for a student who genuinely has them (e.g. the tag
        is present in the v3 students feed but absent from the studentReports
        payload). The previous code assigned unconditionally, so that empty
        response blanked out real tags/notes. Now the fetched value is used only
        when non-empty; otherwise the student keeps whatever it already had (e.g.
        tags carried over from the v3 students endpoint)."""
        fetched_tags = extras.get("tags") or []
        student["tags"] = fetched_tags or (student.get("tags") or [])

        fetched_note = (extras.get("additionalNote") or "").strip()
        student["additionalNote"] = fetched_note or (student.get("additionalNote") or "")

        # Registration fields: prefer a non-empty fetched value, else keep existing.
        for k in ("reg_enabled", "reg_status"):
            v = extras.get(k, "")
            if v not in (None, "") or not student.get(k):
                student[k] = v if v is not None else ""

    def _cache_entry(student: dict) -> dict:
        return {
            "tags":           student.get("tags") or [],
            "additionalNote": student.get("additionalNote") or "",
            "reg_enabled":    student.get("reg_enabled", ""),
            "reg_status":     student.get("reg_status", ""),
        }

    # ── Try loading enriched snapshot from cache ─────────────────────────────
    if use_cache:
        cached_enriched = _cache_get("sp_enriched", "", TTL_ENRICHED)
        if cached_enriched is not None and isinstance(cached_enriched, dict):
            current_ids = {s.get("_id") or "" for s in all_students if s.get("_id")}
            missing     = current_ids - set(cached_enriched.keys())
            if not missing:
                # Apply cached enrichment to each student (empty values never wipe
                # the tags the student already carries from the v3 students feed).
                stale = []          # students whose cached note is still empty
                for student in all_students:
                    sid = student.get("_id") or ""
                    if sid and sid in cached_enriched:
                        _apply_extras(student, cached_enriched[sid])
                    # A cached entry with an empty note looks like a failed/partial
                    # earlier fetch (or a note added on the portal since the cache
                    # was built). The completeness check alone would keep serving
                    # that empty forever, so re-fetch just those students to recover
                    # any note now available — this is what makes the note load
                    # again without a full cache wipe.
                    if sid and not (student.get("additionalNote") or "").strip():
                        stale.append(student)
                if not stale:
                    print(f"[Enrich] {total} student(s) enriched from cache")
                    return all_students
                print(f"[Enrich] {total} student(s) from cache; re-fetching "
                      f"{len(stale)} with an empty note to recover any note/tag "
                      f"missed by an earlier partial sync …")
                updated = dict(cached_enriched)
                for idx, student in enumerate(stale, 1):
                    sid = student.get("_id") or ""
                    extras = _fetch_profile_extras(sid)
                    _apply_extras(student, extras)
                    if extras.get("_ok"):          # never persist a failed fetch
                        updated[sid] = _cache_entry(student)
                    if PROFILE_REQUEST_DELAY > 0:
                        time.sleep(PROFILE_REQUEST_DELAY)
                if use_cache:
                    _cache_put("sp_enriched", "", updated)
                print(f"[Enrich] {total} student(s) enriched "
                      f"(cache + {len(stale)} refreshed)")
                return all_students
            else:
                print(f"[Enrich] Cache has {len(cached_enriched)} students, "
                      f"but {len(missing)} new student(s) found — re-enriching all")

    # ── Fallback: per-student API calls ──────────────────────────────────────
    print(f"[Enrich] Fetching tags, additionalNote & registration data for {total} student(s) …")
    enriched_map = {}   # student_id → {tags, additionalNote, reg_enabled, reg_status}

    for idx, student in enumerate(all_students, 1):
        sid = student.get("_id") or ""
        if not sid:
            continue
        extras = _fetch_profile_extras(sid)
        _apply_extras(student, extras)     # empty fetch never wipes existing tags/note

        # Collect for cache (merged values — so tags carried from the v3 feed are
        # preserved even when the profile call returned an empty tag list).
        enriched_map[sid] = _cache_entry(student)

        if idx % 25 == 0 or idx == total:
            print(f"  [{idx}/{total}] enriched")
        if PROFILE_REQUEST_DELAY > 0:
            time.sleep(PROFILE_REQUEST_DELAY)

    # ── Save enriched snapshot to cache ──────────────────────────────────────
    if use_cache and enriched_map:
        _cache_put("sp_enriched", "", enriched_map)

    print(f"[Enrich] Done — {total} student(s) enriched")
    return all_students


def fetch_all_students(use_cache: bool = True) -> list:
    """
    Paginate /institutes/v3/{id}/students (ALL statuses) until an empty page.
    Raises RuntimeError if ANY page fails — caller must abort before writing.

    NOTE: the previous implementation passed status=ACCEPTED, which silently
    dropped valid students who exist on the portal but are not yet in the ACCEPTED
    state (e.g. invited / pending / enrolled). The status filter is removed so no
    valid student is skipped; results are de-duplicated by _id.
    """
    # ── Check cache first ────────────────────────────────────────────────────
    if use_cache:
        cached = _cache_get("sp_students", "", TTL_STUDENTS_SP)
        if cached is not None:
            print(f"[Fetch Students] {len(cached)} students loaded from cache")
            return cached

    all_students = []
    seen_ids     = set()
    page_number  = 1

    while True:
        # No 'status' filter → students of every status are returned so that valid
        # students who are not yet ACCEPTED are still populated.
        params = {
            "page_size":          PAGE_SIZE,
            "page_number":        page_number,
            "showParents":        "true",
            "showFeedbackData":   "true",
            "showContractStatus": "true",
        }
        try:
            resp = requests.get(
                f"{BASE_URL}/institutes/v3/{INSTITUTE_ID}/students",
                params=params, headers=HEADERS, timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            raise RuntimeError(
                f"Failed on page {page_number} of students fetch: {e}\n"
                f"  → {len(all_students)} students already fetched will NOT be written."
            )

        new_students = (
            (data.get("data") or {}).get("students")
            or data.get("data")
            or data.get("students")
            or (data if isinstance(data, list) else [])
        )
        if not isinstance(new_students, list):
            new_students = []

        # De-duplicate by _id (unfiltered pagination can overlap) and stop when a
        # page contributes no NEW students (guards against a repeating last page).
        added = 0
        for s in new_students:
            _sid = (s or {}).get("_id") or "" if isinstance(s, dict) else ""
            if _sid and _sid in seen_ids:
                continue
            if _sid:
                seen_ids.add(_sid)
            all_students.append(s)
            added += 1

        if not new_students or added == 0:
            break
        page_number += 1

    print(f"[Fetch Students] {len(all_students)} students fetched across {page_number} page(s)")

    # ── Save to cache ────────────────────────────────────────────────────────
    if use_cache:
        _cache_put("sp_students", "", all_students)

    return all_students


def transform_students(all_students: list) -> tuple:
    """Transform raw student records → (student_rows, payment_rows)."""
    synced_at    = now_ist_ymd()
    student_rows = []
    payment_rows = []
    student_set  = set()

    for student in all_students:
        if not student:
            continue

        sid  = student.get("_id") or ""
        snam = _tc(student.get("name") or "")       # TRIM + InitCap

        if sid and sid not in student_set:
            student_set.add(sid)

            # batch_name (tags) — list → comma-separated string
            tags_raw   = student.get("tags") or []
            batch_name = ", ".join(str(t) for t in tags_raw) if isinstance(tags_raw, list) else _trim(str(tags_raw))

            # additionalNote — parse into individual columns
            note_parsed = _parse_additional_note(student.get("additionalNote") or "")

            # candidate_name fallback: when the note has no usable candidate name
            # (missing line, blank, or an 'N/A' placeholder), fall back to the
            # student's own name so the column is never left empty.
            _cn = _trim(note_parsed.get("candidate_name"))
            if _cn.lower() in _NOTE_PLACEHOLDERS:
                note_parsed["candidate_name"] = snam

            # Batch_Timing — derived from the batch_name suffix (E/W/none)
            batch_timing = _derive_batch_timing(batch_name)

            student_rows.append({
                "student_id":   sid,
                "student_name": snam,
                "email":        _lower(student.get("email") or ""),
                "phone":        _trim(student.get("phoneNumber") or ""),
                "batch_name":   batch_name,
                "Batch_Timing": batch_timing,
                **note_parsed,
                "joined_on":    (
                    student.get("joinedOn")
                    or student.get("joined_on")
                    or student.get("joiningDate")
                    or ""
                ),
                "profile_picture": _trim(
                    student.get("profilePicture")
                    or student.get("profilePic")
                    or student.get("photo")
                    or student.get("dp")
                    or ((student.get("userId") or {}).get("profilePicture")
                        if isinstance(student.get("userId"), dict) else "")
                    or ""
                ),
                "reg_enabled":    student.get("reg_enabled") or "",
                "reg_status":     student.get("reg_status") or "",
                "Is_Deleted":     "N",   # present on portal this run
                "synced_at":      synced_at,
            })

        for cls in (student.get("classrooms") or []):
            fee = cls.get("feeSummary")
            if not fee:
                continue
            payment_rows.append({
                "student_id":          sid,
                "student_name":        snam,
                "class_id":            cls.get("_id") or "",
                "class_name":          cls.get("name") or "",
                "total_paid":          (fee.get("totalPaid") or {}).get("value") or 0,
                "total_paid_all_time": (fee.get("totalPaidAllTime") or {}).get("value") or 0,
                "total_due":           (fee.get("totalDue") or {}).get("value") or 0,
                "total_overdue":       (fee.get("totalOverDue") or {}).get("value") or 0,
                "total_remaining":     (fee.get("totalRemaining") or {}).get("value") or 0,
                "currency":            fee.get("currency") or "",
                "due_date":            fee.get("earliestDueDate") or "",
                "synced_at":           synced_at,
            })

    print(f"[Transform B] Students: {len(student_rows)} | Payments: {len(payment_rows)}")
    return student_rows, payment_rows


def fetch_existing_student_rows(service, spreadsheet_id: str, tab_name: str,
                                columns: list) -> dict:
    """
    Read the existing Students tab and return {student_id: {col: value, ...}}.
    Used to detect students that were previously synced but are no longer
    returned by the portal (i.e. deleted on the portal) so they can be flagged
    Is_Deleted='Y' without losing their existing data.
    Returns {} if the tab is empty or unreadable.
    """
    last_col = col_letter(len(columns) - 1)
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:{last_col}",
        ).execute()
    except Exception as e:
        print(f"  [Is_Deleted] Could not read existing '{tab_name}' rows: {e}")
        return {}

    values = result.get("values", [])
    if not values:
        return {}

    header = values[0]
    id_idx = header.index("student_id") if "student_id" in header else 0
    existing = {}
    for row_vals in values[1:]:
        sid = (row_vals[id_idx] if id_idx < len(row_vals) else "").strip()
        if not sid:
            continue
        existing[sid] = {
            header[i]: (row_vals[i] if i < len(row_vals) else "")
            for i in range(len(header))
        }
    return existing


def append_deleted_student_rows(service, student_rows: list) -> list:
    """
    Compare the freshly-fetched students against what's already in the Students
    tab.  Any student present in the sheet but missing from this run's fetch is
    treated as deleted-on-portal: re-emit its existing row data with
    Is_Deleted='Y' so the flag is updated without wiping the rest of the row.
    Returns the (possibly extended) student_rows list.
    """
    fetched_ids = {r.get("student_id", "") for r in student_rows}
    existing    = fetch_existing_student_rows(service, SHEET_ID, STUDENTS_TAB, STUDENTS_COLUMNS)
    synced_at   = now_ist_ymd()

    deleted_rows = []
    for sid, rec in existing.items():
        if sid in fetched_ids:
            continue
        # Preserve existing column values; only flip the flag (and synced_at).
        row = {col: rec.get(col, "") for col in STUDENTS_COLUMNS}
        row["student_id"] = sid
        row["Is_Deleted"] = "Y"
        row["synced_at"]  = synced_at
        deleted_rows.append(row)

    if deleted_rows:
        print(f"  [Is_Deleted] {len(deleted_rows)} student(s) no longer on portal -> flagged 'Y'")
    else:
        print("  [Is_Deleted] No previously-synced students are missing from the portal.")

    return student_rows + deleted_rows


def run_pipeline_b(service, use_cache: bool = True) -> bool:
    """
    Students + Payments pipeline.
    Returns True on success, False on fetch failure (no writes performed).
    """
    print("\n" + "-" * 64)
    print("  Pipeline B  -  Students & Payments")
    print("-" * 64)

    try:
        all_students = fetch_all_students(use_cache=use_cache)
    except RuntimeError as e:
        print(f"\n[Pipeline B] ABORTED - {e}")
        print("[Pipeline B] Sheet not modified. Fix the error and re-run.")
        return False

    # Enrich each student with tags + additionalNote from the profile API
    # (the /institutes/v3/{id}/students list endpoint does not return these fields)
    all_students = enrich_students_with_profile(all_students, use_cache=use_cache)

    student_rows, payment_rows = transform_students(all_students)

    # Flag students that exist in the sheet but were not returned by the portal
    # this run (deleted on portal) with Is_Deleted='Y'; fetched students are 'N'.
    student_rows = append_deleted_student_rows(service, student_rows)

    # Sort newest-first so freshly appended rows already arrive in order; the
    # sheet-level sort below is the authoritative final ordering.
    student_rows.sort(key=lambda r: _trim(r.get("joined_on")), reverse=True)

    upsert_rows(service, SHEET_ID, STUDENTS_TAB, STUDENTS_COLUMNS, student_rows, ["student_id"])
    upsert_rows(service, SHEET_ID, PAYMENTS_TAB, PAYMENTS_COLUMNS, payment_rows, ["student_id"])

    # upsert_rows updates in place / appends at the bottom and never reorders the
    # tab, so explicitly sort the whole Students sheet by joined_on descending
    # (latest joins first). Best-effort — never blocks the run.
    sort_sheet_by_column(
        service, SHEET_ID, STUDENTS_TAB, STUDENTS_COLUMNS,
        "joined_on", descending=True,
    )

    print(f"  Students written : {len(student_rows)}")
    print(f"  Payments written : {len(payment_rows)}")
    return True


# -----------------------------------------------------------------------------
#  PIPELINE C - Instructors  (SCD Type-2)
# -----------------------------------------------------------------------------

def _collect_contacts(primaries: list, extras, keys: tuple) -> list:
    """Build an ordered, de-duplicated list of contact values from a few primary
    fields plus an 'extras' list (of scalars or dicts). `keys` are the dict keys
    to look for when an extra entry is a dict (e.g. ('number', 'value'))."""
    out = []
    seen = set()

    def add(v):
        v = _trim(v)
        key = v.lower()
        if v and key not in seen:
            seen.add(key)
            out.append(v)

    for p in primaries:
        add(p)
    if isinstance(extras, list):
        for item in extras:
            if isinstance(item, dict):
                for k in keys:
                    if item.get(k):
                        add(item.get(k))
                        break
            else:
                add(item)
    return out


def fetch_all_instructors(use_cache: bool = True) -> list:
    """
    Paginate the instructor endpoint until an empty page is returned.
    Raises RuntimeError if ANY page fails - caller must abort before writing.
    """
    if use_cache:
        cached = _cache_get("sp_instructors", "", TTL_INSTRUCTORS_SP)
        if cached is not None:
            print(f"[Fetch Instructors] {len(cached)} instructor(s) loaded from cache")
            return cached

    url = f"{BASE_URL}{INSTRUCTORS_PATH.format(institute_id=INSTITUTE_ID)}"
    all_items = []
    page_number = 1
    total_pages = None

    while True:
        params = {"page_size": PAGE_SIZE, "page_number": page_number, **INSTRUCTORS_PARAMS}
        try:
            resp = requests.get(url, params=params, headers=HEADERS, timeout=30)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            raise RuntimeError(
                f"Failed on page {page_number} of instructors fetch ({url}): {e}\n"
                f"  -> {len(all_items)} instructors already fetched will NOT be written."
            )

        payload = data.get("data") if isinstance(data, dict) else {}
        if isinstance(payload, dict):
            new_items = payload.get("teachers") or payload.get("instructors") or []
        elif isinstance(payload, list):
            new_items = payload
        else:
            new_items = []
        if not isinstance(new_items, list):
            new_items = []

        all_items.extend(new_items)

        if total_pages is None and isinstance(payload, dict):
            total_pages = int(payload.get("page_count") or 1)

        if not new_items or (total_pages is not None and page_number >= total_pages):
            break
        page_number += 1

    print(f"[Fetch Instructors] {len(all_items)} instructor(s) fetched across {page_number} page(s)")

    if use_cache:
        _cache_put("sp_instructors", "", all_items)

    return all_items


def transform_instructors(all_teachers: list) -> list:
    """Map raw teacher records (data.teachers[] from /institutes/v2/{id}/teachers)
    into Instructor business columns. SCD audit columns are added later by
    upsert_scd2().

    NOTE: the v2/teachers payload exposes only ONE email and ONE phone (under the
    nested `userId`) and carries no tags. The alternative-contact / alternative-
    email / tags columns therefore stay blank unless the source object happens to
    carry alternate*/tags arrays (handled defensively below) or a richer per-user
    profile endpoint is wired in later.
    """
    synced_at = now_ist_ymd()
    rows = []
    seen = set()

    for t in all_teachers:
        if not t:
            continue
        user = t.get("userId") if isinstance(t.get("userId"), dict) else {}
        uid = _trim(user.get("_id") or t.get("_id") or "")
        if not uid or uid in seen:
            continue
        seen.add(uid)

        name = _tc(user.get("name") or t.get("name") or "")
        role = _trim(t.get("relation") or t.get("role") or "")

        tags_raw = user.get("tags") or t.get("tags") or []
        if isinstance(tags_raw, list):
            name_list = [str(x).strip() for x in tags_raw if str(x).strip()]
        else:
            name_list = [s.strip() for s in str(tags_raw).split(",") if s.strip()]
        if not name_list and name:
            name_list = [name]   # no tags -> fall back to instructor_details
        name_slots = {
            f"instructor_name_{i}": (name_list[i - 1] if i - 1 < len(name_list) else "")
            for i in range(1, INSTRUCTOR_NAME_SLOTS + 1)
        }

        phones = _collect_contacts(
            [user.get("phoneNumber"), user.get("phone")],
            user.get("alternatePhoneNumbers") or user.get("phoneNumbers")
            or t.get("alternatePhoneNumbers") or [],
            keys=("number", "value", "phoneNumber", "phone"),
        )
        emails = _collect_contacts(
            [user.get("email")],
            user.get("alternateEmails") or user.get("emails")
            or t.get("alternateEmails") or [],
            keys=("email", "value", "address"),
        )
        # Keep the primary contact/email + up to INSTRUCTOR_*_SLOTS alternates so
        # no assigned faculty phone/email is dropped when a row has 3+ instructors.
        phones = (phones + [""] * (INSTRUCTOR_CONTACT_SLOTS + 1))[:INSTRUCTOR_CONTACT_SLOTS + 1]
        emails = (emails + [""] * (INSTRUCTOR_EMAIL_SLOTS + 1))[:INSTRUCTOR_EMAIL_SLOTS + 1]

        classes = t.get("classes")
        class_count = str(len(classes)) if isinstance(classes, list) else ""
        gcal = t.get("googleCalendarLinked")
        gcal = "Y" if gcal is True else ("N" if gcal is False else "")
        joined = to_ist_ymd(t.get("joinedOn")) if t.get("joinedOn") else ""

        rows.append({
            "instructor_id":                uid,
            "instructor_details":           name,
            "role":                         role,
            **name_slots,
            "primary_contact_number":       phones[0],
            **{f"alternative_contact_number_{i}": phones[i]
               for i in range(1, INSTRUCTOR_CONTACT_SLOTS + 1)},
            "primary_email":                _lower(emails[0]),
            **{f"alternative_email_{i}": _lower(emails[i])
               for i in range(1, INSTRUCTOR_EMAIL_SLOTS + 1)},
            "profile_picture":              _trim(user.get("profilePicture") or ""),
            "google_calendar_linked":       gcal,
            "joined_on":                    joined,
            "class_count":                  class_count,
            "synced_at":                    synced_at,
        })

    print(f"[Transform C] Instructors: {len(rows)}")
    return rows


def _extract_identities(obj) -> tuple:
    """From a participant JSON, return (phones, emails) as ordered, de-duplicated
    lists with any primary entries first. Tolerant of several shapes:
      - an 'identities'/'contactIdentities' list of {type, value, isPrimary}
      - explicit emails/phoneNumbers/alternate* arrays
      - singleton email / phoneNumber fields
    """
    phones_primary, phones_other = [], []
    emails_primary, emails_other = [], []

    import re

    def classify(val, typ, primary):
        # Classify purely by the VALUE, not the type label: the portal tags email
        # identities as FIREBASE_ID, so trusting `type` drops alternate emails.
        v = _trim(val)
        if not v:
            return
        if "@" in v and "." in v.rsplit("@", 1)[-1]:
            (emails_primary if primary else emails_other).append(v)
            return
        digits = re.sub(r"[\s\-().]", "", v)
        if re.fullmatch(r"\+?\d{7,15}", digits):
            (phones_primary if primary else phones_other).append(v)
        # anything else (firebase UID, uuid, username) is ignored

    # locate the participant object(s) regardless of nesting
    containers = []
    if isinstance(obj, dict):
        for cand in (obj, obj.get("data")):
            if isinstance(cand, dict):
                containers.append(cand.get("participant") or cand.get("user") or cand)
            elif isinstance(cand, list):
                containers.extend(x for x in cand if isinstance(x, dict))
    elif isinstance(obj, list):
        containers.extend(x for x in obj if isinstance(x, dict))

    for c in containers:
        if not isinstance(c, dict):
            continue
        ids = c.get("identities") or c.get("contactIdentities") or []
        if isinstance(ids, list):
            for it in ids:
                if isinstance(it, dict):
                    typ = it.get("type") or it.get("identityType") or it.get("kind")
                    val = (it.get("value") or it.get("identity")
                           or it.get("displayIdentifier") or it.get("identifier")
                           or it.get("email") or it.get("phoneNumber") or it.get("number"))
                    primary = bool(it.get("isPrimary") or it.get("primary") or it.get("isDefault"))
                    classify(val, typ, primary)
                elif isinstance(it, str):
                    classify(it, None, False)
        for key, typ in (("emails", "EMAIL"), ("alternateEmails", "EMAIL"),
                         ("phoneNumbers", "PHONE"), ("alternatePhoneNumbers", "PHONE")):
            arr = c.get(key)
            if isinstance(arr, list):
                for it in arr:
                    if isinstance(it, dict):
                        classify(it.get("value") or it.get("email")
                                 or it.get("phoneNumber") or it.get("number"),
                                 typ, bool(it.get("isPrimary")))
                    else:
                        classify(it, typ, False)
        classify(c.get("email"), "EMAIL", True)
        classify(c.get("phoneNumber"), "PHONE", True)

    def dedup(seq):
        out, seen = [], set()
        for v in seq:
            k = v.lower()
            if k not in seen:
                seen.add(k)
                out.append(v)
        return out

    return dedup(phones_primary + phones_other), dedup(emails_primary + emails_other)


def _fetch_participant_identities(participant_id: str, use_cache: bool = True) -> tuple:
    """GET the participant resource and return (phones, emails). Non-fatal: returns
    ([], []) on any error so instructor loading is never blocked."""
    if use_cache:
        cached = _cache_get("sp_participant", participant_id, TTL_PARTICIPANTS_SP)
        if cached is not None:
            return cached.get("phones", []), cached.get("emails", [])

    url = f"{BASE_URL}{PARTICIPANT_PATH.format(institute_id=INSTITUTE_ID, participant_id=participant_id)}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        print(f"  [Identities] Skipped {participant_id}: {e}")
        return [], []

    phones, emails = _extract_identities(data)
    if use_cache:
        _cache_put("sp_participant", participant_id, {"phones": phones, "emails": emails})
    return phones, emails


def enrich_instructors_with_identities(rows: list, use_cache: bool = True) -> list:
    """For each instructor row, pull all identities from the participant resource
    and fill the primary + alternative phone/email columns. Best-effort."""
    if not ENRICH_PARTICIPANT_IDENTITIES:
        return rows

    print(f"[Enrich C] Fetching participant identities for {len(rows)} instructor(s) ...")
    enriched = 0
    for r in rows:
        pid = r.get("instructor_id")
        if not pid:
            continue
        phones, emails = _fetch_participant_identities(pid, use_cache=use_cache)
        if not phones and not emails:
            continue
        # Keep the teacher-list primary first, then merge participant identities.
        ph = _collect_contacts([r.get("primary_contact_number", "")], phones,
                               keys=("value", "number", "phoneNumber"))
        em = _collect_contacts([r.get("primary_email", "")], emails,
                               keys=("value", "email", "address"))
        # Keep the primary contact/email + up to INSTRUCTOR_*_SLOTS alternates so
        # no assigned faculty phone/email is dropped when a row has 3+ instructors.
        ph = (ph + [""] * (INSTRUCTOR_CONTACT_SLOTS + 1))[:INSTRUCTOR_CONTACT_SLOTS + 1]
        em = (em + [""] * (INSTRUCTOR_EMAIL_SLOTS + 1))[:INSTRUCTOR_EMAIL_SLOTS + 1]
        r["primary_contact_number"]       = ph[0]
        for _i in range(1, INSTRUCTOR_CONTACT_SLOTS + 1):
            r[f"alternative_contact_number_{_i}"] = ph[_i]
        r["primary_email"]                = _lower(em[0])
        for _i in range(1, INSTRUCTOR_EMAIL_SLOTS + 1):
            r[f"alternative_email_{_i}"]  = _lower(em[_i])
        enriched += 1

    print(f"[Enrich C] Identities merged for {enriched} instructor(s)")
    return rows


def _apply_contact_shift(row: dict) -> dict:
    """Insert DEFAULT_PRIMARY_CONTACT_NUMBER as primary_contact_number and push each
    of the record's original contact numbers down one alternate slot:
        default        -> primary_contact_number
        orig primary   -> alternative_contact_number_1
        orig alt_1     -> alternative_contact_number_2
        orig alt_2     -> alternative_contact_number_3
        orig alt_3     -> alternative_contact_number_4
        orig alt_4     -> alternative_contact_number_5
    The original alternative_contact_number_5 is dropped (only 5 alternate slots
    exist). Applied to every record just before it is inserted/updated. Mutates and
    returns the row dict."""
    # Originals to shift = primary + alt_1 .. alt_(N-1); the last alt is dropped.
    originals = [row.get("primary_contact_number", "")]
    originals += [row.get(f"alternative_contact_number_{i}", "")
                  for i in range(1, INSTRUCTOR_CONTACT_SLOTS)]
    row["primary_contact_number"] = DEFAULT_PRIMARY_CONTACT_NUMBER
    for i in range(1, INSTRUCTOR_CONTACT_SLOTS + 1):
        row[f"alternative_contact_number_{i}"] = originals[i - 1]
    return row


def run_pipeline_c(service, use_cache: bool = True) -> bool:
    """
    Instructors pipeline - loads the Instructor tab using SCD Type-2.
    Returns True on success, False on fetch failure (no writes performed).
    """
    print("\n" + "-" * 64)
    print("  Pipeline C  -  Instructors (SCD Type-2)")
    print("-" * 64)

    try:
        all_instructors = fetch_all_instructors(use_cache=use_cache)
    except RuntimeError as e:
        print(f"\n[Pipeline C] ABORTED - {e}")
        print("[Pipeline C] Sheet not modified. Fix the error and re-run.")
        return False

    instructor_rows = transform_instructors(all_instructors)
    instructor_rows = enrich_instructors_with_identities(instructor_rows, use_cache=use_cache)

    # Contact-number shift applied to EVERY record right before writing: the fixed
    # default becomes primary_contact_number and each original number moves down one
    # alternate slot (see _apply_contact_shift).
    for _r in instructor_rows:
        _apply_contact_shift(_r)

    upsert_scd2(
        service, SHEET_ID, INSTRUCTORS_TAB, INSTRUCTOR_COLUMNS,
        instructor_rows,
        business_key="instructor_id",
        compare_cols=INSTRUCTOR_COMPARE_COLS,
    )

    print(f"  Instructor source rows : {len(instructor_rows)}")
    return True


# -----------------------------------------------------------------------------
#  MAIN
# -----------------------------------------------------------------------------

def main():
    # -- Refresh mode (internal config; replaces the old --force-refresh CLI flag)
    use_cache = REFRESH_MODE != "force-refresh"

    # -- Cache setup ------------------------------------------------------------
    if not use_cache:
        _cache_clear_all()
    _ensure_cache_dirs()

    sep = "=" * 64
    print(f"\n{sep}")
    print("  IntelliBI Student Info Pipeline")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Cache   : {'OFF (--force-refresh)' if not use_cache else 'ON'}")
    print(f"  Sheet   : IntelliBIStudentInfo")
    print(f"  Tabs    : ClassLearnerTeacherEnrolled | Students | Payments | Instructor")
    print(sep)

    service = get_sheets_service(SERVICE_ACCOUNT_FILE)

    ok_a = run_pipeline_a(service, use_cache=use_cache)   # ClassLearnerTeacherEnrolled (sessions API)
    ok_b = run_pipeline_b(service, use_cache=use_cache)   # Students + Payments
    ok_c = run_pipeline_c(service, use_cache=use_cache)   # Instructors (SCD Type-2)

    print(f"\n{sep}")
    print("  Summary")
    print(f"  Pipeline A (ClassLearnerTeacherEnrolled) : {'COMPLETE' if ok_a else 'FAILED - no writes'}")
    print(f"  Pipeline B (Students / Payments)         : {'COMPLETE' if ok_b else 'FAILED - no writes'}")
    print(f"  Pipeline C (Instructor - SCD Type-2)     : {'COMPLETE' if ok_c else 'FAILED - no writes'}")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
