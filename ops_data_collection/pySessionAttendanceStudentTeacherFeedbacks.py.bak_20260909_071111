"""
================================================================================
  Session Attendance & Feedback Pipeline
  Converted from n8n workflow: wfIntelliBISessionAttendanceFeedbacks

  Run from PyCharm: Run > Run 'pySessionAttendanceFeedbacks'
================================================================================

SETUP:
  1. Place service_account.json in the same folder as this script.
  2. Share 'IntellBIAttendance' with the service account email (Editor access).
  3. Create a tab named 'Watermark_Attendance' in the sheet with these headers
     in row 1:  sync_key | load_type | last_sync_time | total_synced
     (The script will create it automatically on first run if the tab is empty.)
  4. pip install -r requirements.txt  (if not already done)

TARGET SHEET:  IntellBIAttendance
SHEET ID:      1TqDjq4gAyo32eRNMbuLd6uu0eCNZb7h1j5YH-q68AhU

TABS WRITTEN:
  Sessions          — one row per unique session
  Attendance        — one row per student per session
  Student_Feedback  — one row per student feedback per session
  Teacher_Feedback  — one row per teacher feedback per session
  Sessions_No_TF    — sessions that have no teacher feedback submitted
  Watermark_Attendance — internal watermark state (one row per sheet per run)

HOW INCREMENTAL LOAD WORKS:
  - First run  (no watermark): fetches all data from 2020-01-01, stores max
    watermark value per sheet.
  - Next runs  (has watermark): fetches data from the watermark date onward,
    filters to only sessions STRICTLY AFTER the stored watermark, appends only
    new rows. Watermark is updated to the new maximum seen this run.
  - Each sheet has its own independent watermark so they can be behind
    different amounts.
  - Data is fetched in 30-day chunks (API uses date-range pagination).

WRITE MODE:
  Pure append (no upsert). The incremental filter ensures no duplicates as
  long as the watermark sheet is intact.
================================================================================
"""

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import sys
import time
import json
import os
import hashlib
import argparse
import requests
from datetime import datetime, timezone, timedelta, date

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

from wise_config import HEADERS   # central headers; rotate API key in config_files/config.py only

INSTITUTE_ID         = "672a0918ae3d6e9fadfbc622"
BASE_URL             = "https://api.wiseapp.live"
SERVICE_ACCOUNT_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json"
)

SHEET_ID             = "1TqDjq4gAyo32eRNMbuLd6uu0eCNZb7h1j5YH-q68AhU"   # IntellBIAttendance
CHUNK_DAYS           = 30    # Split date range into 30-day windows per API call

# ── Force full refresh ────────────────────────────────────────────────────────
# Add tab names here to force a one-time full reload (clears existing data,
# ignores watermark). Remove the tab name after the run completes.
# Example: FORCE_FULL_LOAD_SHEETS = {"Attendance"}
#FORCE_FULL_LOAD_SHEETS = {"Attendance"} # Force Full Load
FORCE_FULL_LOAD_SHEETS = set()

# ─────────────────────────────────────────────────────────────────────────────
#  FILE-BASED API CACHE
#  Caches expensive per-entity API calls (suspended students, attendance detail)
#  to avoid redundant hits.  Use --force-refresh to bypass cache entirely.
#  NOTE: Main session fetch is NOT cached — watermark-based incremental logic
#        must stay intact.
# ─────────────────────────────────────────────────────────────────────────────

_SCRIPT_DIR_CACHE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR         = os.path.join(str(PROJECT_CACHE_DIR), "session_attendance")

TTL_SUSPENDED   = 12 * 3600    # 12 hours — suspension status changes infrequently
TTL_ATT_DETAIL  = 24 * 3600    # 24 hours — past session attendance won't change

# Refresh mode (internal config; replaces the old --force-refresh CLI flag).
#   "cache"         -> use the file cache (default).
#   "force-refresh" -> bypass the cache and fetch everything fresh from the API.
# NOTE: the separate --full-load CLI flag is unchanged.
REFRESH_MODE = "cache"


def _ensure_cache_dirs():
    """Create cache directory structure if missing."""
    for sub in ["", "suspended", "attendance_detail"]:
        path = os.path.join(CACHE_DIR, sub) if sub else CACHE_DIR
        os.makedirs(path, exist_ok=True)


def _cache_path(category: str, key: str = "") -> str:
    """Build the file path for a cache entry."""
    safe_key = hashlib.md5(key.encode()).hexdigest() if key else ""
    if category in ("suspended", "attendance_detail"):
        return os.path.join(CACHE_DIR, category, f"{safe_key or 'data'}.json")
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


# ── Tab names ─────────────────────────────────────────────────────────────────
SESSIONS_TAB         = "Sessions"
ATTENDANCE_TAB       = "Attendance"
STUDENT_FB_TAB       = "Student_Feedback"
TEACHER_FB_TAB       = "Teacher_Feedback"
SESSIONS_NO_TF_TAB   = "Sessions_No_TF"
WATERMARK_TAB        = "Watermark_Attendance"

WATERMARK_SHEETS     = ["Sessions", "Attendance", "Student_Feedback", "Teacher_Feedback"]

# ── Column definitions (order = sheet column order) ───────────────────────────
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

SESSIONS_COLUMNS = [
    "session_id", "course_name", "course_title", "tutor_name", "Instructor_Name",
    "start_time_ist", "end_time_ist", "synced_at",
]
ATTENDANCE_COLUMNS = [
    "session_id", "course_name", "course_title", "student_id", "student_name",
    "email", "tutor_name", "session_start_ist", "session_end_ist",
    "duration", "attendance_percent", "first_join_ist", "last_leave_ist",
    "status", "suspend_status", "synced_at",
]
STUDENT_FB_COLUMNS = [
    "session_id", "course_name", "course_title", "student_id", "student_name",
    "session_datetime", "session_start_ist", "session_end_ist",
    "rating", "comment", "created_at", "synced_at",
]
TEACHER_FB_COLUMNS = [
    "session_id", "course_name", "course_title", "teacher_id", "teacher_name", "Instructor_Name",
    "session_datetime", "session_start_ist", "session_end_ist",
    "topics_covered", "comments", "session_status", "created_at", "synced_at",
]
SESSIONS_NO_TF_COLUMNS = [
    "session_id", "course_name", "course_title", "tutor_name",
    "start_time_ist", "end_time_ist", "synced_at", "remark",
]
WATERMARK_COLUMNS = ["sync_key", "load_type", "last_sync_time", "total_synced"]

IST = timezone(timedelta(hours=5, minutes=30))


# ─────────────────────────────────────────────────────────────────────────────
#  IST HELPERS  (mirrors toIST / istDateOnly helpers in the n8n code node)
# ─────────────────────────────────────────────────────────────────────────────

def to_ist(date_str) -> str:
    """Convert a UTC ISO/timestamp string → 'YYYY-MM-DD HH:MM:SS' in IST."""
    if not date_str:
        return ""
    try:
        # Handle numeric epoch (milliseconds)
        if isinstance(date_str, (int, float)):
            dt = datetime.fromtimestamp(date_str / 1000, tz=timezone.utc)
        else:
            dt = datetime.fromisoformat(str(date_str).replace("Z", "+00:00"))
        return dt.astimezone(IST).strftime("%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError, OSError):
        return ""


def now_ist() -> str:
    """Current datetime as 'YYYY-MM-DD HH:MM:SS' in IST."""
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")


def ist_date_only(ist_str: str) -> str:
    """'YYYY-MM-DD HH:MM:SS' → 'YYYY-MM-DD'."""
    return str(ist_str)[:10] if ist_str else ""


def ist_str_to_date(ist_str: str):
    """Parse 'YYYY-MM-DD HH:MM:SS' IST string to a date object (IST date)."""
    try:
        return datetime.strptime(ist_str.strip(), "%Y-%m-%d %H:%M:%S").date()
    except (ValueError, AttributeError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
#  GOOGLE SHEETS CLIENT
# ─────────────────────────────────────────────────────────────────────────────

def get_sheets_service():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=scopes
        )
        return build("sheets", "v4", credentials=creds, cache_discovery=False)
    except FileNotFoundError:
        print(
            f"\n[ERROR] '{SERVICE_ACCOUNT_FILE}' not found.\n"
            "Place your Google service account JSON key in the same folder "
            "as this script.\n"
        )
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
#  SHARED UTILITIES
# ─────────────────────────────────────────────────────────────────────────────

def append_rows_with_retry(service, sheet_name: str, rows: list, columns: list, max_retries: int = 5):
    """
    Append rows to a sheet tab. Creates the header row on first run if the tab
    is empty. Retries on HTTP 429 with exponential backoff.
    """
    if not rows:
        print(f"[Write → {sheet_name}] No rows to append.")
        return

    # ── Ensure header exists ──────────────────────────────────────────────────
    try:
        result = (
            service.spreadsheets().values()
            .get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!A1:A1")
            .execute()
        )
        header_exists = bool(result.get("values"))
    except HttpError as e:
        print(f"[Write → {sheet_name}] Error checking header: {e}")
        return

    if not header_exists:
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="RAW",
            body={"values": [columns]},
        ).execute()
        print(f"[Write → {sheet_name}] Header row created.")

    # ── Build value matrix in column order ────────────────────────────────────
    value_matrix = [[str(row.get(col, "")) for col in columns] for row in rows]

    # ── Append in chunks to stay within Google Sheets cell limit ─────────────
    APPEND_CHUNK = 5000   # rows per API call (5000 × ~16 cols ≈ 80K cells)
    total_appended = 0

    for chunk_start in range(0, len(value_matrix), APPEND_CHUNK):
        chunk = value_matrix[chunk_start: chunk_start + APPEND_CHUNK]
        delay = 2
        for attempt in range(1, max_retries + 1):
            try:
                service.spreadsheets().values().append(
                    spreadsheetId=SHEET_ID,
                    range=f"{sheet_name}!A1",
                    valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS",
                    body={"values": chunk},
                ).execute()
                total_appended += len(chunk)
                print(
                    f"[Write → {sheet_name}] Appended chunk "
                    f"{chunk_start // APPEND_CHUNK + 1} — "
                    f"{total_appended}/{len(value_matrix)} rows"
                )
                break
            except HttpError as e:
                if e.resp.status == 429 and attempt < max_retries:
                    print(f"  [Rate limit] 429 — waiting {delay}s (attempt {attempt}/{max_retries}) ...")
                    time.sleep(delay)
                    delay *= 2
                else:
                    raise

    print(f"[Write → {sheet_name}] Done — {total_appended} rows appended in total.")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 1 — Read Watermark  (mirrors Read Watermark → Safe Watermark Read nodes)
#
#  Reads all rows from Watermark_Attendance tab.
#  For each SYNC_STATE_<Sheet>, picks the row with the highest last_sync_time.
#  Returns a dict: { "Sessions": "2024-03-01 10:00:00", "Attendance": None, ... }
# ─────────────────────────────────────────────────────────────────────────────

def read_watermarks(service) -> dict:
    """
    Reads the Watermark_Attendance tab and returns the latest last_sync_time
    per sheet as a dict. None means no watermark (first run = full load).
    """
    watermarks = {s: None for s in WATERMARK_SHEETS}

    try:
        result = (
            service.spreadsheets().values()
            .get(spreadsheetId=SHEET_ID, range=f"{WATERMARK_TAB}!A:D")
            .execute()
        )
        rows = result.get("values", [])
    except HttpError as e:
        # Tab may not exist yet on first run — that's fine
        print(f"[Watermark] Could not read watermark tab (first run?): {e}")
        return watermarks

    if len(rows) <= 1:
        print("[Watermark] No watermark data found — full load for all sheets.")
        return watermarks

    # Header: sync_key | load_type | last_sync_time | total_synced
    header = [h.strip() for h in rows[0]]
    try:
        key_idx = header.index("sync_key")
        val_idx = header.index("last_sync_time")
    except ValueError:
        print("[Watermark] Watermark tab header not recognised — full load.")
        return watermarks

    # Pick the latest last_sync_time per sync_key
    best = {}
    for row in rows[1:]:
        if len(row) <= max(key_idx, val_idx):
            continue
        key = str(row[key_idx]).strip()
        val = str(row[val_idx]).strip()
        if key and val and val not in ("", "last_sync_time", "null"):
            if key not in best or val > best[key]:
                best[key] = val

    for sheet in WATERMARK_SHEETS:
        sync_key = f"SYNC_STATE_{sheet}"
        watermarks[sheet] = best.get(sync_key)   # None if never synced

    # Override watermark to None for sheets that need a forced full refresh
    for sheet in FORCE_FULL_LOAD_SHEETS:
        if sheet in watermarks:
            watermarks[sheet] = None
            print(f"[Watermark] FORCE FULL LOAD override for: {sheet}")

    print("[Watermark] Loaded watermarks:")
    for s, w in watermarks.items():
        mode = f"INCREMENTAL filterAfter={w}" if w else "FULL LOAD"
        print(f"  {s}: {mode}")

    return watermarks


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2 — Init: build date chunks  (mirrors Init code node)
#
#  Splits the fetch window into 30-day chunks.
#  The global start date is the earliest watermark across all sheets
#  (or 2020-01-01 for a full load).
# ─────────────────────────────────────────────────────────────────────────────

def build_date_chunks(watermarks: dict) -> list[tuple[str, str]]:
    """
    Returns a list of (start_date, end_date) string tuples (YYYY-MM-DD).
    Each tuple covers at most CHUNK_DAYS days.
    The end of the last chunk is always tomorrow (to include today fully).
    """
    # Determine the earliest start date needed across all sheets
    start_dates = []
    for sheet, wm in watermarks.items():
        if wm:
            d = ist_str_to_date(wm)
            start_dates.append(d if d else date(2024, 11, 1))
        else:
            start_dates.append(date(2024, 11, 1))

    global_start = min(start_dates)
    today        = datetime.now(IST).date()
    tomorrow     = today + timedelta(days=1)

    chunks  = []
    cursor  = global_start
    while cursor <= today:
        chunk_end = min(cursor + timedelta(days=CHUNK_DAYS - 1), today)
        chunks.append((str(cursor), str(chunk_end)))
        cursor = chunk_end + timedelta(days=1)

    # Extend last chunk's end to tomorrow so today's data is always in range
    if chunks:
        chunks[-1] = (chunks[-1][0], str(tomorrow))

    print(
        f"[Init] Date range: {global_start} → {tomorrow} | "
        f"{len(chunks)} chunk(s) of up to {CHUNK_DAYS} days"
    )
    return chunks


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3 — Fetch sessions for all chunks  (mirrors Fetch Attendance HTTP node)
#
#  API: GET /institutes/{INSTITUTE_ID}/sessions
#  Params: paginateBy=DATE, showUnsharedRecording=true, showFeedbackData=true,
#          startDate=<start>, endDate=<end>
#
#  No page-by-page pagination needed — the API returns all sessions
#  within the date range in one response when paginateBy=DATE is set.
# ─────────────────────────────────────────────────────────────────────────────

def fetch_sessions_for_chunk(start_date: str, end_date: str) -> list:
    """Fetch all sessions in the given date range (one API call per chunk).

    The API treats endDate as EXCLUSIVE, so we extend it by 1 day to ensure
    the last day of each chunk is always included.  Sessions on the overlap day
    are deduplicated later in transform() via the seen-set.
    """
    from datetime import date as _date, timedelta as _td
    end_inclusive = str(_date.fromisoformat(end_date) + _td(days=1))

    url = f"{BASE_URL}/institutes/{INSTITUTE_ID}/sessions"
    params = {
        "paginateBy":            "DATE",
        "showUnsharedRecording": "true",
        "showFeedbackData":      "true",
        "showAttendance":        "true",
        "includeParticipants":   "true",
        "startDate":             start_date,
        "endDate":               end_inclusive,   # +1 day so last day is included
    }
    headers = HEADERS

    try:
        resp = requests.get(url, params=params, headers=headers, timeout=60)
        data = resp.json()
    except Exception as e:
        print(f"  [Fetch] Request error ({start_date} → {end_date}): {e}")
        return []

    sessions = (
        (data.get("data") or {}).get("sessions")
        or data.get("sessions")
        or data.get("data")
        or data.get("result")
        or data.get("response")
        or (data if isinstance(data, list) else [])
    )
    if not isinstance(sessions, list):
        sessions = []

    return sessions


def fetch_all_sessions(chunks: list[tuple[str, str]]) -> list:
    """Fetch sessions for all date chunks and return combined flat list."""
    all_sessions = []
    for i, (start, end) in enumerate(chunks, 1):
        chunk_sessions = fetch_sessions_for_chunk(start, end)
        all_sessions.extend(chunk_sessions)
        print(
            f"[Fetch] Chunk {i}/{len(chunks)} | {start} → {end} | "
            f"sessions_this_chunk={len(chunk_sessions)} total_so_far={len(all_sessions)}"
        )

    print(f"[Fetch] All chunks done. Total raw sessions: {len(all_sessions)}")
    return all_sessions


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3b — Participant extraction helpers
# ─────────────────────────────────────────────────────────────────────────────

def _extract_participant_id(p: dict) -> str:
    """
    Robustly extract a student/participant ID.
    Different API versions use different field names; check all known ones.
    """
    return (
        p.get("wiseUserId") or p.get("userId") or p.get("studentId") or
        p.get("_id")        or p.get("id")     or ""
    )


def _is_teacher_participant(p: dict, session_instructor_id: str) -> bool:
    """
    Return True if this participant is an instructor and should be excluded
    from the student attendance list.
    Checks: isTeacher/isInstructor flags, role string, userType string, ID match.
    """
    if p.get("isTeacher") or p.get("isInstructor"):
        return True
    role = str(p.get("role") or p.get("userRole") or "").strip().lower()
    if role in ("teacher", "instructor", "tutor", "host"):
        return True
    user_type = str(p.get("userType") or p.get("type") or "").strip().lower()
    if user_type in ("teacher", "instructor", "tutor", "host"):
        return True
    pid = _extract_participant_id(p)
    if session_instructor_id and pid and pid == session_instructor_id:
        return True
    return False


def _extract_participants_from_any_field(session: dict) -> list:
    """
    Check every known field name that the API might use to store participants.
    Returns the first non-empty list found.
    """
    for field in (
        "participants", "students", "attendees", "attendanceData",
        "participantsList", "attendanceList", "enrolledStudents",
        "sessionStudents", "sessionParticipants", "studentList",
        "sessionAttendance", "attendanceRecords",
    ):
        val = session.get(field)
        if isinstance(val, list) and val:
            return val
    return []


def fetch_session_attendance_detail(sid: str, use_cache: bool = True) -> list:
    """
    Fallback: hit a dedicated per-session attendance endpoint when the session
    list response contains no participant data.
    Tries four URL patterns; returns the first non-empty participant list.
    """
    # ── Check cache first ────────────────────────────────────────────────────
    if use_cache:
        cached = _cache_get("attendance_detail", sid, TTL_ATT_DETAIL)
        if cached is not None:
            if cached:
                print(f"    [AttendanceFallback] {sid} → {len(cached)} participant(s) (from cache)")
            return cached

    headers = HEADERS
    endpoints = [
        f"{BASE_URL}/institutes/{INSTITUTE_ID}/sessions/{sid}/attendance",
        f"{BASE_URL}/institutes/{INSTITUTE_ID}/sessions/{sid}/participants",
        f"{BASE_URL}/sessions/{sid}/attendance",
        f"{BASE_URL}/sessions/{sid}/participants",
    ]
    for url in endpoints:
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code == 404:
                continue
            resp.raise_for_status()
            data = resp.json()
            participants = (
                (data.get("data") or {}).get("participants")
                or (data.get("data") or {}).get("students")
                or (data.get("data") or {}).get("attendees")
                or (data.get("data") or {}).get("attendanceData")
                or data.get("participants")
                or data.get("students")
                or data.get("attendees")
                or data.get("attendanceData")
                or data.get("data")
                or (data if isinstance(data, list) else [])
            )
            if isinstance(participants, list) and participants:
                print(f"    [AttendanceFallback] {sid} → {url} → {len(participants)} participant(s)")
                if use_cache:
                    _cache_put("attendance_detail", sid, participants)
                return participants
        except Exception as e:
            print(f"    [AttendanceFallback] {sid} {url} error: {e}")

    # Cache empty result too so we don't retry failed lookups every run
    if use_cache:
        _cache_put("attendance_detail", sid, [])
    return []


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3c — Fetch suspended students per class (participants API)
# ─────────────────────────────────────────────────────────────────────────────

def fetch_suspended_students(class_ids: set, use_cache: bool = True) -> dict:
    """
    For each class, call GET /teacher/v2/classes/{class_id}/participants
    and read the suspended student IDs from data.classroom.suspendedStudents.
    (Suspended students are NOT listed in the participants array — they are
    stored separately at the classroom level.)
    Only page 1 is needed since suspendedStudents is on the classroom object.

    Returns a dict:  { class_id: set(student_id, ...) }
    Suspension is per-class — a student can be suspended in one class but
    active in another.
    """
    suspended_by_class = {}   # class_id → set of suspended student IDs
    headers = HEADERS
    total_classes = len(class_ids)
    total_suspended = 0
    cache_hits = 0
    print(f"[Suspend] Fetching suspended students for {total_classes} class(es) …")

    for idx, cid in enumerate(sorted(class_ids), 1):
        # ── Check cache first ────────────────────────────────────────────────
        if use_cache:
            cached = _cache_get("suspended", cid, TTL_SUSPENDED)
            if cached is not None:
                cache_hits += 1
                if cached:  # non-empty list
                    suspended_by_class[cid] = set(cached)
                    total_suspended += len(cached)
                if idx % 10 == 0 or idx == total_classes:
                    print(f"  [Suspend] {idx}/{total_classes} classes checked | "
                          f"suspended so far: {total_suspended} | cache hits: {cache_hits}")
                continue

        try:
            resp = requests.get(
                f"{BASE_URL}/teacher/v2/classes/{cid}/participants",
                params={"page_number": 1, "page_size": 1},
                headers=headers, timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            print(f"  [Suspend] Class {cid} error: {e}")
            continue

        # Suspended students are listed under data.classroom.suspendedStudents
        classroom = (data.get("data") or {}).get("classroom") or {}
        suspended_list = classroom.get("suspendedStudents") or []

        class_suspended = set()
        if isinstance(suspended_list, list):
            for entry in suspended_list:
                # Entries can be plain string IDs or dicts with _id/userId
                if isinstance(entry, str) and entry:
                    class_suspended.add(entry)
                elif isinstance(entry, dict):
                    sid = (
                        entry.get("_id") or entry.get("userId") or
                        entry.get("wiseUserId") or entry.get("id") or ""
                    )
                    if sid:
                        class_suspended.add(sid)

        # Cache the result (store as list for JSON serialization)
        if use_cache:
            _cache_put("suspended", cid, list(class_suspended) if class_suspended else [])

        if class_suspended:
            suspended_by_class[cid] = class_suspended
            total_suspended += len(class_suspended)

        if idx % 10 == 0 or idx == total_classes:
            print(f"  [Suspend] {idx}/{total_classes} classes checked | "
                  f"suspended so far: {total_suspended} | cache hits: {cache_hits}")

    print(f"[Suspend] Done — {total_suspended} suspended student(s) across "
          f"{len(suspended_by_class)} class(es) (of {total_classes} total) | "
          f"cache hits: {cache_hits}/{total_classes}")
    return suspended_by_class


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 4 — Transform  (mirrors Transform Attendance code node)
#
#  Per-sheet incremental logic (mirrors isNew() in the n8n code):
#    Full load  (no watermark)  → pass all sessions
#    Incremental (has watermark) → only sessions with start_time_ist STRICTLY
#                                  after the stored watermark (lexicographic
#                                  comparison on 'YYYY-MM-DD HH:MM:SS')
# ─────────────────────────────────────────────────────────────────────────────

def is_new(s_ist: str, sheet: str, watermarks: dict) -> bool:
    """
    Mirrors the isNew() helper in the n8n Transform code node.
    Returns True if this session should be included for the given sheet.
      - No watermark (full load) → always True
      - Has watermark            → only if s_ist > watermark (strict)
    """
    wm = watermarks.get(sheet)
    if not wm:
        return True          # full load: pass all
    if not s_ist:
        return False
    return s_ist > wm        # lexicographic works for 'YYYY-MM-DD HH:MM:SS'


def transform(all_sessions: list, watermarks: dict, synced_at: str,
              suspended_by_class: dict = None, use_cache: bool = True) -> dict:
    """
    Transform raw session objects into rows for each output tab.
    Returns a dict:
      {
        "Sessions":        [...],
        "Attendance":      [...],
        "Student_Feedback": [...],
        "Teacher_Feedback": [...],
        "Sessions_No_TF":  [...],
      }
    Also returns max_ist: the highest start_time_ist seen per sheet this run.

    suspended_by_class: { class_id: set(student_id, ...) }
      — per-class suspension lookup.  A student marked Suspended in one class
        may be Active in another.  Students who are currently suspended but
        attended a session (Present) are treated as Active for that session
        (they were clearly un-suspended on that day).
    """
    if suspended_by_class is None:
        suspended_by_class = {}
    seen = {
        "s":    set(),   # session dedup
        "a":    set(),   # attendance dedup (session_id + student_id)
        "sf":   set(),   # student feedback dedup
        "tf":   set(),   # teacher feedback dedup
        "notf": set(),   # sessions_no_tf dedup
    }

    output = {
        "Sessions":        [],
        "Attendance":      [],
        "Student_Feedback": [],
        "Teacher_Feedback": [],
        "Sessions_No_TF":  [],
    }
    max_ist   = {s: "" for s in WATERMARK_SHEETS}
    counts    = {s: 0  for s in WATERMARK_SHEETS}

    print(f"[Transform] {len(all_sessions)} raw sessions to process")

    for session in all_sessions:
        if not session or not isinstance(session, dict):
            continue

        sid         = session.get("_id")        or session.get("id")        or ""
        s_utc       = session.get("start_time") or session.get("startTime") or session.get("startDate") or ""
        # NOTE: do NOT use "endDate" — the API returns the record's last-modified
        # timestamp in that field (wrong).  Use only genuine end-time fields.
        e_utc       = (
            session.get("end_time")        or
            session.get("endTime")         or
            session.get("completedAt")     or
            session.get("closedAt")        or
            session.get("sessionEndTime")  or
            session.get("meetingEndTime")  or
            ""
        )
        s_ist       = to_ist(s_utc)
        e_ist       = to_ist(e_utc)
        # Sanity check: if e_ist is before or equal to s_ist it is wrong → discard
        if e_ist and s_ist and e_ist <= s_ist:
            e_ist = ""
        s_date_only = ist_date_only(s_ist)

        # Skip sessions whose start date is in the future — these are pre-scheduled
        # sessions that the API returns when queried with endDate = tomorrow+1.
        # They should only be synced once they have actually happened.
        _today_ist = datetime.now(IST).strftime("%Y-%m-%d")
        if s_date_only and s_date_only > _today_ist:
            continue

        class_id_raw = session.get("classId") or {}
        if isinstance(class_id_raw, str):
            class_id_str = class_id_raw
            class_id_raw = {}
        else:
            class_id_str = class_id_raw.get("_id") or ""
        class_id     = class_id_raw   # keep original var name for downstream code
        course_name  = class_id.get("name")    or session.get("className")    or ""
        course_title = class_id.get("subject") or session.get("courseTitle")  or ""
        instructor_name = _fetch_class_instructor_map().get(class_id_str, "")

        user_id = session.get("userId") or {}
        if isinstance(user_id, str):
            # userId is sometimes a plain string ID, not a nested object
            session_instructor_id = user_id
            user_id = {}
        else:
            session_instructor_id = (
                user_id.get("_id") or user_id.get("id") or
                user_id.get("wiseUserId") or ""
            )
        tutor_name = user_id.get("name") or session.get("teacherName") or ""

        has_tf      = bool(session.get("teacherFeedback") or session.get("tutorFeedback"))

        # ── SESSIONS ──────────────────────────────────────────────────────────
        if sid and sid not in seen["s"] and is_new(s_ist, "Sessions", watermarks):
            seen["s"].add(sid)
            counts["Sessions"] += 1
            if s_ist > max_ist["Sessions"]:
                max_ist["Sessions"] = s_ist
            output["Sessions"].append({
                "session_id":     sid,
                "course_name":    course_name,
                "course_title":   course_title,
                "tutor_name":     tutor_name,
                "Instructor_Name": instructor_name,
                "start_time_ist": s_ist,
                "end_time_ist":   e_ist,
                "synced_at":      synced_at,
            })

        # ── ATTENDANCE ────────────────────────────────────────────────────────
        # Check 12 possible field names; fall back to per-session API if needed
        participants = _extract_participants_from_any_field(session)
        if not participants and sid:
            participants = fetch_session_attendance_detail(sid, use_cache=use_cache)
        print(
            f"  [Session] {sid} | {course_name} | "
            f"{s_ist[:10] if s_ist else '?'} | participants={len(participants)}"
        )
        for p in participants:
            if not p or not isinstance(p, dict):
                continue
            # Skip instructor records — 4-layer check including direct ID match
            if _is_teacher_participant(p, session_instructor_id):
                continue
            stid = _extract_participant_id(p)
            key  = f"{sid}_{stid}"
            if stid and key not in seen["a"] and is_new(s_ist, "Attendance", watermarks):
                seen["a"].add(key)
                counts["Attendance"] += 1
                if s_ist > max_ist["Attendance"]:
                    max_ist["Attendance"] = s_ist

                raw_pct = (
                    p.get("absolutePercentAttendance") or
                    p.get("attendancePercent") or 0
                )
                try:
                    pct_float = float(raw_pct)
                    pct_str   = f"{pct_float:.2f}%"
                except (ValueError, TypeError):
                    pct_float = 0.0
                    pct_str   = "0.00%"

                in_duration = p.get("inMeetingDuration") or p.get("duration") or 0
                is_present  = (in_duration > 0) or (pct_float > 0)

                # Determine suspend status:
                #   1. Per-class: only check if student is suspended in THIS class
                #   2. If student actually attended (Present), they were clearly
                #      active on that day — mark Active regardless of current
                #      suspension status (coordinator may have reverted).
                _class_susp = suspended_by_class.get(class_id_str, set())
                p_suspended = (stid in _class_susp) and not is_present

                output["Attendance"].append({
                    "session_id":         sid,
                    "course_name":        course_name,
                    "course_title":       course_title,
                    "student_id":         stid,
                    "student_name":       p.get("name") or "",
                    "email":              (
                        p.get("wiseUserEmail") or p.get("user_email") or
                        p.get("userEmail")     or p.get("email")      or
                        p.get("emailId")       or ""
                    ),
                    "tutor_name":         tutor_name,
                    "session_start_ist":  s_ist,
                    "session_end_ist":    e_ist,
                    "duration":           in_duration,
                    "attendance_percent": pct_str,
                    "first_join_ist":     to_ist(p.get("firstEntryTime") or p.get("joinTime")),
                    "last_leave_ist":     to_ist(p.get("lastExitTime")   or p.get("leaveTime")),
                    "status":             "Present" if is_present else "Absent",
                    "suspend_status":     "Suspended" if p_suspended else "Active",
                    "synced_at":          synced_at,
                })

        # ── STUDENT FEEDBACK ──────────────────────────────────────────────────
        student_subs = (
            session.get("studentSubmissions") or
            session.get("studentFeedback")    or
            []
        )
        for f in student_subs:
            if not f or not isinstance(f, dict):
                continue
            stid = f.get("userId") or f.get("studentId") or ""
            fkey = f"{sid}_{stid}"
            if stid and fkey not in seen["sf"] and is_new(s_ist, "Student_Feedback", watermarks):
                seen["sf"].add(fkey)
                counts["Student_Feedback"] += 1
                if s_ist > max_ist["Student_Feedback"]:
                    max_ist["Student_Feedback"] = s_ist

                # Find matching participant for the student name
                part = next(
                    (p for p in participants
                     if isinstance(p, dict) and (p.get("wiseUserId") or p.get("userId")) == stid),
                    {}
                )
                output["Student_Feedback"].append({
                    "session_id":        sid,
                    "course_name":       course_name,
                    "course_title":      course_title,
                    "student_id":        stid,
                    "student_name":      part.get("name") or "",
                    "session_datetime":  s_date_only,
                    "session_start_ist": s_ist,
                    "session_end_ist":   e_ist,
                    "rating":            f.get("rating")             or "",
                    "comment":           f.get("comment") or f.get("feedback") or "",
                    "created_at":        to_ist(f.get("createdAt"))  or "",
                    "synced_at":         synced_at,
                })

        # ── TEACHER FEEDBACK ──────────────────────────────────────────────────
        tfb  = session.get("teacherFeedback") or session.get("tutorFeedback")
        tkey = f"{sid}_teacher"
        if tfb and isinstance(tfb, dict) and tkey not in seen["tf"] and is_new(s_ist, "Teacher_Feedback", watermarks):
            seen["tf"].add(tkey)
            counts["Teacher_Feedback"] += 1
            if s_ist > max_ist["Teacher_Feedback"]:
                max_ist["Teacher_Feedback"] = s_ist

            answers = tfb.get("answers") or []
            topics_covered = next(
                (a.get("answer") for a in answers if isinstance(a, dict) and a.get("questionText") == "Topics covered"),
                tfb.get("topicsCovered") or ""
            )
            fb_comments = next(
                (a.get("answer") for a in answers if isinstance(a, dict) and a.get("questionText") == "Comments"),
                tfb.get("comments") or ""
            )
            output["Teacher_Feedback"].append({
                "session_id":        sid,
                "course_name":       course_name,
                "course_title":      course_title,
                "teacher_id":        tfb.get("userId") or "",
                "teacher_name":      tutor_name,
                "Instructor_Name":   instructor_name,
                "session_datetime":  s_date_only,
                "session_start_ist": s_ist,
                "session_end_ist":   e_ist,
                "topics_covered":    topics_covered,
                "comments":          fb_comments,
                "session_status":    tfb.get("sessionStatus") or "",
                "created_at":        to_ist(tfb.get("createdAt")) or "",
                "synced_at":         synced_at,
            })

        # ── SESSIONS WITHOUT TEACHER FEEDBACK ─────────────────────────────────
        if sid and sid not in seen["notf"] and is_new(s_ist, "Sessions", watermarks) and not has_tf:
            seen["notf"].add(sid)
            output["Sessions_No_TF"].append({
                "session_id":     sid,
                "course_name":    course_name,
                "course_title":   course_title,
                "tutor_name":     tutor_name,
                "start_time_ist": s_ist,
                "end_time_ist":   e_ist,
                "synced_at":      synced_at,
                "remark":         "Teacher feedback not submitted",
            })

    print(
        f"[Transform] Sessions: {counts['Sessions']} | "
        f"Attendance: {counts['Attendance']} | "
        f"Student_Feedback: {counts['Student_Feedback']} | "
        f"Teacher_Feedback: {counts['Teacher_Feedback']} | "
        f"Sessions_No_TF: {len(output['Sessions_No_TF'])}"
    )
    return output, counts, max_ist


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5 — Write data to sheet tabs
# ─────────────────────────────────────────────────────────────────────────────

def write_all_tabs(service, transformed: dict):
    """Append rows to each data tab."""
    tab_map = [
        (SESSIONS_TAB,       SESSIONS_COLUMNS,       "Sessions"),
        (ATTENDANCE_TAB,     ATTENDANCE_COLUMNS,     "Attendance"),
        (STUDENT_FB_TAB,     STUDENT_FB_COLUMNS,     "Student_Feedback"),
        (TEACHER_FB_TAB,     TEACHER_FB_COLUMNS,     "Teacher_Feedback"),
        (SESSIONS_NO_TF_TAB, SESSIONS_NO_TF_COLUMNS, "Sessions_No_TF"),
    ]
    for tab_name, columns, key in tab_map:
        rows = transformed.get(key, [])

        # If this tab is marked for forced full refresh, DELETE and recreate
        # the tab to truly free all cells (clear() only removes values but
        # Google Sheets still counts empty rows toward the 10M cell limit).
        if key in FORCE_FULL_LOAD_SHEETS and rows:
            print(f"[Write → {tab_name}] FORCE FULL LOAD — deleting and recreating tab …")
            try:
                # Find the sheet ID for this tab
                meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
                sheet_id = None
                for s in meta.get("sheets", []):
                    if s["properties"]["title"] == tab_name:
                        sheet_id = s["properties"]["sheetId"]
                        break
                if sheet_id is not None:
                    # Delete the tab
                    service.spreadsheets().batchUpdate(
                        spreadsheetId=SHEET_ID,
                        body={"requests": [{"deleteSheet": {"sheetId": sheet_id}}]},
                    ).execute()
                    print(f"[Write → {tab_name}] Old tab deleted.")
                # Recreate the tab
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SHEET_ID,
                    body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
                ).execute()
                print(f"[Write → {tab_name}] New tab created.")
            except HttpError as e:
                print(f"[Write → {tab_name}] Error recreating tab: {e}")

        if rows:
            append_rows_with_retry(service, tab_name, rows, columns)
        else:
            print(f"[Write → {tab_name}] 0 new rows — nothing to append.")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5b — Backfill blank end_time_ist in Sessions from Attendance
#
#  Because the pipeline is append-only with watermarks, sessions synced while
#  still ongoing get blank end_time_ist that is never updated. The Attendance
#  sheet stores session_end_ist per participant row.  This step reads both
#  sheets, finds Sessions rows with blank end_time_ist, looks up a non-blank
#  session_end_ist from Attendance for the same session_id, and batch-updates
#  the Sessions sheet.
# ─────────────────────────────────────────────────────────────────────────────

def backfill_session_end_times(service):
    """Patch blank end_time_ist in Sessions using session_end_ist from Attendance."""
    print("\n[Backfill] Checking for blank end_time_ist in Sessions …")
    try:
        # Read Sessions header + data
        sess_result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{SESSIONS_TAB}!A1:Z",
        ).execute()
        sess_rows = sess_result.get("values", [])
        if len(sess_rows) < 2:
            print("[Backfill] Sessions sheet empty — skipping.")
            return

        sess_header = [h.strip().lower() for h in sess_rows[0]]
        sid_col = sess_header.index("session_id") if "session_id" in sess_header else None
        end_col = sess_header.index("end_time_ist") if "end_time_ist" in sess_header else None
        if sid_col is None or end_col is None:
            print("[Backfill] Required columns not found in Sessions — skipping.")
            return

        # Collect session_ids with blank end_time_ist and their row numbers (1-based)
        blank_sessions = {}  # session_id → sheet row number (1-based, header=row 1)
        for i, row in enumerate(sess_rows[1:], start=2):  # row 2 onward
            sid_val = row[sid_col].strip() if len(row) > sid_col else ""
            end_val = row[end_col].strip() if len(row) > end_col else ""
            if sid_val and end_val in ("", "nan", "NaT", "None", "NAN"):
                blank_sessions[sid_val] = i

        if not blank_sessions:
            print("[Backfill] No blank end_time_ist found — nothing to backfill.")
            return

        print(f"[Backfill] Found {len(blank_sessions)} session(s) with blank end_time_ist.")

        # Read Attendance to get session_end_ist
        att_result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{ATTENDANCE_TAB}!A1:Z",
        ).execute()
        att_rows = att_result.get("values", [])
        if len(att_rows) < 2:
            print("[Backfill] Attendance sheet empty — skipping.")
            return

        att_header = [h.strip().lower() for h in att_rows[0]]
        att_sid_col = att_header.index("session_id") if "session_id" in att_header else None
        att_end_col = att_header.index("session_end_ist") if "session_end_ist" in att_header else None
        if att_sid_col is None or att_end_col is None:
            print("[Backfill] Required columns not found in Attendance — skipping.")
            return

        # Build lookup: session_id → first non-blank session_end_ist
        att_end_map = {}
        for row in att_rows[1:]:
            a_sid = row[att_sid_col].strip() if len(row) > att_sid_col else ""
            a_end = row[att_end_col].strip() if len(row) > att_end_col else ""
            if a_sid and a_end and a_end not in ("", "nan", "NaT", "None", "NAN"):
                if a_sid not in att_end_map:
                    att_end_map[a_sid] = a_end

        # Build batch update for blank sessions that have a match in Attendance
        updates = []
        end_col_letter = chr(ord('A') + end_col) if end_col < 26 else None
        if end_col_letter is None:
            print("[Backfill] end_time_ist column too far right — skipping.")
            return

        for sid_val, row_num in blank_sessions.items():
            att_end = att_end_map.get(sid_val)
            if att_end:
                cell_ref = f"{SESSIONS_TAB}!{end_col_letter}{row_num}"
                updates.append({
                    "range": cell_ref,
                    "values": [[att_end]],
                })

        if not updates:
            print("[Backfill] No matching session_end_ist found in Attendance — nothing to update.")
            return

        # Batch update
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={
                "valueInputOption": "RAW",
                "data": updates,
            },
        ).execute()
        print(f"[Backfill] ✓ Updated end_time_ist for {len(updates)} session(s).")

    except Exception as e:
        print(f"[Backfill] ⚠ Error during backfill: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 6 — Write Watermark  (mirrors Prepare Watermark → Write Watermark nodes)
#
#  Watermark strategy (mirrors the n8n Prepare Watermark code node exactly):
#    Full load   → always write watermark (sets baseline)
#    Incremental → write ONLY if new records were found (count > 0)
#                  If nothing new, preserve old watermark so next run
#                  re-checks the same window.
# ─────────────────────────────────────────────────────────────────────────────

def write_watermarks(service, watermarks: dict, counts: dict, max_ist: dict):
    """
    Appends new watermark rows to Watermark_Attendance tab.
    One row per sheet where a watermark update is warranted.
    """
    rows_to_write = []

    for sheet in WATERMARK_SHEETS:
        wm         = watermarks.get(sheet)
        is_incr    = bool(wm)
        count      = counts.get(sheet, 0)
        new_max    = max_ist.get(sheet, "")

        # Write watermark only when:
        #   Full load   → always (establishes baseline even with 0 records)
        #   Incremental → only if new records were loaded this run
        should_write = (not is_incr) or (count > 0)

        if should_write and new_max:
            rows_to_write.append({
                "sync_key":       f"SYNC_STATE_{sheet}",
                "load_type":      "INCREMENTAL" if is_incr else "FULL",
                "last_sync_time": new_max,
                "total_synced":   str(count),
            })
            print(
                f"[Watermark] WRITE | {sheet} | "
                f"{'INC' if is_incr else 'FULL'} | "
                f"count={count} | max={new_max}"
            )
        else:
            reason = "count=0, watermark preserved" if is_incr else "no max IST found"
            print(f"[Watermark] SKIP  | {sheet} | {reason}")

    if rows_to_write:
        append_rows_with_retry(service, WATERMARK_TAB, rows_to_write, WATERMARK_COLUMNS)
    else:
        print("[Watermark] No watermark rows to write.")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # ── Parse CLI arguments ───────────────────────────────────────────────────
    # Refresh mode is set via the REFRESH_MODE config variable (top of file),
    # replacing the old --force-refresh CLI flag. --full-load remains a CLI flag.
    parser = argparse.ArgumentParser(description="Session Attendance & Feedback Pipeline")
    parser.add_argument(
        "--full-load", action="store_true",
        help="Ignore watermarks and do a full reload of all sheets (deletes & recreates tabs)",
    )
    args = parser.parse_args()
    use_cache = REFRESH_MODE != "force-refresh"

    # --full-load → force full refresh on all data sheets
    if args.full_load:
        FORCE_FULL_LOAD_SHEETS.update(WATERMARK_SHEETS)
        use_cache = False  # also bypass file cache

    # ── Cache setup ───────────────────────────────────────────────────────────
    if not use_cache or args.full_load:
        _cache_clear_all()
    _ensure_cache_dirs()

    separator = "=" * 64
    print(f"\n{separator}")
    print("  Session Attendance & Feedback Pipeline")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    _mode = "FULL LOAD (--full-load)" if args.full_load else ("OFF (--force-refresh)" if not use_cache else "ON")
    print(f"  Cache   : {_mode}")
    print(f"  Sheet   : IntellBIAttendance")
    print(f"  Tabs    : Sessions | Attendance | Student_Feedback")
    print(f"          : Teacher_Feedback | Sessions_No_TF")
    print(f"{separator}\n")

    synced_at = now_ist()
    service   = get_sheets_service()

    # ── Step 1: Read watermarks ────────────────────────────────────────────────
    watermarks = read_watermarks(service)

    # ── Step 2: Build date chunks ──────────────────────────────────────────────
    chunks = build_date_chunks(watermarks)

    # ── Step 3: Fetch all sessions ─────────────────────────────────────────────
    all_sessions = fetch_all_sessions(chunks)

    if not all_sessions:
        print("\n[Main] No sessions returned by API. Nothing to write.")
        print(f"\n{separator}")
        print("  Pipeline complete (0 sessions fetched).")
        print(f"{separator}\n")
        return

    # ── Step 3b: Collect class IDs and fetch suspended students ──────────────
    class_ids = set()
    for session in all_sessions:
        cls_obj = session.get("classId") or {}
        if isinstance(cls_obj, dict):
            cid = cls_obj.get("_id") or ""
        elif isinstance(cls_obj, str):
            cid = cls_obj
        else:
            cid = ""
        if cid:
            class_ids.add(cid)
    suspended_by_class = fetch_suspended_students(class_ids, use_cache=use_cache)

    # ── Step 4: Transform ─────────────────────────────────────────────────────
    transformed, counts, max_ist = transform(all_sessions, watermarks, synced_at,
                                             suspended_by_class, use_cache=use_cache)

    # ── Step 5: Write data tabs ────────────────────────────────────────────────
    write_all_tabs(service, transformed)

    # ── Step 5b: Backfill blank end_time_ist from Attendance ──────────────────
    backfill_session_end_times(service)

    # ── Step 6: Write watermarks ───────────────────────────────────────────────
    write_watermarks(service, watermarks, counts, max_ist)

    total_written = sum(len(v) for v in transformed.values())
    print(f"\n{separator}")
    print(f"  Pipeline complete.")
    print(f"  Sessions written        : {counts.get('Sessions', 0)}")
    print(f"  Attendance written      : {counts.get('Attendance', 0)}")
    print(f"  Student Feedback written: {counts.get('Student_Feedback', 0)}")
    print(f"  Teacher Feedback written: {counts.get('Teacher_Feedback', 0)}")
    print(f"  Sessions_No_TF written  : {len(transformed.get('Sessions_No_TF', []))}")
    print(f"{separator}\n")


if __name__ == "__main__":
    main()
