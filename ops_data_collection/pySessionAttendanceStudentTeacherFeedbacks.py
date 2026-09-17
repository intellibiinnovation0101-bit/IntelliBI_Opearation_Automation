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
# Step 5c late-attendance backfill: how many recent days to re-check FRESH for
# attendance/feedback that finalised AFTER a session was first synced (see the
# backfill_recent_attendance() header). Recovers such rows without moving the
# watermark. Generic; no technology/date/session is hard-coded.
BACKFILL_LOOKBACK_DAYS = 7

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
REFRESH_MODE = "force-refresh"

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
    "start_time_ist", "end_time_ist",
    "Session Scheduled Start", "Session Scheduled End",
    "synced_at",
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

# Candidate API field names for the SCHEDULED session slot (planned start/end).
# The wiseapp session object exposes the scheduled slot under one of these names;
# the first present, non-empty value wins. Kept SEPARATE from the actual/conducted
# start_time_ist / end_time_ist. Actual-only end fields (completedAt / closedAt /
# meetingEndTime) are deliberately NOT used for the scheduled end. If your API uses
# a different key, add it to the front of the relevant list.
_SCHED_START_FIELDS = ["scheduledStartTime", "scheduledStart", "scheduled_start_time",
                       "scheduledStartDate", "plannedStartTime", "sessionStartTime",
                       "startTime", "start_time", "startDate"]
_SCHED_END_FIELDS   = ["scheduledEndTime", "scheduledEnd", "scheduled_end_time",
                       "scheduledEndDate", "plannedEndTime", "sessionEndTime",
                       "endTime", "end_time"]


def _first_present(d, keys):
    """Return the first non-empty value among `keys` in dict `d`, else ''."""
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return ""


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

def _col_letter(idx0: int) -> str:
    """0-based column index → spreadsheet letter(s). 0→A, 25→Z, 26→AA …"""
    s = ""
    n = idx0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _migrate_sheet_layout(service, sheet_name, old_header, new_columns):
    """Re-align an existing tab from `old_header` to `new_columns` BY COLUMN NAME
    and rewrite the header + all rows, so adding/inserting a column never leaves
    historical rows shifted. Columns new in `new_columns` become blank; dropped
    columns are removed. Mirrors the shared upsert helper; only called when the
    header actually differs (a one-time migration)."""
    last_col = _col_letter(max(len(old_header), len(new_columns)) - 1)
    res = (
        service.spreadsheets().values()
        .get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!A1:{last_col}")
        .execute()
    )
    allv = res.get("values", [])
    data_rows = allv[1:] if len(allv) > 1 else []
    pos = {name: i for i, name in enumerate(old_header)}
    realigned = [[(r[pos[c]] if (c in pos and pos[c] < len(r)) else "")
                  for c in new_columns] for r in data_rows]
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": [new_columns] + realigned},
    ).execute()
    print(f"[Write → {sheet_name}] Column layout changed — header + "
          f"{len(realigned)} row(s) re-aligned by name.")


def append_rows_with_retry(service, sheet_name: str, rows: list, columns: list, max_retries: int = 5):
    """
    Append rows to a sheet tab. Creates the header row on first run if the tab
    is empty. Retries on HTTP 429 with exponential backoff.
    """
    # ── Ensure header exists AND matches the current column layout ─────────────
    # If the layout changed (e.g. new columns added), migrate the existing header
    # + rows by column NAME so no row is left shifted. Only rewrites on a genuine
    # difference; a tab whose header already matches is left untouched.
    # NOTE: this runs BEFORE the "no new rows" check, so a column-layout change is
    # applied to the live sheet even on an incremental run that fetched 0 new rows.
    try:
        hdr_res = (
            service.spreadsheets().values()
            .get(spreadsheetId=SHEET_ID, range=f"{sheet_name}!1:1")
            .execute()
        )
        existing_header = (hdr_res.get("values") or [[]])[0]
    except HttpError as e:
        print(f"[Write → {sheet_name}] Error checking header: {e}")
        return

    if not existing_header:
        service.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{sheet_name}!A1",
            valueInputOption="RAW",
            body={"values": [columns]},
        ).execute()
        print(f"[Write → {sheet_name}] Header row created.")
    elif existing_header != columns:
        _migrate_sheet_layout(service, sheet_name, existing_header, columns)

    # No new rows this cycle — the header/layout above is already reconciled.
    if not rows:
        print(f"[Write → {sheet_name}] No new rows to append.")
        return

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
              suspended_by_class: dict = None, use_cache: bool = True,
              seen: dict = None) -> dict:
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
    # dedup sets — normally fresh per run; the Step 5c backfill passes a set
    # pre-seeded from the existing sheet so already-present rows are not re-added.
    if seen is None:
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

        # Scheduled session slot (planned start/end) → IST. Kept separate from the
        # actual/conducted start_time_ist / end_time_ist above.
        sched_s_ist = to_ist(_first_present(session, _SCHED_START_FIELDS))
        sched_e_ist = to_ist(_first_present(session, _SCHED_END_FIELDS))

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
                "Session Scheduled Start": sched_s_ist,
                "Session Scheduled End":   sched_e_ist,
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

        # Sessions is UPSERTED by session_id (dedupe + in-place update), except on
        # a forced full reload where the tab was just recreated empty (plain append
        # is correct and cannot duplicate). All other tabs stay append-only.
        if key == "Sessions" and key not in FORCE_FULL_LOAD_SHEETS:
            upsert_sessions(service, rows)
        elif rows:
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


def sync_session_actual_times(service):
    """ONE SOURCE OF TRUTH for the ACTUAL session start/end.

    Aligns Sessions.start_time_ist / end_time_ist to the authoritative actual
    times recorded per participant in Attendance (session_start_ist /
    session_end_ist) — the same values the portal and the Student/Teacher feedback
    tabs show. The Sessions row can lag: the first same-day sync often records the
    SCHEDULED start, and when the finalised ACTUAL start is EARLIER the start-time
    watermark refuses to re-admit it (is_new is strict >), so Sessions keeps the
    scheduled value while Attendance already holds the actual. This step copies the
    authoritative Attendance value onto the Sessions row so both agree. For each
    session the value from the LATEST-synced attendance row wins. Generic — updates
    only the cells that differ; never touches sessions with no attendance."""
    print("\n[Sync] Aligning Sessions actual start/end with Attendance …")
    try:
        srows = _read_tab_values(service, SESSIONS_TAB)
        arows = _read_tab_values(service, ATTENDANCE_TAB)
        if len(srows) < 2 or len(arows) < 2:
            print("[Sync] Sessions or Attendance empty — skipping.")
            return
        sh = [str(h).strip() for h in srows[0]]
        spos = {h: i for i, h in enumerate(sh)}
        ah = [str(h).strip() for h in arows[0]]
        apos = {h: i for i, h in enumerate(ah)}
        if any(c not in spos for c in ("session_id", "start_time_ist", "end_time_ist")):
            print("[Sync] Sessions columns missing — skipping.")
            return
        if any(c not in apos for c in ("session_id", "session_start_ist", "session_end_ist")):
            print("[Sync] Attendance columns missing — skipping.")
            return

        a_sid, a_ss, a_se = apos["session_id"], apos["session_start_ist"], apos["session_end_ist"]
        a_sy = apos.get("synced_at")

        def _v(row, i):
            return str(row[i]).strip() if (i is not None and i < len(row) and row[i] is not None) else ""

        # authoritative (start, end) per session from the LATEST-synced attendance row
        best = {}   # sid -> [synced_key, start, end]
        for row in arows[1:]:
            sid = _v(row, a_sid)
            if not sid:
                continue
            ss, se = _v(row, a_ss), _v(row, a_se)
            syn = _v(row, a_sy)
            cur = best.get(sid)
            if cur is None:
                best[sid] = [syn, ss, se]
            else:
                if syn >= cur[0]:          # newer sync wins, coalescing non-blank
                    cur[0] = syn
                    if ss:
                        cur[1] = ss
                    if se:
                        cur[2] = se
                else:
                    if not cur[1] and ss:
                        cur[1] = ss
                    if not cur[2] and se:
                        cur[2] = se

        s_sid = spos["session_id"]
        s_start_L = _col_letter(spos["start_time_ist"])
        s_end_L = _col_letter(spos["end_time_ist"])
        updates = []
        n_start = n_end = 0
        for ri, row in enumerate(srows[1:], start=2):
            sid = _v(row, s_sid)
            if not sid or sid not in best:
                continue
            _syn, a_start, a_end = best[sid]
            cur_s = _v(row, spos["start_time_ist"])
            cur_e = _v(row, spos["end_time_ist"])
            if a_start and a_start != cur_s:
                updates.append({"range": f"{SESSIONS_TAB}!{s_start_L}{ri}", "values": [[a_start]]})
                n_start += 1
            if a_end and a_end != cur_e:
                updates.append({"range": f"{SESSIONS_TAB}!{s_end_L}{ri}", "values": [[a_end]]})
                n_end += 1

        if not updates:
            print("[Sync] Sessions actual start/end already match Attendance.")
            return
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID, body={"valueInputOption": "RAW", "data": updates}).execute()
        print(f"[Sync] ✓ Aligned start_time_ist for {n_start} and end_time_ist for {n_end} session(s) "
              f"from Attendance.")
    except Exception as e:
        print(f"[Sync] ⚠ Error aligning session times: {e}")


def _col_letter(idx0: int) -> str:
    """0-based column index → A1 column letters (A, B, … Z, AA, AB …)."""
    s = ""
    n = idx0 + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord("A") + rem) + s
    return s


def _valid_sched_pair(s_str: str, e_str: str) -> bool:
    """True only when both are real 'YYYY-MM-DD HH:MM:SS' timestamps and end>start."""
    try:
        s = datetime.strptime(str(s_str).strip()[:19], "%Y-%m-%d %H:%M:%S")
        e = datetime.strptime(str(e_str).strip()[:19], "%Y-%m-%d %H:%M:%S")
        return e > s
    except (ValueError, TypeError):
        return False


def backfill_session_scheduled_times(service, window_sessions):
    """ROOT-CAUSE FIX for blank/garbage 'Session Scheduled Start/End'.

    Same-day sync often writes a Sessions row before the API has published the
    session's real scheduled slot, so the row lands with a wrong scheduled start
    (a placeholder) and/or a blank scheduled end. Because the pipeline is
    append-only and dedups Sessions by session_id, that first (bad) row is never
    revised — so Scheduled/Diff never populate in the reports.

    This pass re-derives the scheduled slot for the recently re-fetched window
    (same fields the one-time backfill used) and patches ONLY the rows whose
    CURRENT scheduled slot is invalid (blank / null / end<=start), and only when
    a fully valid fresh slot is available. Rows that already hold a valid slot are
    never touched. Fully generic; never advances the watermark; fail-safe."""
    try:
        # 1. Fresh, valid scheduled slot per session_id from the re-fetched window.
        fresh = {}
        for sess in (window_sessions or []):
            if not isinstance(sess, dict):
                continue
            sid = sess.get("_id") or sess.get("id") or ""
            if not sid:
                continue
            s_ist = to_ist(_first_present(sess, _SCHED_START_FIELDS))
            e_ist = to_ist(_first_present(sess, _SCHED_END_FIELDS))
            if _valid_sched_pair(s_ist, e_ist):
                fresh[str(sid)] = (s_ist, e_ist)
        if not fresh:
            print("[Backfill 5d] No valid scheduled slots in the fresh window — skipping.")
            return

        # 2. Read Sessions and locate the columns by NAME (never hard-coded).
        rows = _read_tab_values(service, SESSIONS_TAB)
        if len(rows) < 2:
            print("[Backfill 5d] Sessions sheet empty — skipping.")
            return
        header = [h.strip() for h in rows[0]]
        try:
            sid_i = header.index("session_id")
            ss_i = header.index("Session Scheduled Start")
            se_i = header.index("Session Scheduled End")
        except ValueError:
            print("[Backfill 5d] Scheduled columns not found in Sessions header — skipping.")
            return
        ss_letter, se_letter = _col_letter(ss_i), _col_letter(se_i)

        # 3. Patch only rows whose CURRENT slot is invalid AND we have a valid fresh one.
        updates = []
        for r_i, row in enumerate(rows[1:], start=2):     # sheet row numbers (header = row 1)
            sid = row[sid_i].strip() if len(row) > sid_i else ""
            if not sid or sid not in fresh:
                continue
            cur_s = row[ss_i].strip() if len(row) > ss_i else ""
            cur_e = row[se_i].strip() if len(row) > se_i else ""
            if _valid_sched_pair(cur_s, cur_e):
                continue                                   # already good — leave it
            new_s, new_e = fresh[sid]
            updates.append({"range": f"{SESSIONS_TAB}!{ss_letter}{r_i}", "values": [[new_s]]})
            updates.append({"range": f"{SESSIONS_TAB}!{se_letter}{r_i}", "values": [[new_e]]})

        if not updates:
            print("[Backfill 5d] No sessions needed a scheduled-slot fix.")
            return

        service.spreadsheets().values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={"valueInputOption": "RAW", "data": updates},
        ).execute()
        print(f"[Backfill 5d] ✓ Healed scheduled slot for {len(updates)//2} session(s).")
    except Exception as e:
        print(f"[Backfill 5d] ⚠ Error during scheduled-time heal: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5c — Backfill late-arriving ATTENDANCE / FEEDBACK for recent sessions
#
#  ROOT-CAUSE FIX. The pipeline is append-only with a strict start-time
#  high-watermark (is_new: session_start > watermark). A session synced while its
#  attendance had not yet finalised — e.g. a fresh morning session synced minutes
#  after it started — is written with ZERO attendance rows; once the watermark
#  advances past that session's start time, is_new() can never re-admit it, so its
#  attendance is lost permanently (this is exactly what happened to the 08-Sep
#  06:50 session vs the 07:00 Attendance watermark).
#
#  This pass re-fetches ONLY the last BACKFILL_LOOKBACK_DAYS days FRESH (cache
#  bypassed so the 24h empty-attendance cache can't hide finalised data), seeds the
#  per-row dedup keys from what is already in the sheet, and appends only rows that
#  are genuinely new. It NEVER advances the watermark and CANNOT create duplicates,
#  so the normal incremental engine is completely untouched. Fully generic — it
#  recovers any technology/session whose facts finalise after the first sync.
# ─────────────────────────────────────────────────────────────────────────────

def _read_tab_values(service, tab_name: str) -> list:
    """Read all values from a tab (header + rows). Returns [] on any error."""
    try:
        result = (service.spreadsheets().values()
                  .get(spreadsheetId=SHEET_ID, range=f"{tab_name}!A:ZZ").execute())
        return result.get("values", [])
    except HttpError as e:
        print(f"[Backfill 5c] Could not read '{tab_name}': {e}")
        return []


# ─────────────────────────────────────────────────────────────────────────────
#  SESSIONS UPSERT  (root-cause fix for duplicate session_id rows)
#
#  The pipeline is append-only and dedups Sessions across runs only by a
#  start-time high-watermark (is_new: start_time_ist > watermark). A session's
#  start_time_ist can change between syncs — the first (same-day) sync often
#  records the SCHEDULED start (e.g. 20:00:00) and a later sync the true ACTUAL
#  start (e.g. 20:03:03, a LATER timestamp) — so is_new re-admits the same
#  session_id and a DUPLICATE row is appended (one with the scheduled start, one
#  with the actual). This upsert makes session_id the unique key: it collapses any
#  existing duplicates, updates the row in place when a session re-syncs (keeping
#  the finalised ACTUAL start), and appends only genuinely-new sessions. Fully
#  generic — nothing is keyed to a technology, id or date.
# ─────────────────────────────────────────────────────────────────────────────
_SESS_NULLS = ("", "nan", "nat", "none", "null")


def _sblank(v) -> bool:
    return v is None or str(v).strip().lower() in _SESS_NULLS


def _merge_session_rows(a: dict, b: dict) -> dict:
    """Merge two records for the SAME session_id (each a dict over
    SESSIONS_COLUMNS) into the single best record. A later synced_at is treated as
    more finalised, so it wins for start_time_ist (first sync = scheduled, later =
    actual). Non-blank values are coalesced; the scheduled slot prefers a valid
    (end > start) pair; synced_at keeps the latest."""
    sa = "" if _sblank(a.get("synced_at")) else str(a.get("synced_at")).strip()
    sb = "" if _sblank(b.get("synced_at")) else str(b.get("synced_at")).strip()
    newer, older = (a, b) if sa >= sb else (b, a)

    def pick(col):                       # newer non-blank, else older non-blank, else ""
        v = newer.get(col)
        if not _sblank(v):
            return v
        v2 = older.get(col)
        return v2 if not _sblank(v2) else ""

    out = {c: pick(c) for c in SESSIONS_COLUMNS}
    # scheduled slot: prefer a valid (end>start) pair from newer, else older
    for rec in (newer, older):
        ss, se = rec.get("Session Scheduled Start"), rec.get("Session Scheduled End")
        if _valid_sched_pair(ss, se):
            out["Session Scheduled Start"], out["Session Scheduled End"] = ss, se
            break
    return out


def upsert_sessions(service, fresh_rows: list):
    """Write Sessions as an UPSERT keyed by session_id (see header above): collapse
    existing duplicate session_id rows, update rows whose data changed on re-sync,
    and append new sessions. Falls back to a plain append on any read error so a
    transient failure never loses data."""
    from collections import OrderedDict
    fresh_rows = fresh_rows or []
    try:
        vals = _read_tab_values(service, SESSIONS_TAB)
    except Exception as e:
        print(f"[Sessions] Upsert read failed ({e}); appending instead.")
        if fresh_rows:
            append_rows_with_retry(service, SESSIONS_TAB, fresh_rows, SESSIONS_COLUMNS)
        return

    # Empty tab → write header + fresh rows.
    if not vals:
        if fresh_rows:
            body = [SESSIONS_COLUMNS] + [[r.get(c, "") for c in SESSIONS_COLUMNS] for r in fresh_rows]
            service.spreadsheets().values().update(
                spreadsheetId=SHEET_ID, range=f"{SESSIONS_TAB}!A1",
                valueInputOption="RAW", body={"values": body}).execute()
            print(f"[Sessions] Upsert: empty tab → wrote {len(fresh_rows)} row(s).")
        return

    hdr = [str(h).strip() for h in vals[0]]
    pos = {h: i for i, h in enumerate(hdr)}
    if "session_id" not in pos:
        print("[Sessions] Upsert: no session_id column — appending instead.")
        if fresh_rows:
            append_rows_with_retry(service, SESSIONS_TAB, fresh_rows, SESSIONS_COLUMNS)
        return
    sidx = pos["session_id"]

    def rowdict(row):
        return {c: (row[pos[c]] if c in pos and pos[c] < len(row) else "") for c in SESSIONS_COLUMNS}

    old_data = vals[1:]
    merged = OrderedDict()
    n_dup = 0
    for row in old_data:
        sid = str(row[sidx]).strip() if sidx < len(row) else ""
        if not sid:
            continue
        rd = rowdict(row)
        if sid in merged:
            merged[sid] = _merge_session_rows(merged[sid], rd)
            n_dup += 1
        else:
            merged[sid] = rd

    n_new = n_upd = 0
    for r in fresh_rows:
        sid = str(r.get("session_id", "")).strip()
        if not sid:
            continue
        rr = {c: r.get(c, "") for c in SESSIONS_COLUMNS}
        if sid in merged:
            merged[sid] = _merge_session_rows(merged[sid], rr)
            n_upd += 1
        else:
            merged[sid] = rr
            n_new += 1

    # Nothing to do (no duplicates and no fresh rows) → leave the sheet untouched.
    if n_dup == 0 and not fresh_rows:
        return

    out_rows = [[merged[sid].get(c, "") for c in SESSIONS_COLUMNS] for sid in merged]
    body = [SESSIONS_COLUMNS] + out_rows
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID, range=f"{SESSIONS_TAB}!A1",
        valueInputOption="RAW", body={"values": body}).execute()
    # If dedupe shrank the sheet, clear the now-orphaned trailing rows.
    old_count, new_count = len(old_data), len(out_rows)
    if new_count < old_count:
        service.spreadsheets().values().clear(
            spreadsheetId=SHEET_ID,
            range=f"{SESSIONS_TAB}!A{new_count + 2}:ZZ{old_count + 1}").execute()
    print(f"[Sessions] Upsert: {n_dup} duplicate(s) collapsed, {n_upd} updated, "
          f"{n_new} appended → {new_count} unique session(s).")


def _seed_seen_from_sheet(service) -> dict:
    """Build dedup sets from the rows ALREADY in the sheet, using the exact same
    keys transform() uses, so a re-processed session never duplicates rows."""
    seen = {"s": set(), "a": set(), "sf": set(), "tf": set(), "notf": set()}

    def _idx(header, name):
        for i, h in enumerate(header):
            if str(h).strip() == name:
                return i
        return -1

    def _val(row, i):
        return str(row[i]).strip() if 0 <= i < len(row) else ""

    def _load(tab, set_key, build):
        vals = _read_tab_values(service, tab)
        if len(vals) < 2:
            return
        hdr = [str(h).strip() for h in vals[0]]
        si  = _idx(hdr, "session_id")
        di  = _idx(hdr, "student_id")
        for row in vals[1:]:
            k = build(_val(row, si), _val(row, di))
            if k:
                seen[set_key].add(k)

    _load(SESSIONS_TAB,       "s",    lambda sid, stid: sid or None)
    _load(ATTENDANCE_TAB,     "a",    lambda sid, stid: f"{sid}_{stid}" if sid and stid else None)
    _load(STUDENT_FB_TAB,     "sf",   lambda sid, stid: f"{sid}_{stid}" if sid and stid else None)
    _load(TEACHER_FB_TAB,     "tf",   lambda sid, stid: f"{sid}_teacher" if sid else None)
    _load(SESSIONS_NO_TF_TAB, "notf", lambda sid, stid: sid or None)
    print(f"[Backfill 5c] Seeded dedup from sheet — "
          f"sessions:{len(seen['s'])} attendance:{len(seen['a'])} "
          f"student_fb:{len(seen['sf'])} teacher_fb:{len(seen['tf'])}")
    return seen


def backfill_recent_attendance(service, synced_at: str, use_cache: bool = False):
    """Recover attendance/feedback that finalised after a session's first sync.
    Additive, deduped against the sheet, and watermark-neutral (see Step 5c header)."""
    try:
        today = datetime.now(IST).date()
        start = today - timedelta(days=BACKFILL_LOOKBACK_DAYS)
        print(f"\n[Backfill 5c] Late-attendance backfill — window {start} → {today} (FRESH)")

        # 1. Re-fetch the recent window FRESH (bypass caches so finalised data shows).
        window_sessions = fetch_all_sessions([(str(start), str(today))])
        if not window_sessions:
            print("[Backfill 5c] No sessions in window — nothing to backfill.")
            return

        # Heal blank/garbage scheduled start/end for existing rows in this window
        # (same fresh fetch), so Scheduled/Diff populate in the reports.
        backfill_session_scheduled_times(service, window_sessions)

        # 2. Seed dedup keys from what is already in the sheet.
        seen = _seed_seen_from_sheet(service)

        # 3. Per-class suspension lookup, exactly as the main run does.
        class_ids = set()
        for s in window_sessions:
            c = s.get("classId") or {}
            cid = c.get("_id") if isinstance(c, dict) else (c if isinstance(c, str) else "")
            if cid:
                class_ids.add(cid)
        suspended_by_class = fetch_suspended_students(class_ids, use_cache=use_cache)

        # 4. Transform with FULL-LOAD watermarks (admit all) + seeded dedup, so only
        #    genuinely-new rows survive. use_cache=False → attendance fetched fresh.
        full_wm = {s: None for s in WATERMARK_SHEETS}
        output, counts, _ = transform(window_sessions, full_wm, synced_at,
                                      suspended_by_class, use_cache=use_cache, seen=seen)

        total_new = sum(len(v) for v in output.values())
        if total_new == 0:
            print("[Backfill 5c] No missing rows found — everything already in sheet.")
            return

        print(f"[Backfill 5c] Recovering → Attendance:{counts.get('Attendance', 0)} "
              f"Student_FB:{counts.get('Student_Feedback', 0)} "
              f"Teacher_FB:{counts.get('Teacher_Feedback', 0)} "
              f"Sessions:{counts.get('Sessions', 0)}")

        # 5. Append only. Watermarks are intentionally NOT written here — the normal
        #    incremental engine keeps full ownership of the watermark.
        write_all_tabs(service, output)
    except Exception as e:
        print(f"[Backfill 5c] ⚠ Error during attendance backfill: {e}")


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

    # ── Step 5b2: Align Sessions ACTUAL start/end with Attendance (one source of
    #    truth). Fixes rows where the first same-day sync kept the scheduled start
    #    and the finalised actual (earlier) could not re-admit past the watermark.
    sync_session_actual_times(service)

    # ── Step 5c: Backfill late-arriving attendance/feedback for recent sessions ─
    # Recovers rows for sessions synced before their attendance finalised (the
    # strict start-time watermark can otherwise never re-admit them). Additive,
    # deduped, and watermark-neutral — see backfill_recent_attendance() header.
    backfill_recent_attendance(service, synced_at, use_cache=False)

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
