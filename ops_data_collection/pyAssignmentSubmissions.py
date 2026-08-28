"""
================================================================================
  Assignment Submissions Pipeline
  Converted from user request — mirrors the structure of pyAssignments.py

  FLOW:
    1.  Fetch all classes        →  /institutes/{INSTITUTE_ID}/classes
    2.  Deduplicate classes
    3.  Per class: fetch content timeline
                                 →  /user/classes/{id}/contentTimeline
    4.  Recursively walk timeline; collect every entity where
        entityType / type / contentType == 'assessment'
    5.  Deduplicate assessment IDs across all classes
    6.  Per assessment: fetch full details + submissions
                                 →  /user/getAssessment/{assessment_id}
    7.  Flatten: one row per student submission
    8.  Upsert into Google Sheet (insert new, update changed, skip unchanged)

  TARGET SHEET:  IntelliBIAssessmentSubmission
  SHEET ID:      1E_pOuZfw4BUhQ1bRMmuc8lPDDJRtoXuGkP-qtHC96HU
  TAB:           Submissions

  UPSERT MATCH KEY:  assessment_id + student_id
    Safe to re-run at any time; only new / changed records are written.

SETUP:
  1. Place service_account.json in the same folder as this script.
  2. Share 'IntelliBIAssessmentSubmission' with the service account (Editor).
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
import hashlib
import requests
from datetime import datetime, timezone, timedelta

from utils import get_sheets_service, upsert_rows, clean_sheet, to_ist_dmy, now_ist_ymd


# ─────────────────────────────────────────────────────────────────────────────
#  FILE-BASED API CACHE
#  Stores JSON responses locally to avoid redundant API calls.
#  Each cache entry has a TTL — stale entries are re-fetched automatically.
#  Use --force-refresh to bypass cache entirely for a clean run.
# ─────────────────────────────────────────────────────────────────────────────

_SCRIPT_DIR_CACHE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR         = os.path.join(str(PROJECT_CACHE_DIR))

# TTLs in seconds
TTL_CLASSES   = 12 * 3600      # 12 hours — class list rarely changes
TTL_TIMELINE  = 12 * 3600      # 12 hours — new assessments appear infrequently
TTL_STUDENTS  = 24 * 3600      # 24 hours — student roster is stable
TTL_SESSIONS  = 12 * 3600      # 12 hours — session count changes slowly
# Assessment TTLs are dynamic — see _assessment_ttl()
TTL_ASSESS_OLD    = 7 * 86400  # 7 days  — deadline passed > 30 days ago
TTL_ASSESS_RECENT = 6 * 3600   # 6 hours — deadline within last 30 days
TTL_ASSESS_ACTIVE = 0          # 0 = always fetch fresh (deadline not yet passed)

# Refresh mode (internal config; replaces the old --force-refresh CLI flag).
#   "cache"         -> use the file cache (default).
#   "force-refresh" -> bypass the cache and fetch everything fresh from the API.
REFRESH_MODE = "cache"

def _ensure_cache_dirs():
    """Create cache directory structure if missing."""
    for sub in ["", "timelines", "assessments", "sessions"]:
        path = os.path.join(CACHE_DIR, sub) if sub else CACHE_DIR
        os.makedirs(path, exist_ok=True)

def _cache_path(category: str, key: str = "") -> str:
    """Build the file path for a cache entry."""
    safe_key = hashlib.md5(key.encode()).hexdigest() if key else ""
    if category in ("timelines", "assessments", "sessions"):
        return os.path.join(CACHE_DIR, category, f"{safe_key or 'data'}.json")
    return os.path.join(CACHE_DIR, f"{category}.json")


def _cache_get(category: str, key: str = "", ttl_seconds: int = 3600) -> dict | list | None:
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


def _assessment_ttl(deadline_str: str) -> int:
    """
    Determine cache TTL for an assessment based on its deadline.
      - Deadline > 30 days ago  → 7 days  (old, unlikely to change)
      - Deadline ≤ 30 days ago  → 6 hours (recent, might get graded)
      - Deadline in future/none → 0       (active, always fetch fresh)
    """
    if not deadline_str:
        return TTL_ASSESS_ACTIVE
    # Parse IST date: 'DD/MM/YYYY HH:MM:SS IST' or similar
    s = str(deadline_str).replace(" IST", "").strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            dl = datetime.strptime(s, fmt).date()
            today = datetime.now().date()
            if dl > today:
                return TTL_ASSESS_ACTIVE       # still active
            days_past = (today - dl).days
            if days_past > 30:
                return TTL_ASSESS_OLD          # old
            return TTL_ASSESS_RECENT            # recent
        except ValueError:
            continue
    return TTL_ASSESS_ACTIVE  # unparseable → fetch fresh


# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

from wise_config import HEADERS   # central headers; rotate API key in config_files/config.py only

INSTITUTE_ID         = "672a0918ae3d6e9fadfbc622"
BASE_URL             = "https://api.wiseapp.live"
SERVICE_ACCOUNT_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json"
)

SHEET_ID             = "1E_pOuZfw4BUhQ1bRMmuc8lPDDJRtoXuGkP-qtHC96HU"
SUBMISSIONS_TAB      = "Submissions"
ASSIGNED_TAB         = "Assessment_Assigned"
NOT_ASSIGNED_TAB     = "Assessment_Not_Assigned"

# Column order written to the sheet — keep synced_at last.
SUBMISSIONS_COLUMNS = [
    "assessment_id",
    "assessment_title",
    "assessment_description",
    "class_id",
    "class_name",
    "class_subject",
    "section_title",
    "submission_start_date",
    "submission_deadline",
    "enrolled_count",
    "submitted_count",
    "pending_count",
    "student_id",
    "student_name",
    "student_email",
    "student_phone",
    "submitted_at",
    "submission_status",
    "text_answer",
    "attachment_count",
    "attachment_links",
    "marked_as_solution",
    "maximum_marks",
    "evaluation_marks",
    "evaluation_feedback",
    "assessment_uploaded_link",
    "synced_at",
]

ASSIGNED_COLUMNS = [
    "class_id",
    "assessment_id",
    "class_name",
    "class_subject",
    "assessment_title",
    "assessment_description",
    "submission_start_date",
    "last_modified_date",
    "submission_deadline",
    "maximum_marks",
    "assessment_uploaded_link",
    "course_created_date",
    "course_start_date",
    "sessions_since_start",
    "synced_at",
]

NOT_ASSIGNED_COLUMNS = [
    "class_id",
    "class_name",
    "class_subject",
    "students_enrolled",
    "course_created_date",
    "course_start_date",
    "sessions_since_start",
    "synced_at",
]

# ── Trigger mechanism for Assignment Report ──────────────────────────────────
#  Assignment-level fields whose changes should trigger report regeneration.
#  Changes to student submission fields (submitted_at, evaluation_marks, etc.)
#  do NOT trigger the report — only structural assignment changes matter.
TRIGGER_FIELDS = [
    "class_id", "assessment_id", "class_name", "class_subject",
    "assessment_title", "assessment_description",
    "submission_start_date", "submission_deadline", "maximum_marks",
    "assessment_uploaded_link",
]

_SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
TRIGGER_FILE = os.path.join(CREDENTIALS_DIR, ".assignment_report_trigger.json")

# HEADERS is imported from config (see top of file) -- already includes Accept


# ── Data-cleaning helpers ─────────────────────────────────────────────────────

def _trim(val) -> str:
    """Strip outer whitespace AND collapse internal whitespace; return '' for None.
    e.g. 'Pratik  Samage' → 'Pratik Samage'
    """
    if val is None:
        return ""
    return " ".join(str(val).split())


def _tc(val) -> str:
    """Trim + title-case every word."""
    v = _trim(val)
    return v.title() if v else ""


def _safe_num(val) -> str:
    """Convert numeric to string (2 dp for floats); '' on failure."""
    if val is None or val == "":
        return ""
    try:
        f = float(val)
        return str(int(f)) if f == int(f) else f"{f:.2f}"
    except (TypeError, ValueError):
        return ""


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 1 — Fetch All Classes
# ─────────────────────────────────────────────────────────────────────────────

def fetch_all_classes(use_cache: bool = True) -> list:
    """
    Fetch all classes for the institute.
    Uses local file cache (TTL_CLASSES) to avoid redundant API calls.
    Raises RuntimeError on failure so the caller can abort before any write.
    """
    if use_cache:
        cached = _cache_get("classes", "", TTL_CLASSES)
        if cached is not None:
            print(f"[Fetch Classes] {len(cached)} classes (from cache)")
            return cached

    try:
        resp = requests.get(
            f"{BASE_URL}/institutes/{INSTITUTE_ID}/classes",
            headers=HEADERS, timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        raise RuntimeError(f"Failed to fetch classes: {e}")

    if isinstance(data, dict):
        class_list = (
            (data.get("data") or {}).get("classes")
            or data.get("data")
            or data.get("classes")
            or []
        )
    elif isinstance(data, list):
        class_list = data
    else:
        class_list = []

    if not isinstance(class_list, list):
        class_list = []

    _cache_put("classes", "", class_list)
    print(f"[Fetch Classes] {len(class_list)} classes received (from API → cached)")
    return class_list


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2 — Deduplicate Classes
# ─────────────────────────────────────────────────────────────────────────────

def extract_unique_classes(class_list: list) -> tuple:
    """
    Deduplicate by class _id.
    Returns:
      unique_classes : list of class metadata dicts
      enrolled_map   : dict {class_id: set of enrolled student_ids}
                       built from the joinedRequest field on each class.
    """
    seen         = set()
    unique       = []
    enrolled_map = {}   # class_id → set(student_id)

    for cls in class_list:
        if not cls or not isinstance(cls, dict):
            continue
        cid = cls.get("_id") or cls.get("id") or ""
        if not cid or cid in seen:
            continue
        seen.add(cid)
        unique.append({
            "class_id":          cid,
            "class_name":        cls.get("name")      or cls.get("title")       or "",
            "class_subject":     cls.get("subject")   or cls.get("subjectName") or
                                 cls.get("category")  or cls.get("stream")      or "",
            "course_start_date": to_ist_dmy(
                cls.get("startDate") or cls.get("startTime") or
                cls.get("start_date") or cls.get("createdAt") or ""
            ),
        })

        # joinedRequest is a list of student IDs for this class
        joined = cls.get("joinedRequest") or []
        if isinstance(joined, list):
            enrolled_map[cid] = set(
                str(sid).strip() for sid in joined if sid
            )

    return unique, enrolled_map


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2b — Fetch All Students → build lookup dict
# ─────────────────────────────────────────────────────────────────────────────

def fetch_student_lookup(use_cache: bool = True) -> dict:
    """
    Paginate GET /institutes/v3/{INSTITUTE_ID}/students (status=ACCEPTED).
    Tries camelCase pagination params first (pageSize/pageNumber), then falls
    back to a single unpaginated call if the first attempt returns a 400.
    Returns dict: { student_id → {name, email, phone} }
    Non-fatal: returns whatever was collected (enrich_lookup_for_enrolled fills gaps).
    """
    if use_cache:
        cached = _cache_get("students", "", TTL_STUDENTS)
        if cached is not None:
            print(f"[Students] {len(cached)} students loaded from cache")
            return cached
    PAGE_SIZE    = 500
    all_students = []
    page_number  = 1

    # Candidate pagination param sets — try camelCase first (WiseApp is Node/JS),
    # then snake_case, then no pagination at all.
    PARAM_VARIANTS = [
        {"pageSize": PAGE_SIZE, "pageNumber": page_number},   # camelCase (try first)
        {"page_size": PAGE_SIZE, "page_number": page_number}, # snake_case
        {},                                                    # no pagination (get all)
    ]

    working_variant = None  # index into PARAM_VARIANTS that succeeded on page 1

    # ── Page 1: probe which param variant the API accepts ────────────────────
    for idx, variant in enumerate(PARAM_VARIANTS):
        params = {"status": "ACCEPTED", **variant}
        try:
            resp = requests.get(
                f"{BASE_URL}/institutes/v3/{INSTITUTE_ID}/students",
                params=params, headers=HEADERS, timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            working_variant = idx
            variant_label = list(variant.keys())[0] if variant else "no-pagination"
            print(f"  [Students] Param variant OK: {variant_label}")
            break
        except Exception as e:
            print(f"  [Students] Variant {idx} failed ({list(variant.keys())}): {e}")
            data = None

    if working_variant is None or data is None:
        print(f"  [Students] All param variants failed — lookup will be empty; "
              f"profile fallback will cover enrolled students.")
        return {}

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _extract_students(d):
        students = (
            (d.get("data") or {}).get("students")
            or d.get("data")
            or d.get("students")
            or (d if isinstance(d, list) else [])
        )
        return students if isinstance(students, list) else []

    def _total_count(d):
        """Try to read a totalCount / total field from the response."""
        try:
            return int(
                (d.get("data") or {}).get("totalCount")
                or (d.get("data") or {}).get("total")
                or d.get("totalCount")
                or d.get("total")
                or 0
            )
        except (TypeError, ValueError):
            return 0

    # ── Collect page 1 results ────────────────────────────────────────────────
    page1_students = _extract_students(data)
    all_students.extend(page1_students)
    seen_ids = {s.get("_id") or s.get("id") for s in page1_students if s}
    total_reported = _total_count(data)
    print(f"  [Students] Page 1 → {len(page1_students)} student(s)"
          + (f"  (API total: {total_reported})" if total_reported else ""))

    # ── If no-pagination variant was used, we already have everything ─────────
    if working_variant == len(PARAM_VARIANTS) - 1 or not PARAM_VARIANTS[working_variant]:
        pass  # single call — done
    else:
        # ── Continue paginating with the working variant ──────────────────────
        base_variant = PARAM_VARIANTS[working_variant]
        page_key = "pageNumber" if "pageNumber" in base_variant else "page_number"
        size_key = "pageSize"   if "pageSize"   in base_variant else "page_size"

        MAX_PAGES  = 50   # hard safety cap — avoids infinite loops
        page_number = 2

        while page_number <= MAX_PAGES:
            # Stop early if we've already reached the total the API advertised
            if total_reported and len(all_students) >= total_reported:
                print(f"  [Students] Reached advertised total ({total_reported}); stopping.")
                break

            params = {
                "status": "ACCEPTED",
                size_key: PAGE_SIZE,
                page_key: page_number,
            }
            try:
                resp = requests.get(
                    f"{BASE_URL}/institutes/v3/{INSTITUTE_ID}/students",
                    params=params, headers=HEADERS, timeout=30,
                )
                resp.raise_for_status()
                page_data = resp.json()
            except Exception as e:
                print(f"  [Students] Error on page {page_number}: {e} — "
                      f"using {len(all_students)} fetched so far")
                break

            new_students = _extract_students(page_data)
            if not new_students:
                print(f"  [Students] Empty page {page_number}; done.")
                break

            # Dedup guard: if every ID on this page was already seen, the API
            # is looping (no real pagination) — stop immediately.
            new_ids = {s.get("_id") or s.get("id") for s in new_students if s}
            truly_new = new_ids - seen_ids
            if not truly_new:
                print(f"  [Students] Page {page_number} returned only duplicate IDs — "
                      f"API does not paginate; using {len(all_students)} student(s).")
                break

            seen_ids.update(new_ids)
            all_students.extend(new_students)
            print(f"  [Students] Page {page_number} → {len(new_students)} student(s) "
                  f"({len(all_students)} total so far)")
            page_number += 1

    # ── Build lookup dict ─────────────────────────────────────────────────────
    lookup = {}
    for s in all_students:
        if not s or not isinstance(s, dict):
            continue
        sid = _trim(s.get("_id") or s.get("id") or "")
        if not sid:
            continue
        lookup[sid] = {
            "name":  _tc(s.get("name") or ""),
            "email": _trim(s.get("email") or ""),
            "phone": _trim(s.get("phoneNumber") or s.get("phone") or ""),
        }

    _cache_put("students", "", lookup)
    print(f"[Students] {len(lookup)} students loaded into lookup (from API → cached)")
    return lookup


def enrich_lookup_for_enrolled(lookup: dict, enrolled_map: dict,
                               use_cache: bool = True) -> dict:
    """
    After the bulk ACCEPTED fetch, find every enrolled student whose ID is
    not yet in the lookup (e.g. students with non-ACCEPTED status) and pull
    their basic info from the per-student profile API.
    Mutates and returns the lookup dict.

    Cache strategy: saves the ENRICHED lookup (bulk + fallback combined)
    so subsequent runs skip all 296+ individual profile API calls.
    """
    import time

    # Check for cached enriched lookup first — this is the full lookup
    # (bulk students + profile fallback) from a previous run
    if use_cache:
        cached_enriched = _cache_get("students_enriched", "", TTL_STUDENTS)
        if cached_enriched is not None:
            # Verify it covers all currently enrolled students
            all_enrolled = set()
            for sids in enrolled_map.values():
                all_enrolled.update(sids)
            missing_from_cache = all_enrolled - set(cached_enriched.keys())
            if not missing_from_cache:
                print(f"[Students] {len(cached_enriched)} students loaded from enriched cache "
                      f"(covers all {len(all_enrolled)} enrolled)")
                return cached_enriched
            else:
                # Merge cached data into lookup, then only fetch truly new students
                lookup.update(cached_enriched)
                print(f"[Students] Enriched cache has {len(cached_enriched)} students, "
                      f"but {len(missing_from_cache)} new enrolled student(s) need fetching")

    all_enrolled = set()
    for sids in enrolled_map.values():
        all_enrolled.update(sids)

    missing = all_enrolled - set(lookup.keys())
    total_missing = len(missing)

    if not missing:
        print(f"[Students] All enrolled students present in lookup — no fallback needed.")
        # Save enriched lookup to cache
        _cache_put("students_enriched", "", lookup)
        return lookup

    est_seconds = total_missing * 0.05
    print(f"[Students] {total_missing} enrolled student(s) missing from lookup — "
          f"fetching via profile API ... (est. ~{est_seconds:.0f}s)")

    fetched = 0
    t_start = time.time()

    for idx, sid in enumerate(missing, 1):
        try:
            resp = requests.get(
                f"{BASE_URL}/public/institutes/{INSTITUTE_ID}/studentReports/{sid}",
                params={
                    "showContractData":     "false",
                    "showRegistrationData": "false",
                    "showParent":           "false",
                },
                headers=HEADERS, timeout=15,
            )
            resp.raise_for_status()
            user = (
                (resp.json().get("data") or {})
                .get("studentReport", {})
                .get("user", {})
            )
            if user.get("_id"):
                lookup[sid] = {
                    "name":  _tc(user.get("name") or ""),
                    "email": _trim(user.get("email") or ""),
                    "phone": _trim(user.get("phoneNumber") or user.get("phone") or ""),
                }
                fetched += 1
        except Exception as e:
            print(f"  [Profile fallback] Skipped {sid}: {e}")

        if idx % 10 == 0:
            elapsed  = time.time() - t_start
            rate     = idx / elapsed if elapsed > 0 else 1
            remaining = (total_missing - idx) / rate
            print(f"  [{idx}/{total_missing}] profile fallback — "
                  f"~{remaining:.0f}s remaining ...")

        time.sleep(0.05)   # 50 ms — 3× faster than before

    # Save enriched lookup to cache for next run
    _cache_put("students_enriched", "", lookup)
    print(f"[Students] Lookup now covers {len(lookup)} student(s) "
          f"(+{fetched} via fallback) in {time.time() - t_start:.1f}s → cached")
    return lookup


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2c — Fetch session count for a class (used by Assessment_Not_Assigned)
# ─────────────────────────────────────────────────────────────────────────────

def fetch_class_session_info(class_id: str, use_cache: bool = True) -> tuple:
    """
    Fetch sessions for a class.
    Returns (first_session_date_str, conducted_count) where:
      - first_session_date_str : 'DD/MM/YYYY HH:MM:SS IST' of earliest session, or ""
      - conducted_count         : sessions whose startTime <= now IST (already conducted)
    Returns ("", 0) on any failure or if no sessions found.
    Tries two common WiseApp endpoint patterns.
    """
    if use_cache:
        cached = _cache_get("sessions", class_id, TTL_SESSIONS)
        if cached is not None:
            return (cached[0], cached[1])

    sessions = []
    for url in [
        f"{BASE_URL}/user/classes/{class_id}/sessions",
        f"{BASE_URL}/institutes/{INSTITUTE_ID}/classes/{class_id}/sessions",
    ]:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30)
            if resp.status_code == 404:
                continue
            resp.raise_for_status()
            data = resp.json()
            candidates = (
                (data.get("data") or {}).get("sessions")
                or data.get("sessions")
                or data.get("data")
                or (data if isinstance(data, list) else [])
            )
            if isinstance(candidates, list):
                sessions = candidates
                break
        except Exception:
            continue

    if not sessions:
        return ("", 0)

    IST     = timezone(timedelta(hours=5, minutes=30))
    now_ist = datetime.now(IST)

    # A session is "conducted" if it has a non-null end_time/endTime.
    # Cancelled and missed sessions are scheduled but never ended, so end_time is null.
    conducted_dts = []
    for s in sessions:
        if not s or not isinstance(s, dict):
            continue
        end = s.get("endTime") or s.get("end_time")
        if not end:
            continue   # null end_time → not conducted (cancelled or missed)
        raw = s.get("startTime") or s.get("start_time") or ""
        if not raw:
            continue
        try:
            dt = datetime.fromisoformat(
                str(raw).replace("Z", "+00:00")
            ).astimezone(IST)
            conducted_dts.append(dt)
        except (ValueError, TypeError):
            continue

    if not conducted_dts:
        _cache_put("sessions", class_id, ["", 0])
        return ("", 0)

    first_dt        = min(conducted_dts)
    first_date_str  = first_dt.strftime("%d/%m/%Y %H:%M:%S") + " IST"
    conducted_count = len(conducted_dts)

    _cache_put("sessions", class_id, [first_date_str, conducted_count])
    return (first_date_str, conducted_count)


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3 — Fetch Content Timeline per Class
# ─────────────────────────────────────────────────────────────────────────────

def fetch_content_timeline(class_id: str, use_cache: bool = True) -> dict:
    """Fetch the content timeline for one class. Returns {} on error."""
    if use_cache:
        cached = _cache_get("timelines", class_id, TTL_TIMELINE)
        if cached is not None:
            return cached

    try:
        resp = requests.get(
            f"{BASE_URL}/user/classes/{class_id}/contentTimeline",
            params={"showSequentialLearningDisabledSections": "true"},
            headers=HEADERS, timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        _cache_put("timelines", class_id, data)
        return data
    except Exception as e:
        print(f"  [Timeline] Error for class {class_id}: {e}")
        return {}


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 4 — Extract Assessment IDs from Timeline (recursive entity walk)
# ─────────────────────────────────────────────────────────────────────────────

def _collect_assessment_ids(entity_list: list, section_title: str,
                             class_meta: dict, result: list):
    """
    Recursively walk an entity list; collect items where
    entityType / type / contentType == 'assessment'.
    Mirrors the flattenEntities logic in pyAssignments.py.
    """
    for item in entity_list:
        if not item or not isinstance(item, dict):
            continue

        item_type = (
            item.get("entityType") or item.get("type") or item.get("contentType") or ""
        ).lower()

        if item_type == "assessment":
            # The actual assessment _id can live in several fields
            aid = (
                item.get("entityId")     or   # most common in timeline entities
                item.get("_id")          or
                item.get("id")           or
                item.get("assessmentId") or
                ""
            )
            if aid:
                result.append({
                    "assessment_id":          _trim(aid),
                    "assessment_description": _trim(item.get("description") or ""),
                    "section_title":          section_title,
                    "class_id":               class_meta["class_id"],
                    "class_name":             class_meta["class_name"],
                    "class_subject":          class_meta["class_subject"],
                })

        # Recurse into nested entities regardless of type
        nested = item.get("entities") or item.get("items") or item.get("contents") or []
        if isinstance(nested, list) and nested:
            _collect_assessment_ids(nested, section_title, class_meta, result)


def extract_assessment_refs(raw: dict, cls: dict) -> list:
    """
    Given a raw contentTimeline response and class metadata,
    return list of {assessment_id, section_title, class_*} dicts.
    """
    if isinstance(raw, dict):
        sections = (
            (raw.get("data") or {}).get("timeline")
            or raw.get("data")
            or raw.get("timeline")
            or []
        )
    elif isinstance(raw, list):
        sections = raw
    else:
        sections = []

    if not isinstance(sections, list):
        sections = []

    refs = []
    for section in sections:
        if not section or not isinstance(section, dict):
            continue
        section_title = section.get("name") or section.get("title") or ""
        items = (
            section.get("entities")     or section.get("items") or
            section.get("contents")     or section.get("contentItems") or []
        )
        if isinstance(items, list) and items:
            _collect_assessment_ids(items, section_title, cls, refs)
    return refs


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5 — Fetch Full Assessment Details  (incl. submissions)
# ─────────────────────────────────────────────────────────────────────────────

def fetch_assessment(assessment_id: str, use_cache: bool = True,
                     deadline_hint: str = "") -> dict:
    """
    GET /user/getAssessment/{assessment_id}
    Returns the full assessment object including the submissions list.
    Returns {} on any error (that assessment is skipped gracefully).

    Smart caching: TTL depends on how old the deadline is.
    Pass deadline_hint (from timeline) for better TTL decisions.
    If no hint, tries to extract deadline from the cached response itself.
    """
    if use_cache:
        # Try to get deadline from cached data for smarter TTL
        cache_file = _cache_path("assessments", assessment_id)
        if not deadline_hint and os.path.isfile(cache_file):
            try:
                with open(cache_file, "r", encoding="utf-8") as f:
                    peek = json.load(f)
                payload = (peek.get("data") or peek.get("result") or
                           peek.get("assessment") or peek)
                if isinstance(payload, dict):
                    deadline_hint = to_ist_dmy(payload.get("submitBy") or "")
            except Exception:
                pass

        ttl = _assessment_ttl(deadline_hint)
        if ttl > 0:
            cached = _cache_get("assessments", assessment_id, ttl)
            if cached is not None:
                return cached

    try:
        resp = requests.get(
            f"{BASE_URL}/user/getAssessment/{assessment_id}",
            headers=HEADERS, timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        _cache_put("assessments", assessment_id, data)
        return data
    except Exception as e:
        print(f"  [Assessment] Error fetching {assessment_id}: {e}")
        return {}


def _probe_assessment_structure(all_refs: list):
    """
    Debug helper — probes ALL assessments until it finds one with a
    non-empty 'submission' or 'solutions' list, then prints:
      • The top-level and payload keys
      • The FULL first submission/solution item so we can map every field
    Always also probes the known reference assessment 6943e5b85a180c2b607fcca4.
    """
    import json

    sep = "─" * 60

    # Always probe the known reference ID first
    probe_ids = ["6943e5b85a180c2b607fcca4"] + [r["assessment_id"] for r in all_refs]
    # Deduplicate while preserving order
    seen_p = set()
    unique_probe_ids = []
    for pid in probe_ids:
        if pid not in seen_p:
            seen_p.add(pid)
            unique_probe_ids.append(pid)

    found_with_data = False
    for assessment_id in unique_probe_ids:
        try:
            resp = requests.get(
                f"{BASE_URL}/user/getAssessment/{assessment_id}",
                headers=HEADERS, timeout=30,
            )
            raw = resp.json()
        except Exception as e:
            print(f"  [PROBE] Error fetching {assessment_id}: {e}")
            continue

        if not isinstance(raw, dict):
            continue

        payload = raw.get("data") or raw.get("result") or raw.get("assessment") or raw
        if not isinstance(payload, dict):
            continue

        sub_list = payload.get("submission") or payload.get("solutions") or []
        if not isinstance(sub_list, list) or not sub_list:
            continue   # no submission data — try next assessment

        # Found one with actual submissions!
        found_with_data = True
        print(f"\n{sep}")
        print(f"  [PROBE] Found submissions in assessment: {assessment_id}")
        print(f"  HTTP status  : {resp.status_code}")
        print(f"  payload keys : {list(payload.keys())}")
        print(f"  'submission' : {len(payload.get('submission', []))} item(s)")
        print(f"  'solutions'  : {len(payload.get('solutions', []))} item(s)")
        print(f"\n  First submission/solution item (FULL):")
        print(json.dumps(sub_list[0], indent=4, default=str))
        print(f"{sep}\n")
        break

    if not found_with_data:
        print(f"\n{sep}")
        print("  [PROBE] No assessment with submissions found yet.")
        print("  All 'submission' and 'solutions' lists are empty.")
        print("  Showing structure of reference assessment 6943e5b85a180c2b607fcca4:")
        try:
            resp = requests.get(
                f"{BASE_URL}/user/getAssessment/6943e5b85a180c2b607fcca4",
                headers=HEADERS, timeout=30,
            )
            raw = resp.json()
            payload = raw.get("data") or raw.get("result") or raw.get("assessment") or raw
            print(f"  payload keys: {list(payload.keys()) if isinstance(payload, dict) else type(payload).__name__}")
            print(f"  Full payload (first 1200 chars):")
            print(json.dumps(payload, indent=2, default=str)[:1200])
        except Exception as e:
            print(f"  Error: {e}")
        print(f"{sep}\n")


# ─────────────────────────────────────────────────────────────────────────────
#  STEP 6 — Flatten Submissions from getAssessment Response
# ─────────────────────────────────────────────────────────────────────────────

def flatten_submissions(raw: dict, ref: dict,
                        enrolled_map: dict, student_lookup: dict) -> list:
    """
    Given a /user/getAssessment/{id} response and the timeline metadata (ref),
    return one row dict per student — BOTH submitted and not-submitted.

    Confirmed submission item structure (from probe on 6943e5b85a180c2b607fcca4):
      sub["studentId"]         → dict: {_id, name, email, phoneNumber, ...}
      sub["attachments"]       → list of submitted files
      sub["textAnswer"]        → text / GitHub link submitted by student
      sub["submissionStatus"]  → "submitted"
      sub["markedAsSolution"]  → bool
      sub["createdAt"]         → ISO timestamp of submission

    Not-submitted students are sourced from:
      enrolled_map[class_id]   → set of all enrolled student IDs
      student_lookup[sid]      → {name, email, phone} for each student
    """
    if not raw or not isinstance(raw, dict):
        return []

    # Unwrap top-level wrapper {status, message, data}
    payload = (
        raw.get("data")
        or raw.get("result")
        or raw.get("assessment")
        or raw
    )
    if not isinstance(payload, dict):
        return []

    # ── Assessment-level fields ───────────────────────────────────────────────
    assessment_id = ref["assessment_id"]

    assessment_title = _trim(
        payload.get("topic")          # confirmed: 'topic' is the title
        or payload.get("title")
        or payload.get("name")
        or ""
    )
    # Prefer description from the assessment API payload (fresher source)
    # over the timeline entity description (may be stale in cache)
    assessment_description = _trim(
        payload.get("description") or payload.get("details") or
        ref.get("assessment_description") or ""
    )
    submission_start_date = to_ist_dmy(
        payload.get("startDate") or payload.get("startBy") or
        payload.get("startTime") or payload.get("availableFrom") or
        payload.get("publishedAt") or payload.get("createdAt") or ""
    )
    submission_deadline = to_ist_dmy(payload.get("submitBy") or "")
    maximum_marks = _safe_num(
        payload.get("maxMarks") or payload.get("maximumMarks") or
        payload.get("totalMarks") or payload.get("marks") or
        payload.get("maxScore") or payload.get("totalScore") or ""
    )
    # Institute-uploaded assessment documents (payload-level attachments)
    _inst_attachments = payload.get("attachments") or []
    assessment_uploaded_link = ", ".join(
        f"{_trim(a.get('filename') or '')} :: {_trim(a.get('path') or a.get('url') or '')}"
        for a in _inst_attachments
        if isinstance(a, dict) and (a.get("path") or a.get("url"))
    ) if isinstance(_inst_attachments, list) else ""
    # pending_count is computed after intersection with enrolled_sids (see below)

    class_id      = ref["class_id"]
    class_name    = _trim(ref["class_name"])
    class_subject = _trim(ref["class_subject"])
    section_title = _tc(ref["section_title"])

    # ── Submissions list (confirmed key: 'submission' singular) ───────────────
    sub_list = payload.get("submission") or payload.get("solutions") or []
    if not isinstance(sub_list, list):
        sub_list = []

    rows           = []
    submitted_sids = set()   # track enrolled students who have submitted (dedup guard)
    synced_at      = now_ist_ymd()

    # Pre-compute enrolled set for this class (from joinedRequest)
    joined_sids = enrolled_map.get(class_id, set())

    # Collect all unique submitter IDs from the assessment
    all_submitted_ids = {
        _trim((s.get("studentId") or {}).get("_id") or "")
        for s in sub_list
        if isinstance(s, dict) and _trim((s.get("studentId") or {}).get("_id") or "")
    }

    # True enrolled = union of joinedRequest students + all submitters.
    # joinedRequest can be incomplete — anyone who submitted was clearly enrolled
    # at the time of submission (even if not in joinedRequest today).
    effective_enrolled = joined_sids | all_submitted_ids
    enrolled_count     = str(len(effective_enrolled))
    submitted_count    = str(len(all_submitted_ids))
    pending_count      = str(len(effective_enrolled) - len(all_submitted_ids))

    # Use effective_enrolled as the authoritative enrolled set going forward
    enrolled_sids = effective_enrolled

    # ── 1. Rows for students who HAVE submitted ───────────────────────────────
    for sub in sub_list:
        if not sub or not isinstance(sub, dict):
            continue

        student_obj = sub.get("studentId") or {}
        if not isinstance(student_obj, dict):
            student_obj = {}

        student_id = _trim(student_obj.get("_id") or "")
        if not student_id:
            continue

        # Skip duplicate submissions for the same student (keep first occurrence)
        if student_id in submitted_sids:
            continue
        submitted_sids.add(student_id)

        student_name  = _tc(student_obj.get("name") or "")
        student_email = _trim(student_obj.get("email") or "")
        student_phone = _trim(
            student_obj.get("phoneNumber") or student_obj.get("phone") or ""
        )

        attachments  = sub.get("attachments") or []
        submitted_at = to_ist_dmy(sub.get("createdAt") or "")
        submission_status  = _trim(sub.get("submissionStatus") or "Submitted")
        text_answer        = _trim(sub.get("textAnswer") or "")
        marked_as_solution = "Yes" if sub.get("markedAsSolution") else "No"
        # Evaluation — confirmed API field names from probe:
        #   getMark  → marks awarded by teacher
        #   feedBack → text feedback written by teacher (camelCase with capital B)
        evaluation_marks    = _safe_num(sub.get("getMark") or "")
        evaluation_feedback = _trim(sub.get("feedBack") or "")
        # attachments already extracted above for upload date
        attachment_count   = str(len(attachments)) if isinstance(attachments, list) else "0"
        attachment_links   = ", ".join(
            _trim(
                a.get("link") or a.get("url") or a.get("fileUrl") or
                a.get("attachmentUrl") or a.get("path") or ""
            )
            for a in attachments
            if isinstance(a, dict)
        ) if isinstance(attachments, list) else ""

        rows.append({
            "assessment_id":          assessment_id,
            "assessment_title":       assessment_title,
            "assessment_description": assessment_description,
            "class_id":               class_id,
            "class_name":             class_name,
            "class_subject":          class_subject,
            "section_title":          section_title,
            "submission_start_date":  submission_start_date,
            "submission_deadline":    submission_deadline,
            "enrolled_count":         enrolled_count,
            "submitted_count":        submitted_count,
            "pending_count":          pending_count,
            "student_id":             student_id,
            "student_name":           student_name,
            "student_email":          student_email,
            "student_phone":          student_phone,
            "submitted_at":           submitted_at,
            "submission_status":      submission_status,
            "text_answer":         text_answer,
            "attachment_count":    attachment_count,
            "attachment_links":    attachment_links,
            "marked_as_solution":       marked_as_solution,
            "maximum_marks":            maximum_marks,
            "evaluation_marks":         evaluation_marks,
            "evaluation_feedback":      evaluation_feedback,
            "assessment_uploaded_link": assessment_uploaded_link,
            "synced_at":                synced_at,
        })

    # ── 2. Rows for students who have NOT submitted ───────────────────────────
    not_submitted_sids = enrolled_sids - submitted_sids

    for sid in not_submitted_sids:
        info = student_lookup.get(sid) or {}
        rows.append({
            "assessment_id":          assessment_id,
            "assessment_title":       assessment_title,
            "assessment_description": assessment_description,
            "class_id":               class_id,
            "class_name":             class_name,
            "class_subject":          class_subject,
            "section_title":          section_title,
            "submission_start_date":  submission_start_date,
            "submission_deadline":    submission_deadline,
            "enrolled_count":         enrolled_count,
            "submitted_count":        submitted_count,
            "pending_count":          pending_count,
            "student_id":             sid,
            "student_name":           _tc(info.get("name") or ""),
            "student_email":          _trim(info.get("email") or ""),
            "student_phone":          _trim(info.get("phone") or ""),
            "submitted_at":           "",
            "submission_status":      "Not Submitted",
            "text_answer":         "",
            "attachment_count":    "0",
            "attachment_links":    "",
            "marked_as_solution":       "No",
            "maximum_marks":            maximum_marks,
            "evaluation_marks":         "",
            "evaluation_feedback":      "",
            "assessment_uploaded_link": assessment_uploaded_link,
            "synced_at":                synced_at,
        })

    # ── Sort: submitted rows by submitted_at ascending, not-submitted last ──────
    def _sort_key(r):
        date_str = r.get("submitted_at") or ""
        is_submitted = r.get("submission_status") != "Not Submitted"
        return (0 if is_submitted else 1, date_str)

    rows.sort(key=_sort_key)
    return rows


# ─────────────────────────────────────────────────────────────────────────────
#  TRIGGER MECHANISM — detect assignment-level changes for report generation
# ─────────────────────────────────────────────────────────────────────────────

def _snapshot_submissions(service) -> dict:
    """
    Read the existing Submissions sheet and build a snapshot dict:
        { (assessment_id, student_id): {field: value, ...} }
    Used to detect submission-level changes (new submissions, grade updates, etc.)
    Returns empty dict if the sheet doesn't exist or is empty.
    """
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{SUBMISSIONS_TAB}!A:ZZ",
        ).execute()
        values = result.get("values", [])
    except Exception as e:
        print(f"[Trigger] Warning — could not read existing {SUBMISSIONS_TAB}: {e}")
        return {}

    if len(values) < 2:
        return {}

    header = [h.strip() for h in values[0]]
    snapshot = {}
    for row in values[1:]:
        padded = row + [""] * (len(header) - len(row))
        record = {header[i]: padded[i].strip() for i in range(len(header))}
        key = (str(record.get("assessment_id", "")).strip(),
               str(record.get("student_id", "")).strip())
        if key[0] or key[1]:
            snapshot[key] = {
                "submission_status": record.get("submission_status", ""),
                "submitted_at":     record.get("submitted_at", ""),
                "evaluation_marks": record.get("evaluation_marks", ""),
                "evaluation_feedback": record.get("evaluation_feedback", ""),
                "attachment_count":  record.get("attachment_count", ""),
            }
    return snapshot


def _detect_submission_changes(old_sub_snapshot: dict, new_sub_rows: list) -> tuple:
    """
    Compare old submission snapshot with new submission rows.
    Returns (new_sub_aids, updated_sub_aids):
      new_sub_aids     — assessment_ids where a NEW student submission appeared
      updated_sub_aids — assessment_ids where an existing submission changed
                         (grade update, status change, feedback, etc.)
    """
    new_sub_aids     = set()
    updated_sub_aids = set()

    for row in new_sub_rows:
        key = (str(row.get("assessment_id", "")).strip(),
               str(row.get("student_id", "")).strip())
        aid = str(row.get("assessment_id", "")).strip()
        new_vals = {
            "submission_status": str(row.get("submission_status", "")).strip(),
            "submitted_at":     str(row.get("submitted_at", "")).strip(),
            "evaluation_marks": str(row.get("evaluation_marks", "")).strip(),
            "evaluation_feedback": str(row.get("evaluation_feedback", "")).strip(),
            "attachment_count":  str(row.get("attachment_count", "")).strip(),
        }

        if key not in old_sub_snapshot:
            # Brand-new submission row (new student or first-time submission)
            new_sub_aids.add(aid)
        else:
            old_vals = old_sub_snapshot[key]
            if any(new_vals.get(f, "") != old_vals.get(f, "") for f in new_vals):
                updated_sub_aids.add(aid)

    return new_sub_aids, updated_sub_aids


def _snapshot_last_modified(service) -> dict:
    """
    Read existing Assessment_Assigned sheet and extract last_modified_date values.
    Returns { (class_id, assessment_id): last_modified_date_str }
    """
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{ASSIGNED_TAB}!A:ZZ",
        ).execute()
        values = result.get("values", [])
    except Exception as e:
        print(f"[LastModified] Warning — could not read existing {ASSIGNED_TAB}: {e}")
        return {}

    if len(values) < 2:
        return {}

    header = [h.strip() for h in values[0]]
    lm_map = {}
    for row in values[1:]:
        padded = row + [""] * (len(header) - len(row))
        record = {header[i]: padded[i].strip() for i in range(len(header))}
        key = (str(record.get("class_id", "")).strip(),
               str(record.get("assessment_id", "")).strip())
        if key[0] or key[1]:
            lm_map[key] = record.get("last_modified_date", "")
    return lm_map


def _snapshot_assigned(service) -> dict:
    """
    Read the existing Assessment_Assigned sheet and build a snapshot dict:
        { (class_id, assessment_id): {field: value, ...} }
    Only TRIGGER_FIELDS are captured (excluding synced_at, course_created_date,
    course_start_date, sessions_since_start which are metadata, not assignment data).
    Returns empty dict if the sheet doesn't exist or is empty.
    """
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{ASSIGNED_TAB}!A:ZZ",
        ).execute()
        values = result.get("values", [])
    except Exception as e:
        print(f"[Trigger] Warning — could not read existing {ASSIGNED_TAB}: {e}")
        return {}

    if len(values) < 2:
        return {}

    header = [h.strip() for h in values[0]]
    snapshot = {}
    for row in values[1:]:
        padded = row + [""] * (len(header) - len(row))
        record = {header[i]: padded[i].strip() for i in range(len(header))}
        key = (str(record.get("class_id", "")).strip(),
               str(record.get("assessment_id", "")).strip())
        if key[0] or key[1]:
            snapshot[key] = {f: str(record.get(f, "")).strip() for f in TRIGGER_FIELDS}
    return snapshot


def _detect_assignment_changes(old_snapshot: dict, new_rows: list) -> dict:
    """
    Compare old_snapshot with the new_rows about to be upserted.
    Returns a changes dict with 'new_assignments' and 'changed_assignments' lists,
    or None if no assignment-level changes were detected.
    """
    new_assignments = []
    changed_assignments = []

    print(f"[Trigger] Comparing {len(new_rows)} new row(s) against "
          f"{len(old_snapshot)} existing assignment(s) in sheet ...")

    for row in new_rows:
        key = (str(row.get("class_id", "")).strip(),
               str(row.get("assessment_id", "")).strip())
        new_vals = {f: str(row.get(f, "")).strip() for f in TRIGGER_FIELDS}

        if key not in old_snapshot:
            print(f"[Trigger]   NEW assignment: aid={key[1]}  "
                  f"title={row.get('assessment_title', '')!r}")
            new_assignments.append({
                "assessment_id": row.get("assessment_id", ""),
                "assessment_title": row.get("assessment_title", ""),
                "class_name": row.get("class_name", ""),
            })
        else:
            old_vals = old_snapshot[key]
            changed_fields = [
                f for f in TRIGGER_FIELDS
                if new_vals.get(f, "") != old_vals.get(f, "")
            ]
            if changed_fields:
                print(f"[Trigger]   CHANGED assignment: aid={key[1]}  "
                      f"fields={changed_fields}")
                for cf in changed_fields:
                    print(f"[Trigger]     {cf}: "
                          f"{old_vals.get(cf, '')!r} → {new_vals.get(cf, '')!r}")
                changed_assignments.append({
                    "assessment_id": row.get("assessment_id", ""),
                    "assessment_title": row.get("assessment_title", ""),
                    "class_name": row.get("class_name", ""),
                    "changed_fields": changed_fields,
                })

    print(f"[Trigger] Assignment detection result: "
          f"{len(new_assignments)} new, {len(changed_assignments)} changed")

    if not new_assignments and not changed_assignments:
        return None

    return {
        "new_assignments": new_assignments,
        "changed_assignments": changed_assignments,
    }


def _write_trigger(changes: dict):
    """
    Write the trigger JSON file that pyAssignmentSubmissionsReport.py watches for.
    Includes metadata: what changed, when, and how many changes.
    """
    trigger_data = {
        "triggered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "new_count": len(changes["new_assignments"]),
        "changed_count": len(changes["changed_assignments"]),
        "new_assignments": changes["new_assignments"],
        "changed_assignments": changes["changed_assignments"],
        "new_submission_aids": list(changes.get("new_submission_aids", [])),
        "updated_submission_aids": list(changes.get("updated_submission_aids", [])),
    }
    with open(TRIGGER_FILE, "w", encoding="utf-8") as f:
        json.dump(trigger_data, f, indent=2, ensure_ascii=False)
    # Verify the file was actually written
    if os.path.isfile(TRIGGER_FILE):
        fsize = os.path.getsize(TRIGGER_FILE)
        print(f"[Trigger] ✓ Trigger file created: {TRIGGER_FILE}  ({fsize} bytes)")
    else:
        print(f"[Trigger] ✗ WARNING — file write succeeded but file not found at: {TRIGGER_FILE}")
    print(f"[Trigger]   New: {trigger_data['new_count']} | "
          f"Changed: {trigger_data['changed_count']}")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Refresh mode is set via the REFRESH_MODE config variable (top of file),
    # replacing the old --force-refresh CLI flag.
    use_cache = REFRESH_MODE != "force-refresh"

    sep = "=" * 64
    print(f"\n{sep}")
    print("  Assignment Submissions Pipeline")
    print(f"  Started : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Cache   : {'DISABLED (--force-refresh)' if not use_cache else 'ENABLED'}")
    print(f"  Sheet   : IntelliBIAssessmentSubmission")
    print(f"  Tab     : {SUBMISSIONS_TAB}")
    print(f"{sep}\n")

    # ── Initialize cache directories ─────────────────────────────────────────
    if use_cache:
        _ensure_cache_dirs()
    else:
        _cache_clear_all()
        _ensure_cache_dirs()

    service = get_sheets_service(SERVICE_ACCOUNT_FILE)

    # ── Step 1: Fetch classes — abort on failure, no partial writes ───────────
    try:
        class_list = fetch_all_classes(use_cache=use_cache)
    except RuntimeError as e:
        print(f"[Main] ABORTED — {e}")
        print("[Main] No data written. Fix the error and re-run.")
        return

    if not class_list:
        print("[Main] No classes returned by API. Exiting.")
        return

    # ── Step 2: Deduplicate classes + build enrolled student map ─────────────
    unique_classes, enrolled_map = extract_unique_classes(class_list)
    if not unique_classes:
        print("[Main] No unique classes after dedup. Exiting.")
        return
    print(f"[Classes] {len(unique_classes)} unique class(es) | "
          f"{sum(len(v) for v in enrolled_map.values())} total enrolment records")

    # ── Step 2b: Fetch all students → build name/email/phone lookup ──────────
    # First pass: all ACCEPTED students (bulk API)
    # Second pass: profile API fallback for any enrolled student not in lookup
    student_lookup = fetch_student_lookup(use_cache=use_cache)
    student_lookup = enrich_lookup_for_enrolled(student_lookup, enrolled_map, use_cache=use_cache)

    # ── Steps 3 + 4: Fetch timeline per class → collect assessment IDs ────────
    all_refs  = []   # [{assessment_id, section_title, class_id, ...}]
    seen_aids = set()

    print(f"[Timeline] Fetching content timelines for {len(unique_classes)} class(es) ...")
    for cls in unique_classes:
        raw  = fetch_content_timeline(cls["class_id"], use_cache=use_cache)
        refs = extract_assessment_refs(raw, cls)
        for ref in refs:
            aid = ref["assessment_id"]
            if aid and aid not in seen_aids:
                seen_aids.add(aid)
                all_refs.append(ref)

    print(f"[Timeline] Done — {len(all_refs)} unique assessment(s) discovered")

    if not all_refs:
        print("[Main] No assessments found across all classes. Exiting.")
        return

    # ── Steps 5 + 6: Fetch each assessment → flatten submitted + not-submitted ─
    all_submission_rows  = []
    all_assigned_rows    = []   # one row per assessment → Assessment_Assigned sheet
    total = len(all_refs)
    synced_at = now_ist_ymd()
    cache_hits = 0

    # Lookup: class_id → course_created_date (createdAt from class metadata)
    class_created_map = {
        cls["class_id"]: cls.get("course_start_date") or ""
        for cls in unique_classes
    }
    # Cache: class_id → (course_start_date, sessions_since_start) from sessions API
    session_cache: dict = {}

    print(f"[Submissions] Fetching {total} assessment(s) ...")
    for i, ref in enumerate(all_refs, 1):
        aid = ref["assessment_id"]
        # Use deadline from timeline metadata as hint for smart cache TTL
        deadline_hint = ref.get("submission_deadline", "")
        print(f"  [{i:>3}/{total}] {aid} ...", end=" ", flush=True)
        raw  = fetch_assessment(aid, use_cache=use_cache, deadline_hint=deadline_hint)
        rows = flatten_submissions(raw, ref, enrolled_map, student_lookup)
        submitted     = sum(1 for r in rows if r["submission_status"] != "Not Submitted")
        not_submitted = len(rows) - submitted
        print(f"submitted={submitted}  not_submitted={not_submitted}")
        all_submission_rows.extend(rows)

        # Fetch session info for this class (cached so multi-assessment classes
        # only hit the API once)
        cid = ref["class_id"]
        if cid not in session_cache:
            first_date, conducted = fetch_class_session_info(cid, use_cache=use_cache)
            session_cache[cid] = (first_date, conducted)
        course_start_date, sessions_since_start = session_cache[cid]
        course_created_date = class_created_map.get(cid, "")

        # Collect one row for the Assessment_Assigned sheet
        # Pull assessment-level fields from the first row (all share the same values)
        if rows:
            first = rows[0]
            all_assigned_rows.append({
                "class_id":               first["class_id"],
                "assessment_id":          aid,
                "class_name":             first["class_name"],
                "class_subject":          first["class_subject"],
                "assessment_title":       first["assessment_title"],
                "assessment_description": first["assessment_description"],
                "submission_start_date":  first["submission_start_date"],
                "last_modified_date":     "",   # populated later after change detection
                "submission_deadline":    first["submission_deadline"],
                "maximum_marks":          first["maximum_marks"],
                "assessment_uploaded_link": first["assessment_uploaded_link"],
                "course_created_date":    course_created_date,
                "course_start_date":      course_start_date,
                "sessions_since_start":   str(sessions_since_start),
                "synced_at":              synced_at,
            })
        else:
            # Assessment found in timeline but returned no submission rows
            # (e.g. no enrolled students yet).  Extract assessment-level fields
            # directly from the API payload so the Assigned sheet is complete.
            _payload = {}
            if raw and isinstance(raw, dict):
                _payload = (raw.get("data") or raw.get("result")
                            or raw.get("assessment") or raw)
                if not isinstance(_payload, dict):
                    _payload = {}

            _a_title = _trim(
                _payload.get("topic") or _payload.get("title")
                or _payload.get("name") or ""
            )
            _a_desc = _trim(
                _payload.get("description") or _payload.get("details")
                or ref.get("assessment_description") or ""
            )
            _a_start = to_ist_dmy(
                _payload.get("startDate") or _payload.get("startBy")
                or _payload.get("startTime") or _payload.get("availableFrom")
                or _payload.get("publishedAt") or _payload.get("createdAt") or ""
            )
            _a_deadline = to_ist_dmy(_payload.get("submitBy") or "")
            _a_marks = _safe_num(
                _payload.get("maxMarks") or _payload.get("maximumMarks")
                or _payload.get("totalMarks") or _payload.get("marks")
                or _payload.get("maxScore") or _payload.get("totalScore") or ""
            )
            _inst_att = _payload.get("attachments") or []
            _a_link = ", ".join(
                f"{_trim(a.get('filename') or '')} :: {_trim(a.get('path') or a.get('url') or '')}"
                for a in _inst_att
                if isinstance(a, dict) and (a.get("path") or a.get("url"))
            ) if isinstance(_inst_att, list) else ""

            all_assigned_rows.append({
                "class_id":               ref["class_id"],
                "assessment_id":          aid,
                "class_name":             _trim(ref["class_name"]),
                "class_subject":          _trim(ref["class_subject"]),
                "assessment_title":       _a_title,
                "assessment_description": _a_desc,
                "submission_start_date":  _a_start,
                "last_modified_date":     "",   # populated later after change detection
                "submission_deadline":    _a_deadline,
                "maximum_marks":          _a_marks,
                "assessment_uploaded_link": _a_link,
                "course_created_date":    course_created_date,
                "course_start_date":      course_start_date,
                "sessions_since_start":   str(sessions_since_start),
                "synced_at":              synced_at,
            })

    # ── Cache stats ──────────────────────────────────────────────────────────
    if use_cache:
        assess_cache_dir = os.path.join(CACHE_DIR, "assessments")
        cached_count = len(os.listdir(assess_cache_dir)) if os.path.isdir(assess_cache_dir) else 0
        print(f"\n[Submissions] Total: {len(all_submission_rows)} submission row(s) extracted")
        print(f"[Cache] {cached_count} assessment(s) cached on disk")
    else:
        print(f"\n[Submissions] Total: {len(all_submission_rows)} submission row(s) extracted")

    if not all_submission_rows:
        print("[Main] No submission rows to write. Exiting.")
        return

    # ── Snapshot BEFORE upsert: submissions + assigned + last_modified ────────
    print(f"\n[Trigger] Snapshotting existing data for change detection ...")
    old_sub_snapshot = _snapshot_submissions(service)
    print(f"[Trigger] Existing submissions snapshot: {len(old_sub_snapshot)} row(s)")

    old_assigned_snapshot = _snapshot_assigned(service)
    print(f"[Trigger] Existing assigned snapshot: {len(old_assigned_snapshot)} assignment(s)")

    old_lm_map = _snapshot_last_modified(service)
    print(f"[LastModified] Existing last_modified_date map: {len(old_lm_map)} entry(ies)")

    # ── Detect submission-level changes (new submissions vs updated submissions)
    new_sub_aids, updated_sub_aids = _detect_submission_changes(
        old_sub_snapshot, all_submission_rows
    )
    sub_changed_aids = new_sub_aids | updated_sub_aids
    if sub_changed_aids:
        print(f"[Trigger] Submission changes: {len(new_sub_aids)} new + "
              f"{len(updated_sub_aids)} updated across assessment(s)")
        if new_sub_aids:
            print(f"[Trigger]   New submission assessment_ids: {sorted(new_sub_aids)}")
        if updated_sub_aids:
            print(f"[Trigger]   Updated submission assessment_ids: {sorted(updated_sub_aids)}")
    else:
        print(f"[Trigger] No submission-level changes detected "
              f"(compared {len(all_submission_rows)} row(s) vs "
              f"{len(old_sub_snapshot)} existing)")

    # ── Step 7: Upsert to Submissions tab ─────────────────────────────────────
    upsert_rows(
        service, SHEET_ID,
        tab_name   = SUBMISSIONS_TAB,
        columns    = SUBMISSIONS_COLUMNS,
        rows       = all_submission_rows,
        match_keys = ["assessment_id", "student_id"],
    )

    # ── Detect assignment-level changes ──────────────────────────────────────
    changes = _detect_assignment_changes(old_assigned_snapshot, all_assigned_rows)

    # ── Compute last_modified_date for each assigned row ─────────────────────
    today_str = datetime.now().strftime("%m/%d/%Y")
    # Determine which assessment_ids have ANY change (assignment-level or submission)
    changed_assignment_aids = set()
    new_assignment_aids     = set()
    if changes:
        for a in changes.get("new_assignments", []):
            new_assignment_aids.add(a.get("assessment_id", ""))
        for a in changes.get("changed_assignments", []):
            changed_assignment_aids.add(a.get("assessment_id", ""))

    all_modified_aids = new_assignment_aids | changed_assignment_aids | sub_changed_aids

    for row in all_assigned_rows:
        key = (str(row.get("class_id", "")).strip(),
               str(row.get("assessment_id", "")).strip())
        aid = str(row.get("assessment_id", "")).strip()
        if aid in new_assignment_aids:
            # New assignment → set to today
            row["last_modified_date"] = today_str
        elif aid in all_modified_aids:
            # Changed assignment or submission change → set to today
            row["last_modified_date"] = today_str
        else:
            # Unchanged → keep existing value
            row["last_modified_date"] = old_lm_map.get(key, "")

    # ── Sheet: Assessment_Assigned (one row per assessment) ───────────────────
    print(f"\n[Assigned] Writing {len(all_assigned_rows)} row(s) to '{ASSIGNED_TAB}' ...")
    upsert_rows(
        service, SHEET_ID,
        tab_name   = ASSIGNED_TAB,
        columns    = ASSIGNED_COLUMNS,
        rows       = all_assigned_rows,
        match_keys = ["class_id", "assessment_id"],
    )

    # ── Trigger: include BOTH assignment-level AND submission changes ─────────
    if sub_changed_aids:
        if changes is None:
            changes = {"new_assignments": [], "changed_assignments": []}
        # Add submission-changed aids that aren't already in new/changed
        existing_aids = new_assignment_aids | changed_assignment_aids
        for aid in sub_changed_aids:
            if aid not in existing_aids:
                changes["changed_assignments"].append({
                    "assessment_id": aid,
                    "assessment_title": "",
                    "class_name": "",
                    "changed_fields": ["submissions_updated"],
                })
        # Store separate new vs updated submission aid lists for the report
        changes["new_submission_aids"] = list(new_sub_aids)
        changes["updated_submission_aids"] = list(updated_sub_aids)

    if changes:
        try:
            _write_trigger(changes)
        except Exception as e:
            print(f"[Trigger] ✗ ERROR writing trigger file: {e}")
            print(f"[Trigger]   Path: {TRIGGER_FILE}")
    else:
        print(f"[Trigger] No changes detected — no trigger file created.")
        print(f"[Trigger]   (assignment-level: 0 new, 0 changed | "
              f"submission-level: {len(sub_changed_aids)} changed)")
        print(f"[Trigger]   Trigger path: {TRIGGER_FILE}")

    # ── Sheet: Assessment_Not_Assigned (one row per class without assessments) ─
    class_ids_with_assessments = {ref["class_id"] for ref in all_refs}
    not_assigned_classes = [
        cls for cls in unique_classes
        if cls["class_id"] not in class_ids_with_assessments
    ]
    print(f"\n[Not Assigned] {len(not_assigned_classes)} class(es) have no assessments.")

    all_not_assigned_rows = []
    for cls in not_assigned_classes:
        cid = cls["class_id"]
        print(f"  [Sessions] Fetching session info for class {cid} ...", end=" ", flush=True)
        first_date, conducted = fetch_class_session_info(cid)
        print(f"first={first_date or '—'}  conducted={conducted}")
        all_not_assigned_rows.append({
            "class_id":             cid,
            "class_name":           _trim(cls["class_name"]),
            "class_subject":        _trim(cls["class_subject"]),
            "students_enrolled":    str(len(enrolled_map.get(cid, set()))),
            "course_created_date":  cls.get("course_start_date") or "",
            "course_start_date":    first_date,
            "sessions_since_start": str(conducted),
            "synced_at":            synced_at,
        })

    # Clear the sheet before writing to remove any stale or corrupt rows
    # from previous failed runs (upsert cannot clean up rows with garbage keys).
    print(f"[Not Assigned] Clearing '{NOT_ASSIGNED_TAB}' to remove stale rows ...")
    try:
        service.spreadsheets().values().clear(
            spreadsheetId=SHEET_ID,
            range=NOT_ASSIGNED_TAB,
        ).execute()
    except Exception as e:
        print(f"[Not Assigned] Warning — could not clear sheet: {e}")

    print(f"[Not Assigned] Writing {len(all_not_assigned_rows)} row(s) to '{NOT_ASSIGNED_TAB}' ...")
    upsert_rows(
        service, SHEET_ID,
        tab_name   = NOT_ASSIGNED_TAB,
        columns    = NOT_ASSIGNED_COLUMNS,
        rows       = all_not_assigned_rows,
        match_keys = ["class_id"],
    )

    # ── Step 8: Post-write sheet cleaning (trim / purge nulls / dedup) ──────────
    print(f"\n[Clean] Running post-write cleaning on all sheets ...")
    clean_sheet(service, SHEET_ID, SUBMISSIONS_TAB,  SUBMISSIONS_COLUMNS)
    clean_sheet(service, SHEET_ID, ASSIGNED_TAB,     ASSIGNED_COLUMNS)
    clean_sheet(service, SHEET_ID, NOT_ASSIGNED_TAB, NOT_ASSIGNED_COLUMNS)

    print(f"\n{sep}")
    print(f"  Pipeline complete.")
    print(f"  Classes processed    : {len(unique_classes)}")
    print(f"  Assessments found    : {len(all_refs)}")
    print(f"  Submissions written  : {len(all_submission_rows)}")
    print(f"  Assigned sheet rows  : {len(all_assigned_rows)}")
    print(f"  Not-assigned classes : {len(all_not_assigned_rows)}")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
