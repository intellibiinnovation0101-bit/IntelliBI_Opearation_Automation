"""
================================================================================
  IntelliBI Wise Portal — Data Validation & Reporting  (pyWiseDataValidationReport.py)
  ------------------------------------------------------------------------------
  Validates Student, Course and Instructor master data loaded from the Wise
  portal and produces a COORDINATOR-FRIENDLY validation report that contains
  ONLY the records that FAIL one or more validation rules. Records that pass
  every rule are intentionally omitted.

  DATA SOURCES
    - Google Sheet "IntelliBIStudentInfo" tabs (refreshed by the student_info
      load task): Students / ClassLearnerTeacherEnrolled / Instructor.
    - Wise REST API "top-up" for the two student fields that are NOT stored in
      the Students tab: Private Note and Profile Picture (best-effort; reuses
      pyStudentPaymentClassesStudentEnrolled.fetch_all_students()).
    - Faculty Master Sheet (Source 1) + old_instructor_technology_mapping.json
      (Source 2) for instructor / course-tag name validation.

  OUTPUT  (fresh / overwrite on every execution)
    - A Google Sheet (auto-created or reused) inside a configured Drive folder,
      with three tabs:
          IntelliBIWiseStudentValidation
          IntelliBIWiseCourseValidation
          IntelliBIWiseInstructorValidation
    - A local .xlsx copy alongside this script (convenience artifact).

  CONFIG-DRIVEN RULES (all under config_files/)
    course_technologies_mapping.json   -> Tag Name short codes
    technology_mapping.json            -> Course Title technology_name + keywords
    old_instructor_technology_mapping.json -> instructor/faculty names (Source 2)
    batch_tag_suffix_mapping.json      -> allowed Tag Name optional suffixes

  ENTRY POINT
      python Reports/pyWiseDataValidationReport.py
  Registered in config_files/pipeline_config.py (group "report"); runs in the
  daily pipeline (run_pipeline.py) and in run_all.py.
================================================================================
"""
from __future__ import annotations

import os
import sys
import re
import json
from datetime import datetime, date, timezone, timedelta

# ── Operations-project integration plumbing (Layer 2) ────────────────────────
#    ONLY path resolution changed for this project — every validation rule,
#    calculation and output below is the original, unchanged. `common/` goes on
#    sys.path so `_bootstrap` seeds the portable env + config.yaml (exactly like
#    the other Layer-2 scripts); the legacy `config_files/<name>` references
#    resolve through `paths.cfg()` (mapping JSONs -> config/, service account ->
#    credentials/); and the Layer-1 data-collection dir is added to sys.path so
#    the Wise session helper this script imports internally
#    (`fetch_all_sessions` from pyStudentPaymentClassesStudentEnrolled) resolves.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(_SCRIPT_DIR), "common"))
import _bootstrap            # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import cfg, LAYER1_DIR  # noqa: E402
if str(LAYER1_DIR) not in sys.path:
    sys.path.insert(0, str(LAYER1_DIR))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils  import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────
SERVICE_ACCOUNT_FILE = cfg("service_account.json")

COURSE_TECH_MAP_FILE = cfg("course_technologies_mapping.json")
TECH_MAP_FILE        = cfg("technology_mapping.json")
OLD_INSTRUCTOR_FILE  = cfg("old_instructor_technology_mapping.json")
SUFFIX_MAP_FILE      = cfg("batch_tag_suffix_mapping.json")

# ── INPUT: IntelliBIStudentInfo Google Sheet (refreshed by student_info load) ──
SOURCE_SHEET_ID = "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA"
STUDENTS_TAB    = "Students"
COMBINED_TAB    = "ClassLearnerTeacherEnrolled"
INSTRUCTOR_TAB  = "Instructor"

# ── Faculty Master Sheet (Source 1 for tag-name validation) ───────────────────
FACULTY_SHEET_ID   = "1fVFLq0TDxgFS7Ca81h8ScBfgcPhMqhHJDt3URyAkH50"
FACULTY_SHEET_GID  = 1459562660          # tab is resolved from this gid at runtime
FACULTY_NAME_COLUMN_CANDIDATES = ("faculty name", "facultyname", "name", "faculty")

# ── Faculty Onboarding responses (same spreadsheet as the Faculty Master sheet).
#    Instructor phone/email are validated against the LATEST onboarding record.
ONBOARDING_SHEET_ID = FACULTY_SHEET_ID
ONBOARDING_TABS = [
    "IntelliBI — Faculty Onboarding Responses-New",
    "IntelliBI — Faculty Onboarding Responses-Old",
]
# Exact onboarding column headers (identical in both tabs).
ONB_TIMESTAMP_COL = "Timestamp"
ONB_NAME_COL      = "Faculty Name"
ONB_EMAIL_COL     = "Email ID"
ONB_PHONE_COL     = "Phone Number"
# Fallback (secondary) contact columns — used when the primary does not match or
# is blank. Present in the Old tab; absent columns simply read as blank.
ONB_ALT_EMAIL_COL = "Alternative Email ID"
ONB_ALT_PHONE_COL = "Alternative Phone Number"

# Course titles excluded from ALL course validation checks (case-insensitive).
COURSE_EXCLUDED_TITLES = {"sample course", "learning journey with intellibi"}

# ── OUTPUT: written to a Google Sheet (fresh / overwrite each run) ─────────────
# A service account has NO Drive storage of its own, so it cannot CREATE a file
# in a personal "My Drive" folder. To make output reliable, resolution order is:
#   1. OUTPUT_SPREADSHEET_ID  -> write straight to this existing sheet (best).
#   2. else find a sheet named OUTPUT_SPREADSHEET_NAME inside OUTPUT_DRIVE_FOLDER_ID.
#   3. else try to create it in the folder (works only on a Shared Drive).
# Set OUTPUT_SPREADSHEET_ID to a sheet shared with the service account (Editor)
# for the simplest, always-working setup. e.g. the IntelliBIStudentInfo sheet id.
OUTPUT_SPREADSHEET_ID   = "19me_6xEPYYSz903-DHEJjdNr1sQAktWWKW1kw4aR_qs"     # <- paste an existing sheet id here (recommended)
OUTPUT_DRIVE_FOLDER_ID  = "1katHlp9raBx6GUA9ZewdI2k8d1o-VYpb"
OUTPUT_SPREADSHEET_NAME = "IntelliBI Wise Data Validation Report"
TAB_STUDENT    = "IntelliBIWiseStudentValidation"
TAB_COURSE     = "IntelliBIWiseCourseValidation"
TAB_INSTRUCTOR = "IntelliBIWiseInstructorValidation"
LOCAL_XLSX     = os.path.join(_SCRIPT_DIR, "IntelliBIWiseDataValidation.xlsx")

# ── "Current Date" — predefined hardcoded reference for the subtitle check ─────
IST = timezone(timedelta(hours=5, minutes=30))
CURRENT_DATE        = datetime.now(IST).date()      # used for future-session check
CURRENT_DATE_TOKEN  = "Current Date"                # literal token allowed in subtitle
GENERATED_AT        = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

# Private Note presence is inferred from the Students tab: the loader parses each
# student's private note (Wise additionalNote) into these columns, so if ANY is
# populated the student HAS a private note; if all are blank it is Missing.
# NOTE: 'candidate_name' is DELIBERATELY excluded — the loader always fills it with
# the student's own name when the note is blank ("the column is never left empty"),
# so it is not a reliable signal of a real private note and would make every
# student read Valid. Only genuine note-content columns are checked here.
# Profile Picture is read directly from the loader's 'profile_picture' column.
NOTE_PRESENCE_COLUMNS = [
    "highest_education", "passout_year", "how_heard_intellibi",
    "reference_name", "fresher_working_professional", "years_experience",
    "it_non_it", "it_domain", "non_it_industry", "follow_up_comments",
    "is_candidate_active", "is_attendance_required",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Status constants
S_VALID, S_MISSING, S_INVALID, S_NA = "Valid", "Missing", "Invalid", "Not Checked"
S_WARNING = "Warning"   # soft flag (e.g. student not enrolled in any course)

# Instructor names to exclude from the Instructor validation tab (test / non-real
# accounts). Matched case-insensitively against instructor_details and every
# instructor_name_1..5 — a row is skipped if ANY of them matches.
EXCLUDED_INSTRUCTOR_NAMES = {"tester", "intellibi", "neha t"}

# Test / non-real STUDENT names to exclude from the Student validation tab (and
# from all summary counts). Matched case-insensitively against student_name /
# candidate_name (exact match, or a name starting with "test student").
EXCLUDED_STUDENT_NAMES = {"test student", "tester", "test"}

# Dummy batch names whose students are excluded from Student validation entirely.
# EXACT match only (case-insensitive) — "DF" = Dummy Faculty, "DS" = Dummy Student.
# A real batch such as "DS0525" is NOT excluded. Students in these batches are not
# processed and not counted.
EXCLUDED_BATCH_NAMES = {"DF", "DS"}


# ─────────────────────────────────────────────────────────────────────────────
#  GOOGLE SERVICES
# ─────────────────────────────────────────────────────────────────────────────
def google_services():
    """Return (sheets_service, drive_service) authenticated via the service account."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    drive  = build("drive",  "v3", credentials=creds, cache_discovery=False)
    return sheets, drive


def read_tab(sheets, spreadsheet_id, tab, rng="A1:ZZ"):
    """Read a sheet tab into a list of dicts keyed by the header row."""
    resp = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{tab}!{rng}"
    ).execute()
    values = resp.get("values", [])
    if not values:
        return []
    header = [str(h).strip() for h in values[0]]
    rows = []
    for raw in values[1:]:
        raw = list(raw) + [""] * (len(header) - len(raw))
        rows.append({header[i]: (raw[i] if i < len(raw) else "") for i in range(len(header))})
    return rows


# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG / REFERENCE-DATA LOADERS
# ─────────────────────────────────────────────────────────────────────────────
def _load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def load_reference_data():
    """Load all config-driven reference sets used by the validators."""
    ctm = _load_json(COURSE_TECH_MAP_FILE)
    short_codes = set()
    # Per short code: is the MMYY (second batch-tag segment) mandatory?
    #   is_course_mmyy_required = "Y" -> MMYY required (default when absent)
    #   is_course_mmyy_required = "N" -> MMYY optional (tag may be short-code only)
    # A short code may appear more than once (e.g. "DS" = Data Science / Dummy
    # Student); treat MMYY as optional if ANY mapping entry marks it optional, so
    # a legitimately MMYY-less tag is never falsely flagged.
    mmyy_required = {}
    for c in ctm.get("courses", []):
        code = str(c.get("short_code", "")).strip()
        if not code:
            continue
        cu = code.upper()
        short_codes.add(cu)
        req = str(c.get("is_course_mmyy_required", "Y")).strip().upper() != "N"
        mmyy_required[cu] = (mmyy_required[cu] and req) if cu in mmyy_required else req

    tm = _load_json(TECH_MAP_FILE)
    technology_names = {
        str(t.get("technology_name", "")).strip()
        for t in tm.get("technologies", []) if str(t.get("technology_name", "")).strip()
    }
    additional_keywords = [str(k).strip() for k in tm.get("additional_keywords", []) if str(k).strip()]

    oim = _load_json(OLD_INSTRUCTOR_FILE)
    old_instructor_names = {
        str(r.get("name", "")).strip()
        for r in oim.get("instructor_technologies", []) if str(r.get("name", "")).strip()
    }

    sm = _load_json(SUFFIX_MAP_FILE)
    suffix_keys = set(sm.get("suffix_mapping", {}).keys())

    return {
        "short_codes": short_codes,
        "mmyy_required": mmyy_required,
        "technology_names": technology_names,
        "additional_keywords": additional_keywords,
        "old_instructor_names": old_instructor_names,
        "old_instructor_names_lc": {n.lower() for n in old_instructor_names},
        "suffix_keys": suffix_keys,
    }


def load_faculty_names(sheets):
    """Resolve the Faculty Master tab from its gid and return a set of faculty names."""
    names = set()
    tab_title = None
    try:
        meta = sheets.spreadsheets().get(spreadsheetId=FACULTY_SHEET_ID).execute()
        for s in meta.get("sheets", []):
            if s.get("properties", {}).get("sheetId") == FACULTY_SHEET_GID:
                tab_title = s["properties"]["title"]
                break
        if tab_title is None and meta.get("sheets"):
            tab_title = meta["sheets"][0]["properties"]["title"]

        rows = read_tab(sheets, FACULTY_SHEET_ID, tab_title)
        if rows:
            cols = list(rows[0].keys())
            name_col = next((c for c in cols if c.strip().lower() in FACULTY_NAME_COLUMN_CANDIDATES), cols[0])
            for r in rows:
                v = str(r.get(name_col, "")).strip()
                if v:
                    names.add(v)
        print(f"[Faculty] Loaded {len(names)} faculty name(s) from tab '{tab_title}'.")
    except Exception as exc:                              # noqa: BLE001
        print(f"[WARN] Could not read Faculty Master Sheet ({exc}). "
              f"Falling back to old_instructor_technology_mapping.json only.")
    return names, {n.lower() for n in names}


# NOTE: Private Note and Profile Picture are no longer fetched from the Wise API
# here. They are read from the Students tab, which the loader populates (private
# note parsed into NOTE_PRESENCE_COLUMNS; picture into the 'profile_picture'
# column). See build_student_rows().


# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATORS
# ─────────────────────────────────────────────────────────────────────────────
_NAME_WORD    = re.compile(r"^[A-Z][a-z]+$")                      # strict InitCap word
_LAST_WORD    = re.compile(r"^[A-Z][a-z]+(\([A-Z][a-z]+\))?$")    # LastName or LastName(Surname)
_TAG_RE       = re.compile(r"^([A-Za-z]+)(\d{4})(.*)$")   # <ShortCode><MMYY><OptionalSegment>
_TAG_NOMMYY_RE = re.compile(r"^[A-Za-z]+$")               # short-code only (MMYY omitted)
_DATE_PART    = r"\d{2}-[A-Z][a-z]{2}-\d{4}"
_SUBTITLE_RE  = re.compile(rf"^({_DATE_PART})\s+To\s+({_DATE_PART}|Current Date)$")
_EMAIL_RE     = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
_PHONE_RE     = re.compile(r"^\+\d{10,15}$")                      # +<country code><number>


def validate_student_name(name):
    """
    Exactly two parts: 'FirstName LastName' (optionally 'LastName(Surname)').
    Each word InitCap (first letter uppercase, the rest lowercase). More than two
    parts (e.g. 'Akhil Satyanarayan Chukkalwar' or 'Pratik S Patil') is Invalid.
    """
    name = (name or "").strip()
    if not name:
        return S_MISSING, ""
    tokens = name.split()
    if len(tokens) != 2:
        return (S_INVALID,
                "Name must have exactly two parts 'FirstName LastName' "
                f"(optionally 'LastName(Surname)'); found {len(tokens)} part(s).")
    first, last = tokens
    if not _NAME_WORD.match(first):
        return (S_INVALID,
                f"First name '{first}' must be InitCap "
                "(first letter uppercase, remaining lowercase).")
    if not _LAST_WORD.match(last):
        return (S_INVALID,
                f"Last name '{last}' must be InitCap "
                "(optionally 'LastName(Surname)' with each part InitCap).")
    return S_VALID, ""


def validate_email(email):
    """Validate email structure: local part, '@', domain, and a TLD."""
    email = (email or "").strip()
    if not email:
        return S_MISSING, ""
    if _EMAIL_RE.match(email):
        return S_VALID, ""
    if "@" not in email:
        return S_INVALID, f"'{email}' is missing '@'."
    local, _, domain = email.partition("@")
    if not local:
        return S_INVALID, f"'{email}' has no name before '@'."
    if not domain:
        return S_INVALID, f"'{email}' has no domain after '@'."
    if "." not in domain:
        return S_INVALID, f"'{email}' has no top-level domain (e.g. .com)."
    return S_INVALID, f"'{email}' is not a valid email format."


def validate_phone(phone):
    """Validate mobile number: '+', country code, digits only, valid length."""
    phone = (phone or "").strip()
    if not phone:
        return S_MISSING, ""
    if not phone.startswith("+"):
        return S_INVALID, f"'{phone}' must start with '+' and a country code (e.g. +91…)."
    body = phone[1:]
    if not body.isdigit():
        return S_INVALID, f"'{phone}' must contain only digits after '+' (no spaces/dashes)."
    if not 10 <= len(body) <= 15:
        return S_INVALID, f"'{phone}' has an invalid length ({len(body)} digits after '+')."
    return S_VALID, ""


def _split_code_suffix(tag, short_codes, suffix_keys):
    """
    Split an MMYY-less tag (all letters) into (short_code, optional_segment).
    Returns (None, "") if no known short code can be identified.
      'DF'   -> ('DF', '')          'DSE' -> ('DS', 'E')   (DS + suffix E)
    """
    if tag.upper() in short_codes:
        return tag, ""
    # Peel a known suffix off the end, longest first, leaving a known short code.
    for suf in sorted(suffix_keys, key=len, reverse=True):
        if suf and tag.endswith(suf) and len(tag) > len(suf):
            base = tag[: -len(suf)]
            if base.upper() in short_codes:
                return base, suf
    return None, ""


def validate_tag_name(tag_field, ref):
    """
    Validate a (possibly comma-separated) student batch Tag Name field.

    A Tag Name follows '<ShortCode><MMYY><OptionalSegment>'. Whether the MMYY
    (second segment) is mandatory is driven per course by the
    'is_course_mmyy_required' flag in course_technologies_mapping.json:
      "Y" (default) -> MMYY is required; a short-code-only tag is Invalid.
      "N"           -> MMYY is optional; the tag may be the short code alone
                       (with an allowed optional segment).
    """
    tag_field = (tag_field or "").strip()
    if not tag_field:
        return S_MISSING, ""
    tags = [t.strip() for t in tag_field.split(",") if t.strip()]
    if not tags:
        return S_MISSING, ""

    short_codes   = ref["short_codes"]
    mmyy_required = ref.get("mmyy_required", {})
    suffix_keys   = ref["suffix_keys"]

    reasons = []
    for tag in tags:
        m = _TAG_RE.match(tag)
        if m:
            # MMYY present: <ShortCode><MMYY><OptionalSegment>.
            code, mmyy, opt = m.group(1), m.group(2), m.group(3)
            sub = []
            if code.upper() not in short_codes:
                sub.append(f"short code '{code}' not in course_technologies_mapping.json")
            mm = int(mmyy[:2])
            if not 1 <= mm <= 12:
                sub.append(f"invalid month '{mmyy[:2]}' in MMYY")
            if opt and opt not in suffix_keys:
                sub.append(f"optional segment '{opt}' not configured in batch_tag_suffix_mapping.json")
            if sub:
                reasons.append(f"'{tag}': " + "; ".join(sub))
        elif _TAG_NOMMYY_RE.match(tag):
            # No MMYY: only valid when the course allows the second segment omitted.
            code, opt = _split_code_suffix(tag, short_codes, suffix_keys)
            sub = []
            if code is None:
                sub.append(f"short code not in course_technologies_mapping.json")
            else:
                cu = code.upper()
                if mmyy_required.get(cu, True):
                    sub.append(f"MMYY (second segment) is required for short code '{code}'")
                if opt and opt not in suffix_keys:
                    sub.append(f"optional segment '{opt}' not configured in batch_tag_suffix_mapping.json")
            if sub:
                reasons.append(f"'{tag}': " + "; ".join(sub))
        else:
            reasons.append(f"'{tag}' does not match <ShortCode><MMYY><OptionalSegment>")
    if reasons:
        return S_INVALID, " | ".join(reasons)
    return S_VALID, ""


def validate_course_title(title, ref):
    title = (title or "").strip()
    if not title:
        return S_MISSING, ""
    if title in ref["technology_names"]:
        return S_VALID, ""
    low = title.lower()
    for kw in ref["additional_keywords"]:
        if kw.lower() in low:
            return S_VALID, ""
    return (S_INVALID,
            f"'{title}' does not exactly match a technology_name and contains no additional_keyword "
            f"(technology_mapping.json).")


def _parse_dmy(s):
    try:
        return datetime.strptime(s, "%d-%b-%Y").date()
    except ValueError:
        return None


def validate_course_subtitle(subtitle, future_session_exists):
    subtitle = (subtitle or "").strip()
    if not subtitle:
        return S_MISSING, ""
    m = _SUBTITLE_RE.match(subtitle)
    if not m:
        return (S_INVALID,
                f"'{subtitle}' is not in 'DD-Mon-YYYY To DD-Mon-YYYY' or "
                f"'DD-Mon-YYYY To {CURRENT_DATE_TOKEN}' format.")
    start_raw, end_raw = m.group(1), m.group(2)
    if _parse_dmy(start_raw) is None:
        return S_INVALID, f"Start date '{start_raw}' is not a valid DD-Mon-YYYY date."
    if end_raw != CURRENT_DATE_TOKEN and _parse_dmy(end_raw) is None:
        return S_INVALID, f"End date '{end_raw}' is not a valid DD-Mon-YYYY date."
    if end_raw == CURRENT_DATE_TOKEN and not future_session_exists:
        return S_INVALID, "No future session scheduled."
    return S_VALID, ""


def to_first_initial(name):
    """
    Normalize a full name to 'firstname l' (lowercased) for matching:
      'Mahesh Bhairat' -> 'mahesh b'   |   'KUMAR AARADHYA' -> 'kumar a'
      'Pratik Desarda' -> 'pratik d'   |   'Vaibhav B'      -> 'vaibhav b'
    A single-token name returns just the lowercased token.
    """
    parts = [p for p in re.split(r"\s+", (name or "").strip()) if p]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0].lower()
    return f"{parts[0].lower()} {parts[-1][0].lower()}"


def validate_instructor_tag(tag_field, faculty_initials, match_any=False):
    """
    Validate an instructor tag expected as 'FirstName LastInitial' (e.g. 'Vaibhav B')
    against faculty names converted to the same normalized form. Case-insensitive,
    whitespace-trimmed. A tag may list several instructors separated by ' - ',
    comma or slash (e.g. 'Kumar A - Rahul K - Ankit M').

    match_any=False (instructor tab): EVERY listed name must match.
    match_any=True  (course tab):     it is Valid if ANY listed name matches.
    """
    tag_field = (tag_field or "").strip()
    if not tag_field:
        return S_MISSING, ""
    parts = [p.strip() for p in re.split(r"\s+-\s+|[,/]", tag_field) if p.strip()]
    if not parts:
        return S_MISSING, ""

    matched = [p for p in parts if to_first_initial(p) in faculty_initials]
    missing = [p for p in parts if to_first_initial(p) not in faculty_initials]

    if match_any:
        if matched:
            return S_VALID, ""
        return (S_INVALID,
                "No listed instructor matched Faculty Master Sheet or "
                "old_instructor_technology_mapping.json: "
                + ", ".join(f"'{p}'" for p in parts))
    if missing:
        return (S_INVALID,
                "No 'FirstName LastInitial' match in Faculty Master Sheet or "
                "old_instructor_technology_mapping.json for: "
                + ", ".join(f"'{m}'" for m in missing))
    return S_VALID, ""


# ─────────────────────────────────────────────────────────────────────────────
#  RECORD BUILDERS  (return only INVALID/incomplete rows)
# ─────────────────────────────────────────────────────────────────────────────
_BATCH_MMYY_RE = re.compile(r"[A-Za-z]+(\d{2})(\d{2})")   # <ShortCode><MMYY>...


def _is_excluded_student(name, candidate_name=""):
    """True for test / non-real student accounts (excluded from validation)."""
    for nm in (name, candidate_name):
        low = str(nm or "").strip().lower()
        if low and (low in EXCLUDED_STUDENT_NAMES or low.startswith("test student")):
            return True
    return False


def _is_excluded_batch(batch_name):
    """True if ANY comma-separated batch tag EXACTLY equals a dummy batch name
    (DF / DS), case-insensitively. Real batches like 'DS0525' are NOT excluded."""
    for tag in str(batch_name or "").split(","):
        if tag.strip().upper() in EXCLUDED_BATCH_NAMES:
            return True
    return False


def _should_warn_not_enrolled(batch_name):
    """Decide whether the 'Student is not enrolled in any course.' warning applies
    to a not-enrolled student, based on the batch's MMYY vs today. The MMYY is
    converted to a proper (year, month) value (NOT compared as a raw string/number)
    so year boundaries (e.g. Dec 2026 vs Jan 2027) are handled correctly:
        • future batch  → no warning (course assignment may be pending)
        • past batch    → warning
        • current month → warning only when today is BEFORE the 15th
    Handles comma-separated tags and optional suffixes (e.g. 'ADEAI0626E'); the
    LATEST MMYY is used so a student with any current/future batch isn't warned.
    Blank / unparseable batch → no warning."""
    latest = None                              # (year, month) of latest batch tag
    for tag in str(batch_name or "").split(","):
        m = _BATCH_MMYY_RE.search(tag.strip())
        if not m:
            continue
        mm, yy = int(m.group(1)), int(m.group(2))
        if not (1 <= mm <= 12):
            continue
        ym = (2000 + yy, mm)                   # proper (year, month)
        if latest is None or ym > latest:
            latest = ym
    if latest is None:
        return False
    current = (CURRENT_DATE.year, CURRENT_DATE.month)
    if latest > current:                       # future batch
        return False
    if latest < current:                       # previous batch
        return True
    return CURRENT_DATE.day < 15               # current month → warn only before the 15th


def build_student_rows(student_rows, ref, enrolled_ids=None):
    # enrolled_ids: set of student_ids that appear in ClassLearnerTeacherEnrolled.
    # Only enforced when we actually have enrollment data (non-empty set), so a
    # failed/empty combined-tab read never flags every student as un-enrolled.
    apply_enrollment = bool(enrolled_ids)
    out = []
    total = 0
    for r in student_rows:
        if str(r.get("Is_Deleted", "")).strip().upper() == "Y":
            continue
        # (1) Exclude test / non-real students entirely (not counted either).
        if _is_excluded_student(r.get("student_name", ""), r.get("candidate_name", "")):
            continue
        # (2) Skip only explicitly INACTIVE candidates (is_candidate_active == "N").
        #     "Y", blank, null and empty are all included in validation + counts.
        if str(r.get("is_candidate_active", "")).strip().upper() == "N":
            continue
        # (3) Exclude students in dummy batches (batch short code DF / DS) — not
        #     processed and not counted.
        if _is_excluded_batch(r.get("batch_name", "")):
            continue
        total += 1
        sid    = str(r.get("student_id", "")).strip()
        name   = str(r.get("student_name", "")).strip()
        email  = str(r.get("email", "")).strip()
        phone  = str(r.get("phone", "")).strip()
        batch  = str(r.get("batch_name", "")).strip()
        joined = str(r.get("joined_on", "")).strip()
        pic    = str(r.get("profile_picture", "")).strip()

        name_status, name_desc   = validate_student_name(name)
        email_status, email_desc = validate_email(email)
        phone_status, phone_desc = validate_phone(phone)
        tag_status, tag_desc     = validate_tag_name(batch, ref)

        # Enrollment check: warn "not enrolled in any course" ONLY when BOTH
        #   (1) the student is absent from ClassLearnerTeacherEnrolled, AND
        #   (2) the batch (MMYY) should already have started (past batch).
        # A current/future batch — or an unparseable/blank batch — is not warned,
        # because such students may simply not have been enrolled yet.
        if (apply_enrollment and sid and sid not in enrolled_ids
                and _should_warn_not_enrolled(batch)):
            enroll_msg = "Student is not enrolled in any course."
            name_desc = f"{name_desc} | {enroll_msg}" if name_desc else enroll_msg
            if name_status == S_VALID:
                name_status = S_WARNING

        # Private Note present if ANY note-derived column is populated (the loader
        # parses the student's private note into those columns).
        note_present = any(str(r.get(c, "")).strip() for c in NOTE_PRESENCE_COLUMNS)
        note_status = S_VALID if note_present else S_MISSING
        pic_status  = S_VALID if pic else S_MISSING

        failed = (name_status in (S_MISSING, S_INVALID, S_WARNING)
                  or email_status in (S_MISSING, S_INVALID)
                  or phone_status in (S_MISSING, S_INVALID)
                  or tag_status in (S_MISSING, S_INVALID)
                  or note_status == S_MISSING
                  or pic_status == S_MISSING)
        if not failed:
            continue

        out.append({
            "Student Name": name,
            "Batch Name": batch,
            "Student Name Status": name_status,
            "Student Name Validation Description": name_desc,
            "Email ID Status": email_status,
            "Email ID Validation Description": email_desc,
            "Phone Number Status": phone_status,
            "Phone Number Validation Description": phone_desc,
            "Tag Name Status": tag_status,
            "Tag Name Validation Description": tag_desc,
            "Private Note Status": note_status,
            "Profile Picture Status": pic_status,
            "Joined On": joined,
        })
    out.sort(key=lambda d: d.get("Joined On", ""), reverse=True)   # Joined On DESC
    return out, total


def load_future_session_class_ids():
    """
    Return the set of class_ids that have at least one session scheduled ON or
    AFTER CURRENT_DATE — queried straight from the Wise sessions API, because the
    loaded Google Sheets deliberately EXCLUDE future/pre-scheduled sessions and
    the combined tab's end_time is blank for not-yet-held sessions.

    Returns None if the sessions API is unavailable (so the caller can avoid
    raising false 'No future session scheduled' flags).
    """
    try:
        from pyStudentPaymentClassesStudentEnrolled import fetch_all_sessions
        sessions = fetch_all_sessions(use_cache=True)
    except Exception as exc:                              # noqa: BLE001
        print(f"[WARN] Sessions API unavailable ({exc}); "
              f"'No future session scheduled' will NOT be flagged this run.")
        return None

    future = set()
    for s in sessions:
        if not isinstance(s, dict):
            continue
        cls = s.get("classId") or {}
        cid = cls.get("_id") if isinstance(cls, dict) else (cls if isinstance(cls, str) else "")
        if not cid:
            continue
        raw = str(s.get("start_time") or s.get("startTime") or s.get("startDate") or "").strip()
        if len(raw) >= 10:
            try:
                if datetime.strptime(raw[:10], "%Y-%m-%d").date() >= CURRENT_DATE:
                    future.add(cid)
            except ValueError:
                pass
    print(f"[Sessions] {len(future)} class(es) have a session scheduled on/after {CURRENT_DATE}.")
    return future


def build_course_rows(combined_rows, faculty_initials, ref, future_class_ids=None):
    # future_class_ids: set of class_ids with a future session, or None if the
    # sessions API could not be queried (then we never flag 'No future session').
    # Collapse the per-student exploded tab to unique courses (by class_id) and
    # compute the latest session end_time per course for the future-session check.
    courses = {}
    for r in combined_rows:
        cid = str(r.get("class_id", "")).strip()
        if not cid:
            continue
        # Exclude configured course titles from ALL validation (and from totals).
        if str(r.get("class_name", "")).strip().lower() in COURSE_EXCLUDED_TITLES:
            continue
        end_raw = str(r.get("end_time", "")).strip()
        cur = courses.get(cid)
        if cur is None:
            courses[cid] = {
                "title":    str(r.get("class_name", "")).strip(),
                "subtitle": str(r.get("class_subject", "")).strip(),
                "tag":      str(r.get("Instructor_Name", "")).strip(),
                "created":  str(r.get("created_at", "")).strip(),
                "max_end":  end_raw,
            }
        else:
            if end_raw and end_raw > cur["max_end"]:
                cur["max_end"] = end_raw

    out = []
    for cid, c in courses.items():
        # A course has future sessions if EITHER source confirms it: the latest
        # end_time on the combined tab, OR a future session from the Wise API.
        future = False
        if c["max_end"]:
            try:
                future = datetime.strptime(c["max_end"][:10], "%Y-%m-%d").date() >= CURRENT_DATE
            except ValueError:
                future = False
        if not future and future_class_ids is not None and cid in future_class_ids:
            future = True
        # If the sessions API was unavailable, do not raise a false flag.
        if future_class_ids is None and not future:
            future = True

        title_status, title_desc = validate_course_title(c["title"], ref)
        sub_status, sub_desc      = validate_course_subtitle(c["subtitle"], future)
        tag_status, tag_desc      = validate_instructor_tag(c["tag"], faculty_initials, match_any=True)

        failed = (title_status in (S_MISSING, S_INVALID)
                  or sub_status in (S_MISSING, S_INVALID)
                  or tag_status in (S_MISSING, S_INVALID))
        if not failed:
            continue

        out.append({
            "Course Title": c["title"],
            "Course Title Status": title_status,
            "Course Title Validation Description": title_desc,
            "Course Subtitle": c["subtitle"],
            "Course Subtitle Status": sub_status,
            "Course Subtitle Validation Description": sub_desc,
            "Instructor_Name": c["tag"],
            "Course Tag Name Status": tag_status,
            "Course Tag Name Validation Description": tag_desc,
            "Created On": c["created"],
        })
    out.sort(key=lambda d: d.get("Created On", ""), reverse=True)   # Created On DESC
    return out, len(courses)


# ─────────────────────────────────────────────────────────────────────────────
#  FACULTY ONBOARDING (phone / email source of truth for instructor validation)
# ─────────────────────────────────────────────────────────────────────────────
_ONB_TS_FORMATS = (
    "%m/%d/%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
    "%m/%d/%Y %H:%M", "%d/%m/%Y %H:%M", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d",
)


def _parse_onb_ts(s):
    """Parse a Google-Forms Timestamp; unparseable/blank → datetime.min (oldest)."""
    s = str(s or "").strip()
    if not s:
        return datetime.min
    for fmt in _ONB_TS_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return datetime.min


def _norm_phone(s):
    """Normalise a phone to comparable digits (last 10, ignoring +country/spaces).
    Also drops a trailing '.0' left when a number cell is read as a float."""
    s = re.sub(r"\.0+$", "", str(s or "").strip())   # 8087999440.0 -> 8087999440
    d = re.sub(r"\D", "", s)
    return d[-10:] if len(d) >= 10 else d


def _norm_email(s):
    return str(s or "").strip().lower()


def load_onboarding_records(sheets):
    """Read the New + Old Faculty Onboarding tabs into a flat list of records
    {name, email, phone, timestamp, source}. Tab titles are resolved leniently
    (exact match, else by 'onboarding' + new/old keyword) so minor title drift or
    a missing tab never aborts the run."""
    titles = []
    try:
        meta = sheets.spreadsheets().get(spreadsheetId=ONBOARDING_SHEET_ID).execute()
        titles = [s.get("properties", {}).get("title", "") for s in meta.get("sheets", [])]
    except Exception as exc:                                  # noqa: BLE001
        print(f"[WARN] Could not list onboarding tabs ({exc}).")

    def resolve(want):
        for t in titles:
            if t.strip() == want:
                return t
        key = "new" if want.lower().rstrip().endswith("new") else \
              ("old" if want.lower().rstrip().endswith("old") else "")
        for t in titles:
            tl = t.lower()
            if "onboarding" in tl and key and key in tl:
                return t
        return want

    tabs = [resolve(t) for t in ONBOARDING_TABS] if titles else list(ONBOARDING_TABS)
    records = []
    for tab in tabs:
        try:
            rows = read_tab(sheets, ONBOARDING_SHEET_ID, tab)
        except Exception as exc:                              # noqa: BLE001
            print(f"[WARN] Could not read onboarding tab '{tab}' ({exc}).")
            continue
        for r in rows:
            nm = str(r.get(ONB_NAME_COL, "")).strip()
            if not nm:
                continue
            records.append({
                "name":      nm,
                "email":     str(r.get(ONB_EMAIL_COL, "")).strip(),
                "phone":     str(r.get(ONB_PHONE_COL, "")).strip(),
                "alt_email": str(r.get(ONB_ALT_EMAIL_COL, "")).strip(),
                "alt_phone": str(r.get(ONB_ALT_PHONE_COL, "")).strip(),
                "timestamp": str(r.get(ONB_TIMESTAMP_COL, "")).strip(),
                "source":    tab,
            })
    print(f"[Onboarding] Loaded {len(records)} response(s) from {len(tabs)} tab(s).")
    return records


def build_onboarding_index(records):
    """Map normalised 'firstname l' key → the LATEST onboarding record (by
    Timestamp) across both tabs. Older duplicates are ignored, but the number of
    records seen for the key is tracked ('_count') so the caller can flag that a
    duplicate existed and the latest was selected."""
    idx = {}
    for rec in records:
        key = to_first_initial(rec["name"])
        if not key:
            continue
        ts = _parse_onb_ts(rec["timestamp"])
        cur = idx.get(key)
        if cur is None:
            rec2 = dict(rec); rec2["_ts"] = ts; rec2["_count"] = 1
            idx[key] = rec2
        else:
            cnt = cur["_count"] + 1
            if ts >= cur["_ts"]:
                rec2 = dict(rec); rec2["_ts"] = ts; rec2["_count"] = cnt
                idx[key] = rec2
            else:
                cur["_count"] = cnt
    return idx


def _norm_name(s):
    """Normalised full name: trimmed, lower-cased, single-spaced."""
    return " ".join(str(s or "").strip().lower().split())


def build_onboarding_pool(records):
    """Collapse onboarding records to ONE entity per full name, keeping the LATEST
    by Timestamp (handles duplicate submissions). Each entity carries a full-name
    key and an initial key ('firstname l') for the two-tier match, plus its contact
    fields. Different names that share an initial key stay as separate entities."""
    by_full = {}
    for rec in records:
        fk = _norm_name(rec["name"])
        if not fk:
            continue
        ts = _parse_onb_ts(rec["timestamp"])
        cur = by_full.get(fk)
        if cur is None or ts >= cur["_ts"]:
            e = dict(rec)
            e["_ts"]      = ts
            e["full_key"] = fk
            e["init_key"] = to_first_initial(rec["name"])
            by_full[fk] = e
    return list(by_full.values())


def resolve_instructor_matches(occurrences, pool):
    """Two-tier, CONSUME-ONCE matching of instructor names to onboarding entities.

    `occurrences` : ordered list of (key, instructor_name).
    Returns        : {key: entity} for the ones that matched.

    Step 1 — EXACT full-name match ('FirstName LastName' ↔ 'FirstName LastName'),
             applied across ALL instructors first; a matched entity is consumed.
    Step 2 — INITIAL match ('FirstName + Last Initial') for the still-unmatched
             instructors, against the REMAINING (unconsumed) entities only.
    When several entities qualify for one instructor, the LATEST (by Timestamp)
    is chosen. Each onboarding entity can be matched to at most one instructor.
    """
    import difflib
    from collections import defaultdict
    by_full, by_init = defaultdict(list), defaultdict(list)
    for i, e in enumerate(pool):
        by_full[e["full_key"]].append(i)
        by_init[e["init_key"]].append(i)

    consumed, result = set(), {}

    # Step 1 — exact full name (highest priority), consume on match.
    for key, name in occurrences:
        cands = [i for i in by_full.get(_norm_name(name), []) if i not in consumed]
        if cands:
            best = max(cands, key=lambda i: pool[i]["_ts"])
            consumed.add(best)
            result[key] = pool[best]

    # Step 2 — initial-based fallback for the rest. Instructors and onboarding
    # entities that share an initial key ('FirstName L') are paired by BEST
    # full-name similarity first, so a near-identical name (e.g. 'Sumit Chaterjee'
    # ↔ 'Sumit Chatterjee') wins over an unrelated one ('Sumit Kumar Choubey'),
    # avoiding the cross-match. Timestamp (latest) breaks ties. Consume-once.
    def _sim(a, b):
        return difflib.SequenceMatcher(None, _norm_name(a), _norm_name(b)).ratio()

    candidate_pairs = []
    for key, name in occurrences:
        if key in result:
            continue
        for i in by_init.get(to_first_initial(name), []):
            if i in consumed:
                continue
            candidate_pairs.append((_sim(name, pool[i]["name"]), pool[i]["_ts"], key, i))
    # Highest similarity first, then latest Timestamp.
    candidate_pairs.sort(key=lambda p: (p[0], p[1]), reverse=True)
    for _sim_v, _ts, key, i in candidate_pairs:
        if key in result or i in consumed:
            continue
        consumed.add(i)
        result[key] = pool[i]

    return result


# ─────────────────────────────────────────────────────────────────────────────
#  INSTRUCTOR VALIDATION  (per-check rows; failures only)
# ─────────────────────────────────────────────────────────────────────────────
def build_instructor_rows(instructor_rows, onb_pool):
    """Validate each ACTIVE instructor row using STRICT position-based mapping
    (instructor_name_k ↔ alternative_contact_number_k ↔ alternative_email_k) against
    the latest Faculty Onboarding record. Emits one row per flagged check with the
    detailed columns below. Returns (rows, processed_instructors).

    Matching is by 'FirstName LastName' or 'FirstName L' (to_first_initial), case-
    insensitive and trimmed. Phone/Email in the onboarding record are the source of
    truth; a slot is NEVER compared against a different sequence's contact details.
    """
    out = []
    total = 0

    # ── Phase A: collect the processable instructor contexts (active, non-test) ─
    contexts = []
    for r in instructor_rows:
        if str(r.get("Is_Active", "Y")).strip().upper() != "Y":
            continue
        iid    = str(r.get("instructor_id", "")).strip()
        detail = str(r.get("instructor_details", "")).strip()
        names  = [str(r.get(f"instructor_name_{i}", "")).strip() for i in range(1, 6)]
        # 1. Ignore test records (e.g. "Tester") + blank name columns.
        if any(nm.strip().lower() in EXCLUDED_INSTRUCTOR_NAMES
               for nm in [detail, *names] if nm):
            continue
        # 2. Count assigned instructors (non-blank instructor_name_1..5). The
        #    summary counts each INSTRUCTOR ENTRY validated, not the instructor_id
        #    row (a row with 3 names = 3 validations). A row with no names still
        #    counts as one processed record (flagged as an invalid assignment).
        assigned = [(i, names[i - 1]) for i in range(1, 6) if names[i - 1]]
        total += len(assigned) or 1
        contexts.append({"r": r, "iid": iid, "detail": detail, "assigned": assigned})

    # ── Phase B: two-tier, consume-once matching across ALL instructors ───────
    #    (exact full name first, then initial-based on the remaining entities).
    occurrences = [((ci, slot), nm)
                   for ci, ctx in enumerate(contexts)
                   for slot, nm in ctx["assigned"]]
    matches = resolve_instructor_matches(occurrences, onb_pool)

    # ── Phase C: validation (per-check failure rows) ──────────────────────────
    for ci, ctx in enumerate(contexts):
        r, iid, detail, assigned = ctx["r"], ctx["iid"], ctx["detail"], ctx["assigned"]

        def add(name_for_row, vtype, expected, actual, msg, status="Fail", _iid=iid):
            out.append({
                "Instructor ID": _iid,
                "Instructor Name": name_for_row,
                "Validation Type": vtype,
                "Expected Value": expected,
                "Actual Value": actual,
                "Validation Status": status,
                "Detailed Validation Message": msg,
            })

        if not assigned:
            add(detail or iid, "Instructor Assignment",
                "At least one instructor_name_1..5", "(none)",
                "Invalid instructor assignment — no instructor names present "
                "(missing instructor details).")
            continue

        # Each slot's matched onboarding entity (position-based; used also for the
        # wrong-sequence cross-checks). Matching honoured exact→initial + consume.
        onb = {slot: matches.get((ci, slot)) for slot, nm in assigned}

        for slot, nm in assigned:
            # STRICT position-based contact fields for this instructor.
            phone_val = str(r.get(f"alternative_contact_number_{slot}", "")).strip()
            email_val = str(r.get(f"alternative_email_{slot}", "")).strip()
            rec = onb[slot]

            # 8. Required contact detail present in the Instructor sheet.
            if not phone_val:
                add(nm, "Phone Number Missing (Instructor)",
                    f"alternative_contact_number_{slot} not blank", "(blank)",
                    f"Required phone number alternative_contact_number_{slot} is "
                    f"blank / missing for '{nm}'.")
            if not email_val:
                add(nm, "Email ID Missing (Instructor)",
                    f"alternative_email_{slot} not blank", "(blank)",
                    f"Required email alternative_email_{slot} is blank / missing "
                    f"for '{nm}'.")

            # 5. Match against onboarding responses.
            if rec is None:
                add(nm, "Instructor Not Found",
                    "Match in Faculty Onboarding (New/Old)", "Not found",
                    f"Instructor '{nm}' not found in either onboarding sheet.")
                continue

            # Onboarding source of truth — PRIMARY and ALTERNATIVE (fallback).
            onb_phone     = str(rec.get("phone", "")).strip()
            onb_alt_phone = str(rec.get("alt_phone", "")).strip()
            onb_email     = str(rec.get("email", "")).strip()
            onb_alt_email = str(rec.get("alt_email", "")).strip()

            # 7/11. Phone: PASS if the instructor value matches the PRIMARY OR the
            #       ALTERNATIVE onboarding phone; FAIL only if it matches neither.
            if not onb_phone and not onb_alt_phone:
                add(nm, "Phone Number Missing (Onboarding)",
                    "Phone Number or Alternative Phone Number", "(both blank)",
                    f"Phone Number and Alternative Phone Number are both missing in "
                    f"the Faculty Onboarding record for '{nm}'.")
            elif phone_val and not _matches_any(
                    phone_val, _split_phones(onb_phone) + _split_phones(onb_alt_phone), "phone"):
                wrong = _wrong_sequence_slot(slot, phone_val, assigned, onb, "phone")
                exp = _expected_str(onb_phone, onb_alt_phone)
                if wrong:
                    add(nm, "Phone Number Wrong Sequence", exp, phone_val,
                        f"alternative_contact_number_{slot} for '{nm}' matches neither "
                        f"'{nm}'s primary nor alternative onboarding phone, but matches "
                        f"instructor_name_{wrong} — contact mapped to the wrong sequence.")
                else:
                    add(nm, "Phone Number Mismatch", exp, phone_val,
                        f"Phone number mismatch for '{nm}': alternative_contact_number_"
                        f"{slot} matches neither the primary Phone Number nor the "
                        f"Alternative Phone Number in the latest onboarding record.")

            # 7/11. Email: PASS if it matches the PRIMARY OR the ALTERNATIVE onboarding
            #       email; FAIL only if it matches neither.
            if not onb_email and not onb_alt_email:
                add(nm, "Email ID Missing (Onboarding)",
                    "Email ID or Alternative Email ID", "(both blank)",
                    f"Email ID and Alternative Email ID are both missing in the "
                    f"Faculty Onboarding record for '{nm}'.")
            elif email_val and not _matches_any(
                    email_val, _split_emails(onb_email) + _split_emails(onb_alt_email), "email"):
                wrong = _wrong_sequence_slot(slot, email_val, assigned, onb, "email")
                exp = _expected_str(onb_email, onb_alt_email)
                if wrong:
                    add(nm, "Email ID Wrong Sequence", exp, email_val,
                        f"alternative_email_{slot} for '{nm}' matches neither '{nm}'s "
                        f"primary nor alternative onboarding email, but matches "
                        f"instructor_name_{wrong} — email mapped to the wrong sequence.")
                else:
                    add(nm, "Email ID Mismatch", exp, email_val,
                        f"Email address mismatch for '{nm}': alternative_email_{slot} "
                        f"matches neither the primary Email ID nor the Alternative "
                        f"Email ID in the latest onboarding record.")

    # Only failed records are included in the output.
    out = [d for d in out if d.get("Validation Status") == "Fail"]
    out.sort(key=lambda d: (d.get("Instructor Name", "").lower(),
                            d.get("Validation Type", "")))
    return out, total


def _split_emails(value):
    """Onboarding Email ID / Alternative Email ID cells may hold several addresses
    separated by '/'. Split into trimmed individual addresses (any count)."""
    return [e.strip() for e in str(value or "").split("/") if e.strip()]


def _split_phones(value):
    """Onboarding Phone Number / Alternative Phone Number cells may hold several
    numbers separated by '/'. Split into trimmed individual numbers (any count)."""
    return [p.strip() for p in str(value or "").split("/") if p.strip()]


def _matches_any(value, candidates, kind):
    """True if `value` equals ANY non-blank candidate under phone/email normalisation."""
    norm = _norm_phone if kind == "phone" else _norm_email
    v = norm(value)
    if not v:
        return False
    return any(cand and norm(cand) == v for cand in candidates)


def _expected_str(primary, alternative):
    """Human-readable 'primary / alternative' for the Expected Value column."""
    return f"{primary or '(blank)'} / {alternative or '(blank)'}"


def _wrong_sequence_slot(slot, value, assigned, onb, kind):
    """Return the OTHER slot number whose onboarding contact (`kind`='phone'|'email')
    — primary OR alternative — equals `value` (value placed in the wrong sequence).
    Else None."""
    norm = _norm_phone if kind == "phone" else _norm_email
    target = norm(value)
    if not target:
        return None
    alt_key = "alt_phone" if kind == "phone" else "alt_email"
    for j, _nm in assigned:
        if j == slot:
            continue
        rec = onb.get(j)
        if not rec:
            continue
        if kind == "email":
            cands = _split_emails(rec.get("email", "")) + _split_emails(rec.get("alt_email", ""))
        else:
            cands = _split_phones(rec.get("phone", "")) + _split_phones(rec.get("alt_phone", ""))
        if any(norm(c) == target for c in cands):
            return j
    return None


def _instructor_tab_stats(total_entries, rows):
    """KPI stats for the instructor validation log. `total_entries` is the number of
    instructor entries validated; `rows` are per-check FAILURES."""
    # `total_entries` = number of instructor entries validated (non-blank names).
    # A failing ENTRY is a distinct (Instructor ID, Instructor Name) among failures.
    fail_rows = [r for r in rows if r.get("Validation Status") == "Fail"]
    failed = len({(r.get("Instructor ID", ""), r.get("Instructor Name", "")) for r in fail_rows})
    valid  = max(total_entries - failed, 0)
    pv = (valid / total_entries * 100) if total_entries else 0.0
    return {"total": total_entries, "valid": valid, "invalid": failed,
            "missing": 0, "warning": 0, "failed": len(fail_rows), "pv": pv, "pi": 0.0}


# Output column order per sheet (also the header row).
STUDENT_COLUMNS = [
    "Student Name", "Batch Name",
    "Student Name Status", "Student Name Validation Description",
    "Email ID Status", "Email ID Validation Description",
    "Phone Number Status", "Phone Number Validation Description",
    "Tag Name Status", "Tag Name Validation Description",
    "Private Note Status", "Profile Picture Status", "Joined On",
]
COURSE_COLUMNS = [
    "Course Title", "Course Title Status", "Course Title Validation Description",
    "Course Subtitle", "Course Subtitle Status", "Course Subtitle Validation Description",
    "Instructor_Name",
    "Course Tag Name Status", "Course Tag Name Validation Description", "Created On",
]

INSTRUCTOR_COLUMNS = [
    "Instructor ID", "Instructor Name", "Validation Type",
    "Expected Value", "Actual Value", "Validation Status",
    "Detailed Validation Message",
]


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL (local artifact)
# ─────────────────────────────────────────────────────────────────────────────
# ── Professional colour palette (shared by Excel + Google Sheet) ─────────────
C_NAV     = "1A2E5A"   # title / column-header bar
C_BLUEMID = "2E5AA8"   # subtitle band
C_WHITE   = "FFFFFF"
C_SUBHDR  = "DCE6F4"   # KPI: processed
C_ZEBRA   = "EEF3FA"   # alternate data row
C_BORDER  = "B7C2D0"
C_GREEN   = "C6EFCE";  GREEN_TX = "1B5E20"   # Valid
C_RED     = "FFC7CE";  RED_TX   = "9C0006"   # Invalid
C_AMBER   = "FFEB9C";  AMBER_TX = "7A4F01"   # Missing
C_WARN    = "FFD9A0";  WARN_TX  = "8A4B08"   # Warning

_FONT        = "Calibri"
_THIN        = Side(style="thin", color=C_BORDER)
_CELL_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _split_ranges(n, parts):
    """Split columns 1..n into `parts` contiguous (start, end) blocks for the KPI strip."""
    cuts = [round(n * k / parts) for k in range(parts + 1)]
    out = []
    for i in range(parts):
        start = cuts[i] + 1
        out.append((start, max(cuts[i + 1], start)))
    return out


def _hex_rgb(hex_color):
    """'1A2E5A' -> {'red':..,'green':..,'blue':..} (0-1 floats) for the Sheets API."""
    h = hex_color.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}


def _kpi_cells(stats):
    """The KPI tiles: (text, bg_hex, text_hex)."""
    return [
        (f"Records Processed:  {stats['total']}",                 C_SUBHDR, C_NAV),
        (f"Valid:  {stats['valid']}   ({stats['pv']:.1f}%)",      C_GREEN,  GREEN_TX),
        (f"Invalid:  {stats['invalid']}",                         C_RED,    RED_TX),
        (f"Missing:  {stats['missing']}",                         C_AMBER,  AMBER_TX),
        (f"Warning:  {stats['warning']}",                         C_WARN,   WARN_TX),
    ]


def _tab_stats(total, rows):
    """Record-level counts for the header. Each flagged record is bucketed once by
    severity: Invalid (any Invalid field) > Missing (any Missing) > Warning."""
    invalid = sum(1 for r in rows if any(v == S_INVALID for v in r.values()))
    missing = sum(1 for r in rows
                  if not any(v == S_INVALID for v in r.values())
                  and any(v == S_MISSING for v in r.values()))
    warning = sum(1 for r in rows
                  if not any(v == S_INVALID for v in r.values())
                  and not any(v == S_MISSING for v in r.values())
                  and any(v == S_WARNING for v in r.values()))
    valid  = max(total - len(rows), 0)
    failed = len(rows)
    pv = (valid / total * 100) if total else 0.0
    pi = (failed / total * 100) if total else 0.0
    return {"total": total, "valid": valid, "invalid": invalid, "missing": missing,
            "warning": warning, "failed": failed, "pv": pv, "pi": pi}


def _xl_font(bold=False, size=10, color="000000"):
    return Font(name=_FONT, bold=bold, size=size, color=color)


def _col_kind(header):
    """Classify a column for alignment / width purposes."""
    h = header.lower()
    if h.endswith("status"):
        return "status"
    if "description" in h or "remarks" in h or "message" in h:
        return "desc"
    if h in ("joined on", "created on"):
        return "date"
    return "text"


def _write_xlsx_sheet(ws, report_title, columns, rows, stats):
    n = len(columns)
    last = get_column_letter(n)

    # ── Report header block ──────────────────────────────────────────────────
    # Row 1 — title bar
    ws.merge_cells(f"A1:{last}1")
    t = ws.cell(row=1, column=1, value=report_title)
    t.font = _xl_font(bold=True, size=16, color=C_WHITE)
    t.fill = PatternFill("solid", fgColor=C_NAV)
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 38

    # Row 2 — subtitle band
    ws.merge_cells(f"A2:{last}2")
    s = ws.cell(row=2, column=1, value=(
        f"Data Quality Validation Report        •        Generated: {GENERATED_AT} IST"))
    s.font = _xl_font(bold=True, size=10, color=C_WHITE)
    s.fill = PatternFill("solid", fgColor=C_BLUEMID)
    s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

    # Row 3 — KPI tiles (Processed / Valid / Invalid / Missing / Warning), colour-highlighted
    _kpis = _kpi_cells(stats)
    for (text, bg, tx), (c0, c1) in zip(_kpis, _split_ranges(n, len(_kpis))):
        if c1 > c0:
            ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c1)
        cell = ws.cell(row=3, column=c0, value=text)
        cell.font = _xl_font(bold=True, size=11, color=tx)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _CELL_BORDER
    ws.row_dimensions[3].height = 26

    ws.row_dimensions[4].height = 6      # spacer

    # ── Column header row ────────────────────────────────────────────────────
    hdr_row = 5
    for col, h in enumerate(columns, 1):
        c = ws.cell(row=hdr_row, column=col, value=h)
        c.font = _xl_font(bold=True, size=10, color=C_WHITE)
        c.fill = PatternFill("solid", fgColor=C_NAV)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = _CELL_BORDER
    ws.row_dimensions[hdr_row].height = 30

    # ── Data rows: zebra + status colour + alignment + borders ───────────────
    for i, rowd in enumerate(rows):
        r = hdr_row + 1 + i
        zebra = C_ZEBRA if (i % 2) else C_WHITE
        for col, h in enumerate(columns, 1):
            val = str(rowd.get(h, ""))
            c = ws.cell(row=r, column=col, value=val)
            c.border = _CELL_BORDER
            center = _col_kind(h) in ("status", "date")
            c.alignment = Alignment(
                horizontal="center" if center else "left",
                vertical="center" if center else "top",
                wrap_text=not center)
            if val == S_VALID:
                c.fill = PatternFill("solid", fgColor=C_GREEN); c.font = _xl_font(bold=True, color=GREEN_TX)
            elif val == S_INVALID:
                c.fill = PatternFill("solid", fgColor=C_RED);   c.font = _xl_font(bold=True, color=RED_TX)
            elif val == S_MISSING:
                c.fill = PatternFill("solid", fgColor=C_AMBER); c.font = _xl_font(bold=True, color=AMBER_TX)
            elif val == S_WARNING:
                c.fill = PatternFill("solid", fgColor=C_WARN);  c.font = _xl_font(bold=True, color=WARN_TX)
            else:
                c.fill = PatternFill("solid", fgColor=zebra);   c.font = _xl_font()

    last_row = max(hdr_row, hdr_row + len(rows))
    ws.auto_filter.ref = f"A{hdr_row}:{last}{last_row}"
    ws.freeze_panes = f"A{hdr_row + 1}"

    # ── Content-aware column widths (wider for descriptions) ─────────────────
    for ci, h in enumerate(columns, 1):
        kind = _col_kind(h)
        content_max = len(h)
        for rowd in rows:
            for seg in str(rowd.get(h, "")).split("\n"):
                content_max = max(content_max, len(seg))
        if kind == "desc":
            width = min(60, max(30, content_max + 2))
        elif kind == "status":
            width = max(13, min(16, content_max + 2))
        elif kind == "date":
            width = 20
        else:
            width = max(14, min(42, content_max + 2))
        ws.column_dimensions[get_column_letter(ci)].width = width


def save_local_xlsx(student, course, instructor):
    """DISABLED (Google Drive only) — no local Excel copy is written.

    Kept as a no-op so any caller stays intact. The validation results are
    delivered to the Google Sheet output instead.
    """
    return


# ─────────────────────────────────────────────────────────────────────────────
#  GOOGLE SHEET OUTPUT  (fresh / overwrite each run)
# ─────────────────────────────────────────────────────────────────────────────
class OutputSetupError(RuntimeError):
    """Raised when the output Google Sheet cannot be located or created."""


def resolve_output_spreadsheet(drive):
    """Resolve the target spreadsheet id (see OUTPUT_* config for the order)."""
    # 1) Explicit id wins — write straight to an existing, shared sheet.
    if OUTPUT_SPREADSHEET_ID.strip():
        print(f"[Output] Using configured OUTPUT_SPREADSHEET_ID ({OUTPUT_SPREADSHEET_ID.strip()}).")
        return OUTPUT_SPREADSHEET_ID.strip()

    # 2) Find an existing sheet by name inside the target folder.
    try:
        q = (f"name = '{OUTPUT_SPREADSHEET_NAME}' and "
             f"'{OUTPUT_DRIVE_FOLDER_ID}' in parents and "
             f"mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false")
        res = drive.files().list(
            q=q, fields="files(id, name)", supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = res.get("files", [])
        if files:
            print(f"[Output] Reusing '{OUTPUT_SPREADSHEET_NAME}' ({files[0]['id']}).")
            return files[0]["id"]
    except Exception as exc:                              # noqa: BLE001
        print(f"[WARN] Could not search the output folder ({exc}).")

    # 3) Last resort: create it in the folder (only succeeds on a Shared Drive).
    try:
        meta = {
            "name": OUTPUT_SPREADSHEET_NAME,
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "parents": [OUTPUT_DRIVE_FOLDER_ID],
        }
        created = drive.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
        print(f"[Output] Created '{OUTPUT_SPREADSHEET_NAME}' ({created['id']}).")
        return created["id"]
    except Exception as exc:                              # noqa: BLE001
        raise OutputSetupError(
            "The service account could not create the output spreadsheet. A service "
            "account has no Drive storage, so it cannot create files in a personal "
            "'My Drive' folder. Fix ONE of the following and re-run:\n"
            f"   (a) Create a blank Google Sheet named '{OUTPUT_SPREADSHEET_NAME}' "
            "inside the target folder, share it (or the folder) with the service "
            "account as Editor; or\n"
            "   (b) Set OUTPUT_SPREADSHEET_ID (top of this script) to an existing "
            "sheet that is shared with the service account; or\n"
            "   (c) Move the folder onto a Google Shared Drive and add the service "
            "account as a member.\n"
            "   Service account email: see 'client_email' in "
            "config_files/service_account.json.\n"
            f"   Underlying error: {exc}"
        ) from exc


def _sheet_id_for_tab(sheets, spreadsheet_id, tab):
    """Ensure a tab exists; return its numeric sheetId. Renames a leftover 'Sheet1' if present."""
    meta = sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta.get("sheets", [])}
    if tab in existing:
        return existing[tab]
    # Reuse a default 'Sheet1' for the first tab instead of leaving it around.
    if "Sheet1" in existing and len(existing) == 1:
        sid = existing["Sheet1"]
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": sid, "title": tab}, "fields": "title"}}]},
        ).execute()
        return sid
    resp = sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": tab}}}]},
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def _safe_batch(sheets, spreadsheet_id, requests, label=""):
    """Run a batchUpdate; never fatal — formatting issues must not lose the data."""
    if not requests:
        return
    try:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    except Exception as exc:                              # noqa: BLE001
        print(f"[WARN] Formatting step '{label}' skipped: {exc}")


def _border():
    return {"style": "SOLID", "color": {"red": 0.72, "green": 0.76, "blue": 0.82}}


def _col_px(header, rows):
    """Pixel width for a Google Sheet column (content-aware, kind-capped)."""
    kind = _col_kind(header)
    content = len(header)
    for r in rows:
        for seg in str(r.get(header, "")).split("\n"):
            content = max(content, len(seg))
    if kind == "desc":
        chars = min(58, max(30, content))
    elif kind == "status":
        chars = 15
    elif kind == "date":
        chars = 20
    else:
        chars = min(40, max(14, content))
    return int(chars * 7 + 18)


def write_google_tab(sheets, spreadsheet_id, tab, columns, rows, stats, report_title):
    sid = _sheet_id_for_tab(sheets, spreadsheet_id, tab)
    n   = len(columns)
    hdr_idx  = 4                              # 0-based header row index (5th row)
    last_row = hdr_idx + 1 + len(rows)

    subtitle = f"Data Quality Validation Report        •        Generated: {GENERATED_AT} IST"
    kpis = _kpi_cells(stats)
    kpi_ranges = _split_ranges(n, len(kpis))
    kpi_row = [""] * n
    for (text, _bg, _tx), (c0, c1) in zip(kpis, kpi_ranges):
        kpi_row[c0 - 1] = text            # value sits in the first cell of each block
    matrix = [[report_title], [subtitle], kpi_row, [], list(columns)]
    for r in rows:
        matrix.append([str(r.get(c, "")) for c in columns])

    # Fresh overwrite: clear values then write.
    sheets.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=tab).execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=f"{tab}!A1",
        valueInputOption="RAW", body={"values": matrix}).execute()

    # Fetch existing decorations so re-runs stay idempotent (no stacking).
    merges, cf_count, bandings = [], 0, []
    try:
        meta = sheets.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId),merges,conditionalFormats,bandedRanges)").execute()
        for s in meta.get("sheets", []):
            if s.get("properties", {}).get("sheetId") == sid:
                merges   = s.get("merges") or []
                cf_count = len(s.get("conditionalFormats") or [])
                bandings = s.get("bandedRanges") or []
                break
    except Exception:                                    # noqa: BLE001
        pass

    NAVY    = _hex_rgb(C_NAV)
    BLUEMID = _hex_rgb(C_BLUEMID)
    WHITE   = {"red": 1, "green": 1, "blue": 1}

    def rng(r0, r1, c0=0, c1=n):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": c0, "endColumnIndex": c1}

    # ── 1) Clear previous decorations ────────────────────────────────────────
    clear = [{"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}} for _ in range(cf_count)]
    clear += [{"deleteBanding": {"bandedRangeId": b["bandedRangeId"]}}
              for b in bandings if b.get("bandedRangeId") is not None]
    clear += [{"unmergeCells": {"range": m}} for m in merges]
    _safe_batch(sheets, spreadsheet_id, clear, "clear")

    # ── 2) Structural styling (header block, header row, freeze, filter, borders) ─
    struct = [
        # Title bar (row 0) and subtitle bar (row 1), merged across all columns.
        {"mergeCells": {"range": rng(0, 1), "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": rng(1, 2), "mergeType": "MERGE_ALL"}},
        {"repeatCell": {"range": rng(0, 1),
            "cell": {"userEnteredFormat": {"backgroundColor": NAVY, "horizontalAlignment": "LEFT",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True, "fontSize": 16, "foregroundColor": WHITE}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}},
        {"repeatCell": {"range": rng(1, 2),
            "cell": {"userEnteredFormat": {"backgroundColor": BLUEMID, "horizontalAlignment": "LEFT",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": WHITE}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}},
        {"repeatCell": {"range": rng(hdr_idx, hdr_idx + 1),
            "cell": {"userEnteredFormat": {"backgroundColor": NAVY, "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP",
                "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": WHITE}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)"}},
        {"repeatCell": {"range": rng(hdr_idx + 1, max(last_row, hdr_idx + 2)),
            "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"}},
            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": hdr_idx + 1}},
            "fields": "gridProperties.frozenRowCount"}},
        {"setBasicFilter": {"filter": {"range": rng(hdr_idx, max(last_row, hdr_idx + 1))}}},
        {"updateBorders": {"range": rng(hdr_idx, max(last_row, hdr_idx + 1)),
            "top": _border(), "bottom": _border(), "left": _border(), "right": _border(),
            "innerHorizontal": _border(), "innerVertical": _border()}},
    ]
    # KPI tiles on row index 2 — merged colour blocks (Processed / Valid / Invalid / Missing)
    for (text, bg_hex, tx_hex), (c0, c1) in zip(kpis, kpi_ranges):
        if c1 > c0:
            struct.append({"mergeCells": {"range": rng(2, 3, c0 - 1, c1), "mergeType": "MERGE_ALL"}})
        struct.append({"repeatCell": {"range": rng(2, 3, c0 - 1, c1),
            "cell": {"userEnteredFormat": {"backgroundColor": _hex_rgb(bg_hex),
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": _hex_rgb(tx_hex)}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
    # explicit per-column widths + centre-align status/date columns
    for ci, h in enumerate(columns):
        struct.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": ci, "endIndex": ci + 1},
            "properties": {"pixelSize": _col_px(h, rows)}, "fields": "pixelSize"}})
        if _col_kind(h) in ("status", "date"):
            struct.append({"repeatCell": {
                "range": rng(hdr_idx + 1, max(last_row, hdr_idx + 2), ci, ci + 1),
                "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                "fields": "userEnteredFormat.horizontalAlignment"}})
    _safe_batch(sheets, spreadsheet_id, struct, "structure")

    # ── 3) Zebra banding over data rows ──────────────────────────────────────
    if rows:
        _safe_batch(sheets, spreadsheet_id, [{"addBanding": {"bandedRange": {
            "range": rng(hdr_idx + 1, last_row),
            "rowProperties": {"firstBandColor": WHITE,
                              "secondBandColor": {"red": 0.933, "green": 0.953, "blue": 0.980}}}}}],
            "banding")

    # ── 4) Status colour coding (Valid/Invalid/Missing/Warning) — highest priority ─
    data_range = rng(hdr_idx + 1, max(last_row, hdr_idx + 2))
    cf = [{"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [data_range],
            "booleanRule": {
                "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": text}]},
                "format": {"backgroundColor": color, "textFormat": {"bold": True}}}}}}
          for text, color in ((S_VALID,   {"red": 0.776, "green": 0.937, "blue": 0.804}),
                              (S_INVALID, {"red": 1.0,   "green": 0.780, "blue": 0.812}),
                              (S_MISSING, {"red": 1.0,   "green": 0.922, "blue": 0.612}),
                              (S_WARNING, _hex_rgb(C_WARN)),
                              # Instructor validation log uses Pass / Fail.
                              ("Pass",    {"red": 0.776, "green": 0.937, "blue": 0.804}),
                              ("Fail",    {"red": 1.0,   "green": 0.780, "blue": 0.812}))]
    _safe_batch(sheets, spreadsheet_id, cf, "status-colours")

    print(f"[Output] Wrote tab '{tab}' ({len(rows)} record(s)).")


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main() -> int:
    print("=" * 78)
    print(f"  IntelliBI Wise Data Validation  |  {GENERATED_AT} IST")
    print("=" * 78)

    ref = load_reference_data()
    print(f"[Config] short_codes={len(ref['short_codes'])} technology_names={len(ref['technology_names'])} "
          f"keywords={len(ref['additional_keywords'])} suffixes={len(ref['suffix_keys'])} "
          f"old_instructors={len(ref['old_instructor_names'])}")

    sheets, drive = google_services()

    # ── Input data ────────────────────────────────────────────────────────────
    student_src    = read_tab(sheets, SOURCE_SHEET_ID, STUDENTS_TAB)
    combined_src   = read_tab(sheets, SOURCE_SHEET_ID, COMBINED_TAB)
    instructor_src = read_tab(sheets, SOURCE_SHEET_ID, INSTRUCTOR_TAB)
    print(f"[Input] students={len(student_src)} class-rows={len(combined_src)} instructors={len(instructor_src)}")

    faculty_names, faculty_lc = load_faculty_names(sheets)
    # Convert faculty (Source 1) + old-instructor (Source 2) names to the
    # normalized 'firstname l' form used to match Course/Instructor tag names.
    faculty_initials = {to_first_initial(n) for n in faculty_names if to_first_initial(n)}
    faculty_initials |= {to_first_initial(n) for n in ref["old_instructor_names"]
                         if to_first_initial(n)}
    print(f"[Faculty] {len(faculty_initials)} normalized 'FirstName LastInitial' name(s) for matching.")

    # Faculty Onboarding responses → pool of entities (latest per name) for the
    # two-tier consume-once instructor matching (exact full name → initial).
    onb_pool = build_onboarding_pool(load_onboarding_records(sheets))
    print(f"[Onboarding] {len(onb_pool)} unique faculty entity(ies) for matching.")

    # ── Validate (invalid-only rows + totals for % summary) ───────────────────
    future_class_ids = load_future_session_class_ids()

    # Students enrolled in ≥1 course = those present in ClassLearnerTeacherEnrolled.
    enrolled_ids = {str(r.get("student_id", "")).strip() for r in combined_src
                    if str(r.get("student_id", "")).strip()}
    print(f"[Enrollment] {len(enrolled_ids)} student(s) enrolled in at least one course.")

    student_rows,    student_total    = build_student_rows(student_src, ref, enrolled_ids)
    course_rows,     course_total     = build_course_rows(combined_src, faculty_initials, ref, future_class_ids)
    instructor_rows, instructor_total = build_instructor_rows(instructor_src, onb_pool)

    student_stats    = _tab_stats(student_total,    student_rows)
    course_stats     = _tab_stats(course_total,     course_rows)
    instructor_stats = _instructor_tab_stats(instructor_total, instructor_rows)

    def _log(label, s):
        print(f"[Validate] {label:<11} -> processed={s['total']} valid={s['valid']} "
              f"invalid={s['invalid']} missing={s['missing']}  "
              f"(Valid {s['pv']:.1f}% / Invalid {s['pi']:.1f}%)")
    _log("Students", student_stats)
    _log("Courses", course_stats)
    _log("Instructors", instructor_stats)

    T_STUDENT    = "IntelliBI Wise  —  Student Data Validation"
    T_COURSE     = "IntelliBI Wise  —  Course Data Validation"
    T_INSTRUCTOR = "IntelliBI Wise  —  Instructor Data Validation"

    # ── Output (1): local Excel artifact — DISABLED (Google Drive only) ───────
    #    Previously wrote a local .xlsx copy here; the Google Sheet below is now
    #    the sole output destination.

    # ── Output (2): Google Sheet (fresh / overwrite). Non-fatal on setup gaps ─
    try:
        spreadsheet_id = resolve_output_spreadsheet(drive)
        write_google_tab(sheets, spreadsheet_id, TAB_STUDENT,    STUDENT_COLUMNS,    student_rows,    student_stats,    T_STUDENT)
        write_google_tab(sheets, spreadsheet_id, TAB_COURSE,     COURSE_COLUMNS,     course_rows,     course_stats,     T_COURSE)
        write_google_tab(sheets, spreadsheet_id, TAB_INSTRUCTOR, INSTRUCTOR_COLUMNS, instructor_rows, instructor_stats, T_INSTRUCTOR)
        print(f"[Output] Google Sheet: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
    except OutputSetupError as exc:
        print("\n[ACTION REQUIRED] " + str(exc))
        print("[Output] Validation completed, but the Google Sheet could not be "
              "written (see action required above). No local copy is produced.")
    except Exception as exc:                              # noqa: BLE001
        print(f"[WARN] Google Sheet output failed: {exc}")
        print("[Output] Validation completed, but the Google Sheet output failed. "
              "No local copy is produced.")

    print("[Done] Validation report generated.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
