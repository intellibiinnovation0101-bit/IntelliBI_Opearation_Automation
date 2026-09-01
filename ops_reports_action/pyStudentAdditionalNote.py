"""
IntelliBI Student Data Generator
=================================
Reads student admission responses from one Google Sheet,
generates formatted data, and writes it to the student list Google Sheet.

⚠ IMPORTANT — DESTINATION SHEET IS IMPORTRANGE-DRIVEN ⚠
The destination "Note List" tab pulls its student_id / student_name / email /
phone columns (A:D) live from the portal via an IMPORTRANGE() array formula in
cell A1. Those columns are NOT stored values — they are a formula spill.
Therefore this script:
  * writes ONLY the Generated_Data and IsRecordUpdated columns,
  * NEVER inserts or deletes rows,
  * NEVER writes into columns A:D.
Inserting/deleting rows or writing hard values into A:D collides with the
IMPORTRANGE spill and makes the entire imported student list vanish (#REF!).

IsRecordUpdated (a manually-maintained, positionally-STORED column) is NOT part
of the IMPORTRANGE spill, so writing to it is safe. However, because the
student name/email columns (A:D) spill live from IMPORTRANGE, whenever the
source gains/loses/reorders students every name shifts row position while the
stored Generated_Data / IsRecordUpdated columns stay anchored — so they drift
out of alignment with the student they belong to. Each run this script already
re-aligns Generated_Data to the current names; it now ALSO re-aligns
IsRecordUpdated the same way, so each student keeps their exact existing
IsRecordUpdated value (Yes / No / blank / etc.) sitting on their current row.
The value is NEVER invented or changed — it is recovered from the sheet's own
prior state (the candidate name embedded in each existing note) and simply moved
to follow its student. See _build_isrecordupdated_owner_map() below.

SETUP (one-time):
1. pip install gspread google-auth
2. Create a Google Cloud Service Account:
   - Go to https://console.cloud.google.com/
   - Create a project (or use existing)
   - Enable "Google Sheets API"
   - Go to "Credentials" -> "Create Credentials" -> "Service Account"
   - Download the JSON key file
   - Save it as "service_account.json in the same folder as this script
3. Share BOTH Google Sheets with the service account email
   (found in credentials.json as "client_email") — give "Editor" access.

USAGE:
   python intellibi_generate_data.py
"""

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import gspread
from google.oauth2.service_account import Credentials
import os
import sys
import re
import time
import random
from datetime import datetime
from difflib import SequenceMatcher
from itertools import permutations


# ─── TRANSIENT-ERROR RETRY (Google Sheets 429 / 5xx) ─────────────────────────
def install_gspread_retry(max_retries=6, base_delay=2.0, max_delay=60.0):
    """Make every gspread API call resilient to transient Google errors.

    Google's Sheets/Drive backend intermittently returns HTTP 429 (rate limit)
    or 5xx — 500 / 502 / 503 ('The service is currently unavailable') / 504.
    These are temporary, server-side, and almost always succeed on retry. The
    script previously had no retry, so a single such blip aborted the whole
    'Student Additional Note' job. We patch gspread's low-level HTTP request so
    EVERY call (open_by_key, get_all_values, batch_update, worksheets, …) retries
    automatically with exponential backoff + jitter instead of failing the run.
    """
    RETRY_STATUS = {429, 500, 502, 503, 504}

    # gspread >= 6 routes requests through http_client.HTTPClient.request;
    # gspread 5.x routes through client.Client.request. Patch whichever exists.
    try:
        from gspread.http_client import HTTPClient as _Target
    except Exception:  # pragma: no cover - older gspread
        from gspread.client import Client as _Target

    if getattr(_Target, "_intellibi_retry_installed", False):
        return

    _orig_request = _Target.request

    def _request_with_retry(self, *args, **kwargs):
        delay = base_delay
        last_err = None
        for attempt in range(max_retries + 1):
            try:
                return _orig_request(self, *args, **kwargs)
            except gspread.exceptions.APIError as err:
                status = None
                try:
                    status = err.response.status_code
                except Exception:
                    status = None
                if status in RETRY_STATUS and attempt < max_retries:
                    sleep_for = min(delay, max_delay) + random.uniform(0, 1)
                    print(f"  ⚠ Google API {status} (transient) — retry "
                          f"{attempt + 1}/{max_retries} in {sleep_for:.1f}s...")
                    time.sleep(sleep_for)
                    delay *= 2
                    last_err = err
                    continue
                raise
        if last_err:
            raise last_err

    _Target.request = _request_with_retry
    _Target._intellibi_retry_installed = True


# ─── CONFIG ──────────────────────────────────────────────────────────────────

CREDENTIALS_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json")

# Sources: Student Admission Responses (multiple sheets)
# Each source is read, cleaned, and merged before matching to the destination.
SOURCES = [
    {
        # Already-Enrolled Students
        "sheet_id": "1ZcOf2rkxy7ueQpGHedLFoQnpckZmcZKWYRpFOTjq7c8",
        "worksheet": "Form Responses 1",
    },
    {
        # Newly-Enrolled Students
        "sheet_id": "1oaXxg3JdtxFp8lFWijIMZKaMZvS0SiglI1K2JTrN2fs",
        "worksheet": "IntelliBI — Student Admission Responses-New Enroll",
    },
]

# Destination: Student List with Generated_Data column
DEST_SHEET_ID = "1DNwt_ll9OWXo-Xsdv4Hq4tsafyFvpWUeDvM73OlRzds"
DEST_WORKSHEET_INDEX = 0  # First sheet (index 0)

# ── Enhancement 1: Alumni Students ───────────────────────────────────────────
# Additional student records (Alumni) NOT present in the existing admission
# sources. Same column layout as the admission response sheets, so the very
# same note-generation logic applies. Read-only and used ONLY as a fallback —
# existing-source matches always take precedence, so existing students are
# completely unaffected.
ALUMNI_SOURCE = {
    "sheet_id": "1QYQJrDymvg7cNO7d4q5HTT4jwC6jvz4EEygp_0inYqA",
    "worksheet_gid": 1157109356,
}

# ── Enhancement 2: Follow-Up Comments ────────────────────────────────────────
# Per-student follow_up_comments, matched by email/name. Purely additive to the
# generated note; its absence (sheet/row/value missing) never blocks a note.
FOLLOWUP_SOURCE = {
    "sheet_id": "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA",
    "worksheet_gid": 0,
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Manual name overrides: source response name (lowercase) -> target student list name (lowercase)
# Only for truly unmatchable cases — most swaps are handled automatically now
MANUAL_NAME_MAP = {}

# Fuzzy match threshold (0.0 to 1.0) — applied to first name AND last name separately
# 0.80 catches spelling variations (Monika/Monica) but rejects different people (Santosh/Satish)
FUZZY_THRESHOLD = 0.80

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def normalize_name(name):
    """Lowercase, trim, and collapse multiple spaces."""
    return re.sub(r'\s+', ' ', str(name).strip().lower())


def normalize_email(email):
    """Lowercase + trim an email for reliable exact-match comparison.
    Returns "" for blank/invalid values (anything without an '@')."""
    e = str(email).strip().lower()
    return e if "@" in e else ""


def get_name_permutations(name):
    """Generate all permutations of name parts for rearranged name matching.
    e.g. 'niketh reddy loka' -> {'niketh reddy loka', 'niketh loka reddy', 'loka niketh reddy', ...}
    """
    parts = name.split()
    if len(parts) <= 1:
        return {name}
    return {' '.join(p) for p in permutations(parts)}


def get_swapped_variants(name):
    """Generate swapped first/last name variants.
    In India people write 'Firstname Lastname' or 'Lastname Firstname' interchangeably.
    For 3-part names like 'Kutal Nimisha Rohan', generates:
      - swap first & last: 'rohan nimisha kutal'
      - just first & last swapped as 2-word: 'nimisha kutal', 'kutal nimisha'
      - any part as first + any other part as last
    """
    parts = name.split()
    variants = set()
    if len(parts) == 2:
        # Simple swap: "kutal nimisha" -> "nimisha kutal"
        variants.add(f"{parts[1]} {parts[0]}")
    elif len(parts) >= 3:
        # Swap first and last
        swapped = parts.copy()
        swapped[0], swapped[-1] = swapped[-1], swapped[0]
        variants.add(' '.join(swapped))
        # Try all 2-word combos (any part as first, any other as last)
        for i in range(len(parts)):
            for j in range(len(parts)):
                if i != j:
                    variants.add(f"{parts[i]} {parts[j]}")
    return variants


def fuzzy_match(name, candidates):
    """Find best fuzzy match using smart first+last name comparison.
    Both first name AND last name must independently be >= 0.80 similar.
    This avoids false matches like 'Santosh Shinde' ↔ 'Satish Shinde'.
    """
    name_parts = name.split()
    if len(name_parts) < 2:
        return None

    best_score = 0
    best_match = None
    for candidate in candidates:
        cand_parts = candidate.split()
        if len(cand_parts) < 2:
            continue
        first_score = SequenceMatcher(None, name_parts[0], cand_parts[0]).ratio()
        last_score = SequenceMatcher(None, name_parts[-1], cand_parts[-1]).ratio()
        if first_score >= FUZZY_THRESHOLD and last_score >= FUZZY_THRESHOLD:
            combined = (first_score + last_score) / 2
            if combined > best_score:
                best_score = combined
                best_match = candidate
    return best_match


def email_contains_name(email, name_parts):
    """Check if email prefix contains all parts of a name.
    e.g. 'nehakardile1125@gmail.com' contains 'neha' and 'kardile'
    Useful for matching married women whose last name changed.
    """
    if not email or '@' not in email:
        return False
    prefix = email.split('@')[0].lower()
    # Remove digits and special chars for cleaner matching
    prefix_clean = re.sub(r'[^a-z]', '', prefix)
    return all(part in prefix_clean for part in name_parts)


def safe(val, default="N/A"):
    """Return trimmed string value, or default if empty/None."""
    if val is None or str(val).strip() == "" or str(val).strip().lower() == "nan":
        return default
    return str(val).strip()


def clean_year(val):
    """Convert year value like '2026.0' or '2026' to clean string."""
    s = safe(val, "N/A")
    if s == "N/A":
        return s
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return s


def clean_experience(val):
    """Format experience: '0.0' -> '0', '4.7' -> '4.7'."""
    s = safe(val, "0")
    try:
        f = float(s)
        return str(int(f)) if f == int(f) else str(f)
    except (ValueError, TypeError):
        return s


def concat_fields(primary, other):
    """Concatenate two fields, handling N/A and empty values."""
    p = safe(primary, "")
    o = safe(other, "")
    if p and o:
        return f"{p}, {o}"
    return p or o or "N/A"


def generate_record(row_dict, candidate_name="N/A", follow_up_comments=""):
    """Generate the formatted data string for one student response.

    `follow_up_comments` (Enhancement 2) is appended to the existing
    'Follow-Up Comments:' line. It defaults to "" so the output is byte-for-byte
    identical to the previous behaviour whenever no comment is available."""
    # Highest Education
    highest_ed = safe(row_dict.get("Highest Education"), "N/A")
    if highest_ed == "Other":
        highest_ed = safe(row_dict.get("Other Education (Please Specify)"), "Other")

    # Passout Year
    passout_year = clean_year(row_dict.get("Passout Year"))

    # How did you hear about IntelliBI
    hear_about = safe(row_dict.get("How did you hear about IntelliBI?"), "N/A")

    # Reference Name
    ref_name = safe(row_dict.get("Reference Name"), "")

    # Fresher / Working Professional
    fresher_wp = safe(row_dict.get("Fresher / Working Professional"), "N/A")

    # No. of Years Experience
    experience = clean_experience(row_dict.get("No. of Years Experience"))

    # IT / Non IT
    it_nonit = safe(row_dict.get("IT / Non IT"), "N/A")

    # IT Domain (concat with Other)
    it_domain = concat_fields(
        row_dict.get("IT Domain"),
        row_dict.get("IT Domain - Other (Please Specify)"),
    )

    # Non-IT Industry (concat with Other)
    nonit_industry = concat_fields(
        row_dict.get("Non-IT Industry"),
        row_dict.get("Non-IT Industry - Other (Please Specify)"),
    )

    return (
        f"Candidate Name: {candidate_name}\n"
        f"Highest Education: {highest_ed}\n"
        f"Passout Year: {passout_year}\n"
        f"How did you hear about IntelliBI: {hear_about}\n"
        f"Reference Name: {ref_name}\n"
        f"Fresher / Working Professional: {fresher_wp}\n"
        f"No. of Years Experience: {experience}\n"
        f"IT / Non IT: {it_nonit}\n"
        f"IT Domain: {it_domain}\n"
        f"Non-IT Industry: {nonit_industry}\n"
        f"IsCandidateActive:Y\n"
        f"IsAttendanceRequired:Y\n"
        f"Follow-Up Comments:{safe(follow_up_comments, '')}"
    )


def get_records(ws):
    """Like worksheet.get_all_records(), but tolerant of duplicate header names.
    gspread's get_all_records() raises on duplicate headers; the new-enroll sheet
    has repeated 'First/Second Installment Date/Amount' columns. We build dicts
    from raw values instead (for duplicate headers, the later column wins — none
    of the fields we actually use are duplicated)."""
    values = ws.get_all_values()
    if not values:
        return []
    headers = [h.strip() for h in values[0]]
    records = []
    for row in values[1:]:
        rec = {}
        for i, h in enumerate(headers):
            rec[h] = row[i] if i < len(row) else ""
        records.append(rec)
    return records


def resolve_worksheet(sheet, worksheet_name):
    """Return the worksheet matching worksheet_name, tolerant of case/space
    differences, falling back to the first worksheet if none matches."""
    try:
        return sheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        all_ws = sheet.worksheets()
        available = [w.title for w in all_ws]
        target_norm = re.sub(r'\s+', ' ', worksheet_name.strip().lower())
        match = next(
            (w for w in all_ws
             if re.sub(r'\s+', ' ', w.title.strip().lower()) == target_norm),
            None,
        )
        if match is None:
            match = all_ws[0]
        print(f"  ⚠ Worksheet '{worksheet_name}' not found. Available tabs: {available}")
        print(f"  → Using '{match.title}' instead.")
        return match


def read_and_clean_source(gc, source):
    """Open a source sheet, clean the Student Name column (fix extra spaces,
    Title Case), and return the cleaned records as a list of dicts."""
    sheet_id = source["sheet_id"]
    worksheet_name = source["worksheet"]
    print(f"Reading admission responses from {sheet_id} ('{worksheet_name}')...")
    source_sheet = gc.open_by_key(sheet_id)
    source_ws = resolve_worksheet(source_sheet, worksheet_name)
    source_data = get_records(source_ws)
    print(f"  Found {len(source_data)} responses.")

    # ── Locate Student Name column ───────────────────────────────────────
    source_all_values = source_ws.get_all_values()
    source_headers = source_all_values[0] if source_all_values else []
    source_name_col = -1
    for i, h in enumerate(source_headers):
        if h.strip() == "Student Name":
            source_name_col = i
            break

    if source_name_col == -1:
        print("  ⚠ No 'Student Name' column found — skipping name cleanup for this source.")
        return source_data

    # ── Fix multiple spaces in Student Name column ───────────────────────
    print("  Cleaning student names (fixing multiple spaces)...")
    name_fixes = []
    for row_idx, row in enumerate(source_all_values[1:], start=2):
        if source_name_col < len(row):
            original = row[source_name_col]
            cleaned = re.sub(r'\s+', ' ', original.strip())
            if original != cleaned:
                name_fixes.append({
                    "range": gspread.utils.rowcol_to_a1(row_idx, source_name_col + 1),
                    "values": [[cleaned]]
                })
    if name_fixes:
        source_ws.batch_update(name_fixes, value_input_option="RAW")
        print(f"    Fixed {len(name_fixes)} names with extra spaces.")
    else:
        print("    No names needed fixing.")

    # ── Normalize source names to Title Case ─────────────────────────────
    # Only fix casing (lowercase/UPPERCASE → Title Case). Do NOT change spelling,
    # drop middle names, or reorder name parts — the student typed their own name.
    print("  Normalizing source names to Title Case (Firstname Lastname)...")
    source_all_values = source_ws.get_all_values()  # Re-read after space fix
    source_name_updates = []
    for row_idx, row in enumerate(source_all_values[1:], start=2):
        if source_name_col >= len(row):
            continue
        src_original = row[source_name_col].strip()
        title_cased = ' '.join(word.capitalize() for word in src_original.split())
        if src_original != title_cased:
            source_name_updates.append({
                "range": gspread.utils.rowcol_to_a1(row_idx, source_name_col + 1),
                "values": [[title_cased]]
            })
    if source_name_updates:
        source_ws.batch_update(source_name_updates, value_input_option="RAW")
        print(f"    Fixed {len(source_name_updates)} names to Title Case:")
        for u in source_name_updates:
            print(f"      {u['range']}: → '{u['values'][0][0]}'")
    else:
        print("    All source names already in correct format.")

    # Re-read after renaming
    return get_records(source_ws)


def _open_worksheet_by_gid(gc, sheet_id, gid):
    """Return the worksheet with the given gid, falling back to the first tab.
    Using the gid (from the sheet URL) avoids fragile tab-name matching."""
    sheet = gc.open_by_key(sheet_id)
    try:
        return sheet.get_worksheet_by_id(int(gid))
    except Exception:
        match = next((w for w in sheet.worksheets() if str(w.id) == str(gid)), None)
        if match is not None:
            return match
        print(f"  ⚠ Worksheet gid {gid} not found in {sheet_id}. Using first tab.")
        return sheet.get_worksheet(0)


def read_alumni_source(gc, source):
    """Enhancement 1: read Alumni Student records (read-only).

    The Alumni sheet shares the exact column layout of the admission response
    sheets, so its rows can feed `generate_record` unchanged. We deliberately do
    NOT write anything back to this sheet (no name cleanup) — matching uses
    `normalize_name`, which already lowercases/normalises, so no in-sheet edits
    are needed and the Alumni sheet stays untouched."""
    print(f"Reading Alumni students from {source['sheet_id']} (gid {source['worksheet_gid']})...")
    ws = _open_worksheet_by_gid(gc, source["sheet_id"], source["worksheet_gid"])
    records = get_records(ws)
    print(f"  Found {len(records)} Alumni record(s).")
    return records


def build_followup_lookup(gc, source):
    """Enhancement 2: read the Follow-Up sheet and build lookups mapping a
    student (by email and by several name variants) to their
    `follow_up_comments` value.

    Returns a dict of lookups. On any failure it returns an empty structure so
    that a missing/unreadable Follow-Up sheet can never block note generation."""
    empty = {"email": {}, "name": {}, "first_last": {}, "perm": {},
             "swapped": {}, "names_list": []}
    try:
        print(f"Reading follow-up comments from {source['sheet_id']} (gid {source['worksheet_gid']})...")
        ws = _open_worksheet_by_gid(gc, source["sheet_id"], source["worksheet_gid"])
        records = get_records(ws)
    except Exception as e:
        print(f"  ⚠ Could not read follow-up sheet ({e}). Continuing with blank follow-up comments.")
        return empty

    lk = {"email": {}, "name": {}, "first_last": {}, "perm": {},
          "swapped": {}, "names_list": []}
    for r in records:
        # follow_up_comments may legitimately be blank — store as-is (trimmed).
        fu = str(r.get("follow_up_comments", "") or "").strip()
        email = normalize_email(r.get("email", ""))
        if email:
            lk["email"][email] = fu
        name = normalize_name(safe(r.get("student_name"), ""))
        if name:
            lk["name"].setdefault(name, fu)
            parts = name.split()
            if len(parts) >= 3:
                lk["first_last"].setdefault(f"{parts[0]} {parts[-1]}", fu)
            for perm in get_name_permutations(name):
                lk["perm"].setdefault(perm, fu)
            for variant in get_swapped_variants(name):
                lk["swapped"].setdefault(variant, fu)
    lk["names_list"] = list(lk["name"].keys())
    print(f"  Loaded follow-up data for {len(lk['name'])} student(s).")
    return lk


def lookup_followup(fu_lookup, name_key, email_key):
    """Return the follow_up_comments for a student, or "" when the student is
    not found (or is found with a blank value). Mirrors the destination matching
    priority: exact email -> exact name -> first+last -> permutation -> swapped
    -> fuzzy. Never raises."""
    if email_key and email_key in fu_lookup["email"]:
        return fu_lookup["email"][email_key]
    if name_key in fu_lookup["name"]:
        return fu_lookup["name"][name_key]
    if name_key in fu_lookup["first_last"]:
        return fu_lookup["first_last"][name_key]
    if name_key in fu_lookup["perm"]:
        return fu_lookup["perm"][name_key]
    if name_key in fu_lookup["swapped"]:
        return fu_lookup["swapped"][name_key]
    fuzzy_result = fuzzy_match(name_key, fu_lookup["names_list"])
    if fuzzy_result:
        return fu_lookup["name"][fuzzy_result]
    return ""


# ─── IsRecordUpdated preservation ────────────────────────────────────────────

# Matches the first line produced by generate_record(): "Candidate Name: <name>".
_CANDIDATE_NAME_RE = re.compile(r"Candidate Name:\s*(.*)")


def _embedded_candidate_name(generated_value):
    """Return the 'Candidate Name:' embedded in an existing Generated_Data cell,
    or None when the cell is empty / has no such line. This is the identity the
    row's IsRecordUpdated value truly belongs to (the note and the flag were
    written together and drift together relative to the IMPORTRANGE names)."""
    if not generated_value:
        return None
    m = _CANDIDATE_NAME_RE.search(str(generated_value))
    if not m:
        return None
    return normalize_name(m.group(1))


def _isrecordupdated_recoverable(dest_data, generated_col, iru_col):
    """Decide whether each student's IsRecordUpdated value can be safely
    recovered from the note it was written beside.

    The recovery in _build_isrecordupdated_owner_map() keys a row's flag by the
    candidate name embedded in that SAME row's Generated_Data note. That is only
    correct while the two stored columns are still co-aligned with each other —
    i.e. they have drifted together relative to the IMPORTRANGE names. We detect
    that by presence-agreement: in a co-aligned sheet a row has a note exactly
    when it has a flag (they were written together), so 'note present XOR flag
    present' is ~0. If the columns have been re-aligned independently (e.g. a
    prior run rewrote the notes but not the flags), that XOR spikes and the
    per-note owner would be wrong — so we refuse to touch the column and ask for
    a one-time re-seed instead of corrupting it.

    Returns (recoverable: bool, xor_rows: int, flag_rows: int)."""
    if iru_col == -1 or generated_col == -1:
        return (False, 0, 0)
    xor = 0
    flags = 0
    for row in dest_data[1:]:
        has_note = generated_col < len(row) and str(row[generated_col]).strip() != ""
        has_flag = iru_col < len(row) and str(row[iru_col]).strip() != ""
        if has_flag:
            flags += 1
        if has_note != has_flag:
            xor += 1
    # Reliable when co-aligned: allow a tiny slack for legitimate note-only /
    # flag-only rows, but reject the wholesale mismatch of a corrupted column.
    threshold = max(5, int(0.05 * flags))
    return (xor <= threshold, xor, flags)


def _build_isrecordupdated_owner_map(dest_data, name_col, generated_col, iru_col):
    """Recover each student's existing IsRecordUpdated value, keyed by student
    identity, from the sheet's CURRENT (pre-write) state.

    IsRecordUpdated is co-located with Generated_Data (both are stored columns
    written/maintained together), so the reliable owner of a row's flag is the
    candidate name embedded in that same row's existing note. We therefore key
    primarily by that embedded name. For the rare row that carries a flag but no
    note, we fall back to that row's own student_name (without ever overwriting a
    note-based mapping). The returned dict maps normalized student name ->
    existing IsRecordUpdated value, letting the caller place each value back on
    the student's current row so nothing is lost when the names shift."""
    if iru_col == -1:
        return {}
    owner = {}
    # Pass 1 — authoritative: attribute each flag to its note's embedded student.
    for row in dest_data[1:]:
        gen = row[generated_col] if generated_col < len(row) else ""
        emb = _embedded_candidate_name(gen)
        if emb:
            flag = row[iru_col] if iru_col < len(row) else ""
            owner[emb] = flag
    # Pass 2 — fallback for flags on note-less rows: attribute to the row's own
    # name, but never clobber an authoritative note-based mapping.
    for row in dest_data[1:]:
        gen = row[generated_col] if generated_col < len(row) else ""
        if _embedded_candidate_name(gen):
            continue
        flag = row[iru_col] if iru_col < len(row) else ""
        key = normalize_name(row[name_col]) if name_col < len(row) else ""
        if key and str(flag).strip() != "":
            owner.setdefault(key, flag)
    return owner


# ─── MAIN ────────────────────────────────────────────────────────────────────

def main():
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Starting IntelliBI data generation...")

    # Make all Google Sheets calls resilient to transient 429/5xx (e.g. 503
    # 'service currently unavailable') by retrying with exponential backoff.
    install_gspread_retry()

    # Auth
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"ERROR: service_account.json not found at {CREDENTIALS_FILE}")
        print("See script header for setup instructions.")
        sys.exit(1)

    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)

    # ── Read & clean all Sources (Admission Responses) ───────────────────
    source_data = []
    for source in SOURCES:
        source_data.extend(read_and_clean_source(gc, source))
    print(f"Total combined responses across {len(SOURCES)} source(s): {len(source_data)}")

    # Build lookups for matching:
    # 1. Exact name (normalized)
    # 2. First+last name (skip middle)
    # 3. All permutations of name parts (for rearranged names)
    # 4. Swapped first/last name variants (Indian naming convention)
    # 5. Manual overrides
    response_lookup = {}
    response_lookup_first_last = {}
    response_lookup_permutations = {}
    response_lookup_swapped = {}
    response_lookup_email = {}   # exact email (Email ID) -> row  — most reliable join key
    for row in source_data:
        # Index by exact email FIRST — independent of name spelling/format.
        # Sources that carry an "Email ID" column (e.g. the new-enroll sheet)
        # get a bullet-proof match here; later rows (newer responses) win.
        email = normalize_email(row.get("Email ID", ""))
        if email:
            response_lookup_email[email] = row
        name = normalize_name(safe(row.get("Student Name"), ""))
        if name:
            response_lookup[name] = row
            parts = name.split()
            if len(parts) >= 3:
                first_last = f"{parts[0]} {parts[-1]}"
                response_lookup_first_last[first_last] = row
            # Store all permutations for rearranged name matching
            for perm in get_name_permutations(name):
                response_lookup_permutations[perm] = row
            # Store swapped first/last variants (Indian naming: firstname lastname or lastname firstname)
            for variant in get_swapped_variants(name):
                response_lookup_swapped[variant] = row
            # Manual override mapping
            if name in MANUAL_NAME_MAP:
                response_lookup[MANUAL_NAME_MAP[name]] = row

    # ── Enhancement 1: Merge Alumni Students (fallback only) ─────────────
    # Alumni rows are added with setdefault() so an existing-source row ALWAYS
    # wins on any colliding key. This guarantees existing students keep matching
    # exactly as before; Alumni records only fill in students that had no match.
    try:
        alumni_data = read_alumni_source(gc, ALUMNI_SOURCE)
    except Exception as e:
        print(f"  ⚠ Could not read Alumni sheet ({e}). Continuing without Alumni records.")
        alumni_data = []
    for row in alumni_data:
        email = normalize_email(row.get("Email ID", ""))
        if email:
            response_lookup_email.setdefault(email, row)
        name = normalize_name(safe(row.get("Student Name"), ""))
        if name:
            response_lookup.setdefault(name, row)
            parts = name.split()
            if len(parts) >= 3:
                response_lookup_first_last.setdefault(f"{parts[0]} {parts[-1]}", row)
            for perm in get_name_permutations(name):
                response_lookup_permutations.setdefault(perm, row)
            for variant in get_swapped_variants(name):
                response_lookup_swapped.setdefault(variant, row)

    # ── Enhancement 2: Build Follow-Up Comments lookup ───────────────────
    # Defensive: a missing/unreadable sheet yields an empty lookup, so notes are
    # still generated (with blank follow-up comments).
    followup_lookup = build_followup_lookup(gc, FOLLOWUP_SOURCE)

    # ── Read Destination (Student List) ──────────────────────────────────
    # NOTE: columns A:D (student_id/name/email/phone) are an IMPORTRANGE spill.
    # We read their computed values for matching but NEVER write to them, and we
    # NEVER insert/delete rows. Only the Generated_Data column is written.
    print("Reading student list...")
    dest_sheet = gc.open_by_key(DEST_SHEET_ID)
    dest_ws = dest_sheet.get_worksheet(DEST_WORKSHEET_INDEX)
    dest_data = dest_ws.get_all_values()

    if not dest_data:
        print("ERROR: Destination sheet is empty.")
        sys.exit(1)

    headers = dest_data[0]
    print(f"  Found {len(dest_data) - 1} students. Columns: {headers}")

    # Find column indices
    def find_col(name):
        for i, h in enumerate(headers):
            if h.strip().lower() == name.lower():
                return i
        return -1

    name_col = find_col("student_name")
    email_col = find_col("email")
    generated_col = find_col("Generated_Data")
    iru_col = find_col("IsRecordUpdated")

    if name_col == -1:
        print("ERROR: 'student_name' column not found in destination sheet.")
        sys.exit(1)

    # If Generated_Data column doesn't exist, add it
    if generated_col == -1:
        headers.append("Generated_Data")
        generated_col = len(headers) - 1
        dest_ws.update_cell(1, generated_col + 1, "Generated_Data")
        print("  Added 'Generated_Data' column.")

    # ── Preserve IsRecordUpdated across the IMPORTRANGE name-shift ────────
    # Capture each student's EXISTING IsRecordUpdated value (keyed by identity)
    # from the sheet's current pre-write state, so we can place it back on the
    # student's current row after the names have shifted. This never invents or
    # changes a value — it only moves each existing value to follow its student.
    # We only do this when the flag column is still safely recoverable (see
    # _isrecordupdated_recoverable); otherwise we leave the column completely
    # untouched so a corrupted state can never be entrenched.
    iru_owner_map = {}
    iru_recoverable = False
    if iru_col == -1:
        print("  ⚠ 'IsRecordUpdated' column not found — leaving it untouched.")
    else:
        iru_recoverable, iru_xor, iru_flags = _isrecordupdated_recoverable(
            dest_data, generated_col, iru_col
        )
        if iru_recoverable:
            iru_owner_map = _build_isrecordupdated_owner_map(
                dest_data, name_col, generated_col, iru_col
            )
            print(f"  Preserving IsRecordUpdated for {len(iru_owner_map)} student(s).")
        else:
            print(
                "  ⚠ IsRecordUpdated looks misaligned with Generated_Data "
                f"({iru_xor} of {iru_flags} flagged rows disagree). Leaving the "
                "column UNTOUCHED to avoid corrupting it. Re-seed it once from a "
                "known-good snapshot (see IsRecordUpdated_corrected sheet), then "
                "future runs will keep it aligned automatically."
            )

    # ── Generate Data ────────────────────────────────────────────────────
    print("Generating data...")
    updates_generated = []
    updates_iru = []          # IsRecordUpdated cells re-aligned to follow students
    matched = 0
    skipped = 0

    for row_idx, row in enumerate(dest_data[1:], start=2):  # row_idx = sheet row (1-indexed)
        # Pad row if shorter than headers
        while len(row) < len(headers):
            row.append("")

        name_key = normalize_name(row[name_col])
        dest_email = normalize_email(row[email_col]) if (email_col != -1 and email_col < len(row)) else ""

        # ── Re-align IsRecordUpdated to this student's current row ────────
        # Independent of note matching: every student keeps the exact flag they
        # already had. We only queue a write when the value actually differs
        # (i.e. it had drifted), so unchanged cells are never rewritten. Only
        # runs when the column is safely recoverable (see gate above).
        if iru_col != -1 and iru_recoverable and name_key:
            desired_iru = iru_owner_map.get(name_key, "")
            current_iru = row[iru_col] if iru_col < len(row) else ""
            if str(desired_iru).strip() != str(current_iru).strip():
                updates_iru.append({"row": row_idx, "col": iru_col + 1, "val": desired_iru})

        # Match priority: EXACT EMAIL -> exact name -> first+last -> permutation
        #                 -> swapped -> fuzzy -> email-contains-name
        # Exact email is checked first because it is unique and independent of
        # how the student's name was typed in either sheet (handles cases like
        # dest "Hitesh Patil" vs source "Hitesh Dipak Patil").
        resp_row = None
        match_type = ""
        if dest_email and dest_email in response_lookup_email:
            resp_row = response_lookup_email[dest_email]
            match_type = "email-exact"
        elif name_key in response_lookup:
            resp_row = response_lookup[name_key]
            match_type = "exact"
        elif name_key in response_lookup_first_last:
            resp_row = response_lookup_first_last[name_key]
            match_type = "first+last"
        elif name_key in response_lookup_permutations:
            resp_row = response_lookup_permutations[name_key]
            match_type = "permutation"
        elif name_key in response_lookup_swapped:
            resp_row = response_lookup_swapped[name_key]
            match_type = "swapped"
        else:
            # Fuzzy match
            fuzzy_result = fuzzy_match(name_key, response_lookup.keys())
            if fuzzy_result:
                resp_row = response_lookup[fuzzy_result]
                match_type = f"fuzzy({fuzzy_result})"
            elif email_col != -1 and email_col < len(row):
                # Email-based match: check if email contains response name parts
                # Handles married name changes (e.g. nehakardile@gmail -> Neha Kardile)
                student_email = row[email_col].strip().lower()
                if student_email:
                    for resp_name, resp_data in response_lookup.items():
                        resp_parts = resp_name.split()
                        if len(resp_parts) >= 2 and email_contains_name(student_email, resp_parts):
                            resp_row = resp_data
                            match_type = f"email({resp_name})"
                            break

        if resp_row is not None:
            dest_student_name = row[name_col].strip() if name_col < len(row) else "N/A"
            # Enhancement 2: pull this student's follow_up_comments (blank if the
            # student/value is absent — never blocks the note).
            follow_up_val = lookup_followup(followup_lookup, name_key, dest_email)
            record = generate_record(
                resp_row,
                candidate_name=dest_student_name,
                follow_up_comments=follow_up_val,
            )
            updates_generated.append({"row": row_idx, "col": generated_col + 1, "val": record})
            matched += 1
            if match_type not in ("exact",):
                source_name = safe(resp_row.get("Student Name"), "?")
                print(f"    ↳ '{name_key}' matched via {match_type} → '{source_name}'")
        else:
            # Only clear if previously had data
            current_val = row[generated_col] if generated_col < len(row) else ""
            if current_val == "":
                skipped += 1
            else:
                updates_generated.append({"row": row_idx, "col": generated_col + 1, "val": ""})
                skipped += 1

    print(f"  Matched: {matched}, No response: {skipped}")

    # ── Report unmatched source names ────────────────────────────────────
    # Build set of all destination names + all match variants
    dest_names = set()
    dest_names_list = []
    for row in dest_data[1:]:
        if name_col < len(row):
            n = normalize_name(row[name_col])
            if n:
                dest_names.add(n)
                dest_names_list.append(n)
                parts = n.split()
                if len(parts) >= 3:
                    dest_names.add(f"{parts[0]} {parts[-1]}")
                # Add all permutations and swapped variants
                for perm in get_name_permutations(n):
                    dest_names.add(perm)
                for variant in get_swapped_variants(n):
                    dest_names.add(variant)

    # Build email set from destination for email-based matching check
    dest_emails = []
    dest_emails_exact = set()   # normalized exact emails present in the student list
    if email_col != -1:
        for row in dest_data[1:]:
            if email_col < len(row):
                dest_emails.append(row[email_col].strip().lower())
                ne = normalize_email(row[email_col])
                if ne:
                    dest_emails_exact.add(ne)

    unmatched_source = []
    for name in response_lookup:
        # Exact-email match takes precedence — if this source response's email
        # exists in the student list, it IS matched (regardless of name).
        src_email = normalize_email(response_lookup[name].get("Email ID", ""))
        if src_email and src_email in dest_emails_exact:
            continue
        # Check manual map
        if name in MANUAL_NAME_MAP and MANUAL_NAME_MAP[name] in dest_names:
            continue
        # Check exact, first+last, permutations
        parts = name.split()
        first_last = f"{parts[0]} {parts[-1]}" if len(parts) >= 3 else None
        if name in dest_names:
            continue
        if first_last and first_last in dest_names:
            continue
        # Check fuzzy
        fuzzy_result = fuzzy_match(name, dest_names_list)
        if fuzzy_result:
            continue
        # Check email-based match
        email_matched = False
        if len(parts) >= 2:
            for email in dest_emails:
                if email_contains_name(email, parts):
                    email_matched = True
                    break
        if email_matched:
            continue
        original = safe(response_lookup[name].get("Student Name"), name)
        unmatched_source.append(original)

    if unmatched_source:
        print(f"\n  ⚠ {len(unmatched_source)} response(s) NOT found in student list:")
        for i, n in enumerate(sorted(unmatched_source), 1):
            print(f"    {i}. {n}")
    else:
        print("\n  ✓ All responses matched to a student in the list.")

    # ── Batch Update to Google Sheet (Generated_Data + IsRecordUpdated) ──
    if not updates_generated and not updates_iru:
        print("No updates to write.")
        return

    print(
        f"Writing {len(updates_generated)} Generated_Data cell(s) and "
        f"{len(updates_iru)} re-aligned IsRecordUpdated cell(s) to Google Sheet..."
    )

    # Use batch update for efficiency. This writes individual cells by row index
    # in the Generated_Data column and the IsRecordUpdated column — it never
    # inserts/deletes rows and never touches the IMPORTRANGE columns (A:D). The
    # IsRecordUpdated writes only re-place each student's own existing value onto
    # their current row (queued above only when the value had drifted).
    cells_to_update = []
    for u in updates_generated:
        cells_to_update.append({
            "range": f"{gspread.utils.rowcol_to_a1(u['row'], u['col'])}",
            "values": [[u["val"]]]
        })
    for u in updates_iru:
        cells_to_update.append({
            "range": f"{gspread.utils.rowcol_to_a1(u['row'], u['col'])}",
            "values": [[u["val"]]]
        })

    # gspread batch_update
    dest_ws.batch_update(cells_to_update, value_input_option="RAW")

    print(
        f"[{datetime.now():%Y-%m-%d %H:%M:%S}] Done! Updated {matched} student "
        f"records; re-aligned {len(updates_iru)} IsRecordUpdated value(s)."
    )


if __name__ == "__main__":
    main()
