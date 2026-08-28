"""
================================================================================
  IntelliBI - Admission Formalities Reports  (pyAdmissionFormalitiesReport.py)
  ------------------------------------------------------------------------------
  Two NEW, independent reports built on the existing project plumbing. This
  script does NOT modify any existing report or business logic.

  REPORT 1 - Historical Admission Formalities Report
      Consolidated admission-form / signature status of ALL active students.

  REPORT 2 - Current Student Admission Formalities Report  (run daily)
      Only students who are still PENDING admission formalities (joined on/after
      CURRENT_REPORT_START_DATE and not yet signed), with row colour-coding.

  SOURCES
      Source 1 : Google Sheet "IntelliBIStudentInfo"  tab "Students"
      Source 2 : Google Sheet "Zoho Sign - Signers"   tab "Signers"

  MATCHING (per student, case-insensitive, trimmed)
      1. Email ID   (highest priority)
      2. Student Name (fallback)
      First match wins.

  Reuses utils.get_sheets_service / ensure_tab_exists / _normalize /
  sort_sheet_by_column / retry helpers. Reads via the service account in
  config_files/service_account.json.

  PREREQUISITES
      * Run zoho_sign_to_gsheet.py first so the Signers tab has the latest
        statuses (recipient_status / sent_date / signed_date / expiry_date).
      * The service account must have EDITOR access on BOTH output sheets.

  ENTRY POINT
      python Reports/pyAdmissionFormalitiesReport.py
================================================================================
"""
from __future__ import annotations

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import os
import sys
import time
from datetime import datetime, date, timezone, timedelta

# ── Resolve project root so utils.py / config_files are importable ────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR   = os.path.dirname(_SCRIPT_DIR)              # IntelliBI Automation/
if _BASE_DIR not in sys.path:
    sys.path.insert(0, _BASE_DIR)

import utils   # shared Google-Sheets helpers (service account based)

# ═════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═════════════════════════════════════════════════════════════════════════════
GENERATE_HISTORICAL_ADMISSION_FORMALITIES_REPORT = True
GENERATE_CURRENT_ADMISSION_FORMALITIES_REPORT    = True

# Current report: only students who joined on/after this date (YYYY-MM-DD).
CURRENT_REPORT_START_DATE = "2026-06-01"

# Current report is a "still-pending" snapshot. With pure upsert, a student who
# signs later would linger in the sheet. When True, rows whose student is no
# longer pending (signed / filtered out) are removed so the report stays a true
# "currently pending" list. Set False for strict insert/update/skip-only.
CURRENT_REPORT_PRUNE_STALE = True

# ── Source sheets ─────────────────────────────────────────────────────────────
STUDENT_SHEET_ID = "1Eq7Q3Gota7nYiaorm1L0NoouVfYtS7JkbBp4U5MWzVA"   # IntelliBIStudentInfo
STUDENTS_TAB     = "Students"
SIGNERS_SHEET_ID = "1Pu3I4Ujoo-G_l8nifFS3b5P_cRbWewTlX0J2W3aOi6I"   # Zoho Sign - Signers
SIGNERS_TAB      = "Signers"

# ── Output sheets (must already exist; SA needs Editor access) ────────────────
HISTORICAL_SHEET_ID = "1oA-d_mMuLkbcO93jGEhv1melOs7gHga9HSiYOB4Udo4"
HISTORICAL_TAB      = "Historical Admission Formalities"
CURRENT_SHEET_ID    = "19me_6xEPYYSz903-DHEJjdNr1sQAktWWKW1kw4aR_qs"
CURRENT_TAB         = "Current Student Admission Formalities"

SA_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json")

IST = timezone(timedelta(hours=5, minutes=30))

# ── Report layout (exact column order requested) ──────────────────────────────
REPORT_COLUMNS = [
    "Student Name",
    "Email ID",
    "Phone Number",
    "Batch Name",
    "Joined On",
    "Request Form Name",
    "Recipient Status",
    "Request Status",
    "Sent Date",
    "Signed Date",
    "Expiry Date",
]
# "Joined On" is stored as YYYY-MM-DD, so a plain descending text sort = newest first.
SORT_COLUMN = "Joined On"

FORM_NOT_SENT = "Form Not Sent"

# ── Row colours (Sheets API RGB 0..1) — used by both reports ──────────────────
COL_GREEN       = {"red": 0.714, "green": 0.843, "blue": 0.659}   # signed / completed (done)
COL_LIGHT_GREEN = {"red": 0.851, "green": 0.918, "blue": 0.827}   # viewed / opened
COL_AMBER       = {"red": 1.000, "green": 0.898, "blue": 0.600}   # sent, not opened
COL_RED         = {"red": 0.957, "green": 0.800, "blue": 0.800}   # not sent/expired/declined
COL_WHITE       = {"red": 1.0,   "green": 1.0,   "blue": 1.0}


# ═════════════════════════════════════════════════════════════════════════════
#  NORMALISATION HELPERS
# ═════════════════════════════════════════════════════════════════════════════
def norm(value) -> str:
    """Trim, collapse internal whitespace, lowercase. '' for null/blank."""
    if value is None:
        return ""
    return " ".join(str(value).split()).strip().lower()


def norm_email(value) -> str:
    return str(value or "").strip().lower()


def to_ist_date(value):
    """Parse an ISO/date-ish string into an IST date (or None)."""
    s = str(value or "").strip()
    if not s:
        return None
    txt = s.replace("Z", "+00:00")
    dt = None
    try:
        dt = datetime.fromisoformat(txt)
    except ValueError:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%b-%Y",
                    "%d/%m/%Y", "%m/%d/%Y"):
            try:
                dt = datetime.strptime(s.split("+")[0].strip(), fmt)
                break
            except ValueError:
                continue
    if dt is None:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(IST).date()


# ═════════════════════════════════════════════════════════════════════════════
#  SHEET READ
# ═════════════════════════════════════════════════════════════════════════════
def read_records(service, sheet_id: str, tab: str) -> list:
    """Read a tab into a list of dicts keyed by header. [] if empty/missing."""
    try:
        result = utils._gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=sheet_id, range=f"{tab}").execute(),
            label=f"read → {tab}",
        )
    except Exception as e:  # noqa: BLE001
        print(f"[Read → {tab}] Could not read: {e}")
        return []
    values = result.get("values", [])
    if not values:
        return []
    header = values[0]
    records = []
    for row in values[1:]:
        row = list(row) + [""] * (len(header) - len(row))
        records.append({header[i]: row[i] for i in range(len(header))})
    return records


# ═════════════════════════════════════════════════════════════════════════════
#  SIGNER INDEX  (best record per email / per name)
# ═════════════════════════════════════════════════════════════════════════════
def signer_field(rec: dict, *names, default=""):
    """First present, non-empty field among candidate names (schema-tolerant)."""
    for n in names:
        if n in rec and str(rec[n]).strip():
            return rec[n]
    return default


def build_signer_indexes(signer_records: list):
    """Return (by_email, by_name). Each maps key -> the most-recent signer row."""
    by_email, by_name = {}, {}

    def recency(rec):
        d = to_ist_date(signer_field(rec, "created_date", "sent_date", "last_updated"))
        return d or date.min

    for rec in signer_records:
        email = norm_email(signer_field(rec, "recipient_email"))
        name  = norm(signer_field(rec, "recipient_name"))
        if email:
            if email not in by_email or recency(rec) >= recency(by_email[email]):
                by_email[email] = rec
        if name:
            if name not in by_name or recency(rec) >= recency(by_name[name]):
                by_name[name] = rec
    return by_email, by_name


def signer_to_report_fields(rec):
    """Extract the report fields from a matched signer record (schema-tolerant)."""
    return {
        "Request Form Name": signer_field(rec, "request_name"),
        "Recipient Status":  signer_field(rec, "recipient_status", "action_status"),
        "Request Status":    signer_field(rec, "request_status"),
        "Sent Date":         signer_field(rec, "sent_date"),
        "Signed Date":       signer_field(rec, "signed_date", "signed_time"),
        "Expiry Date":       signer_field(rec, "expiry_date"),
    }


# ═════════════════════════════════════════════════════════════════════════════
#  BUILD REPORT ROWS
# ═════════════════════════════════════════════════════════════════════════════
def is_active_student(stu: dict) -> bool:
    return norm(stu.get("is_candidate_active")) == "y" and norm(stu.get("Is_Deleted")) == "n"


def build_rows(students, by_email, by_name, *, current_only, start_date):
    """
    Return (rows, stats). One row per qualifying student, matched to a signer.
    stats tracks match method + pending/without-request counts.
    """
    rows = []
    stats = {"processed": 0, "by_email": 0, "by_name": 0, "failed": 0,
             "without_request": 0}

    for stu in students:
        if not is_active_student(stu):
            continue

        joined_dt = to_ist_date(stu.get("joined_on"))
        if current_only:
            if joined_dt is None or joined_dt < start_date:
                continue

        # ── Matching: email first, then name ────────────────────────────────
        email_key = norm_email(stu.get("email"))
        name_key  = norm(stu.get("student_name"))
        match, method = None, "failed"
        if email_key and email_key in by_email:
            match, method = by_email[email_key], "email"
        elif name_key and name_key in by_name:
            match, method = by_name[name_key], "name"

        if match is not None:
            fields = signer_to_report_fields(match)
        else:
            fields = {"Request Form Name": "", "Recipient Status": FORM_NOT_SENT,
                      "Request Status": "", "Sent Date": "", "Signed Date": "",
                      "Expiry Date": ""}

        # ── Current report: drop students who already completed signing ─────
        if current_only and norm(fields["Recipient Status"]) == "signed":
            continue

        stats["processed"] += 1
        stats[{"email": "by_email", "name": "by_name", "failed": "failed"}[method]] += 1
        if match is None:
            stats["without_request"] += 1

        joined_sort = joined_dt.strftime("%Y-%m-%d") if joined_dt else ""
        rows.append({
            "Student Name":      stu.get("student_name", ""),
            "Email ID":          stu.get("email", ""),
            "Phone Number":      stu.get("phone", ""),
            "Batch Name":        stu.get("batch_name", ""),
            "Joined On":         joined_sort,
            "Request Form Name": fields["Request Form Name"],
            "Recipient Status":  fields["Recipient Status"],
            "Request Status":    fields["Request Status"],
            "Sent Date":         fields["Sent Date"],
            "Signed Date":       fields["Signed Date"],
            "Expiry Date":       fields["Expiry Date"],
        })
    return rows, stats


def report_key(row: dict) -> str:
    """Upsert identity: Email ID if present, else name::<student name>."""
    email = norm_email(row.get("Email ID"))
    return email if email else "name::" + norm(row.get("Student Name"))


# ═════════════════════════════════════════════════════════════════════════════
#  UPSERT (INSERT / UPDATE / SKIP) with a custom key + count summary
# ═════════════════════════════════════════════════════════════════════════════
def upsert_report(service, sheet_id, tab, columns, rows, key_of):
    summary = {"inserted": 0, "updated": 0, "skipped": 0, "failed": 0}
    last_col = utils.col_letter(len(columns) - 1)
    utils.ensure_tab_exists(service, sheet_id, tab)

    try:
        result = utils._gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=sheet_id, range=f"{tab}!A1:{last_col}").execute(),
            label=f"read → {tab}")
        existing = result.get("values", [])
    except Exception as e:  # noqa: BLE001
        print(f"[Write → {tab}] Read failed: {e}")
        summary["failed"] = len(rows)
        return summary

    # Header (+ re-align existing data if the layout changed)
    if existing:
        header, data_rows = existing[0], existing[1:]
        if header != columns:
            data_rows = utils._realign_rows_to_columns(header, columns, data_rows)
            utils._gsheets_call_with_retry(
                lambda: service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f"{tab}!A1:{last_col}{len(data_rows) + 1}",
                    valueInputOption="RAW",
                    body={"values": [columns] + data_rows}).execute(),
                label=f"migrate layout → {tab}")
            header = columns
    else:
        header, data_rows = columns, []
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id, range=f"{tab}!A1",
            valueInputOption="RAW", body={"values": [columns]}).execute()

    # Index existing rows by key
    existing_index = {}
    for i, vals in enumerate(data_rows):
        rec = {header[j]: (vals[j] if j < len(vals) else "") for j in range(len(header))}
        k = key_of(rec)
        if k:
            existing_index[k] = (i + 2, vals)   # +2: row 1 = header

    batch_update, to_append = [], []
    for row in rows:
        incoming = [str(row.get(c, "")) for c in columns]
        k = key_of(row)
        if k in existing_index:
            sheet_row, old_vals = existing_index[k]
            padded = list(old_vals) + [""] * (len(columns) - len(old_vals))
            changed = any(utils._normalize(incoming[i]) != utils._normalize(padded[i])
                          for i in range(len(columns)))
            if changed:
                batch_update.append({"range": f"{tab}!A{sheet_row}:{last_col}{sheet_row}",
                                     "values": [incoming]})
            else:
                summary["skipped"] += 1
        else:
            to_append.append(incoming)

    if batch_update:
        try:
            utils.sheets_batch_update_with_retry(service, sheet_id, batch_update)
            summary["updated"] = len(batch_update)
        except Exception as e:  # noqa: BLE001
            print(f"[Write → {tab}] Update failed: {e}")
            summary["failed"] += len(batch_update)
    if to_append:
        try:
            utils._gsheets_call_with_retry(
                lambda r=to_append: service.spreadsheets().values().append(
                    spreadsheetId=sheet_id, range=f"{tab}!A1", valueInputOption="RAW",
                    insertDataOption="INSERT_ROWS", body={"values": r}).execute(),
                label=f"append → {tab}")
            summary["inserted"] = len(to_append)
        except Exception as e:  # noqa: BLE001
            print(f"[Write → {tab}] Append failed: {e}")
            summary["failed"] += len(to_append)
    return summary


# ═════════════════════════════════════════════════════════════════════════════
#  TAB GID + PRUNE + COLOURING
# ═════════════════════════════════════════════════════════════════════════════
def get_tab_gid(service, sheet_id, tab):
    meta = utils._gsheets_call_with_retry(
        lambda: service.spreadsheets().get(spreadsheetId=sheet_id).execute(),
        label=f"meta → {tab}")
    for s in meta.get("sheets", []):
        if s.get("properties", {}).get("title") == tab:
            return s["properties"]["sheetId"]
    return None


def prune_stale_rows(service, sheet_id, tab, columns, key_of, keep_keys):
    """Delete data rows whose key is no longer in keep_keys. Returns count."""
    last_col = utils.col_letter(len(columns) - 1)
    result = utils._gsheets_call_with_retry(
        lambda: service.spreadsheets().values()
        .get(spreadsheetId=sheet_id, range=f"{tab}!A1:{last_col}").execute(),
        label=f"read → {tab}")
    existing = result.get("values", [])
    if len(existing) <= 1:
        return 0
    header = existing[0]
    del_indices = []   # 0-based sheet row indices
    for i, vals in enumerate(existing[1:], start=1):
        rec = {header[j]: (vals[j] if j < len(vals) else "") for j in range(len(header))}
        if key_of(rec) not in keep_keys:
            del_indices.append(i)
    if not del_indices:
        return 0
    gid = get_tab_gid(service, sheet_id, tab)
    if gid is None:
        return 0
    requests = [{"deleteDimension": {"range": {
        "sheetId": gid, "dimension": "ROWS",
        "startIndex": idx, "endIndex": idx + 1}}}
        for idx in sorted(del_indices, reverse=True)]   # bottom-up
    utils._gsheets_call_with_retry(
        lambda: service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body={"requests": requests}).execute(),
        label=f"prune → {tab}")
    return len(del_indices)


def apply_header_filter(service, sheet_id, tab, columns):
    """Put a basic filter (header-row dropdowns) over the whole data range."""
    try:
        gid = get_tab_gid(service, sheet_id, tab)
        if gid is None:
            return
        result = utils._gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=sheet_id, range=f"{tab}!A1:A").execute(),
            label=f"count → {tab}")
        n_rows = len(result.get("values", []))
        if n_rows < 1:
            return
        request = {"setBasicFilter": {"filter": {"range": {
            "sheetId": gid,
            "startRowIndex": 0, "endRowIndex": n_rows,
            "startColumnIndex": 0, "endColumnIndex": len(columns)}}}}
        utils._gsheets_call_with_retry(
            lambda: service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body={"requests": [request]}).execute(),
            label=f"filter → {tab}")
    except Exception as e:  # noqa: BLE001
        print(f"[Filter → {tab}] Skipped (non-fatal): {e}")


def row_colour(recipient_status, request_status):
    rs, qs = norm(recipient_status), norm(request_status)
    # Hard-fail / attention states first.
    if (rs in {"form not sent", "declined", "authentication failed", "failed"}
            or qs in {"expired", "declined", "recalled / cancelled", "failed"}):
        return COL_RED
    # Done — recipient signed or the whole request completed.
    if rs == "signed" or qs == "completed":
        return COL_GREEN
    if rs in {"sent - not opened", "unopened"}:
        return COL_AMBER
    if rs.startswith("viewed"):
        return COL_LIGHT_GREEN
    return COL_WHITE


def apply_colours(service, sheet_id, tab, columns):
    """Colour each data row of the Current report by its Recipient/Request status."""
    last_col = utils.col_letter(len(columns) - 1)
    result = utils._gsheets_call_with_retry(
        lambda: service.spreadsheets().values()
        .get(spreadsheetId=sheet_id, range=f"{tab}!A1:{last_col}").execute(),
        label=f"read → {tab}")
    existing = result.get("values", [])
    if len(existing) <= 1:
        return
    header = existing[0]
    try:
        ri = header.index("Recipient Status")
        qi = header.index("Request Status")
    except ValueError:
        return
    gid = get_tab_gid(service, sheet_id, tab)
    if gid is None:
        return
    requests = []
    for i, vals in enumerate(existing[1:], start=1):
        rs = vals[ri] if ri < len(vals) else ""
        qs = vals[qi] if qi < len(vals) else ""
        colour = row_colour(rs, qs)
        requests.append({"repeatCell": {
            "range": {"sheetId": gid, "startRowIndex": i, "endRowIndex": i + 1,
                      "startColumnIndex": 0, "endColumnIndex": len(columns)},
            "cell": {"userEnteredFormat": {"backgroundColor": colour}},
            "fields": "userEnteredFormat.backgroundColor"}})
    # chunk to keep each batch reasonable
    for j in range(0, len(requests), 200):
        chunk = requests[j:j + 200]
        utils._gsheets_call_with_retry(
            lambda c=chunk: service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id, body={"requests": c}).execute(),
            label=f"colour → {tab}")


# ═════════════════════════════════════════════════════════════════════════════
#  REPORT RUNNERS
# ═════════════════════════════════════════════════════════════════════════════
def log_summary(title, rows, stats, up, extra=None):
    print("\n" + "=" * 62)
    print(f"  {title}")
    print("=" * 62)
    print(f"  Total Students Processed        : {stats['processed']}")
    print(f"  Records Inserted                : {up['inserted']}")
    print(f"  Records Updated                 : {up['updated']}")
    print(f"  Records Skipped (No Changes)    : {up['skipped']}")
    print(f"  Records Without Zoho Sign Req.  : {stats['without_request']}")
    print(f"  Matching by Email               : {stats['by_email']}")
    print(f"  Matching by Name                : {stats['by_name']}")
    print(f"  Matching Failed                 : {stats['failed']}")
    if up.get("failed"):
        print(f"  Records Failed                  : {up['failed']}")
    if extra:
        for k, v in extra.items():
            print(f"  {k:<32}: {v}")
    print("=" * 62)


def run_report(service, *, title, sheet_id, tab, students, by_email, by_name,
               current_only, start_date, colour, prune):
    t0 = time.time()
    rows, stats = build_rows(students, by_email, by_name,
                             current_only=current_only, start_date=start_date)
    # Sort in Python (desc by Joined On) so a first-time sheet is already ordered
    rows.sort(key=lambda r: r.get("Joined On", ""), reverse=True)

    up = upsert_report(service, sheet_id, tab, REPORT_COLUMNS, rows, report_key)

    extra = {}
    if prune:
        keep = {report_key(r) for r in rows}
        pruned = prune_stale_rows(service, sheet_id, tab, REPORT_COLUMNS,
                                  report_key, keep)
        extra["Records Pruned (no longer pending)"] = pruned

    # Physically order the tab (best-effort, reuses shared helper)
    utils.sort_sheet_by_column(service, sheet_id, tab, REPORT_COLUMNS,
                               SORT_COLUMN, descending=True)
    if colour:
        apply_colours(service, sheet_id, tab, REPORT_COLUMNS)
    apply_header_filter(service, sheet_id, tab, REPORT_COLUMNS)

    extra["Execution Time"] = f"{time.time() - t0:.1f}s"
    log_summary(title, rows, stats, up, extra)
    print(f"  Sheet: https://docs.google.com/spreadsheets/d/{sheet_id}/edit")


def main():
    print("=== Admission Formalities Reports ===")
    service = utils.get_sheets_service(SA_FILE)

    print("Reading source sheets ...")
    students = read_records(service, STUDENT_SHEET_ID, STUDENTS_TAB)
    signers  = read_records(service, SIGNERS_SHEET_ID, SIGNERS_TAB)
    print(f"  Students: {len(students)} | Signer rows: {len(signers)}")
    by_email, by_name = build_signer_indexes(signers)

    start_date = to_ist_date(CURRENT_REPORT_START_DATE) or date(2026, 7, 1)

    if GENERATE_HISTORICAL_ADMISSION_FORMALITIES_REPORT:
        run_report(service, title="REPORT 1 - Historical Admission Formalities",
                   sheet_id=HISTORICAL_SHEET_ID, tab=HISTORICAL_TAB,
                   students=students, by_email=by_email, by_name=by_name,
                   current_only=False, start_date=start_date,
                   colour=True, prune=False)
    else:
        print("Historical report skipped (flag is False).")

    if GENERATE_CURRENT_ADMISSION_FORMALITIES_REPORT:
        run_report(service, title="REPORT 2 - Current Student Admission Formalities",
                   sheet_id=CURRENT_SHEET_ID, tab=CURRENT_TAB,
                   students=students, by_email=by_email, by_name=by_name,
                   current_only=True, start_date=start_date,
                   colour=True, prune=CURRENT_REPORT_PRUNE_STALE)
    else:
        print("Current report skipped (flag is False).")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
