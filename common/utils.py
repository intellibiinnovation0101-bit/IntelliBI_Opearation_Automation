"""
================================================================================
  IntelliBI Pipeline — Shared Utilities
  Imported by all pipeline scripts. Do not run directly.
================================================================================

CONTENTS:
  IST helpers         — timezone conversion functions (3 format variants)
  Sheets client       — get_sheets_service()
  Column helper       — col_letter()
  Write helpers       — append_rows_with_retry(), sheets_batch_update_with_retry(),
                        upsert_rows()
================================================================================
"""

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import os
import sys
import time
from datetime import datetime, timezone, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ─────────────────────────────────────────────────────────────────────────────
#  IST TIMEZONE
# ─────────────────────────────────────────────────────────────────────────────

IST = timezone(timedelta(hours=5, minutes=30))


def _parse_to_ist(date_str, handle_epoch: bool = False):
    """Parse a date string (or epoch ms) to an IST-aware datetime. Returns None on failure."""
    if not date_str:
        return None
    try:
        if handle_epoch and isinstance(date_str, (int, float)):
            return datetime.fromtimestamp(date_str / 1000, tz=timezone.utc).astimezone(IST)
        return datetime.fromisoformat(str(date_str).replace("Z", "+00:00")).astimezone(IST)
    except (ValueError, TypeError, OSError):
        return None


def to_ist_ymd(date_str) -> str:
    """UTC ISO string → 'YYYY-MM-DD HH:MM:SS IST'  (used by Student Info pipeline)."""
    dt = _parse_to_ist(date_str)
    if dt:
        return dt.strftime("%Y-%m-%d %H:%M:%S") + " IST"
    return str(date_str) if date_str else ""


def to_ist_dmy(date_str) -> str:
    """UTC ISO string → 'DD/MM/YYYY HH:MM:SS IST'  (used by Assignments pipeline)."""
    dt = _parse_to_ist(date_str)
    if dt:
        return dt.strftime("%d/%m/%Y %H:%M:%S") + " IST"
    return str(date_str) if date_str else ""


def to_ist_session(date_str) -> str:
    """UTC ISO or epoch-ms → 'YYYY-MM-DD HH:MM:SS'  (no suffix; used by Session pipeline)."""
    dt = _parse_to_ist(date_str, handle_epoch=True)
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def now_ist_ymd() -> str:
    """Current IST time as 'YYYY-MM-DD HH:MM:SS IST'."""
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S") + " IST"


def now_ist_session() -> str:
    """Current IST time as 'YYYY-MM-DD HH:MM:SS' (no suffix)."""
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")


def now_ist_date() -> str:
    """Current IST date as 'YYYY-MM-DD' (used for SCD-2 effective dates)."""
    return datetime.now(IST).strftime("%Y-%m-%d")


# ─────────────────────────────────────────────────────────────────────────────
#  GOOGLE SHEETS CLIENT
# ─────────────────────────────────────────────────────────────────────────────

_DEFAULT_SERVICE_ACCOUNT_FILE = os.path.join(CREDENTIALS_DIR, "service_account.json"
)


def get_sheets_service(service_account_file: str = _DEFAULT_SERVICE_ACCOUNT_FILE):
    """Build and return an authenticated Google Sheets API client."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        creds = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=scopes
        )
        return build("sheets", "v4", credentials=creds, cache_discovery=False)
    except FileNotFoundError:
        print(
            f"\n[ERROR] '{service_account_file}' not found.\n"
            "Place your Google service account JSON key in the same folder as the script.\n"
        )
        sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
#  TAB CREATION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def ensure_tab_exists(service, spreadsheet_id: str, tab_name: str):
    """
    Create the named tab if it does not already exist in the spreadsheet.
    Called automatically by append_rows_with_retry and upsert_rows so callers
    never have to pre-create tabs manually.
    """
    try:
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
        if tab_name not in existing:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
            ).execute()
            print(f"[Sheet] Tab '{tab_name}' did not exist — created.")
    except HttpError as e:
        print(f"[Sheet] Could not ensure tab '{tab_name}' exists: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  RETRY HELPER FOR GOOGLE API CALLS
# ─────────────────────────────────────────────────────────────────────────────

def _gsheets_call_with_retry(fn, label: str = "", max_retries: int = 5):
    """
    Execute a zero-argument callable that makes a Google Sheets API request.
    Retries on:
      - HTTP 429  (rate limit)     — exponential backoff: 2, 4, 8, 16, 32 s
      - HTTP 500/503 (server error) — same backoff (transient Google-side errors)
    Raises on any other error or after max_retries exhausted.
    """
    delay = 2
    for attempt in range(1, max_retries + 1):
        try:
            return fn()
        except HttpError as e:
            status = e.resp.status
            if status in (429, 500, 503) and attempt < max_retries:
                reason = "rate limit (429)" if status == 429 else f"server error ({status})"
                print(
                    f"  [Retry] {label} — {reason}, "
                    f"waiting {delay}s (attempt {attempt}/{max_retries}) ..."
                )
                time.sleep(delay)
                delay *= 2
            else:
                raise


# ─────────────────────────────────────────────────────────────────────────────
#  COLUMN LETTER HELPER
# ─────────────────────────────────────────────────────────────────────────────

def col_letter(zero_based_index: int) -> str:
    """0-based column index → spreadsheet letter(s).  0→A, 25→Z, 26→AA …"""
    result = ""
    n = zero_based_index + 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ─────────────────────────────────────────────────────────────────────────────
#  APPEND ROWS WITH RETRY
# ─────────────────────────────────────────────────────────────────────────────

def append_rows_with_retry(
    service,
    spreadsheet_id: str,
    tab_name: str,
    rows: list,
    columns: list,
    max_retries: int = 5,
):
    """
    Append rows to a sheet tab. Creates the header row on first run if empty.
    Deduplicates within the batch (by all columns except synced_at) before writing.
    Retries on HTTP 429 with exponential backoff (2 → 4 → 8 → 16 → 32 s).
    """
    if not rows:
        print(f"[Write → {tab_name}] No rows to append.")
        return

    # ── In-batch deduplication ────────────────────────────────────────────────
    rows, removed = dedup_rows(rows, columns)
    if removed:
        print(f"[Write → {tab_name}] Dedup: {removed} duplicate row(s) removed from batch.")
    if not rows:
        print(f"[Write → {tab_name}] All rows were duplicates — nothing to append.")
        return

    # ── Ensure tab exists before any read/write ───────────────────────────────
    ensure_tab_exists(service, spreadsheet_id, tab_name)

    # Ensure header exists
    try:
        result = _gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:A1")
            .execute(),
            label=f"check header → {tab_name}",
        )
        header_exists = bool(result.get("values"))
    except HttpError as e:
        print(f"[Write → {tab_name}] Error checking header after retries: {e}")
        return

    if not header_exists:
        _gsheets_call_with_retry(
            lambda: service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{tab_name}!A1",
                valueInputOption="RAW",
                body={"values": [columns]},
            ).execute(),
            label=f"write header → {tab_name}",
        )
        print(f"[Write → {tab_name}] Header row created.")

    value_matrix = [[str(row.get(col, "")) for col in columns] for row in rows]

    _gsheets_call_with_retry(
        lambda: service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{tab_name}!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": value_matrix},
        ).execute(),
        label=f"append → {tab_name}",
        max_retries=max_retries,
    )
    print(f"[Write → {tab_name}] Appended {len(rows)} row(s).")


# ─────────────────────────────────────────────────────────────────────────────
#  BATCH UPDATE WITH RETRY
# ─────────────────────────────────────────────────────────────────────────────

def sheets_batch_update_with_retry(
    service,
    spreadsheet_id: str,
    data: list,
    max_retries: int = 5,
):
    """
    Send values().batchUpdate() in chunks of 500 ranges per call.
    Retries on HTTP 429 with exponential backoff (2 → 4 → 8 → 16 → 32 s).
    """
    CHUNK_SIZE = 500
    for chunk_start in range(0, len(data), CHUNK_SIZE):
        chunk = data[chunk_start: chunk_start + CHUNK_SIZE]
        _gsheets_call_with_retry(
            lambda c=chunk: service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"valueInputOption": "RAW", "data": c},
            ).execute(),
            label="batchUpdate",
            max_retries=max_retries,
        )


# ─────────────────────────────────────────────────────────────────────────────
#  DEDUPLICATION HELPER
# ─────────────────────────────────────────────────────────────────────────────

def dedup_rows(rows: list, columns: list, exclude_cols: set = None) -> tuple:
    """
    Remove rows that are identical across all meaningful columns.

    Rules:
      - 'synced_at' is always excluded from comparison (it changes every run
        but does not indicate a new record).
      - Any columns in exclude_cols are also excluded from comparison.
      - _normalize() is applied so whitespace/null variants don't cause
        false negatives.
      - First occurrence wins; order is preserved.

    Returns:
        (deduped_rows, removed_count)
    """
    skip = {"synced_at"}
    if exclude_cols:
        skip |= set(exclude_cols)

    check_cols = [c for c in columns if c not in skip]
    seen: set  = set()
    deduped    = []

    for row in rows:
        key = tuple(_normalize(str(row.get(col, ""))) for col in check_cols)
        if key not in seen:
            seen.add(key)
            deduped.append(row)

    return deduped, len(rows) - len(deduped)


# ─────────────────────────────────────────────────────────────────────────────
#  UPSERT ROWS (appendOrUpdate)
# ─────────────────────────────────────────────────────────────────────────────

def _normalize(val: str) -> str:
    """Strip whitespace; treat 'none', 'null', 'n/a' as empty string."""
    v = str(val).strip()
    return "" if v.lower() in ("none", "null", "n/a") else v


def _realign_rows_to_columns(old_header: list, new_columns: list, data_rows: list) -> list:
    """Re-map existing sheet rows from `old_header` layout to `new_columns` layout,
    matched by column NAME. Columns new in `new_columns` become blank; columns that
    no longer exist are dropped. Used when a tab's column layout changes so that
    inserting/reordering columns never leaves existing rows shifted under the new
    header (which previously corrupted every historical/unchanged row)."""
    pos = {name: i for i, name in enumerate(old_header)}
    out = []
    for row in data_rows:
        vals = list(row)
        out.append([
            (vals[pos[c]] if (c in pos and pos[c] < len(vals)) else "")
            for c in new_columns
        ])
    return out


def upsert_rows(
    service,
    spreadsheet_id: str,
    tab_name: str,
    columns: list,
    rows: list,
    match_keys: list,
):
    """
    Upsert row-dicts into a Google Sheet tab (mirrors appendOrUpdate in n8n).

      Match found + business data changed  → batchUpdate
      Match found + nothing changed        → skip
      No match                             → bulk append

    All updates in ONE batchUpdate call; all inserts in ONE append call.
    Change detection uses normalize() to ignore whitespace and null-string variants.
    """
    if not rows:
        print(f"[Write → {tab_name}] No rows to write.")
        return

    last_col = col_letter(len(columns) - 1)

    # ── Ensure tab exists before any read/write ───────────────────────────────
    ensure_tab_exists(service, spreadsheet_id, tab_name)

    # Read full tab (with retry for transient 429/500/503 errors)
    try:
        result = _gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:{last_col}")
            .execute(),
            label=f"read → {tab_name}",
        )
        existing_values = result.get("values", [])
    except HttpError as e:
        print(f"[Write → {tab_name}] Error reading sheet after retries: {e}")
        return

    # Parse header + data rows
    if existing_values:
        header    = existing_values[0]
        data_rows = existing_values[1:]
        # If the column layout has changed since the last run (new columns added,
        # renamed, or reordered), re-align EVERY existing data row from the old
        # header to the new column order (by name) and rewrite header + rows so no
        # row is left shifted. (Previously only the header was updated, which
        # silently corrupted every row not rewritten later in this run.)
        if header != columns:
            data_rows = _realign_rows_to_columns(header, columns, data_rows)
            _gsheets_call_with_retry(
                lambda: service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"{tab_name}!A1:{last_col}{len(data_rows) + 1}",
                    valueInputOption="RAW",
                    body={"values": [columns] + data_rows},
                ).execute(),
                label=f"migrate layout → {tab_name}",
            )
            header = columns
            print(f"[Write → {tab_name}] Column layout changed — header + "
                  f"{len(data_rows)} row(s) re-aligned.")
    else:
        header    = columns
        data_rows = []
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{tab_name}!A1",
            valueInputOption="RAW",
            body={"values": [columns]},
        ).execute()
        print(f"[Write → {tab_name}] Header row created.")

    # Build lookup: match-key tuple → (sheet row number, row values)
    key_col_indices = [header.index(k) for k in match_keys if k in header]
    existing_index: dict = {}
    for i, row_vals in enumerate(data_rows):
        key_tuple = tuple(
            (row_vals[idx] if idx < len(row_vals) else "").strip()
            for idx in key_col_indices
        )
        if any(key_tuple):   # skip blank rows
            existing_index[key_tuple] = (i + 2, row_vals)   # +2: row 1 is header

    # Value columns = all columns except match keys and synced_at
    value_col_indices = [
        i for i, col in enumerate(columns)
        if col not in match_keys and col != "synced_at"
    ]

    # ── In-batch deduplication before processing ─────────────────────────────
    rows, removed = dedup_rows(rows, columns)
    if removed:
        print(f"[Write → {tab_name}] Dedup: {removed} duplicate row(s) removed from batch.")

    batch_update_data = []
    rows_to_append    = []
    updated_count     = 0
    skipped_count     = 0

    for row_dict in rows:
        incoming = [str(row_dict.get(col, "")) for col in columns]
        key      = tuple(_normalize(row_dict.get(k, "")) for k in match_keys)

        if key in existing_index:
            sheet_row, existing_vals = existing_index[key]
            padded = existing_vals + [""] * (len(columns) - len(existing_vals))
            changed = any(
                _normalize(incoming[i]) != _normalize(padded[i])
                for i in value_col_indices
            )
            if changed:
                batch_update_data.append({
                    "range":  f"{tab_name}!A{sheet_row}:{last_col}{sheet_row}",
                    "values": [incoming],
                })
                updated_count += 1
            else:
                skipped_count += 1
        else:
            rows_to_append.append(incoming)

    if batch_update_data:
        sheets_batch_update_with_retry(service, spreadsheet_id, batch_update_data)

    appended_count = 0
    if rows_to_append:
        _gsheets_call_with_retry(
            lambda r=rows_to_append: service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{tab_name}!A1",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": r},
            ).execute(),
            label=f"append new rows → {tab_name}",
        )
        appended_count = len(rows_to_append)

    print(
        f"[Write → {tab_name}] "
        f"Updated: {updated_count} | Appended: {appended_count} | Skipped (unchanged): {skipped_count}"
    )


def overwrite_rows(service, spreadsheet_id: str, tab_name: str, columns: list, rows: list):
    """
    FULL, header-aligned overwrite of a tab — for tabs that are completely rebuilt
    every run (e.g. ClassLearnerTeacherEnrolled). Clears the tab, then writes the
    header followed by all rows STRICTLY in `columns` order. This guarantees every
    value lands under its correct header and eliminates any column drift left over
    from an earlier layout change (e.g. a column inserted in the middle).

    Use this instead of upsert_rows when the incoming `rows` represent the complete
    current state of the tab (not an incremental delta).
    """
    ensure_tab_exists(service, spreadsheet_id, tab_name)

    rows, removed = dedup_rows(rows, columns)
    if removed:
        print(f"[Write → {tab_name}] Dedup: {removed} duplicate row(s) removed from batch.")

    # Build the value matrix strictly by column name → guaranteed alignment.
    matrix = [list(columns)] + [[str(r.get(col, "")) for col in columns] for r in rows]

    # Clear the entire tab first so no stale (mis-aligned) rows survive.
    _gsheets_call_with_retry(
        lambda: service.spreadsheets().values()
        .clear(spreadsheetId=spreadsheet_id, range=tab_name).execute(),
        label=f"clear → {tab_name}",
    )
    _gsheets_call_with_retry(
        lambda: service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1",
            valueInputOption="RAW", body={"values": matrix},
        ).execute(),
        label=f"overwrite → {tab_name}",
    )
    print(f"[Write → {tab_name}] Overwrote {len(rows)} row(s) in aligned column order "
          f"({len(columns)} columns).")


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET SORTER
# ─────────────────────────────────────────────────────────────────────────────

def sort_sheet_by_column(service, spreadsheet_id: str, tab_name: str,
                         columns: list, sort_column: str, descending: bool = True):
    """
    Sort the data rows (everything below the header) of a tab by a single column.

    upsert_rows() updates matched rows in place and appends new rows at the
    bottom, so it never reorders the tab. Call this afterwards to keep a tab
    physically ordered (e.g. Students by joined_on, latest first).

    Best-effort: any failure is logged and swallowed so it never blocks a run.
    """
    try:
        if sort_column not in columns:
            print(f"[Sort → {tab_name}] Column '{sort_column}' not in layout — skipped.")
            return
        col_index = columns.index(sort_column)

        # Resolve the numeric sheetId (gid) for this tab.
        meta = _gsheets_call_with_retry(
            lambda: service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute(),
            label=f"meta → {tab_name}",
        )
        sheet_id = None
        for s in meta.get("sheets", []):
            if s.get("properties", {}).get("title") == tab_name:
                sheet_id = s["properties"]["sheetId"]
                break
        if sheet_id is None:
            print(f"[Sort → {tab_name}] Tab not found — skipped.")
            return

        # Count current rows (header + data) from column A.
        result = _gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:A").execute(),
            label=f"count → {tab_name}",
        )
        n_rows = len(result.get("values", []))
        if n_rows <= 2:
            print(f"[Sort → {tab_name}] Nothing to sort ({max(n_rows - 1, 0)} data row(s)).")
            return

        request = {
            "sortRange": {
                "range": {
                    "sheetId":          sheet_id,
                    "startRowIndex":    1,            # skip the header row
                    "endRowIndex":      n_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex":   len(columns),
                },
                "sortSpecs": [{
                    "dimensionIndex": col_index,
                    "sortOrder":      "DESCENDING" if descending else "ASCENDING",
                }],
            }
        }
        _gsheets_call_with_retry(
            lambda: service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [request]},
            ).execute(),
            label=f"sort → {tab_name}",
        )
        print(f"[Sort → {tab_name}] Sorted {n_rows - 1} data row(s) by "
              f"'{sort_column}' {'DESC' if descending else 'ASC'}.")
    except Exception as e:
        print(f"[Sort → {tab_name}] Skipped (non-fatal): {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  SHEET CLEANER
# ─────────────────────────────────────────────────────────────────────────────

def clean_sheet(service, spreadsheet_id: str, tab_name: str, columns: list):
    """
    Post-write cleaning for a Google Sheet tab.  Three operations in order:

      1. Trim   — strip leading/trailing whitespace from every cell value.
      2. Purge  — drop rows where ALL cell values are empty after trimming.
      3. Dedup  — drop rows that are fully duplicate (synced_at excluded from
                  comparison since it changes every run but is not business data).

    Rewrites the sheet in-place: clear → write header + cleaned rows.
    Skips the rewrite entirely if nothing changed.
    """
    n_cols   = len(columns)
    last_col = col_letter(n_cols - 1)

    # ── Read the full sheet ───────────────────────────────────────────────────
    try:
        result = _gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:{last_col}")
            .execute(),
            label=f"clean read → {tab_name}",
        )
    except HttpError as e:
        print(f"[Clean → {tab_name}] Error reading sheet: {e}")
        return

    all_values = result.get("values", [])
    if len(all_values) <= 1:
        print(f"[Clean → {tab_name}] Nothing to clean (0 or 1 rows).")
        return

    data_rows      = all_values[1:]   # skip header row
    original_count = len(data_rows)

    # ── 1. Trim every cell + pad / truncate each row to the exact column width ─
    cleaned = []
    for row in data_rows:
        padded = (row + [""] * n_cols)[:n_cols]          # ensure exactly n_cols cells
        cleaned.append([str(cell).strip() for cell in padded])

    # ── 2. Drop rows where every cell is empty after trimming ─────────────────
    cleaned = [row for row in cleaned if any(row)]
    null_removed = original_count - len(cleaned)

    # ── 3. Deduplicate — exclude synced_at column from comparison ─────────────
    try:
        synced_idx = columns.index("synced_at")
    except ValueError:
        synced_idx = -1   # column not present — compare all

    seen:    set  = set()
    deduped: list = []
    for row in cleaned:
        key = tuple(v for i, v in enumerate(row) if i != synced_idx)
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    dup_removed = len(cleaned) - len(deduped)

    print(
        f"[Clean → {tab_name}] "
        f"Rows before: {original_count} | "
        f"All-null removed: {null_removed} | "
        f"Duplicates removed: {dup_removed} | "
        f"Rows after: {len(deduped)}"
    )

    if null_removed == 0 and dup_removed == 0:
        print(f"[Clean → {tab_name}] Sheet already clean — no rewrite needed.")
        return

    # ── Rewrite: clear sheet then write header + cleaned rows ─────────────────
    try:
        _gsheets_call_with_retry(
            lambda: service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=tab_name,
            ).execute(),
            label=f"clean clear → {tab_name}",
        )
        write_values = [columns] + deduped
        _gsheets_call_with_retry(
            lambda v=write_values: service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{tab_name}!A1",
                valueInputOption="RAW",
                body={"values": v},
            ).execute(),
            label=f"clean write → {tab_name}",
        )
        print(f"[Clean → {tab_name}] Rewritten with {len(deduped)} clean row(s).")
    except HttpError as e:
        print(f"[Clean → {tab_name}] Error during rewrite: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  SLOWLY-CHANGING-DIMENSION (TYPE 2) UPSERT
# ─────────────────────────────────────────────────────────────────────────────

def upsert_scd2(
    service,
    spreadsheet_id: str,
    tab_name: str,
    columns: list,
    rows: list,
    business_key: str,
    compare_cols: list,
    surrogate_key_col: str = "Surrogate_Key",
    active_flag_col: str = "Is_Active",
    version_col: str = "Record_Version",
    start_date_col: str = "Start_Effective_Date",
    end_date_col: str = "End_Effective_Date",
    max_retries: int = 5,
):
    """
    Load row-dicts into a Google Sheet tab using Slowly-Changing-Dimension Type-2.

    For each incoming record (one current snapshot per `business_key`):
      • No active row for the key      → INSERT new row
                                          (new Surrogate_Key, Is_Active='Y',
                                           Record_Version=1, Start=today, End=NULL)
      • Active row exists, unchanged    → SKIP (compare only `compare_cols`)
      • Active row exists, changed       → EXPIRE the active row
                                          (Is_Active='N', End_Effective_Date=today)
                                          and INSERT a new version
                                          (new Surrogate_Key, Is_Active='Y',
                                           Record_Version=prev+1, Start=today, End=NULL)

    All expirations go out in one batchUpdate; all inserts in one append, so full
    history is preserved. Records absent from `rows` are left untouched (a missing
    instructor is treated as "not updated", not deleted).

    The SCD/audit columns (surrogate_key, active flag, version, start/end dates) are
    managed entirely by this function — callers only supply business attributes.
    """
    if not rows:
        print(f"[SCD2 → {tab_name}] No rows to write.")
        return

    last_col = col_letter(len(columns) - 1)
    ensure_tab_exists(service, spreadsheet_id, tab_name)

    # ── Read existing tab ────────────────────────────────────────────────────
    try:
        result = _gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:{last_col}")
            .execute(),
            label=f"read → {tab_name}",
        )
        existing_values = result.get("values", [])
    except HttpError as e:
        print(f"[SCD2 → {tab_name}] Error reading sheet after retries: {e}")
        return

    today = now_ist_date()

    if existing_values:
        header = existing_values[0]
        data_rows = existing_values[1:]
        if header != columns:
            # Re-align EVERY existing data row from the old header to the new column
            # order (by name) BEFORE indexing active rows below — otherwise a
            # mid-schema column insert shifts historical rows and makes the
            # active-row / version lookups read the wrong cells (which duplicates
            # current instructors and corrupts history).
            data_rows = _realign_rows_to_columns(header, columns, data_rows)
            _gsheets_call_with_retry(
                lambda: service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"{tab_name}!A1:{last_col}{len(data_rows) + 1}",
                    valueInputOption="RAW",
                    body={"values": [columns] + data_rows},
                ).execute(),
                label=f"migrate layout → {tab_name}",
            )
            header = columns
            print(f"[SCD2 → {tab_name}] Column layout changed — header + "
                  f"{len(data_rows)} row(s) re-aligned.")
    else:
        data_rows = []
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{tab_name}!A1",
            valueInputOption="RAW",
            body={"values": [columns]},
        ).execute()
        print(f"[SCD2 → {tab_name}] Header row created.")

    idx = {c: i for i, c in enumerate(columns)}
    bkey_i = idx[business_key]
    sk_i = idx[surrogate_key_col]
    active_i = idx[active_flag_col]
    ver_i = idx[version_col]

    # ── Index current active rows + find the highest surrogate key ───────────
    active_index = {}   # business key → (sheet_row_number, padded_values)
    max_sk = 0
    for i, row_vals in enumerate(data_rows):
        padded = row_vals + [""] * (len(columns) - len(row_vals))
        try:
            sk_num = int(float(padded[sk_i]))
            max_sk = max(max_sk, sk_num)
        except (ValueError, TypeError):
            pass
        key = padded[bkey_i].strip()
        if key and padded[active_i].strip().upper() == "Y":
            active_index[key] = (i + 2, padded)   # +2: row 1 is header

    def build_row(row_dict, sk, version):
        out = []
        for c in columns:
            if c == surrogate_key_col:
                out.append(str(sk))
            elif c == active_flag_col:
                out.append("Y")
            elif c == version_col:
                out.append(str(version))
            elif c == start_date_col:
                out.append(today)
            elif c == end_date_col:
                out.append("")            # NULL → blank
            else:
                out.append(str(row_dict.get(c, "")))
        return out

    # ── In-batch de-dup of incoming rows by business key (keep first) ─────────
    seen_keys = set()
    deduped_rows = []
    for r in rows:
        k = str(r.get(business_key, "")).strip()
        if not k or k in seen_keys:
            continue
        seen_keys.add(k)
        deduped_rows.append(r)

    batch_update_data = []
    rows_to_append = []
    next_sk = max_sk
    inserted = updated = skipped = 0

    for row_dict in deduped_rows:
        key = str(row_dict.get(business_key, "")).strip()

        if key in active_index:
            sheet_row, existing_vals = active_index[key]
            changed = any(
                _normalize(str(row_dict.get(c, ""))) != _normalize(existing_vals[idx[c]])
                for c in compare_cols
            )
            if not changed:
                skipped += 1
                continue

            # Expire the current active version
            expired = list(existing_vals)
            expired[active_i] = "N"
            expired[idx[end_date_col]] = today
            batch_update_data.append({
                "range": f"{tab_name}!A{sheet_row}:{last_col}{sheet_row}",
                "values": [expired],
            })

            # Insert the new version
            try:
                prev_ver = int(float(existing_vals[ver_i] or 0))
            except (ValueError, TypeError):
                prev_ver = 0
            next_sk += 1
            rows_to_append.append(build_row(row_dict, next_sk, prev_ver + 1))
            updated += 1
        else:
            next_sk += 1
            rows_to_append.append(build_row(row_dict, next_sk, 1))
            inserted += 1

    if batch_update_data:
        sheets_batch_update_with_retry(service, spreadsheet_id, batch_update_data)

    if rows_to_append:
        _gsheets_call_with_retry(
            lambda r=rows_to_append: service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=f"{tab_name}!A1",
                valueInputOption="RAW",
                insertDataOption="INSERT_ROWS",
                body={"values": r},
            ).execute(),
            label=f"append new versions → {tab_name}",
            max_retries=max_retries,
        )

    print(
        f"[SCD2 → {tab_name}] "
        f"New: {inserted} | Versioned (changed): {updated} | Unchanged: {skipped}"
    )

