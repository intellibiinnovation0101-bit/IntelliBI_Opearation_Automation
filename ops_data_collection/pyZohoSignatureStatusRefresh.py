#!/usr/bin/env python3
"""
pyZohoSignatureStatusRefresh.py
===============================
Pull the people who have SIGNED your Zoho Sign documents and sync them into a
Google Sheet inside a Drive folder — built to live in the IntelliBI Automation
project and REUSE its shared plumbing, exactly like pyInteraktUsers.py:

  * utils.get_sheets_service()                  -> service-account Sheets client
  * interakt_common.get_or_create_spreadsheet() -> find-or-create sheet in folder
  * utils.upsert_rows(...)                       -> upsert-by-key (no duplicates,
                                                    auto header + column re-align)

Because it uses upsert_rows with match_keys, the "create with header if missing,
append if present" behaviour is automatic — and re-running never duplicates a
signer; changed values (e.g. a status moving to completed) are updated in place.

Data source: Zoho Sign India data-center (https://sign.zoho.in).

--------------------------------------------------------------------------------
CONFIG YOU STILL NEED TO PROVIDE  ->  config_files/zoho_credentials.json
--------------------------------------------------------------------------------
{
  "data_center":   "in",
  "client_id":     "1000.xxxxxxxx",
  "client_secret": "xxxxxxxx",
  "refresh_token": "1000.xxxxxxxx.xxxxxxxx"
}
Generate these once at https://api-console.zoho.in/ (Self Client, scope
ZohoSign.documents.READ) — see README_zoho_sign.md.

The Google service account (config_files/service_account.json) is already in
place and shared with the Drive folder, so nothing to do on the Google side.
--------------------------------------------------------------------------------
"""

from __future__ import annotations

# --- IntelliBI Operations Automation portability bootstrap (auto-inserted) ---
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap  # noqa: E402  (sys.path + env defaults + config.yaml)
from paths import CREDENTIALS_DIR, CONFIG_DIR, LOGS_DIR, CACHE_DIR as PROJECT_CACHE_DIR  # noqa: E402
# --- end bootstrap ---

import json
import logging
import os
import sys
import time
from datetime import datetime, timedelta, timezone

import requests

import utils                       # shared Google auth + upsert (same as Exotel/Interakt)
import interakt_common as ic       # reused only for get_or_create_spreadsheet()

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))

# Target Google Drive folder (the one you provided) — already shared with the
# service account as Editor.
DRIVE_FOLDER_ID  = "1bW_ABRmvM8brREcE0JvMJ-1UmqmLbVEK"
SPREADSHEET_NAME = "Zoho Sign - Signers"
TAB_NAME         = "Signers"

# Status filter. "" (empty) = pull EVERY request in every state
# (draft / inprogress / completed / expired / declined / recalled).
# Set to e.g. "completed" to restrict.
REQUEST_STATUS   = ""

SA_FILE    = os.path.join(CREDENTIALS_DIR, "service_account.json")
CRED_FILE  = os.path.join(CREDENTIALS_DIR, "zoho_credentials.json")
STATE_FILE = os.path.join(CREDENTIALS_DIR, "zoho_sign_spreadsheet.json")
LOG_FILE   = os.path.join(str(LOGS_DIR), "zoho_sign_sync.log")

IST = timezone(timedelta(hours=5, minutes=30))

# Human-friendly labels for the request-level status.
REQUEST_STATUS_LABELS = {
    "draft":      "Draft",
    "inprogress": "In Progress - Waiting for Signature",
    "completed":  "Completed",
    "expired":    "Expired",
    "declined":   "Declined",
    "recalled":   "Recalled / Cancelled",
}
# Human-friendly labels for the per-recipient action status.
ACTION_STATUS_LABELS = {
    "SIGNED":   "Signed",
    "VIEWED":   "Viewed / Opened (not signed)",
    "UNOPENED": "Sent - Not Opened",
    "NOACTION": "No Action Required",
    "DECLINED": "Declined",
    "AUTHENTICATION_FAILED": "Authentication Failed",
}

# Columns written to the sheet. request_id + recipient_email are the match keys.
COLUMNS = [
    "request_id",
    "recipient_email",
    "recipient_name",
    "recipient_status",     # friendly per-recipient status (their individual state)
    "action_type",          # SIGN / VIEW / APPROVE
    "request_name",
    "request_status",       # friendly request-level status
    "sign_percentage",      # % of signers completed on the request
    "created_date",         # request created
    "sent_date",            # submitted / sent for signature
    "last_updated",         # last modified
    "signed_date",          # completion date (for completed requests)
    "expiry_date",          # expires on
    "owner_email",
    "folder_name",
    "synced_at",            # utils treats this as non-value: it won't trigger updates
]
MATCH_KEYS = ["request_id", "recipient_email"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"),
              logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("zoho_sign")


# ─────────────────────────────────────────────────────────────────────────────
# Zoho Sign API
# ─────────────────────────────────────────────────────────────────────────────
def load_credentials() -> dict:
    if not os.path.exists(CRED_FILE):
        sys.exit(
            f"[ERROR] Zoho credentials not found: {CRED_FILE}\n"
            f"        Create it with client_id / client_secret / refresh_token "
            f"(see README_zoho_sign.md)."
        )
    with open(CRED_FILE, "r", encoding="utf-8") as fh:
        return json.load(fh)


def get_access_token(cred: dict) -> str:
    """Exchange the long-lived refresh token for a 1-hour access token."""
    dc = cred.get("data_center", "in")
    resp = requests.post(
        f"https://accounts.zoho.{dc}/oauth/v2/token",
        params={
            "refresh_token": cred["refresh_token"],
            "client_id":     cred["client_id"],
            "client_secret": cred["client_secret"],
            "grant_type":    "refresh_token",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if "access_token" not in data:
        sys.exit(f"[ERROR] Zoho token request failed: {data}")
    return data["access_token"]


def epoch_ms_to_ist(value) -> str:
    """Zoho timestamps are epoch milliseconds -> 'YYYY-MM-DD HH:MM:SS' IST."""
    try:
        ms = int(value)
        if ms <= 0:
            return ""
        return datetime.fromtimestamp(ms / 1000, tz=IST).strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return ""


def iter_requests(access_token: str, data_center: str):
    """Yield every signing request, paging through the Zoho Sign API."""
    base = f"https://sign.zoho.{data_center}/api/v1"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    start_index, row_count = 1, 100

    while True:
        # NOTE: Zoho Sign's search_columns does NOT accept request_status
        # (returns 9043 "Extra key found"), so we page through everything and
        # filter by status in main().
        page_context = {
            "row_count":   row_count,
            "start_index": start_index,
            "sort_column": "created_time",
            "sort_order":  "DESC",
        }

        resp = requests.get(
            f"{base}/requests",
            headers=headers,
            params={"data": json.dumps({"page_context": page_context})},
            timeout=60,
        )
        if resp.status_code != 200:
            log.error("List requests failed: HTTP %s | %s",
                      resp.status_code, resp.text[:300])
            resp.raise_for_status()
        payload = resp.json()

        batch = payload.get("requests", []) or []
        for req in batch:
            yield req

        if len(batch) < row_count:
            break
        start_index += row_count
        time.sleep(0.3)


def get_request_detail(access_token: str, data_center: str, request_id):
    """Fetch full detail for one request (used when the list omits actions)."""
    if not request_id:
        return None
    url = f"https://sign.zoho.{data_center}/api/v1/requests/{request_id}"
    r = requests.get(url, headers={"Authorization": f"Zoho-oauthtoken {access_token}"},
                     timeout=60)
    if r.status_code != 200:
        log.warning("Detail fetch failed for %s: HTTP %s | %s",
                    request_id, r.status_code, r.text[:200])
        return None
    data = (r.json() or {}).get("requests")
    if isinstance(data, list):
        return data[0] if data else None
    return data


def flatten_signers(req: dict, synced_at: str):
    """One dict per recipient on a request (keyed by column name for upsert_rows)."""
    request_id     = str(req.get("request_id", ""))
    request_name   = req.get("request_name", "")
    raw_status     = (req.get("request_status") or "").lower()
    status_label   = REQUEST_STATUS_LABELS.get(raw_status, raw_status or "")
    owner_email    = req.get("owner_email", "") or req.get("owner_email_id", "")
    folder_name    = req.get("folder_name", "")
    sign_pct       = req.get("sign_percentage", "")

    created_date = epoch_ms_to_ist(req.get("created_time"))
    sent_date    = epoch_ms_to_ist(req.get("sign_submitted_time"))
    last_updated = epoch_ms_to_ist(req.get("modified_time"))
    expiry_date  = epoch_ms_to_ist(req.get("expire_by"))
    # Per-signer signing timestamps are not exposed in the list API; for a
    # completed request, action_time is the completion (final-signature) time.
    signed_date  = epoch_ms_to_ist(req.get("action_time")) if raw_status == "completed" else ""

    for action in req.get("actions", []) or []:
        email = (action.get("recipient_email") or "").strip()
        if not email:
            continue
        raw_action = (action.get("action_status") or "").upper()
        yield {
            "request_id":       request_id,
            "recipient_email":  email,
            "recipient_name":   action.get("recipient_name", ""),
            "recipient_status": ACTION_STATUS_LABELS.get(raw_action, raw_action),
            "action_type":      action.get("action_type", ""),
            "request_name":     request_name,
            "request_status":   status_label,
            "sign_percentage":  sign_pct,
            "created_date":     created_date,
            "sent_date":        sent_date,
            "last_updated":     last_updated,
            "signed_date":      signed_date,
            "expiry_date":      expiry_date,
            "owner_email":      owner_email,
            "folder_name":      folder_name,
            "synced_at":        synced_at,
        }


# ─────────────────────────────────────────────────────────────────────────────
# Upsert with an explicit INSERT / UPDATE / SKIP / FAILED summary
# ─────────────────────────────────────────────────────────────────────────────
def upsert_with_summary(service, spreadsheet_id, tab_name, columns, rows, match_keys):
    """
    Same INSERT/UPDATE/SKIP logic as utils.upsert_rows, but returns a count
    summary. Decision rule:
        key not found              -> INSERT
        key found + a value changed-> UPDATE
        key found + all equal      -> SKIP (no write)
    Comparison reuses utils._normalize (trims spaces; blank/null treated equal)
    and ignores the 'synced_at' column so the run-timestamp alone never counts
    as a change.
    """
    summary = {"fetched": len(rows), "inserted": 0, "updated": 0,
               "skipped": 0, "failed": 0, "deduped": 0}
    if not rows:
        return summary

    last_col = utils.col_letter(len(columns) - 1)
    utils.ensure_tab_exists(service, spreadsheet_id, tab_name)

    # Read the whole tab
    try:
        result = utils._gsheets_call_with_retry(
            lambda: service.spreadsheets().values()
            .get(spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1:{last_col}")
            .execute(),
            label=f"read → {tab_name}")
        existing_values = result.get("values", [])
    except Exception as e:  # noqa: BLE001
        log.error("Read failed for %s: %s", tab_name, e)
        summary["failed"] = len(rows)
        return summary

    # Header + data (re-align existing rows if the column layout changed)
    if existing_values:
        header, data_rows = existing_values[0], existing_values[1:]
        if header != columns:
            data_rows = utils._realign_rows_to_columns(header, columns, data_rows)
            utils._gsheets_call_with_retry(
                lambda: service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=f"{tab_name}!A1:{last_col}{len(data_rows) + 1}",
                    valueInputOption="RAW",
                    body={"values": [columns] + data_rows}).execute(),
                label=f"migrate layout → {tab_name}")
            header = columns
            log.info("Column layout changed — header + %d row(s) re-aligned.",
                     len(data_rows))
    else:
        header, data_rows = columns, []
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1",
            valueInputOption="RAW", body={"values": [columns]}).execute()

    # Index existing rows by the match key
    key_idx = [header.index(k) for k in match_keys if k in header]
    existing_index = {}
    for i, vals in enumerate(data_rows):
        key = tuple((vals[j] if j < len(vals) else "").strip() for j in key_idx)
        if any(key):
            existing_index[key] = (i + 2, vals)   # +2: row 1 is the header

    # Columns that count as "data" for change detection (exclude keys + synced_at)
    value_idx = [i for i, c in enumerate(columns)
                 if c not in match_keys and c != "synced_at"]

    # De-duplicate the incoming batch itself first
    rows, removed = utils.dedup_rows(rows, columns)
    summary["deduped"] = removed

    batch_update, to_append = [], []
    for row_dict in rows:
        incoming = [str(row_dict.get(c, "")) for c in columns]
        key = tuple(utils._normalize(row_dict.get(k, "")) for k in match_keys)
        if key in existing_index:
            sheet_row, existing_vals = existing_index[key]
            padded = existing_vals + [""] * (len(columns) - len(existing_vals))
            changed = any(utils._normalize(incoming[i]) != utils._normalize(padded[i])
                          for i in value_idx)
            if changed:
                batch_update.append({
                    "range": f"{tab_name}!A{sheet_row}:{last_col}{sheet_row}",
                    "values": [incoming]})
            else:
                summary["skipped"] += 1
        else:
            to_append.append(incoming)

    # Apply UPDATES
    if batch_update:
        try:
            utils.sheets_batch_update_with_retry(service, spreadsheet_id, batch_update)
            summary["updated"] = len(batch_update)
        except Exception as e:  # noqa: BLE001
            log.error("Update batch failed: %s", e)
            summary["failed"] += len(batch_update)

    # Apply INSERTS
    if to_append:
        try:
            utils._gsheets_call_with_retry(
                lambda r=to_append: service.spreadsheets().values().append(
                    spreadsheetId=spreadsheet_id, range=f"{tab_name}!A1",
                    valueInputOption="RAW", insertDataOption="INSERT_ROWS",
                    body={"values": r}).execute(),
                label=f"append new rows → {tab_name}")
            summary["inserted"] = len(to_append)
        except Exception as e:  # noqa: BLE001
            log.error("Append failed: %s", e)
            summary["failed"] += len(to_append)

    return summary


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main() -> int:
    log.info("=== Zoho Sign → Google Sheet ===")
    cred = load_credentials()
    dc = cred.get("data_center", "in")
    log.info("Data center: %s | status filter: %s", dc, REQUEST_STATUS or "ALL")

    # 1) Zoho: gather signer rows
    token = get_access_token(cred)
    synced_at = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")

    rows, requests_seen, matched = [], 0, 0
    status_counts = {}
    for req in iter_requests(token, dc):
        requests_seen += 1
        if REQUEST_STATUS and (req.get("request_status") or "") != REQUEST_STATUS:
            continue
        matched += 1
        s = (req.get("request_status") or "?").lower()
        status_counts[s] = status_counts.get(s, 0) + 1
        # The list response may not include the signer 'actions'; fetch detail if so.
        if not (req.get("actions") or []):
            detail = get_request_detail(token, dc, req.get("request_id"))
            if detail:
                req = detail
        rows.extend(flatten_signers(req, synced_at))
    log.info("Requests scanned: %d | matched status '%s': %d | recipient rows: %d",
             requests_seen, REQUEST_STATUS or "ALL", matched, len(rows))
    log.info("Request status breakdown: %s",
             ", ".join(f"{k}={v}" for k, v in sorted(status_counts.items())))

    if not rows:
        log.info("No signer rows found — nothing to write.")
        return 0

    # 2) Google: find-or-create the sheet in the Drive folder
    spreadsheet_id = ic.get_or_create_spreadsheet(
        SA_FILE, DRIVE_FOLDER_ID, SPREADSHEET_NAME, TAB_NAME, state_file=STATE_FILE
    )

    # 3) Upsert with an explicit INSERT / UPDATE / SKIP / FAILED summary
    service = utils.get_sheets_service(SA_FILE)
    s = upsert_with_summary(service, spreadsheet_id, TAB_NAME, COLUMNS, rows,
                            match_keys=MATCH_KEYS)

    log.info("---------------- SYNC SUMMARY ----------------")
    log.info("Records Fetched              : %d", s["fetched"])
    log.info("Records Inserted (new)       : %d", s["inserted"])
    log.info("Records Updated (changed)    : %d", s["updated"])
    log.info("Records Unchanged / Skipped  : %d", s["skipped"])
    log.info("Records Failed               : %d", s["failed"])
    if s["deduped"]:
        log.info("Duplicate rows in batch dropped: %d", s["deduped"])
    log.info("----------------------------------------------")

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    log.info("Done. Sheet: %s", url)
    return 0


if __name__ == "__main__":
    sys.exit(main())
