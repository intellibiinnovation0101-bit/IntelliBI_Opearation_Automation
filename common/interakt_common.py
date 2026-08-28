"""
interakt_common.py
==================
Shared helpers for the Interakt -> Google Sheet integration. Designed to sit
inside the IntelliBI Automation project and REUSE its existing Google auth
(utils.get_sheets_service / utils.upsert_rows). Nothing here duplicates the
Sheets read/write/upsert logic that already lives in utils.py.

Contents
  * InteraktClient            - authenticated, paginated access to the Interakt
                                "Get Users" API (the only bulk data-pull
                                endpoint Interakt exposes) with retry/backoff.
  * flatten_user              - nested user object -> flat spreadsheet row.
  * order_columns             - stable column order (key + common cols first).
  * get_or_create_spreadsheet - find-or-create the target spreadsheet INSIDE a
                                Drive folder (the one piece utils.py does not
                                cover, since utils uses the spreadsheets-only
                                scope). The created file's id is cached so
                                repeat runs reuse one sheet.

Docs reference:
  Get Users API -> POST https://api.interakt.ai/v1/public/apis/users/
  Auth          -> HTTP Basic: header "Authorization: Basic <SECRET_KEY>"
  Pagination    -> query params offset (0) & limit (max 100); response carries
                   "has_next_page": true|false
  Filtering     -> body {"filters":[{"trait":"created_at_utc"|"modified_at_utc",
                   "op":"gt"|"lt","val":"<ISO-UTC>","supr_op":"and"}]}
"""

from __future__ import annotations

import json
import logging
import os
import time
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional

import requests

log = logging.getLogger("interakt")

# ---------------------------------------------------------------------------
# Interakt API client
# ---------------------------------------------------------------------------

INTERAKT_USERS_URL = "https://api.interakt.ai/v1/public/apis/users/"
PAGE_LIMIT = 100          # Interakt hard maximum per page
MAX_RETRIES = 5
BACKOFF_BASE = 2.0        # seconds; exponential: BASE, BASE*2, BASE*4, ...
REQUEST_TIMEOUT = 60      # seconds


class InteraktError(RuntimeError):
    """Raised for non-retryable Interakt API failures (e.g. bad credentials)."""


class InteraktClient:
    """Thin, paginating client for the Interakt Get Users API."""

    def __init__(self, api_key: str, base_url: str = INTERAKT_USERS_URL,
                 page_limit: int = PAGE_LIMIT):
        if not api_key or api_key.startswith("YOUR_"):
            raise InteraktError(
                "Interakt API key is missing. Set INTERAKT_API_KEY or put your "
                "Secret Key (Settings > Developer Setting > Secret Key) into "
                "config_files/interakt_credentials.json."
            )
        self.api_key = api_key.strip()
        self.base_url = base_url
        self.page_limit = min(page_limit, PAGE_LIMIT)
        self.session = requests.Session()
        # Interakt expects the raw secret key after the literal word "Basic".
        self.session.headers.update({
            "Authorization": f"Basic {self.api_key}",
            "Content-Type": "application/json",
        })

    # -- low level ----------------------------------------------------------
    def _post(self, offset: int, body: Dict[str, Any]) -> Dict[str, Any]:
        """POST one page with retry/backoff on 429 and 5xx."""
        params = {"offset": offset, "limit": self.page_limit}
        last_exc: Optional[Exception] = None

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self.session.post(
                    self.base_url, params=params, json=body,
                    timeout=REQUEST_TIMEOUT,
                )
            except requests.RequestException as exc:      # network hiccup
                last_exc = exc
                wait = BACKOFF_BASE * (2 ** (attempt - 1))
                log.warning("Network error (offset=%s, attempt %s/%s): %s "
                            "-> retrying in %.0fs",
                            offset, attempt, MAX_RETRIES, exc, wait)
                time.sleep(wait)
                continue

            if resp.status_code == 200:
                return resp.json()

            # Retryable: rate limit or server error
            if resp.status_code == 429 or resp.status_code >= 500:
                retry_after = resp.headers.get("Retry-After")
                wait = (float(retry_after) if retry_after
                        else BACKOFF_BASE * (2 ** (attempt - 1)))
                log.warning("HTTP %s (offset=%s, attempt %s/%s) -> "
                            "retrying in %.0fs",
                            resp.status_code, offset, attempt, MAX_RETRIES,
                            wait)
                time.sleep(wait)
                continue

            # Non-retryable client error (401/403 bad key, 400 bad request...)
            raise InteraktError(
                f"Interakt API returned HTTP {resp.status_code} at offset "
                f"{offset}: {resp.text[:500]}"
            )

        raise InteraktError(
            f"Exhausted {MAX_RETRIES} retries at offset {offset}. "
            f"Last error: {last_exc}"
        )

    # -- public -------------------------------------------------------------
    @staticmethod
    def build_filters(created_after: Optional[datetime] = None,
                      modified_after: Optional[datetime] = None
                      ) -> Dict[str, Any]:
        """Build the request body's filter list from optional UTC datetimes."""
        filters: List[Dict[str, Any]] = []
        if created_after is not None:
            filters.append({
                "trait": "created_at_utc",
                "op": "gt",
                "val": _iso_utc(created_after),
            })
        if modified_after is not None:
            filters.append({
                "trait": "modified_at_utc",
                "op": "gt",
                "val": _iso_utc(modified_after),
            })
        # supr_op links multiple conditions with a logical AND.
        if len(filters) > 1:
            for f in filters[:-1]:
                f["supr_op"] = "and"
        return {"filters": filters}

    def iter_users(self, body: Optional[Dict[str, Any]] = None
                   ) -> Iterable[Dict[str, Any]]:
        """
        Yield every user object across all pages. Pagination stops when the API
        reports has_next_page == false OR a page is empty (defensive guard).
        """
        body = body or {"filters": []}
        offset = 0
        page_no = 0
        total = 0

        while True:
            page_no += 1
            payload = self._post(offset, body)
            users = _extract_user_list(payload)
            got = len(users)
            total += got
            log.info("Fetched page %s (offset=%s): %s users (running total %s)",
                     page_no, offset, got, total)

            for user in users:
                yield user

            has_next = _has_next_page(payload)
            if not has_next or got == 0:
                break
            offset += self.page_limit
            time.sleep(0.2)   # gentle pacing under the per-minute cap

        log.info("Get Users complete: %s users across %s page(s).",
                 total, page_no)


# ---------------------------------------------------------------------------
# Response parsing / flattening
# ---------------------------------------------------------------------------

# Interakt public Get Users wraps the list at data.customers with
# data.has_next_page. Keep the other keys for resilience across API variants.
_LIST_KEYS = ("customers", "data", "users", "results", "records", "items")


def _extract_user_list(payload: Any) -> List[Dict[str, Any]]:
    """Locate the list of user objects regardless of the exact wrapper key."""
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return []
    for key in _LIST_KEYS:
        val = payload.get(key)
        if isinstance(val, list):
            return val
        if isinstance(val, dict):
            for k2 in _LIST_KEYS:
                if isinstance(val.get(k2), list):
                    return val[k2]
    return []


def _has_next_page(payload: Any) -> bool:
    """Read has_next_page from the top level or from a nested data envelope."""
    if isinstance(payload, dict):
        if "has_next_page" in payload:
            return bool(payload["has_next_page"])
        data = payload.get("data")
        if isinstance(data, dict) and "has_next_page" in data:
            return bool(data["has_next_page"])
    return False


# ---------------------------------------------------------------------------
# Custom fields (user-created attributes in the Interakt portal)
# ---------------------------------------------------------------------------
# Mapping of the Interakt trait key -> the friendly Google-Sheet column header.
# Keys are the exact attribute keys Interakt returns inside customer `traits`
# (confirmed from GET /customers/attributes/). Order here == column order in
# the sheet. These columns are ALWAYS written (blank when a lead has no value).
# NOTE: "next_process" is not yet defined in the Interakt account (16 custom
# fields exist; this one isn't among them) — its column stays blank until the
# field is created in Interakt with this key.
CUSTOM_FIELDS = [
    ("course_interested_in", "Course Interested In"),
    ("course_advised", "Course Advised"),
    ("candidate_type", "Candidate Type"),
    ("current_role", "Current Role"),
    ("total_experience", "Total Experience"),
    ("current_company", "Current Company"),
    ("current_location", "Current Location"),
    ("career_goal", "Career Goal / Requirement"),
    ("next_process", "Next Process"),
    ("google_meet_schedule", "Google Meet Scheduled"),
    ("google_meet_datetime", "Google Meet Date & Time"),
    ("walk-in_scheduled", "Walk-in Scheduled"),
    ("walk-in_date_&_time", "Walk-in Date & Time"),
    ("declined_meet_walkin", "Declined Both Google Meet and Walk-in"),
    ("admission_status", "Admission Status"),
    ("next_follow-up_date:", "Next Follow-up Date"),
    ("counsellor_notes", "Counsellor Notes"),
]
CUSTOM_LABELS = dict(CUSTOM_FIELDS)                    # trait key -> label
CUSTOM_LABEL_ORDER = [label for _, label in CUSTOM_FIELDS]
CUSTOM_LABEL_SET = set(CUSTOM_LABEL_ORDER)


def stringify_value(value: Any) -> Any:
    """Render an API value as a flat, sheet-friendly scalar."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (list, tuple)):
        return ", ".join(str(stringify_value(v)) for v in value)
    if isinstance(value, dict):
        return "; ".join(f"{k}={stringify_value(v)}" for k, v in value.items())
    return value


def custom_fields_from_traits(traits: Dict[str, Any]) -> Dict[str, Any]:
    """Pull the user-created custom fields out of a `traits` dict, keyed by the
    friendly column label. Only non-empty values are returned (so a blank in one
    source never overwrites a value from another)."""
    out: Dict[str, Any] = {}
    if not isinstance(traits, dict):
        return out
    for key, label in CUSTOM_FIELDS:
        if key in traits and traits[key] not in (None, ""):
            out[label] = stringify_value(traits[key])
    return out


def flatten_user(user: Dict[str, Any]) -> Dict[str, Any]:
    """
    Flatten one nested user object into a single-level dict suitable for a
    spreadsheet row. Nested dicts expand with dotted keys; lists are joined.
    Standard/internal traits become 'trait_<key>' columns; user-created custom
    fields (see CUSTOM_FIELDS) become their friendly-named columns.
    """
    flat: Dict[str, Any] = {}
    for key, value in user.items():
        if key == "traits" and isinstance(value, dict):
            for tk, tv in value.items():
                if tk in CUSTOM_LABELS:
                    flat[CUSTOM_LABELS[tk]] = stringify_value(tv)
                else:
                    flat[f"trait_{tk}"] = stringify_value(tv)
        elif isinstance(value, dict):
            for sk, sv in value.items():
                flat[f"{key}.{sk}"] = stringify_value(sv)
        else:
            flat[key] = stringify_value(value)
    return flat


# Preferred left-most columns when present (everything else sorted after).
_PREFERRED_ORDER = [
    "id", "user_id", "userId", "phone_number", "phoneNumber",
    "country_code", "countryCode", "trait_name", "trait_email",
    "tags", "created_at_utc", "modified_at_utc",
]


def order_columns(rows: List[Dict[str, Any]], key_col: str) -> List[str]:
    """Stable column order: key col, then preferred cols, then the rest sorted."""
    seen = set().union(*(r.keys() for r in rows)) if rows else set()
    ordered: List[str] = []
    if key_col in seen:
        ordered.append(key_col)
    for col in _PREFERRED_ORDER:
        if col in seen and col not in ordered:
            ordered.append(col)
    for col in sorted(seen):
        if col not in ordered:
            ordered.append(col)
    return ordered


def _iso_utc(dt: datetime) -> str:
    """Format a datetime as Interakt-style ISO-8601 UTC with milliseconds."""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    dt = dt.astimezone(timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ---------------------------------------------------------------------------
# Spreadsheet find-or-create inside a Drive folder
# ---------------------------------------------------------------------------
# utils.get_sheets_service() only has the spreadsheets scope, so it cannot
# place a file in a Drive folder. This helper uses a service account credential
# scoped for BOTH sheets + drive to create the spreadsheet with a named tab and
# move it into the target folder. The resulting id is cached to state_file so
# subsequent runs reuse the same sheet (no duplicates).

_SHEETS_DRIVE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_MIME = "application/vnd.google-apps.spreadsheet"


def _build_services(service_account_file: str):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=_SHEETS_DRIVE_SCOPES)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return sheets, drive


def _load_cached_id(state_file: Optional[str]) -> Optional[str]:
    if state_file and os.path.exists(state_file):
        try:
            with open(state_file, "r", encoding="utf-8") as fh:
                return json.load(fh).get("spreadsheet_id")
        except Exception:
            return None
    return None


def _save_cached_id(state_file: Optional[str], spreadsheet_id: str) -> None:
    if not state_file:
        return
    try:
        os.makedirs(os.path.dirname(state_file), exist_ok=True)
        with open(state_file, "w", encoding="utf-8") as fh:
            json.dump({"spreadsheet_id": spreadsheet_id}, fh, indent=2)
    except Exception as exc:
        log.warning("Could not cache spreadsheet id to %s: %s", state_file, exc)


def get_or_create_spreadsheet(service_account_file: str, folder_id: str,
                              name: str, tab_name: str,
                              state_file: Optional[str] = None) -> str:
    """
    Return the spreadsheet id for `name` inside `folder_id`, creating it if it
    does not exist. Resolution order:
      1. cached id in state_file (verified to still exist & not trashed)
      2. Drive search by name within the folder
      3. create a new spreadsheet (with `tab_name` as its first tab) and move
         it into the folder
    """
    from googleapiclient.errors import HttpError
    sheets, drive = _build_services(service_account_file)

    # 1) cached id ---------------------------------------------------------
    cached = _load_cached_id(state_file)
    if cached:
        try:
            meta = drive.files().get(
                fileId=cached, fields="id, trashed, parents",
                supportsAllDrives=True).execute()
            if not meta.get("trashed", False):
                log.info("Reusing cached spreadsheet id %s.", cached)
                return cached
        except HttpError:
            log.info("Cached spreadsheet id %s no longer valid — recreating.",
                     cached)

    # 2) search folder by name --------------------------------------------
    safe_name = name.replace("'", "\\'")
    query = (f"name = '{safe_name}' and '{folder_id}' in parents and "
             f"mimeType = '{SPREADSHEET_MIME}' and trashed = false")
    try:
        files = drive.files().list(
            q=query, fields="files(id, name)", supportsAllDrives=True,
            includeItemsFromAllDrives=True).execute().get("files", [])
    except HttpError as exc:
        raise InteraktError(
            f"Drive search failed. Is the folder shared with the service "
            f"account as Editor? Folder={folder_id}. Details: {exc}")
    if files:
        found = files[0]["id"]
        log.info("Found existing spreadsheet '%s' in folder (%s).", name, found)
        _save_cached_id(state_file, found)
        return found

    # 3) create + move into folder ----------------------------------------
    created = sheets.spreadsheets().create(body={
        "properties": {"title": name},
        "sheets": [{"properties": {"title": tab_name}}],
    }).execute()
    spreadsheet_id = created["spreadsheetId"]

    prev_parents = drive.files().get(
        fileId=spreadsheet_id, fields="parents",
        supportsAllDrives=True).execute().get("parents", [])
    drive.files().update(
        fileId=spreadsheet_id,
        addParents=folder_id,
        removeParents=",".join(prev_parents) if prev_parents else None,
        fields="id, parents", supportsAllDrives=True).execute()

    log.info("Created new spreadsheet '%s' (%s) in folder %s.",
             name, spreadsheet_id, folder_id)
    _save_cached_id(state_file, spreadsheet_id)
    return spreadsheet_id
