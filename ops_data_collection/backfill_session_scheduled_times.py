"""
================================================================================
  ONE-TIME BACKFILL — Session Scheduled Start / End (IntellBIAttendance → Sessions)
  (ops_data_collection / backfill_session_scheduled_times.py)
  ------------------------------------------------------------------------------
  WHY
    pySessionAttendanceStudentTeacherFeedbacks.py is incremental (watermark), so
    it never re-writes HISTORICAL Sessions rows. This standalone utility runs ONCE
    to:
      1. Add the two new columns to the Sessions tab header (after end_time_ist):
             Session Scheduled Start | Session Scheduled End
      2. Re-fetch ALL sessions from the API (full history) and fill those two
         columns (IST) for every EXISTING Sessions row, matched by session_id.
    After this runs, the normal daily collector keeps the columns (header now
    matches) and fills them for new sessions automatically.

  SAFE
    * Reuses the collector's own auth, API fetch, IST + scheduled-field logic —
      no duplicated business logic.
    * Realigns existing rows BY COLUMN NAME, so no data is shifted; all other
      columns are preserved exactly.
    * If the scheduled field cannot be found for ANY session (wrong API key), it
      ABORTS before writing and prints sample session keys, so it never
      blank-wipes the sheet.

  RUN (once, from the project folder):
      .venv\\Scripts\\python.exe ops_data_collection\\backfill_session_scheduled_times.py

  Options (edit below):
      START_DATE    — earliest date to fetch (default 2024-11-01, the collector's floor)
      INSPECT_ONLY  — True → only print diagnostics (session keys + match counts), no write
================================================================================
"""
from __future__ import annotations

import os
import sys
from datetime import date, datetime, timedelta

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

# Reuse the collector wholesale (auth, API fetch, IST helpers, scheduled fields).
import pySessionAttendanceStudentTeacherFeedbacks as S

# ── options ───────────────────────────────────────────────────────────────────
START_DATE   = date(2024, 11, 1)   # earliest date to fetch (matches the collector floor)
INSPECT_ONLY = False               # True = diagnostics only, no write

SCHED_START_COL = "Session Scheduled Start"
SCHED_END_COL   = "Session Scheduled End"


def _build_chunks(start: date):
    today = datetime.now(S.IST).date()
    chunks, cur = [], start
    while cur <= today:
        end = min(cur + timedelta(days=S.CHUNK_DAYS - 1), today)
        chunks.append((str(cur), str(end)))
        cur = end + timedelta(days=1)
    if chunks:
        chunks[-1] = (chunks[-1][0], str(today + timedelta(days=1)))
    return chunks


def build_scheduled_map(all_sessions):
    """session_id -> (sched_start_ist, sched_end_ist). Also returns diagnostics."""
    smap = {}
    n_start = n_end = 0
    sample_keys = []
    for sess in all_sessions:
        if not isinstance(sess, dict):
            continue
        if not sample_keys:
            sample_keys = sorted(sess.keys())
        sid = sess.get("_id") or sess.get("id") or ""
        if not sid:
            continue
        ss = S.to_ist(S._first_present(sess, S._SCHED_START_FIELDS))
        se = S.to_ist(S._first_present(sess, S._SCHED_END_FIELDS))
        if ss:
            n_start += 1
        if se:
            n_end += 1
        smap[str(sid)] = (ss, se)
    return smap, n_start, n_end, sample_keys


def run_backfill():
    print("=" * 70)
    print("  Backfill: Session Scheduled Start / End  →  Sessions tab")
    print("=" * 70)

    chunks = _build_chunks(START_DATE)
    print(f"[Fetch] {START_DATE} → today  ({len(chunks)} chunk(s))")
    all_sessions = S.fetch_all_sessions(chunks)
    smap, n_start, n_end, sample_keys = build_scheduled_map(all_sessions)
    print(f"[Map] sessions fetched={len(all_sessions)} | unique ids={len(smap)} | "
          f"with scheduled start={n_start} | with scheduled end={n_end}")
    print(f"[Diag] sample session field names: {sample_keys}")

    if smap and n_start == 0 and n_end == 0:
        print("\n[ABORT] No scheduled start/end found via the candidate field names")
        print("        (_SCHED_START_FIELDS / _SCHED_END_FIELDS in the collector).")
        print("        Nothing was written. Please share the sample field names above")
        print("        so the exact scheduled key can be pinned, then re-run.")
        sys.exit(2)

    service = S.get_sheets_service()

    # ── read the whole Sessions tab ───────────────────────────────────────────
    res = service.spreadsheets().values().get(
        spreadsheetId=S.SHEET_ID, range=S.SESSIONS_TAB).execute()
    values = res.get("values", [])
    if not values:
        print(f"[Sheet] '{S.SESSIONS_TAB}' is empty — writing header only.")
        if not INSPECT_ONLY:
            service.spreadsheets().values().update(
                spreadsheetId=S.SHEET_ID, range=f"{S.SESSIONS_TAB}!A1",
                valueInputOption="RAW", body={"values": [S.SESSIONS_COLUMNS]}).execute()
        return

    header = values[0]
    data = values[1:]
    pos = {name: i for i, name in enumerate(header)}
    sid_i = pos.get("session_id")
    if sid_i is None:
        print("[ABORT] 'session_id' column not found in the Sessions header — cannot match rows.")
        sys.exit(2)

    new_cols = S.SESSIONS_COLUMNS      # already includes the two scheduled columns
    filled = matched = 0
    out_rows = []
    for r in data:
        sid = str(r[sid_i]) if sid_i < len(r) else ""
        ss, se = smap.get(sid, ("", ""))
        if sid in smap:
            matched += 1
        if ss or se:
            filled += 1
        row_out = []
        for c in new_cols:
            if c == SCHED_START_COL:
                row_out.append(ss)
            elif c == SCHED_END_COL:
                row_out.append(se)
            elif c in pos and pos[c] < len(r):
                row_out.append(r[pos[c]])
            else:
                row_out.append("")
        out_rows.append(row_out)

    print(f"[Plan] existing rows={len(data)} | matched to API session={matched} | "
          f"rows getting scheduled values={filled}")

    if INSPECT_ONLY:
        print("[Inspect-only] No changes written. Set INSPECT_ONLY=False to apply.")
        # show a few example rows
        for row in out_rows[:3]:
            print("   e.g.", row[pos.get("session_id", 0)] if "session_id" in pos else "",
                  "->", row[new_cols.index(SCHED_START_COL)], "/", row[new_cols.index(SCHED_END_COL)])
        return

    # ── rewrite header + rows IN PLACE (same row count, only columns added) ───
    # update-only (no clear first) so the sheet is never momentarily empty.
    body = [new_cols] + out_rows
    service.spreadsheets().values().update(
        spreadsheetId=S.SHEET_ID, range=f"{S.SESSIONS_TAB}!A1",
        valueInputOption="RAW", body={"values": body}).execute()
    print(f"[Done] Sessions tab rewritten: header + {len(out_rows)} row(s); "
          f"{filled} row(s) now carry Scheduled Start/End.")
    print("       New sessions from the next daily run will fill these automatically.")


if __name__ == "__main__":
    if "--inspect" in sys.argv:
        INSPECT_ONLY = True
    run_backfill()
