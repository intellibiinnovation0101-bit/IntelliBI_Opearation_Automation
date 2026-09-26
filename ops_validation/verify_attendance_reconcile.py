"""
Verification of the Step 5c attendance reconcile fix (run from the project root:
  python ops_validation\\verify_attendance_reconcile.py). Runs the REAL transform() and
backfill_recent_attendance() of pySessionAttendanceStudentTeacherFeedbacks.py
against an in-memory fake Google Sheet + fake LMS API, reproducing the reported
case generically (a session first synced with every participant at duration 0,
later finalised by the LMS) alongside other technologies that must not change.
"""
import sys, types, os, json, copy
from datetime import datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))

# ── stub the project bootstrap modules the script imports ────────────────────
for name in ("_bootstrap",):
    sys.modules[name] = types.ModuleType(name)
paths = types.ModuleType("paths")
paths.CREDENTIALS_DIR = paths.CONFIG_DIR = paths.LOGS_DIR = paths.CACHE_DIR = HERE
sys.modules["paths"] = paths
wc = types.ModuleType("wise_config"); wc.HEADERS = {}
sys.modules["wise_config"] = wc

sys.path.insert(0, os.path.join(os.path.dirname(HERE), "ops_data_collection"))
import pySessionAttendanceStudentTeacherFeedbacks as S   # the live script

IST = timezone(timedelta(hours=5, minutes=30))
def utc(ist_str):
    """'YYYY-MM-DD HH:MM:SS' IST → ISO UTC string as the API would return it."""
    d = datetime.strptime(ist_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=IST)
    return d.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")


# ── fake Google Sheets service (values.get / update / append / batchUpdate) ──
class FakeValues:
    def __init__(self, sheet): self.sheet = sheet
    def get(self, spreadsheetId, range):
        tab = range.split("!")[0]
        rows = self.sheet.tabs.get(tab, [])
        rng = range.split("!")[1] if "!" in range else ""
        if rng == "1:1":
            rows = rows[:1]
        return _Exec({"values": copy.deepcopy(rows)})
    def update(self, spreadsheetId, range, valueInputOption, body):
        tab, a1 = range.split("!")
        rows = self.sheet.tabs.setdefault(tab, [])
        r0 = int("".join(ch for ch in a1.split(":")[0] if ch.isdigit()) or 1) - 1
        for i, v in enumerate(body["values"]):
            while len(rows) <= r0 + i: rows.append([])
            rows[r0 + i] = [str(x) for x in v]
        return _Exec({})
    def append(self, spreadsheetId, range, valueInputOption, insertDataOption, body):
        tab = range.split("!")[0]
        self.sheet.tabs.setdefault(tab, []).extend([[str(x) for x in r] for r in body["values"]])
        return _Exec({})
    def batchUpdate(self, spreadsheetId, body):
        for d in body["data"]:
            self.update(spreadsheetId, d["range"], "RAW", {"values": d["values"]})
        self.sheet.batch_calls += 1
        return _Exec({})
    def clear(self, spreadsheetId, range): return _Exec({})

class _Exec:
    def __init__(self, r): self.r = r
    def execute(self): return self.r

class FakeSheet:
    def __init__(self): self.tabs = {}; self.batch_calls = 0
    def spreadsheets(self): return self
    def values(self): return FakeValues(self)
    def get(self, spreadsheetId): return _Exec({"sheets": []})

def rows_of(sheet, tab):
    v = sheet.tabs[tab]; hdr = v[0]
    return [dict(zip(hdr, r + [""] * (len(hdr) - len(r)))) for r in v[1:]]


# ── fake LMS API ─────────────────────────────────────────────────────────────
def make_session(sid, course, title, start_ist, end_ist, participants, tutor="T One"):
    return {"_id": sid, "start_time": utc(start_ist), "end_time": utc(end_ist) if end_ist else "",
            "classId": {"_id": "cls_" + course, "name": course, "subject": title},
            "userId": {"_id": "teacher1", "name": tutor}, "participants": participants}

def part(stid, name, dur, pct, join=None, leave=None):
    p = {"wiseUserId": stid, "name": name, "inMeetingDuration": dur, "absolutePercentAttendance": pct}
    if join:  p["firstEntryTime"] = utc(join)
    if leave: p["lastExitTime"]   = utc(leave)
    return p

today = datetime.now(IST).date()
D  = str(today)                                   # the "reported" day
D1 = str(today - timedelta(days=1))
API = {"sessions": []}
S.fetch_sessions_for_chunk = lambda a, b: copy.deepcopy(API["sessions"])
S.fetch_suspended_students  = lambda class_ids, use_cache=True: {}
S._fetch_class_instructor_map = lambda: {}
S.fetch_session_attendance_detail = lambda sid, use_cache=True: []
S.backfill_session_scheduled_times = lambda service, ws: None
S._cache_get = lambda *a, **k: None
S._cache_put = lambda *a, **k: None

# Technology A (the reported pattern): 23 students, first synced with 0 duration.
A_students = [(f"stuA{i:02d}", f"Student A{i:02d}") for i in range(23)]
sess_A_early = make_session("sessA", "TechA", f"{D} To Current Date", f"{D} 06:52:22", "",
                            [part(s, "", 0, 0) for s, _ in A_students])
# Technology B: already correct at first sync (must never change).
B_students = [(f"stuB{i:02d}", f"Student B{i:02d}") for i in range(10)]
sess_B = make_session("sessB", "TechB", "batch B", f"{D1} 19:00:10", f"{D1} 21:05:00",
                      [part(s, n, 6000 if i % 2 == 0 else 0, 90 if i % 2 == 0 else 0,
                            f"{D1} 19:02:00" if i % 2 == 0 else None,
                            f"{D1} 21:00:00" if i % 2 == 0 else None)
                       for i, (s, n) in enumerate(B_students)])
# Technology C: a student marked Present in the sheet that a later fetch (API glitch)
# reports as duration 0 — must NOT be downgraded.
sess_C = make_session("sessC", "TechC", "batch C", f"{D1} 07:00:00", f"{D1} 09:00:00",
                      [part("stuC01", "C One", 5000, 80, f"{D1} 07:01:00", f"{D1} 08:59:00")])

sheet = FakeSheet()
# First sync (pre-finalised): main engine writes everything as it is at that time.
API["sessions"] = [sess_A_early, sess_B, sess_C]
out, counts, _ = S.transform(API["sessions"], {s: None for s in S.WATERMARK_SHEETS}, "2026-01-01 00:00:00", {})
S.write_all_tabs(sheet, out)

att0 = rows_of(sheet, "Attendance")
assert sum(r["status"] == "Present" for r in att0 if r["session_id"] == "sessA") == 0
assert sum(r["status"] == "Present" for r in att0 if r["session_id"] == "sessB") == 5
print(f"[setup] sheet after first sync: {len(att0)} attendance rows; TechA Present=0 (frozen snapshot)")

# ── LMS finalises TechA: 13 Present / 10 Absent; C glitches to 0; B unchanged ─
A_final = [part(s, n, 7000 if i < 13 else 0, 95 if i < 13 else 0,
                f"{D} 06:55:00" if i < 13 else None, f"{D} 11:10:00" if i < 13 else None)
           for i, (s, n) in enumerate(A_students)]
sess_A_final = make_session("sessA", "TechA", f"{D} To Current Date", f"{D} 06:52:22", f"{D} 11:12:10", A_final)
# plus one brand-new late enrolment that was not in the first snapshot
sess_A_final["participants"].append(part("stuA99", "Late Joiner", 100, 5, f"{D} 10:00:00", f"{D} 10:02:00"))
sess_C_glitch = copy.deepcopy(sess_C); sess_C_glitch["participants"] = [part("stuC01", "C One", 0, 0)]
API["sessions"] = [sess_A_final, sess_B, sess_C_glitch]

n_before = len(rows_of(sheet, "Attendance"))
refreshed = S.backfill_recent_attendance(sheet, "2026-01-02 00:00:00", use_cache=False)
att1 = rows_of(sheet, "Attendance")

a_rows = [r for r in att1 if r["session_id"] == "sessA" and r["student_id"] != "stuA99"]
present_A = [r for r in a_rows if r["status"] == "Present"]
# all 23 TechA rows change (13 Absent→Present; the 10 Absent gain the real session end)
assert refreshed == 23, refreshed
assert len(present_A) == 13, len(present_A)
assert all(r["session_end_ist"] == f"{D} 11:12:10" for r in a_rows)
assert all(r["student_name"].startswith("Student A") for r in a_rows if r["student_id"] != "stuA99")
assert all(r["first_join_ist"] == f"{D} 06:55:00" and r["duration"] == "7000" for r in present_A)
assert all(r["synced_at"] == "2026-01-02 00:00:00" for r in a_rows if r["student_id"] != "stuA99")
# late joiner appended exactly once, no duplicates of any key
late = [r for r in att1 if r["student_id"] == "stuA99"]
assert len(late) == 1 and late[0]["status"] == "Present"
keys = [(r["session_id"], r["student_id"]) for r in att1]
assert len(keys) == len(set(keys)), "duplicate attendance keys!"
assert len(att1) == n_before + 1
# TechB untouched (still 5 Present, original synced_at)
b_rows = [r for r in att1 if r["session_id"] == "sessB"]
assert sum(r["status"] == "Present" for r in b_rows) == 5
assert all(r["synced_at"] == "2026-01-01 00:00:00" for r in b_rows)
# TechC never downgraded
c = [r for r in att1 if r["session_id"] == "sessC"][0]
assert c["status"] == "Present" and c["duration"] == "5000" and c["synced_at"] == "2026-01-01 00:00:00"
print(f"[pass] TechA: Present 0 → {len(present_A)} of {len(a_rows)} (13 expected); "
      f"TechB unchanged; TechC not downgraded; late joiner appended; no duplicates.")

# ── idempotence: a second run with the same LMS data must change nothing ─────
calls = sheet.batch_calls
refreshed2 = S.backfill_recent_attendance(sheet, "2026-01-03 00:00:00", use_cache=False)
assert refreshed2 == 0 and sheet.batch_calls == calls
assert rows_of(sheet, "Attendance") == att1
print("[pass] second run with identical LMS data: 0 refreshed, sheet byte-identical.")

# ── column order independence: shuffle the sheet's Attendance columns ────────
hdr = sheet.tabs["Attendance"][0]
order = list(range(len(hdr)))[::-1]
sheet.tabs["Attendance"] = [[r[i] if i < len(r) else "" for i in order] for r in sheet.tabs["Attendance"]]
# LMS now says one more TechA student is present
sess_A_final["participants"][13] = part(A_students[13][0], A_students[13][1], 3000, 40, f"{D} 08:00:00", f"{D} 09:00:00")
refreshed3 = S.backfill_recent_attendance(sheet, "2026-01-04 00:00:00", use_cache=False)
att3 = rows_of(sheet, "Attendance")
assert refreshed3 == 1
r = [x for x in att3 if x["student_id"] == A_students[13][0]][0]
assert r["status"] == "Present" and r["duration"] == "3000" and r["attendance_percent"] == "40.00%"
assert r["course_name"] == "TechA" and r["session_id"] == "sessA"
print("[pass] reversed column order: the one changed row landed under the right headers.")

# ── refreshed count drives the Sessions re-align in main (unit check) ────────
assert isinstance(refreshed3, int)
print("ALL CHECKS PASSED")
