"""
Verification of the Daily report date of pyCoordinatorTaskListReport.py
(run from the project root:  python ops_validation\\verify_coordinator_daily_date.py).

DAILY_DATE = "YYYY-MM-DD" must produce the Daily report for exactly that day, in
AUTO mode (the scheduler's setting) and in flag mode; DAILY_DATE = None = today.
Runs the REAL generate() → _plan_jobs() → _generate_daily() → upload_datewise()
→ email_results() with only the data load, the source loaders and the Google /
SMTP calls replaced, run date fixed to 03-Oct-2026 (IST), and checks that the
pinned day is used for: job planning, data window + as-of aggregates, report
period label, file name, Drive folder, tab headings, e-mail subject and body.
Weekly / Monthly / Manual planning is checked unchanged.
No Google access needed.
"""
import io
import os
import sys
import types
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.modules.setdefault("_bootstrap", types.ModuleType("_bootstrap"))
ec = types.ModuleType("email_config")
ec.GMAIL_SENDER = "info@intellibiinnovationstechnologies.in"
ec.GMAIL_APP_PASS = "test-only"
sys.modules["email_config"] = ec                       # never the real credentials
for p in ("common", "ops_reports_action", "co-ordinator reports"):
    sys.path.insert(0, os.path.join(ROOT, p))

import openpyxl                                         # noqa: E402
import pandas as pd                                     # noqa: E402
import utils as _utils                                  # noqa: E402
import pyCoordinatorTaskListReport as BC                # noqa: E402

FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


RUN_DAY = date(2026, 10, 3)                             # the "today" of the run (Saturday)


def J(jobs):
    return [(j["kind"], str(j["date"])) if j["kind"] == "daily"
            else (j["report_type"], str(j["start"]), str(j["end"])) for j in jobs]


def flags(auto, daily_date, **kw):
    BC.GENERATE_AUTO, BC.DAILY_DATE = auto, daily_date
    BC.GENERATE_DAILY = kw.get("daily", True)
    BC.GENERATE_WEEKLY = kw.get("weekly", False)
    BC.GENERATE_MONTHLY = kw.get("monthly", False)
    BC.GENERATE_MANUAL = kw.get("manual", False)
    BC.WEEKLY_REFERENCE_DATE = kw.get("weekly_ref")
    BC.MONTHLY_MONTH = BC.MONTHLY_YEAR = None


# =============================================================================
print("\n== 1. Job planning ==")
flags(True, "2026-10-02")
check("AUTO + DAILY_DATE 2026-10-02 → Daily 02-Oct (not the run date)",
      J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-02")])
flags(True, None)
check("AUTO + DAILY_DATE None → Daily today", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-03")])
flags(True, "")
check("AUTO + DAILY_DATE '' → Daily today", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-03")])
flags(True, " 2026-10-02 ")
check("surrounding spaces tolerated", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-02")])
flags(True, date(2026, 9, 30))
check("a date object is accepted", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-09-30")])
flags(False, "2026-10-02")
check("flag mode + DAILY_DATE → that day", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-02")])
flags(False, None)
check("flag mode + None → today", J(BC._plan_jobs(RUN_DAY)), [("daily", "2026-10-03")])
for bad in ("02-10-2026", "2026-13-01", "yesterday"):
    flags(True, bad)
    try:
        BC._plan_jobs(RUN_DAY)
        got = "no error"
    except ValueError as exc:
        got = "ValueError" if bad in str(exc) else str(exc)
    check(f"malformed DAILY_DATE {bad!r} → clear error, never silently today", got, "ValueError")

print("\n== 2. Weekly / Monthly / Manual planning unchanged ==")
MON, LAST = date(2026, 10, 5), date(2026, 10, 31)
flags(True, "2026-10-02")
check("AUTO Monday: Weekly still = previous week of the RUN date",
      J(BC._plan_jobs(MON)), [("daily", "2026-10-02"), ("weekly", "2026-09-28", "2026-10-04")])
check("AUTO last day of month: Monthly still = the run date's month",
      J(BC._plan_jobs(LAST)), [("daily", "2026-10-02"), ("monthly", "2026-10-01", "2026-10-31")])
flags(True, None)
check("AUTO Monday with None: unchanged",
      J(BC._plan_jobs(MON)), [("daily", "2026-10-05"), ("weekly", "2026-09-28", "2026-10-04")])
flags(False, "2026-10-02", daily=False, weekly=True, monthly=True, weekly_ref="2026-09-16")
check("flag mode Weekly/Monthly ignore DAILY_DATE",
      J(BC._plan_jobs(RUN_DAY)), [("weekly", "2026-09-14", "2026-09-20"), ("monthly", "2026-10-01", "2026-10-31")])
flags(False, None, daily=False, manual=True)
BC.MANUAL_START_DATE, BC.MANUAL_END_DATE = "2026-09-01", "2026-09-10"
check("Manual unchanged", J(BC._plan_jobs(RUN_DAY)), [("manual", "2026-09-01", "2026-09-10")])
BC.MANUAL_START_DATE = BC.MANUAL_END_DATE = None


# =============================================================================
print("\n== 3. REAL generate() for DAILY_DATE = 2026-10-02 on a 03-Oct run ==")
def at(d, h, m=0):
    return datetime(d.year, d.month, d.day, h, m)


D1, D2, D3 = date(2026, 10, 1), date(2026, 10, 2), date(2026, 10, 3)
ATT = pd.DataFrame({"_dt": [at(D1, 11), at(D1, 13), at(D2, 20), at(D3, 9)],
                    "_date": [D1, D1, D2, D3], "tag": ["before", "in-1", "in-2", "next-day"]})
LOADS, CAPT, UPLOADS, MAILS = [], {}, [], []
_utils.get_sheets_service = lambda *a, **k: None
BC.AR._ist_today = lambda: RUN_DAY
BC.AR._ist_now = lambda: datetime(2026, 10, 3, 10, 36, 20)


def _load_all(service, start, end):
    LOADS.append(end)
    return tuple(ATT.copy() if i == 1 else pd.DataFrame() for i in range(5))   # att_agg = #1


def _cap_student(ws, att_daily, susp_daily, fb_daily, agg_map, report_date, yest_date=None):
    CAPT["att_daily"] = list(att_daily["tag"]) if att_daily is not None and len(att_daily) else []
    CAPT["student_tab_date"] = report_date
    BC.ds_title(ws, 5, "Learner Attendance Follow-Ups", report_date.strftime("%d-%b-%Y"), "x")


def _cap_aggr(att, fb, cfg):
    CAPT["agg_rows"] = list(att["tag"]) if att is not None and len(att) else []
    return {}


BC.AR.load_all_data = _load_all
BC.build_student_detail_ext = _cap_student
BC.build_aggregates = _cap_aggr
BC.load_instructor_phones = lambda s: {}
BC.load_assignment_followups = lambda s, d: CAPT.setdefault("assignment_date", d) and {}
BC.load_admission_formalities = lambda s: []
BC.load_wise_validation = lambda: {"student": [], "course": [], "instructor": []}


def _no_iv():
    raise RuntimeError("interview sources not available in the test")


BC._iv_services = _no_iv
_real_build_interview = BC.build_interview_reminders


def _upload_report(folders, filename, buf, base_prefix):
    UPLOADS.append({"folders": list(folders), "filename": filename, "xlsx": buf.getvalue()})
    return "https://docs.google.com/spreadsheets/d/test/edit"


BC.upload_report = _upload_report
BC.CE.send = lambda subject, body, recipients, sender=None, star=False: MAILS.append(
    (subject, body, list(recipients))) or True
BC.SEND_EMAIL = True
BC.PROTECT_SHEETS = False

flags(True, "2026-10-02")
res = BC.generate()
check("one report generated, not failed", [bool(r.get("failed")) for r in res], [False])
r = res[0]
check("result report_date", r["report_date"], "2026-10-02")
check("history loaded up to the run date (as before; slicing does the rest)", LOADS, [RUN_DAY])
check("data window = 01-Oct 12:00 PM → 02-Oct 11:59 PM (03-Oct rows excluded)",
      CAPT["att_daily"], ["in-1", "in-2"])
check("aggregates as of 02-Oct (03-Oct excluded)", CAPT["agg_rows"], ["before", "in-1", "in-2"])
check("assignment follow-ups loaded for 02-Oct", CAPT.get("assignment_date"), D2)
up = UPLOADS[0]
check("Drive folder = Daily 02-Oct-2026",
      up["folders"], ["Daily Coordinator Reports", "Daily 02-Oct-2026"])
check("file name carries the 02-Oct period",
      up["filename"],
      f"{BC.REPORT_BASENAME}_01-Oct-2026_12.00_PM_-_02-Oct-2026_11.59_PM.xlsx")
wb = openpyxl.load_workbook(io.BytesIO(up["xlsx"]))
heads = {ws.title: str(ws["A1"].value or "") for ws in wb.worksheets}
check("every tab heading shows 02-Oct-2026",
      {t: ("02-Oct-2026" in h) for t, h in heads.items()}, {t: True for t in heads})
check("no tab heading shows the run date 03-Oct-2026", [t for t, h in heads.items() if "03-Oct-2026" in h], [])
subj, body, rcpts = MAILS[0]
check("e-mail subject", subj, "Daily Batch Coordinator Report - 02-Oct-2026")
import re as _re                                                          # noqa: E402
_content = _re.split(r"Generated", body)[0]          # the footer stamps the real generation time
check("e-mail body names 02-Oct-2026, not 03-Oct-2026", ("02-Oct-2026" in _content, "03-Oct-2026" in _content),
      (True, False))
check("… and is not worded as today's list", ("Today's" in body, "The Coordinator task list for" in body),
      (False, True))
check("footer still shows the real generation time", "Generated" in body, True)
check("recipients unchanged", rcpts, BC.EMAIL_RECIPIENTS)

print("\n== 4. DAILY_DATE = None → today, live window up to now ==")
LOADS.clear(); CAPT.clear(); UPLOADS.clear(); MAILS.clear()
flags(True, None)
res = BC.generate()
check("report for today", res[0]["report_date"], "2026-10-03")
check("window 02-Oct 12:00 PM → now", CAPT["att_daily"], ["in-2", "next-day"])
check("Drive folder Daily 03-Oct-2026", UPLOADS[0]["folders"][-1], "Daily 03-Oct-2026")
check("subject", MAILS[0][0], "Daily Batch Coordinator Report - 03-Oct-2026")
check("today's report keeps the existing wording", "Today's Coordinator task list for" in MAILS[0][1], True)

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
