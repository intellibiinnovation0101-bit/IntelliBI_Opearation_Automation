"""
Verification of ops_reports_action/pyAssignmentSubmissionPerformanceReport.py
(run from the project root:  python ops_validation\\verify_assignment_submission_reports.py).

Offline: the Submissions sheet is a synthetic DataFrame (made-up learners), Google
Drive is in memory, SMTP is replaced. Checks:
  1. periods — AUTO Daily = system date − 1, Weekly on Monday (previous Mon–Sun),
     Monthly on the 1st (previous month); flag mode; DAILY_DATE; Manual validation
     (valid, end = system date − 1, end = today, future end, start > end, missing)
  2. eligibility — deadline in period AND passed; boundary (== now / 1 s later);
     date-only deadline; missing deadline; outside the period
  3. submission — submitted / not submitted / status spellings / late vs on time /
     duplicate rows (submitted wins) / several assignments, technologies, batches
  4. reconciliation — Performance report vs Non-Submission report, cell by cell
  5. Drive — <root>/<Type> Assignment Report/<period>/<name> under BOTH roots for
     Daily / Weekly / Monthly / Manual; a re-run overwrites the period's reports in place
  6. generate() end to end on a Monday run (system date − 1, Weekly), e-mail,
     run summary, scheduler wiring
  7. Gmail Star (★) — real coordinator_email.send + common/gmail_star.py against
     an in-memory Gmail (one mailbox per account): every Daily / Weekly / Monthly
     / Manual e-mail starred in info@ only; failed send → no star attempt; star
     failure → e-mail and run unaffected
No Google access needed.
"""
import io
import os
import re
import sys
import types
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
for p in ("common", "ops_reports_action", "co-ordinator reports"):
    sys.path.insert(0, os.path.join(ROOT, p))
ec = types.ModuleType("email_config")
ec.GMAIL_SENDER = "info@intellibiinnovationstechnologies.in"
ec.GMAIL_APP_PASS = "test-only"
sys.modules["email_config"] = ec                       # never the real credentials

import pandas as pd                                     # noqa: E402
import openpyxl                                         # noqa: E402
import pyAssignmentSubmissionPerformanceReport as M     # noqa: E402
_REAL_SEND = M.CE.send                                  # the real coordinator_email.send

FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


def J(jobs):
    return [(j["kind"], str(j["start"]), str(j["end"])) for j in jobs]


def flags(auto=False, daily=False, weekly=False, monthly=False, manual=False, daily_date=None,
          weekly_ref=None, month=None, year=None, ms=None, me=None):
    M.GENERATE_AUTO, M.GENERATE_DAILY, M.GENERATE_WEEKLY = auto, daily, weekly
    M.GENERATE_MONTHLY, M.GENERATE_MANUAL = monthly, manual
    M.DAILY_DATE, M.WEEKLY_REFERENCE_DATE = daily_date, weekly_ref
    M.MONTHLY_MONTH, M.MONTHLY_YEAR = month, year
    M.MANUAL_START_DATE, M.MANUAL_END_DATE = ms, me


MON, THU1, WED = date(2026, 10, 5), date(2026, 10, 1), date(2026, 10, 7)

# =============================================================================
print("\n== 1. Reporting periods ==")
flags(auto=True)
check("AUTO Wednesday 07-Oct → Daily 06-Oct only", J(M.plan_jobs(WED)[0]), [("Daily", "2026-10-06", "2026-10-06")])
check("AUTO Monday 05-Oct → Daily 04-Oct + Weekly 28-Sep..04-Oct",
      J(M.plan_jobs(MON)[0]), [("Daily", "2026-10-04", "2026-10-04"), ("Weekly", "2026-09-28", "2026-10-04")])
check("AUTO 1st (Thu 01-Oct) → Daily 30-Sep + Monthly Sep",
      J(M.plan_jobs(THU1)[0]), [("Daily", "2026-09-30", "2026-09-30"), ("Monthly", "2026-09-01", "2026-09-30")])
check("AUTO 1st of March (leap year) → February 1..29",
      J(M.plan_jobs(date(2028, 3, 1))[0])[1], ("Monthly", "2028-02-01", "2028-02-29"))
check("AUTO 1st of January → previous December",
      J(M.plan_jobs(date(2027, 1, 1))[0])[1], ("Monthly", "2026-12-01", "2026-12-31"))
check("AUTO last day of month (31-Oct) → no Monthly", len(M.plan_jobs(date(2026, 10, 31))[0]), 1)
check("AUTO ignores the other flags and dates (also Manual)",
      (flags(auto=True, manual=True, daily_date="2026-01-01", ms="2026-01-01", me="2026-01-02"),
       J(M.plan_jobs(WED)[0]))[1], [("Daily", "2026-10-06", "2026-10-06")])
flags(daily=True)
check("flag Daily, DAILY_DATE None → system date − 1", J(M.plan_jobs(WED)[0]), [("Daily", "2026-10-06", "2026-10-06")])
flags(daily=True, daily_date="2026-10-02")
check("flag Daily, explicit DAILY_DATE", J(M.plan_jobs(WED)[0]), [("Daily", "2026-10-02", "2026-10-02")])
flags(daily=True, daily_date="2026-10-09")
jobs, errs = M.plan_jobs(WED)
check("future DAILY_DATE → clear error, no job", (J(jobs), any("future" in e for e in errs)), ([], True))
flags(weekly=True, weekly_ref="2026-09-16")
check("flag Weekly, reference date → its Mon–Sun", J(M.plan_jobs(WED)[0]), [("Weekly", "2026-09-14", "2026-09-20")])
flags(monthly=True, month=8, year=2026)
check("flag Monthly Aug-2026", J(M.plan_jobs(WED)[0]), [("Monthly", "2026-08-01", "2026-08-31")])
flags(manual=True, ms="2026-08-21", me="2026-09-22")
check("Manual valid historical range", J(M.plan_jobs(WED)[0]), [("Manual", "2026-08-21", "2026-09-22")])
flags(manual=True, ms="2026-10-01", me="2026-10-06")
check("Manual End = system date − 1 → allowed", J(M.plan_jobs(WED)[0]), [("Manual", "2026-10-01", "2026-10-06")])
for label, ms, me, frag in [("End = today", "2026-10-01", "2026-10-07", "today"),
                            ("future End", "2026-10-01", "2026-10-20", "a future date"),
                            ("Start > End", "2026-10-05", "2026-10-02", "is after"),
                            ("Start missing", None, "2026-10-02", "not both set"),
                            ("End missing", "2026-10-01", None, "not both set"),
                            ("malformed", "01-10-2026", "2026-10-02", "YYYY-MM-DD")]:
    flags(manual=True, ms=ms, me=me)
    jobs, errs = M.plan_jobs(WED)
    check(f"Manual {label} → clear error, no report", (jobs, any(frag in e for e in errs)), ([], True))
flags(manual=True, ms="2026-10-01", me="2026-10-07")
check("… error names the allowed last date", "must be on or before 06-Oct-2026" in M.plan_jobs(WED)[1][0], True)

# =============================================================================
print("\n== 2./3. Eligibility and submission rules ==")
ROWS = []


def add(aid, title, tech, batch, deadline, sid, name, status, sub_at="", marks="", cid=None):
    ROWS.append({"assessment_id": aid, "assessment_title": title, "class_id": cid or "c-" + aid,
                 "class_name": tech, "class_subject": batch,
                 "submission_start_date": "28/09/2026 10:00:00 IST", "submission_deadline": deadline,
                 "student_id": sid, "student_name": name, "student_email": f"{sid}@example.com",
                 "student_phone": "+919000000000", "submitted_at": sub_at,
                 "submission_status": status, "evaluation_marks": marks, "maximum_marks": "10"})


B1, B2 = "01-Sep-2026 To Current Date", "15-Sep-2026 To Current Date"
# A1 SQL / B1 — deadline Sun 04-Oct 23:45: 3 learners, 1 on time, 1 late, 1 not submitted
add("A1", "SQL Joins", "SQL", B1, "04/10/2026 23:45:00 IST", "s1", "Learner One", "submitted", "04/10/2026 20:00:00 IST", "8")
add("A1", "SQL Joins", "SQL", B1, "04/10/2026 23:45:00 IST", "s2", "Learner Two", "Not Submitted")
add("A1", "SQL Joins", "SQL", B1, "04/10/2026 23:45:00 IST", "s3", "Learner Three", "graded", "05/10/2026 09:00:00 IST", "6")
# duplicate rows: s2 again as submitted (wins), s1 again as not submitted (ignored)
add("A1", "SQL Joins", "SQL", B1, "04/10/2026 23:45:00 IST", "s2", "Learner Two", "submitted", "04/10/2026 23:00:00 IST")
add("A1", "SQL Joins", "SQL", B1, "04/10/2026 23:45:00 IST", "s1", "Learner One", "Not Submitted")
# A2 Power BI / B2 — deadline Sat 03-Oct 18:00 (date-only spelling variants)
add("A2", "PBI Visuals", "Power BI", B2, "03/10/2026 18:00:00 IST", "s4", "Learner Four", "not submitted")
add("A2", "PBI Visuals", "Power BI", B2, "03/10/2026 18:00:00 IST", "s5", "Learner Five", "Not Submitted")
add("A2", "PBI Visuals", "Power BI", B2, "03/10/2026 18:00:00 IST", "s1", "Learner One", "Submitted", "", "")
# A3 SQL / B2 — deadline date only (= end of day) 28-Sep
add("A3", "SQL Basics", "SQL", B2, "28/09/2026", "s6", "Learner Six", "Not Submitted")
# A4 — deadline TODAY 05-Oct 23:45 (not yet passed at the 10:30 run)
add("A4", "Python Loops", "Python", B1, "05/10/2026 23:45:00 IST", "s7", "Learner Seven", "Not Submitted")
# A5 — no deadline; A6 — outside the week (27-Sep)
add("A5", "No Deadline", "Python", B1, "", "s8", "Learner Eight", "Not Submitted")
add("A6", "Last Week", "Python", B1, "27/09/2026 23:45:00 IST", "s9", "Learner Nine", "Not Submitted")
# A7 — deadline exactly at the run time (boundary: passed) and one second later
add("A7", "Boundary Now", "Excel", B1, "05/10/2026 10:30:00 IST", "s10", "Learner Ten", "Not Submitted")
add("A8", "Boundary +1s", "Excel", B1, "05/10/2026 10:30:01 IST", "s11", "Learner Eleven", "Not Submitted")
# same assessment id in two classes = two assignments
add("A9", "Shared Quiz", "Excel", B1, "02/10/2026 23:45:00 IST", "s12", "Learner Twelve", "Submitted", "02/10/2026 21:00:00 IST", cid="cx")
add("A9", "Shared Quiz", "Excel", B2, "02/10/2026 23:45:00 IST", "s13", "Learner Thirteen", "Not Submitted", cid="cy")
DF = pd.DataFrame(ROWS)
NOW = datetime(2026, 10, 5, 10, 30)
ACTIVE = {"s1@example.com": "Y", "s2@example.com": "Y", "s4@example.com": "N", "s5@example.com": "Y"}

wk = M.CP.make_job("Weekly", date(2026, 9, 28), date(2026, 10, 4))
ds = M.build_dataset(DF, wk["start"], wk["end"], NOW, ACTIVE)
A = {(a["assessment_id"], a["class_id"]): a for a in ds["assignments"]}
check("week: eligible assignments (A1, A2, A3, A9×2) — not A4..A8",
      sorted(A), sorted([("A1", "c-A1"), ("A2", "c-A2"), ("A3", "c-A3"), ("A9", "cx"), ("A9", "cy")]))
a1 = A[("A1", "c-A1")]
check("A1 duplicates: 3 learners, 3 submitted (s2's submitted row wins), 0 not submitted",
      (a1["expected"], a1["submitted"], a1["not_submitted"]), (3, 3, 0))
check("A1 timing: on time 2 (s1 20:00, s2 23:00), late 1 (s3 next morning)", (a1["on_time"], a1["late"]), (2, 1))
check("duplicate rows counted", ds["excluded"]["duplicates"], 2)
check("A1 avg marks over submitted learners with marks", a1["avg_marks"], 7.0)
a2 = A[("A2", "c-A2")]
check("A2: 'not submitted' / 'Not Submitted' both not submitted; 'Submitted' counted",
      (a2["expected"], a2["submitted"], a2["not_submitted"], a2["rate"]), (3, 1, 2, 33.3))
check("A2 submitted without a time → 'Not recorded'",
      [r["timing"] for r in a2["rows"] if r["status"] == "Submitted"], ["Not recorded"])
check("A3 date-only deadline treated as 23:59:59", A[("A3", "c-A3")]["deadline"], datetime(2026, 9, 28, 23, 59, 59))
check("A9 in two classes = two assignments with their own batches",
      sorted((A[k]["batch"], A[k]["not_submitted"]) for k in A if k[0] == "A9"), [(B1, 0), (B2, 1)])
check("no-deadline assignment reported as excluded", list(ds["excluded"]["no_deadline"].values()), ["No Deadline"])

day = M.CP.make_job("Daily", date(2026, 10, 5), date(2026, 10, 5))
dd = M.build_dataset(DF, day["start"], day["end"], NOW, ACTIVE)
check("today's deadlines: 10:30:00 passed (boundary) — 10:30:01 and 23:45 not yet",
      ([a["assessment_id"] for a in dd["assignments"]],
       sorted(t for t, _ in dd["excluded"]["not_yet_due"].values())),
      (["A7"], ["Boundary +1s", "Python Loops"]))
d4 = M.build_dataset(DF, date(2026, 10, 4), date(2026, 10, 4), NOW, ACTIVE)
check("Daily 04-Oct (system date − 1 on 05-Oct): only A1 (deadline 04-Oct 23:45)",
      [a["assessment_id"] for a in d4["assignments"]], ["A1"])
check("active flag from the Students tab (Yes / No / Not found)",
      sorted({(r["learner"], r["active"]) for r in ds["rows"] if r["learner"] in
              ("Learner Two", "Learner Four", "Learner Six")}),
      [("Learner Four", "No"), ("Learner Six", "Not found"), ("Learner Two", "Yes")])
s = M.summarise(ds)
check("week totals", (s["assignments"], s["expected"], s["submitted"], s["not_submitted"], s["rate"]),
      (5, 9, 5, 4, 55.6))
check("technologies / batches", (s["technologies"], s["batches"]), (3, 5))

# =============================================================================
print("\n== 4. Reconciliation: Performance report vs Non-Submission report ==")


def sheet_rows(ws, hdr_text="#"):
    out, hdr = [], None
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            if row and row[0] == hdr_text and len([v for v in row if v]) > 3:
                hdr = [str(v) for v in row]
            continue
        if row[0] is None or not isinstance(row[0], int):
            continue
        out.append(dict(zip(hdr, row)))
    return out


perf_wb, ps = M.build_performance_workbook(wk, ds)
ns_wb, ns = M.build_non_submission_workbook(wk, ds)
check("Performance tabs", perf_wb.sheetnames, ["Summary", "Assignment Performance", "Learner Submission Detail"])
check("Non-Submission tabs", ns_wb.sheetnames, ["Non-Submitters", "By Learner", "Assignment Reconciliation"])
ap = sheet_rows(perf_wb["Assignment Performance"])
nsr = sheet_rows(ns_wb["Non-Submitters"])


def detail_blocks(ws):
    """Learner Submission Detail → [(banner text, header list, [row dicts])]."""
    blocks, hdr = [], None
    for row in ws.iter_rows(values_only=True):
        v0 = row[0]
        if isinstance(v0, str) and v0.strip().startswith("Class:") and "📝" in v0:
            blocks.append([v0.strip(), None, []])
        elif v0 == "#" and blocks:
            hdr = [str(v) for v in row]
            blocks[-1][1] = hdr
        elif isinstance(v0, int) and blocks:
            blocks[-1][2].append(dict(zip(blocks[-1][1], row)))
    return blocks


DBLK = detail_blocks(perf_wb["Learner Submission Detail"])
detail = [dict(r, Assignment=b.split("|")[2].replace("📝", "").strip()) for b, _h, rs in DBLK for r in rs]
rec = sheet_rows(ns_wb["Assignment Reconciliation"])
per_asg_perf = {(r["Assignment"], r["Batch"]): r["Not Submitted"] for r in ap}
per_asg_ns = {}
for r in nsr:
    per_asg_ns[(r["Assignment"], r["Batch"])] = per_asg_ns.get((r["Assignment"], r["Batch"]), 0) + 1
check("per assignment: Not Submitted (performance) = rows listed (non-submission)",
      {k: v for k, v in per_asg_perf.items() if v}, per_asg_ns)
check("same learners: detail 'Not Submitted' rows = Non-Submitters rows",
      sorted((r["Assignment"], r["Learner"]) for r in detail if r["Status"] == "Not Submitted"),
      sorted((r["Assignment"], r["Learner"]) for r in nsr))
check("reconciliation tab = performance tab (expected / submitted / not submitted)",
      [(r["Assignment"], r["Expected"], r["Submitted"], r["Not Submitted (listed)"]) for r in rec],
      [(r["Assignment"], r["Expected"], r["Submitted"], r["Not Submitted"]) for r in ap])
check("totals agree", (ps["not_submitted"], ns["pending"], len(nsr)), (4, 4, 4))
check("expected = submitted + not submitted on every assignment",
      all(r["Expected"] == r["Submitted"] + r["Not Submitted"] for r in ap), True)
check("detail rows = expected learners", len(detail), ps["expected"])
bl = sheet_rows(ns_wb["By Learner"])
check("By Learner: one row per learner, counts add up",
      (len(bl), sum(r["Pending Assignments"] for r in bl)), (ns["learners"], 4))
sm = perf_wb["Summary"]
kpi_vals = [sm.cell(row=4, column=c).value for c in (1, 3, 4, 6, 8)]
check("Summary KPIs (Attendance-report order: totals, %, not submitted, submitted)",
      kpi_vals, [5, 9, "55.6%", 4, 5])

print("\n== 4b. Colour code = pyAttendaceFeedbackReport ==")
AR = M.AR


def fill(c):
    return (c.fill.fgColor.rgb or "")[-6:].upper()


def cells_of(ws, header, hdr_text="#"):
    hdr_row = next(r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == hdr_text)
    col = next(c for c in range(1, ws.max_column + 1) if ws.cell(row=hdr_row, column=c).value == header)
    return [ws.cell(row=r, column=col) for r in range(hdr_row + 1, ws.max_row + 1)
            if isinstance(ws.cell(row=r, column=1).value, int)]


pct_ok = all(fill(c) == AR._att_bg(c.value, "Present").upper()
             for ws, h in ((perf_wb["Assignment Performance"], "Submission %"),
                           (perf_wb["Summary"], "Submission %"),
                           (ns_wb["Assignment Reconciliation"], "Submission %"))
             for c in cells_of(ws, h))
check("Submission % cells use the Attendance report's Att % fills (AR._att_bg)", pct_ok, True)
check("… text dark green at ≥ 75 %, dark red below",
      all((c.font.color.rgb[-6:] == AR.C_GREEN_DARK) == (c.value >= 75)
          for c in cells_of(perf_wb["Assignment Performance"], "Submission %")), True)
print("\n== 4c. Learner Submission Detail: columns, assignment banners, row colours ==")
DCOLS = ["#", "Learner", "Email", "Phone", "Submitted At", "Status", "Marks Obtained",
         "Max Marks", "Feedback", "Active Learner"]
check("only the 10 requested columns, in order, under every assignment banner",
      {tuple(h) for _b, h, _r in DBLK}, {tuple(DCOLS)})
_want_banners = [f"Class: {a['technology']}  |  Duration: {a['batch']}  |  📝  {a['title']}  |  Start: "
                 f"{a['start']:%m/%d/%Y}  |  Deadline: {a['deadline']:%m/%d/%Y}" for a in ds["assignments"]]
check("one banner per assignment, values from the data (class, duration, name, start, deadline)",
      [b for b, _h, _r in DBLK], _want_banners)
check("example banner: Class | Duration | 📝 Assignment | Start | Deadline",
      DBLK[[b for b, *_ in DBLK].index(next(b for b in _want_banners if "SQL Joins" in b))][0],
      "Class: SQL  |  Duration: 01-Sep-2026 To Current Date  |  📝  SQL Joins  |  "
      "Start: 09/28/2026  |  Deadline: 10/04/2026")
check("Duration = each class's class_subject in the Submissions data (assignment + class)",
      [b.split("|")[1].replace("Duration:", "").strip() for b, *_ in DBLK],
      [DF.loc[(DF["assessment_id"].astype(str) == str(a["assessment_id"]))
              & (DF["class_id"].astype(str) == str(a["class_id"])), "class_subject"].iloc[0].strip()
       for a in ds["assignments"]])
_blank = dict(ds["assignments"][0], batch="—")
check("blank class duration → 'Duration: —'", "|  Duration: —  |" in M._assignment_banner(_blank), True)
check("each block lists exactly that assignment's learners",
      [len(r) for _b, _h, r in DBLK], [a["expected"] for a in ds["assignments"]])
_a1 = next(rs for b, _h, rs in DBLK if "SQL Joins" in b)
_one = next(r for r in _a1 if r["Learner"] == "Learner One")
check("Marks Obtained / Max Marks / Feedback / Submitted At mapped from the data",
      (_one["Marks Obtained"], _one["Max Marks"], _one["Feedback"], _one["Submitted At"]),
      (8, "10", "", datetime(2026, 10, 4, 20, 0)))
det = perf_wb["Learner Submission Detail"]
_row_style = {}
for r in range(1, det.max_row + 1):
    if isinstance(det.cell(row=r, column=1).value, int):
        cells = [det.cell(row=r, column=c) for c in range(1, 11)]
        _row_style[(det.cell(row=r, column=2).value, det.cell(row=r, column=6).value,
                    det.cell(row=r, column=10).value, r)] = (
            {fill(c) for c in cells}, {bool(c.font.strike) for c in cells})
_by = lambda st, act: [v for (l, s_, a_, _r), v in _row_style.items() if s_ == st and a_ == act]
check("Submitted (active) → whole row green, no strikethrough",
      {(tuple(f), tuple(k)) for f, k in _by("Submitted", "Yes")}, {((AR.C_GREEN_PALE,), (False,))})
check("Not Submitted (active) → whole row red, no strikethrough",
      {(tuple(f), tuple(k)) for f, k in _by("Not Submitted", "Yes")}, {((AR.C_RED_PALE,), (False,))})
check("Active Learner = No takes priority → whole row amber + strikethrough",
      {(tuple(f), tuple(k)) for f, k in _by("Not Submitted", "No")}, {((AR.C_AMBER,), (True,))})
_sub_inactive = M._learner_row_style({"active": "No", "status": "Submitted"})
check("… also for a submitted inactive learner",
      (_sub_inactive[0], _sub_inactive[1](False).strike), (AR.C_AMBER, True))
check("'Not found' / 'Unknown' activity is not treated as inactive",
      [M._learner_row_style({"active": a, "status": "Submitted"})[0] for a in ("Not found", "Unknown")],
      [AR.C_GREEN_PALE, AR.C_GREEN_PALE])
tot = [r for r in range(1, perf_wb["Assignment Performance"].max_row + 1)
       if perf_wb["Assignment Performance"].cell(row=r, column=1).value == "⬛  TOTAL"]
check("TOTAL row in the Attendance report's light-blue band",
      (len(tot), fill(perf_wb["Assignment Performance"].cell(row=tot[0], column=8))), (1, AR.C_BLUE_LITE))
check("section banners in the Attendance report's blue",
      fill(next(c for row in sm.iter_rows() for c in row if c.value and "BATCH-WISE" in str(c.value))),
      AR.C_BLUE_MID)
check("Non-Submitters rows in the absent reds",
      {fill(c) for c in cells_of(ns_wb["Non-Submitters"], "Learner")}, {AR.C_RED_LITE, AR.C_RED_PALE})
_sum_txt = [str(c.value) for row in sm.iter_rows() for c in row if c.value]
check("Summary: no Technology-wise table, no NOTES section; Batch-wise table kept",
      (any("TECHNOLOGY-WISE" in t for t in _sum_txt), any("NOTES" in t for t in _sum_txt),
       any("BATCH-WISE" in t for t in _sum_txt)), (False, False, True))
_last = max(r for r in range(1, sm.max_row + 1)
            if any(sm.cell(row=r, column=c).value not in (None, "") for c in range(1, 9)))
check("Summary ends with the Batch-wise TOTAL row (no trailing rows or gaps)",
      sm.cell(row=_last, column=1).value, "⬛  TOTAL")
_bw = sheet_rows(sm)
check("Batch-wise rows (technology + batch) and totals", (len(_bw), sum(r["Expected"] for r in _bw)), (5, 9))
_det = perf_wb["Learner Submission Detail"]
_hdr_rows = [r for r in range(1, _det.max_row + 1) if _det.cell(row=r, column=1).value == "#"]
check("filters kept on every detail tab; Learner Submission Detail filter spans all its blocks",
      [bool(perf_wb["Assignment Performance"].auto_filter.ref),
       bool(ns_wb["Non-Submitters"].auto_filter.ref), bool(ns_wb["By Learner"].auto_filter.ref),
       _det.auto_filter.ref], [True, True, True, f"A{_hdr_rows[0]}:J{_det.max_row}"])
check("Learner Submission Detail: no frozen rows or columns", _det.freeze_panes, None)
check("Assignment Performance: no frozen rows or columns", perf_wb["Assignment Performance"].freeze_panes, None)
check("deadlines stored as real date-times (sortable)",
      isinstance(ap[0]["Deadline (IST)"], datetime), True)
empty = M.build_dataset(DF, date(2026, 10, 6), date(2026, 10, 6), NOW, ACTIVE)
ew, es = M.build_performance_workbook(M.CP.make_job("Daily", date(2026, 10, 6), date(2026, 10, 6)), empty)
nw, _n = M.build_non_submission_workbook(M.CP.make_job("Daily", date(2026, 10, 6), date(2026, 10, 6)), empty)
txt = " ".join(str(c.value) for ws in list(ew) + list(nw) for row in ws.iter_rows() for c in row if c.value)
check("empty period: clear message, no blank tables", "nothing to evaluate" in txt and
      "nothing to follow up" in txt, True)

# =============================================================================
print("\n== 5. Google Drive layout ==")


class _Req:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class FakeDrive:
    def __init__(self):
        self.items = {}
        self.n = 0
        for rid, nm in ((M.PERFORMANCE_ROOT_FOLDER_ID, "Assignment Submission Report"),
                        (M.NON_SUBMISSION_ROOT_FOLDER_ID, "Assignment Not Submitted Students List")):
            self.items[rid] = {"id": rid, "name": nm, "mimeType": M.FOLDER_MIME, "parents": []}

    def files(self):
        return self

    def list(self, q=None, **kw):
        def ok(it):
            for cl in q.split(" and "):
                cl = cl.strip()
                m = re.match(r"^'(.+)' in parents$", cl)
                if m and m.group(1) not in it["parents"]:
                    return False
                m = re.match(r"^name='(.*)'$", cl)
                if m and it["name"] != m.group(1).replace("\\'", "'"):
                    return False
                m = re.match(r"^name contains '(.*)'$", cl)
                if m and m.group(1).replace("\\'", "'") not in it["name"]:
                    return False
                m = re.match(r"^mimeType='(.*)'$", cl)
                if m and it["mimeType"] != m.group(1):
                    return False
                if cl.replace(" ", "") == "trashed=false" and it.get("trashed"):
                    return False
            return True
        return _Req(lambda: {"files": [dict(it) for it in self.items.values() if ok(it)]})

    def create(self, body=None, media_body=None, **kw):
        def go():
            self.n += 1
            fid = f"f{self.n}"
            data = media_body.getbytes(0, media_body.size()) if media_body is not None else None
            self.items[fid] = {"id": fid, "name": body["name"], "mimeType": body["mimeType"],
                               "parents": list(body.get("parents", [])), "data": data}
            return {"id": fid, "webViewLink": f"https://docs.google.com/spreadsheets/d/{fid}/edit"}
        return _Req(go)

    refuse_content_update = False

    def update(self, fileId=None, body=None, media_body=None, **kw):
        def go():
            it = self.items[fileId]
            if media_body is not None:
                if self.refuse_content_update:
                    raise RuntimeError("403 content update refused")
                it["data"] = media_body.getbytes(0, media_body.size())
                it["updates"] = it.get("updates", 0) + 1
            for k, v in (body or {}).items():
                it[k] = v
            return {"id": fileId, "webViewLink": f"https://docs.google.com/spreadsheets/d/{fileId}/edit"}
        return _Req(go)

    def path(self, fid):
        out, cur = [], self.items[fid]
        while cur["parents"]:
            out.append(cur["name"])
            cur = self.items[cur["parents"][0]]
        return cur["name"] + "/" + "/".join(reversed(out))

    def files_under(self):
        return sorted(self.path(i) for i, it in self.items.items()
                      if it["mimeType"] == M.SHEET_MIME and not it.get("trashed"))


drive = FakeDrive()
MAILS = []
M.CE.send = lambda subject, body, rcpts, sender=None, star=False: MAILS.append(
    (subject, body, list(rcpts), sender, star)) or True
jobs = [M.CP.make_job("Daily", date(2026, 10, 4), date(2026, 10, 4)), wk,
        M.CP.make_job("Monthly", date(2026, 9, 1), date(2026, 9, 30)),
        M.CP.make_job("Manual", date(2026, 8, 21), date(2026, 9, 22))]
res = M.run_jobs(jobs, DF, NOW, ACTIVE, drive, upload=True, email=True)
check("all four built", [bool(r.get("failed")) for r in res], [False] * 4)
_hdr_ok, _hdr_style = [], set()
for _r in res:
    _j = _r["job"]
    for _wb in _r["workbooks"]:
        for _ws in _wb.worksheets:
            _a1 = _ws["A1"]
            _hdr_ok.append(bool(re.fullmatch(
                r".+  \|  " + re.escape(f"{_j['kind']}  ·  {_j['label']}")
                + r"  \|  Generated On: 05-Oct-2026 10:30 AM", str(_a1.value))))
            _hdr_style.add((_a1.font.b, _a1.font.sz, _a1.alignment.horizontal,
                            any(str(m).startswith("A1:") for m in _ws.merged_cells.ranges)))
check("every tab header (Daily / Weekly / Monthly / Manual, both reports) = existing header "
      "+ '  |  Generated On: <IST run time>'", (len(_hdr_ok), all(_hdr_ok)), (24, True))
check("… header styling, alignment and merge unchanged", _hdr_style, {(True, 14.0, "center", True)})
P, N = "Assignment Submission Report", "Assignment Not Submitted Students List"
want = []
for root, base in ((P, M.PERFORMANCE_BASENAME), (N, M.NON_SUBMISSION_BASENAME)):
    want += [f"{root}/Daily Assignment Report/Daily 04-Oct-2026/{base}_Daily_04-Oct-2026",
             f"{root}/Weekly Assignment Report/Weekly 28-Sep-2026 to 04-Oct-2026/{base}_Weekly_28-Sep-2026_to_04-Oct-2026",
             f"{root}/Monthly Assignment Report/Monthly Sep-2026/{base}_Monthly_September_2026",
             f"{root}/Manual Assignment Report/Manual 21-Aug-2026 to 22-Sep-2026/{base}_Manual_21-Aug-2026_to_22-Sep-2026"]
check("each report in its root / type folder / period folder", drive.files_under(), sorted(want))
check("uploaded as native Google Sheets from an in-memory workbook",
      all(it["mimeType"] == M.SHEET_MIME and it["data"][:2] == b"PK"
          for it in drive.items.values() if it["mimeType"] != M.FOLDER_MIME), True)
up_wb = openpyxl.load_workbook(io.BytesIO([it for it in drive.items.values()
                                           if it["name"].endswith("_Weekly_28-Sep-2026_to_04-Oct-2026")
                                           and "Non_Submission" in it["name"]][0]["data"]))
check("uploaded Non-Submission workbook lists the 4 non-submitters",
      len(sheet_rows(up_wb["Non-Submitters"])), 4)
_daily_ids = {it["name"]: fid for fid, it in drive.items.items() if it["name"].endswith("_Daily_04-Oct-2026")}
_files_before = drive.files_under()
DF2 = DF[~((DF["assessment_id"] == "A1") & (DF["student_id"] == "s3"))]       # data changed since
_r = M.run_jobs(jobs[:1], DF2, NOW, ACTIVE, drive, upload=True, email=False)
check("re-run overwrites the existing reports in place: same files, same names, no '- Version'",
      (drive.files_under(), {it["name"]: fid for fid, it in drive.items.items()
                             if it["name"].endswith("_Daily_04-Oct-2026")}),
      (_files_before, _daily_ids))
_upd = openpyxl.load_workbook(io.BytesIO(drive.items[_daily_ids[f"{M.PERFORMANCE_BASENAME}_Daily_04-Oct-2026"]]["data"]))
check("… with the new content (A1 now 2 learners)",
      [r["Expected"] for r in sheet_rows(_upd["Assignment Performance"])], [2])
check("… and the same links", _r[0]["perf_link"].split("/d/")[1].split("/")[0],
      _daily_ids[f"{M.PERFORMANCE_BASENAME}_Daily_04-Oct-2026"])
drive.refuse_content_update = True
M.run_jobs(jobs[:1], DF, NOW, ACTIVE, drive, upload=True, email=False)
drive.refuse_content_update = False
check("in-place overwrite refused → replaced: still one report per name, old copy in the trash",
      (drive.files_under() == _files_before,
       sorted(it["name"] for it in drive.items.values() if it.get("trashed"))),
      (True, sorted(_daily_ids)))
M.OVERWRITE_EXISTING = False
M.run_jobs(jobs[:1], DF, NOW, ACTIVE, drive, upload=True, email=False)
M.OVERWRITE_EXISTING = True
check("OVERWRITE_EXISTING = False keeps the old file and adds '- Version 2'",
      sorted(n.rsplit("/", 1)[1] for n in drive.files_under() if "Version" in n),
      sorted(f"{b}_Daily_04-Oct-2026 - Version 2" for b in (M.PERFORMANCE_BASENAME, M.NON_SUBMISSION_BASENAME)))

print("\n== E-mail ==")
check("one e-mail per period", len(MAILS), 4)
subj, body, rcpts, sender, star = MAILS[1]
check("subject", subj, "Weekly Assignment Submission Report - 28-Sep-2026 to 04-Oct-2026")
check("from info@ to info@ + intellibihropsb2ch (as the Coordinator Performance report)",
      (sender, rcpts), ("info@intellibiinnovationstechnologies.in",
                        ["info@intellibiinnovationstechnologies.in", "intellibihropsb2ch@gmail.com"]))
links = re.findall(r"https://docs\.google\.com/spreadsheets/d/(f\d+)/edit", body)
check("links to both reports of that period",
      sorted({drive.items[f]["name"].split("_Report_")[0] for f in links}),
      ["IntelliBI_Assignment_Non_Submission", "IntelliBI_Assignment_Submission_Performance"])
check("body shows the period's KPIs", all(x in body for x in ("55.6%", "Not Submitted", "By Technology")), True)
_rows = re.findall(r"<tr style='background:(#[0-9A-F]+);color:#1a2a48'><td[^>]*>([^<]+)</td>"
                   r".*?<td[^>]*color:(#[0-9A-F]+);font-weight:700'>([^<]+)</td></tr>", body)
_want = []
for _k, _c, _n in M.group_rows(M.build_dataset(DF, wk["start"], wk["end"], NOW, ACTIVE),
                               lambda a: a["technology"]):
    _bg, _tx = M.AR._email_att_color(_c["rate"])
    _want.append((_bg, _k, _tx, f"{_c['rate']:.1f}%"))
check("By Technology rows colour-coded with the Attendance e-mail's bands (AR._email_att_color)",
      _rows, _want)
check("… same columns / order as before, with the colour legend",
      ("<th" in body and "Green &gt; 70%" in body and "Amber 55–70%" in body and "Red &lt; 55%" in body,
       [r[1] for r in _rows]), (True, [k for k, *_ in M.group_rows(
           M.build_dataset(DF, wk["start"], wk["end"], NOW, ACTIVE), lambda a: a["technology"])]))
_b = M._tech_table_html([("T70", {"expected": 10, "submitted": 7, "not_submitted": 3, "rate": 70.0}, 1),
                         ("T71", {"expected": 100, "submitted": 71, "not_submitted": 29, "rate": 71.0}, 1),
                         ("T55", {"expected": 20, "submitted": 11, "not_submitted": 9, "rate": 55.0}, 1),
                         ("T54", {"expected": 100, "submitted": 54, "not_submitted": 46, "rate": 54.0}, 1),
                         ("T0", {"expected": 0, "submitted": 0, "not_submitted": 0, "rate": None}, 0)])
check("band boundaries: >70 green, 70 & 55 amber, <55 red, no expected = neutral",
      re.findall(r"<tr style='background:(#[0-9A-Fa-f]+)", _b),
      ["#FFF4D6", "#E8F5E9", "#FFF4D6", "#FDE0E0", "#ffffff", "#FFF4D6"])

_TOT = re.compile(r"<tr style='background:(#[0-9A-Fa-f]+);color:#1a2a48;font-weight:700'>"
                  r"((?:<td[^>]*>[^<]*</td>)+)</tr>")


def _total_row(html):
    m = _TOT.findall(html)
    return [(bg, re.findall(r"<td[^>]*>([^<]*)</td>", cells)) for bg, cells in m]


_ds_wk = M.build_dataset(DF, wk["start"], wk["end"], NOW, ACTIVE)
_s_wk = M.summarise(_ds_wk)
_tb = M.AR._email_att_color(_s_wk["rate"])[0]
check("e-mail By Technology ends with one bold TOTAL row = the period's overall figures",
      _total_row(body), [(_tb, ["TOTAL", str(_s_wk["assignments"]), str(_s_wk["expected"]),
                                str(_s_wk["submitted"]), str(_s_wk["not_submitted"]),
                                f"{_s_wk['rate']:.1f}%"])])
check("… TOTAL is the last row of the table",
      body.split("By Technology", 1)[1].split("</table>")[0].rsplit("<tr", 1)[1].count(">TOTAL<"), 1)
_w = M._tech_table_html([("Python", {"expected": 10, "submitted": 10, "not_submitted": 0, "rate": 100.0}, 2),
                         ("SQL", {"expected": 90, "submitted": 36, "not_submitted": 54, "rate": 40.0}, 3)])
check("TOTAL % from the summed counts (46/100 = 46.0%, red), not the average of 100% and 40% (70%)",
      _total_row(_w), [("#FDE0E0", ["TOTAL", "5", "100", "46", "54", "46.0%"])])
check("TOTAL % cell uses the band's strong text colour, like the technology rows",
      re.findall(r"color:(#[0-9A-Fa-f]+);font-weight:700'>46\.0%", _w), [M.AR._email_att_color(46.0)[1]])
check("… a single technology: TOTAL repeats its figures",
      _total_row(M._tech_table_html([("SQL", {"expected": 9, "submitted": 7, "not_submitted": 2,
                                              "rate": 77.8}, 1)])),
      [("#E8F5E9", ["TOTAL", "1", "9", "7", "2", "77.8%"])])
check("… nothing expected anywhere: TOTAL shows '—' on white",
      _total_row(M._tech_table_html([("T0", {"expected": 0, "submitted": 0, "not_submitted": 0,
                                             "rate": None}, 0)])),
      [("#ffffff", ["TOTAL", "0", "0", "0", "0", "—"])])

# =============================================================================
print("\n== 6. generate() end to end on a Monday 10:30 run ==")
import utils as _utils                                      # noqa: E402
_utils.get_sheets_service = lambda *a, **k: None
M.load_submissions = lambda service: DF
M.ASG.load_active_status_map = lambda service: ACTIVE
drive2 = FakeDrive()
M._drive_client = lambda: drive2
M.now_ist = lambda: datetime(2026, 10, 5, 10, 30)
MAILS.clear()
flags(auto=True)
out = io.StringIO()
import contextlib                                           # noqa: E402
with contextlib.redirect_stdout(out):
    res, errs = M.generate()
check("planned Daily 04-Oct (system date − 1) + Weekly 28-Sep..04-Oct",
      [(r["job"]["kind"], r["job"]["label"]) for r in res],
      [("Daily", "04-Oct-2026"), ("Weekly", "28-Sep-2026 to 04-Oct-2026")])
check("Daily evaluates only assignments due 04-Oct", res[0]["summary"]["assignments"], 1)
check("4 reports uploaded, 2 e-mails", (len(drive2.files_under()), len(MAILS)), (4, 2))
with contextlib.redirect_stdout(io.StringIO()):
    res2, _ = M.generate()
check("run again the same day: both periods rebuilt, reports overwritten (still 4 files), e-mailed again",
      ([bool(r.get("failed")) for r in res2], len(drive2.files_under()), len(MAILS)), ([False, False], 4, 4))
flags(manual=True, ms="2026-10-01", me="2026-10-05")
with contextlib.redirect_stdout(io.StringIO()):
    res3, errs3 = M.generate()
check("Manual end = today → nothing generated, configuration error reported",
      (res3, len(errs3)), ([], 1))
with contextlib.redirect_stdout(io.StringIO()):
    rc = M.main([])
check("… main() exits 1", rc, 1)
import exec_summary                                         # noqa: E402
summ = exec_summary.summarize("pyAssignmentSubmissionPerformanceReport", out.getvalue())
check("run summary KPIs", [k for k, _v in summ["kpis"]],
      ["Reports uploaded", "Eligible assignments", "Submitted", "Not submitted", "E-mailed"])
rra = open(os.path.join(ROOT, "scripts", "run_reports_action.py"), encoding="utf-8").read()
check("scheduled in the Morning batch after the Assignment Submissions refresh",
      bool(re.search(r'\("pyAssignmentSubmissionPerformanceReport", "[^"]+",\s*\["pyAssignmentSubmissions"\]\)', rra)),
      True)
src = open(M.__file__, encoding="utf-8").read()
check("nothing written to local disk", re.findall(r"open\([^)]*['\"][wa]b?['\"]|to_excel\(|MediaFileUpload", src), [])

# =============================================================================
print("\n== 7. Gmail Star (★) on the report e-mails ==")
import email as _email                                       # noqa: E402
import imaplib                                               # noqa: E402
import smtplib                                               # noqa: E402
import api_retry                                             # noqa: E402
import gmail_star as GS                                      # noqa: E402

api_retry._sleep = lambda s: None
GS.FIND_DELAYS = (0, 0)
GS.time.sleep = lambda s: None
BOXES, IMAP_LOGINS, STATE = {}, [], {"smtp_fail": False, "imap_fail": False}


class FakeSMTP:
    def __init__(self, *a, **k):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def login(self, user, pw):
        if STATE["smtp_fail"]:
            raise smtplib.SMTPAuthenticationError(535, b"5.7.8 Username and Password not accepted")

    def sendmail(self, sender, rcpts, raw):
        m = _email.message_from_string(raw)
        for a in {sender.lower(), *[r.lower() for r in rcpts]}:    # each mailbox has its own copy
            box = BOXES.setdefault(a, [])
            box.append({"uid": str(100 + len(box)).encode(), "msgid": m["Message-ID"],
                        "subject": str(_email.header.make_header(_email.header.decode_header(m["Subject"]))),
                        "flags": set()})


class FakeIMAP:
    def __init__(self, *a, **k):
        self.user = None

    def login(self, user, pw):
        IMAP_LOGINS.append(user.lower())
        if STATE["imap_fail"]:
            raise imaplib.IMAP4.error("[ALERT] IMAP access is disabled for your domain.")
        if (user, pw) != (ec.GMAIL_SENDER, ec.GMAIL_APP_PASS):
            raise imaplib.IMAP4.error("[AUTHENTICATIONFAILED] Invalid credentials")
        self.user = user.lower()
        return "OK", [b""]

    def list(self):
        return "OK", [b'(\\HasNoChildren \\All) "/" "[Gmail]/All Mail"']

    def select(self, *a, **k):
        return "OK", [b"1"]

    def noop(self):
        return "OK", [b""]

    def uid(self, cmd, *args):
        box = BOXES.get(self.user, [])
        if cmd == "SEARCH":
            q = " ".join(str(a) for a in args if a)
            m = re.search(r"rfc822msgid:([^\s\"]+)", q)
            hits = [x["uid"] for x in box if m and x["msgid"].strip("<>") == m.group(1)]
            return "OK", [b" ".join(hits)]
        if cmd == "STORE":
            for x in box:
                if x["uid"] == args[0]:
                    x["flags"].add(args[2].strip("()"))
            return "OK", [b""]
        raise AssertionError(cmd)

    def logout(self):
        return "BYE", [b""]


smtplib.SMTP_SSL = FakeSMTP
imaplib.IMAP4_SSL = FakeIMAP
M.CE.send = _REAL_SEND
M.STAR_EMAIL_IN_GMAIL = True
INFO, HR = "info@intellibiinnovationstechnologies.in", "intellibihropsb2ch@gmail.com"


def starred(addr):
    return [x["subject"] for x in BOXES.get(addr, []) if "\\Flagged" in x["flags"]]


check("STAR_EMAIL_IN_GMAIL is on by default", open(M.__file__, encoding="utf-8").read().count(
    "\nSTAR_EMAIL_IN_GMAIL = True"), 1)
out = io.StringIO()
with contextlib.redirect_stdout(out):
    res = M.run_jobs(jobs, DF, NOW, ACTIVE, None, upload=False, email=True)
subjects = [f"{j['kind']} Assignment Submission Report - {j['label']}" for j in jobs]
check("all four e-mails sent (Daily / Weekly / Monthly / Manual)", [r.get("emailed") for r in res], [True] * 4)
check("each one Starred in the info@ mailbox (STARRED = IMAP \\Flagged)", starred(INFO), subjects)
check("the other recipient receives each e-mail, not starred",
      ([x["subject"] for x in BOXES[HR]], starred(HR)), (subjects, []))
check("only the info@ mailbox is ever opened for starring", sorted(set(IMAP_LOGINS)), [INFO])
check("subject / recipients unchanged; no ★ text added",
      ("★" in "".join(x["subject"] for x in BOXES[INFO]), sorted(BOXES)), (False, sorted([INFO, HR])))
check("star logged per e-mail", out.getvalue().count(f"[Email] ★ starred in Gmail ({INFO})"), 4)

BOXES.clear(); IMAP_LOGINS.clear(); STATE["smtp_fail"] = True
with contextlib.redirect_stdout(io.StringIO()):
    r = M.run_jobs(jobs[:1], DF, NOW, ACTIVE, None, upload=False, email=True)
STATE["smtp_fail"] = False
check("send failed → no star attempt (no IMAP login), report still built",
      (r[0].get("emailed"), IMAP_LOGINS, bool(r[0].get("failed"))), (False, [], False))

GS._GIVE_UP.clear(); STATE["imap_fail"] = True
out = io.StringIO()
with contextlib.redirect_stdout(out):
    r = M.run_jobs(jobs[:1], DF, NOW, ACTIVE, None, upload=False, email=True)
STATE["imap_fail"] = False; GS._GIVE_UP.clear()
check("star failure → e-mail still sent, report not failed, one clear warning",
      (r[0].get("emailed"), bool(r[0].get("failed")), "[Email] ★ not starred" in out.getvalue(),
       starred(INFO)), (True, False, True, []))

BOXES.clear(); IMAP_LOGINS.clear(); M.STAR_EMAIL_IN_GMAIL = False
with contextlib.redirect_stdout(io.StringIO()):
    M.run_jobs(jobs[:1], DF, NOW, ACTIVE, None, upload=False, email=True)
M.STAR_EMAIL_IN_GMAIL = True
check("STAR_EMAIL_IN_GMAIL = False → sent, nothing starred, no IMAP", (len(BOXES[INFO]), IMAP_LOGINS), (1, []))

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
