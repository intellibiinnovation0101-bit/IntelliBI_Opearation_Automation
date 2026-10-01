"""
Verification of pyCoordinatorTaskPerformanceReport.py (run from the project root:
  python ops_validation\\verify_coordinator_task_performance.py).

Builds REAL Coordinator task tabs with pyBatchCoordinatorDailyAttendanceReport's own
builders (Admission Formalities, Assignment Follow-Ups, Wise Validation), fills the
follow-up columns the way the Coordinator would, across two report days and several
report versions, and checks the performance ledger against hand-computed answers:
  * one task per day however many versions list it (no double counting)
  * a task actioned in an early version and dropped later is still counted;
    an un-actioned task that dropped off is not
  * Done? = Yes in ANY version completes the task; earliest valid stamp is used
  * #REF! stamp -> completed, excluded from timing
  * on time / late / missed / open, Timely %, Days on List, median time
No Google access needed.
"""
import os
import sys
import types
from datetime import date, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.modules["_bootstrap"] = types.ModuleType("_bootstrap")
paths = types.ModuleType("paths")
paths.CREDENTIALS_DIR = paths.CONFIG_DIR = paths.LOGS_DIR = paths.CACHE_DIR = HERE
sys.modules["paths"] = paths
ec = types.ModuleType("email_config"); ec.GMAIL_SENDER = "x@x"; ec.GMAIL_APP_PASS = "x"
sys.modules["email_config"] = ec
for p in ("common", "ops_reports_action", "co-ordinator reports"):
    sys.path.insert(0, os.path.join(ROOT, p))

import openpyxl                                          # noqa: E402
import pyBatchCoordinatorDailyAttendanceReport as BC     # noqa: E402  (the live task report)
import pyCoordinatorTaskPerformanceReport as P           # noqa: E402  (the new report)

FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


def adm(name, email):
    return {"Student Name": name, "Email ID": email, "Phone Number": "+919000000001",
            "Batch Name": "DAAI1026", "Joined On": "2026-10-01", "Request Form Name": "",
            "Recipient Status": "Form Not Sent", "Request Status": "", "Sent Date": "",
            "Signed Date": "", "Expiry Date": ""}


def asg(name, lvl="1st"):
    return {"student_name": name, "student_email": name.lower().replace(" ", ".") + "@x.com",
            "student_phone": "+919000000002", "deadline_str": "08-Oct-2026", "reminder_level": lvl,
            "reminder_label": "1st Reminder", "assigned_date_str": "04-Oct-2026",
            "class_subject": "01-Oct-2026 To Current Date", "maximum_marks": ""}


def report(day, adm_rows, asg_map, wise_students=None):
    """A Coordinator daily workbook built by the real builders."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    BC.build_admission_formalities(wb.create_sheet("Learner Admission Formalities"), adm_rows, day)
    by_tech = {"Power BI": [(title, f"id-{title}", [asg(n) for n in names])
                            for title, names in asg_map.items()]}
    BC.build_assignment_followups(wb.create_sheet("Learner Assignment Follow-Ups"), by_tech, day)
    if wise_students is not None:
        data = {"student": [{"Student Name": n, "Batch Name": "DAAI1026", "Joined On": "",
                             "Student Name Status": "Valid", "Email ID Status": "Valid",
                             "Phone Number Status": "Valid", "Tag Name Status": "Missing",
                             "Private Note Status": "Valid", "Profile Picture Status": "Valid"}
                            for n in wise_students], "course": [], "instructor": []}
        BC.build_wise_validation(wb.create_sheet("Wise & Interview Feedback Validation"), data, day)
    return wb


def fill(wb, tab, who, done=None, stamp=None, action=None, section=None):
    """Type a follow-up into the row whose cells contain `who` (below the banner
    containing `section`, when given) — what the Coordinator does in the sheet."""
    ws = wb[tab]
    hdr, in_section = None, section is None
    for r in range(1, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        txt = [str(v or "") for v in vals]
        if section and any(section in t for t in txt[:1]):
            in_section = True
        if any(t.strip() == "Why Flagged" for t in txt):
            hdr = {t.replace("✎", "").strip(): i + 1 for i, t in enumerate(txt) if t}
            continue
        if hdr and in_section and any(t == who for t in txt):
            if done is not None:
                ws.cell(row=r, column=hdr["Follow-Up Done?"]).value = done
            if stamp is not None:
                ws.cell(row=r, column=hdr["Follow-Up DateTime"]).value = stamp
            if action is not None:
                ws.cell(row=r, column=hdr["Action Taken"]).value = action
            return
    raise AssertionError(f"row {who!r} not found in {tab}")


def clear_formulas(wb):
    """Google stores the stamp's VALUE; an untouched row exports blank."""
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=IF("):
                    c.value = None


D1, D2 = date(2026, 10, 5), date(2026, 10, 6)
ADM, ASG = "Learner Admission Formalities", "Learner Assignment Follow-Ups"

# ── Day 1, Version 1 (09:00) ────────────────────────────────────────────────
d1v1 = report(D1, [adm("Asha Rao", "asha@x.com"), adm("Bala K", "bala@x.com"), adm("Chitra M", "chitra@x.com")],
              {"Project X": ["Lina P", "Mohan S"], "Project Y": ["Lina P"]})
clear_formulas(d1v1)
fill(d1v1, ADM, "Asha Rao", "Yes", datetime(2026, 10, 5, 10, 0), "Call")
fill(d1v1, ADM, "Bala K", "No", datetime(2026, 10, 5, 10, 30), "WhatsApp")
fill(d1v1, ASG, "Lina P", "Yes", "#REF!", "WhatsApp Group", section="Project X")
# ── Day 1, Version 2 (12:00): Asha dropped off, Dev added, Mohan (X) dropped off ─
d1v2 = report(D1, [adm("Bala K", "bala@x.com"), adm("Chitra M", "chitra@x.com"), adm("Dev N", "dev@x.com")],
              {"Project X": ["Lina P"], "Project Y": ["Lina P"]})
clear_formulas(d1v2)
fill(d1v2, ADM, "Bala K", "Yes", datetime(2026, 10, 5, 13, 0), "Call")
fill(d1v2, ADM, "Chitra M", "Yes", datetime(2026, 10, 6, 10, 0), "Call")      # done next day -> late
# ── Day 2, Version 1 (09:00) ────────────────────────────────────────────────
d2v1 = report(D2, [adm("Chitra M", "chitra@x.com"), adm("Dev N", "dev@x.com")],
              {"Project Y": ["Lina P"]}, wise_students=["Wasim Q"])
clear_formulas(d2v1)
fill(d2v1, ADM, "Dev N", "Yes", datetime(2026, 10, 6, 11, 0), "Form Send")

BOOKS = {"d1v1": d1v1, "d1v2": d1v2, "d2v1": d2v1}
versions = [
    {"id": "d1v1", "name": "d1v1", "day": D1, "version": 1, "created": datetime(2026, 10, 5, 9, 0), "modified_raw": ""},
    {"id": "d1v2", "name": "d1v2", "day": D1, "version": 2, "created": datetime(2026, 10, 5, 12, 0), "modified_raw": ""},
    {"id": "d2v1", "name": "d2v1", "day": D2, "version": 1, "created": datetime(2026, 10, 6, 9, 0), "modified_raw": ""},
]
now = datetime(2026, 10, 6, 18, 0)
ledger = P.build_ledger(versions, lambda v: P.parse_report_workbook(BOOKS[v["id"]]), now)
T = {(t["day"], t["group"], t["label"].split("  ·  ")[0], t.get("context", "")): t for t in ledger["tasks"]}


def status(day, group, who, ctx_has=""):
    m = [t for (d, g, w, c), t in T.items() if d == day and g == group and w == who and ctx_has in c]
    return m[0]["status"] if len(m) == 1 else f"{len(m)} matches"


check("day-1 tasks (no double counting across 2 versions; dropped un-actioned task excluded)",
      sum(1 for t in ledger["tasks"] if t["day"] == D1), 6)
check("Asha: actioned in V1, dropped in V2 -> still counted, on time", status(D1, "admission", "Asha Rao"), P.ST_ON_TIME)
check("Bala: No in V1, Yes in V2 -> completed on time", status(D1, "admission", "Bala K"), P.ST_ON_TIME)
check("Chitra day 1: Yes stamped next day -> late", status(D1, "admission", "Chitra M"), P.ST_LATE)
check("Dev day 1: never done on day 1 -> missed", status(D1, "admission", "Dev N"), P.ST_MISSED)
check("Lina / Project X: Yes with #REF! stamp -> completed, time not recorded",
      status(D1, "assignment", "Lina P", "Project X"), P.ST_UNKNOWN)
check("Lina / Project Y day 1 -> missed", status(D1, "assignment", "Lina P", "Project Y"), P.ST_MISSED)
check("Mohan / Project X (listed in V1 only, never actioned) -> not a task",
      status(D1, "assignment", "Mohan S", "Project X"), "0 matches")
check("Dev day 2 -> on time", status(D2, "admission", "Dev N"), P.ST_ON_TIME)
check("Chitra day 2 -> open (due today)", status(D2, "admission", "Chitra M"), P.ST_OPEN)
check("Wise record day 2 -> open", status(D2, "wise", "Wasim Q"), P.ST_OPEN)

lina_y2 = [t for t in ledger["tasks"] if t["day"] == D2 and t["group"] == "assignment"][0]
check("Days on List: Lina / Project Y pending on both days", lina_y2["days_on_list"], 2)
chitra2 = [t for t in ledger["tasks"] if t["day"] == D2 and t["label"].startswith("Chitra")][0]
check("Days on List: Chitra day 2 (day-1 task was completed) ", chitra2["days_on_list"], 1)

s = P.summarise(ledger["tasks"])
check("total tasks", s["tasks"], 10)
check("completed", s["completed"], 5)
check("pending (open + missed)", (s["pending"], s["open"], s["missed"]), (5, 3, 2))
check("completion %", round(s["completion_pct"], 1), 50.0)
check("timely % = on time / tasks with measurable timing = 3 / 9", round(s["timely_pct"], 1), 33.3)
check("median time to complete (1h, 4h, 2h)", s["median_ttc"], 2.0)
check("pending 2+ report days", s["carried"], 1)
check("unique items", s["unique_items"], 7)

cov = {c["day"]: c for c in ledger["coverage"]}
check("coverage day 1: versions / dropped / multi-version / no-time / Yes-No conflicts",
      (cov[D1]["versions"], cov[D1]["dropped"], cov[D1]["multi_version"], cov[D1]["yes_no_time"],
       cov[D1]["conflicts"]), (2, 1, 1, 1, 1))

# the workbook builds end-to-end for both a day and a multi-day period
wb, tasks, summ = P.build_workbook(ledger, D1, D2, "Test", False, now)
check("workbook tabs", wb.sheetnames,
      ["Dashboard", "Progress Trend", "Pending & Overdue", "Task Register", "Data Coverage & Rules"])
wb, tasks, summ = P.build_workbook(ledger, D2, D2, "Test", True, now)
check("daily workbook task count", len(tasks), 4)

# a report version WITHOUT follow-up columns is never scored
old = report(date(2026, 10, 4), [adm("Old L", "old@x.com")], {})
for ws in old.worksheets:            # drop the follow-up headers, as in pre-column reports
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("✎"):
                c.value = None
led2 = P.build_ledger([{"id": "o", "name": "o", "day": date(2026, 10, 4), "version": 1,
                        "created": datetime(2026, 10, 4, 9, 0), "modified_raw": ""}],
                      lambda v: P.parse_report_workbook(old), now)
check("pre-follow-up report: 0 scored tasks, listed as not trackable",
      (len(led2["tasks"]), led2["coverage"][0]["untracked_tasks"]), (0, 1))

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
