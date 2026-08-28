"""
================================================================================
  Execution-summary framework  (common/exec_summary.py)   — shared by both the
  Sales and Operations pipelines.
  ------------------------------------------------------------------------------
  Turns a script's raw stdout (already captured by the runner and written in
  full to the log file) into a small, BUSINESS-READABLE summary for the
  completion e-mail:

        { "title": "Sessions & Attendance",
          "kpis":  [("New Sessions", 12), ("Attendance rows", 148), ...],
          "note":  "Report generated & e-mailed"  # optional }

  Design notes
  ------------
  * NO business logic is changed anywhere. We only READ what each script already
    prints (its own summary lines) and map the meaningful numbers to labels.
  * The registry below is keyed by script stem and lists, per script, exactly
    which numbers are meaningful for THAT script (the request: "do not force the
    same metrics on every script").
  * Counts represent the CURRENT run (new / updated / skipped / sent / generated)
    — parsed from the script's own end-of-run summary lines, not the whole
    dataset — so the e-mail never reports "total existing" as "loaded".
  * Unknown scripts fall back to a clean status-only line.

  This module contains extractors for BOTH projects; a stem that isn't present
  in a given project simply never matches — harmless.
================================================================================
"""
from __future__ import annotations

import re

# Friendly section titles (fallback: prettified stem).
TITLES = {
    # ── Operations — Layer 1 ────────────────────────────────────────────────
    "pyStudentPaymentClassesStudentEnrolled": "Student Info & Payments",
    "pySessionAttendanceStudentTeacherFeedbacks": "Sessions & Attendance",
    "pyAssignmentSubmissions": "Assignment Submissions (load)",
    "pyZohoSignatureStatusRefresh": "Zoho Sign Status",
    # ── Operations — Layer 2 ────────────────────────────────────────────────
    "pyAttendaceFeedbackReport": "Attendance & Feedback Report",
    "pyAssignmentSubmissionsReport": "Assignment Submissions Report",
    "pyAssignmentSubmissionEmailReminder": "Assignment Reminders",
    "pyBatchPlanner": "Batch Planner",
    "pyStudentProfileReport": "Student Profile Reports",
    "pyAdmissionFormalitiesReport": "Admission Formalities Report",
    "pyStudentAdditionalNote": "Student Additional Note",
    # ── Sales ───────────────────────────────────────────────────────────────
    "pyInteraktUsers": "Interakt WhatsApp Users",
    "pyExotelInboxScrape": "Exotel Inbox Scrape",
    "pyExotelCallDetails": "Exotel Call Details",
    "pyConsolidateLeadsLoad": "Lead Consolidation",
    "pyConsolidatedLeadPerformanceReport": "Lead Performance Report",
    "pyLeadFollowUpAnalysisReport": "Follow-Up Analysis Report",
}

# Rule dicts:
#   k    : "last" | "first" | "sum" | "count" | "flag" | "note_if"
#   label: KPI label (not used for note_if)
#   re   : regex; group(1) is the integer for numeric kinds
#   zero : "keep" to show a 0 value (default: omit KPI when value is 0)
#   yes/no: for "flag" (text shown when the pattern is / isn't present)
#   note : for "note_if" (text set as the section note when present)
_R = {
    "pyStudentPaymentClassesStudentEnrolled": [
        {"k": "last", "label": "Students (full refresh)", "re": r"Students written\s*:\s*(\d+)"},
        {"k": "last", "label": "Payment records", "re": r"Payments written\s*:\s*(\d+)"},
        {"k": "last", "label": "Class-enrolment rows", "re": r"Combined rows written\s*:\s*(\d+)"},
        {"k": "last", "label": "Instructors", "re": r"\[Fetch Instructors\]\s*(\d+)\s+instructor\(s\) fetched"},
        {"k": "last", "label": "Students flagged removed", "re": r"(\d+)\s+student\(s\) no longer on portal"},
    ],
    "pySessionAttendanceStudentTeacherFeedbacks": [
        {"k": "last", "label": "New Sessions", "re": r"\[Write . Sessions\][^\n]*?Done . (\d+) rows appended", "zero": "keep"},
        {"k": "last", "label": "New Attendance rows", "re": r"\[Write . Attendance\][^\n]*?Done . (\d+) rows appended", "zero": "keep"},
        {"k": "last", "label": "New Student Feedback", "re": r"\[Write . Student_Feedback\][^\n]*?Done . (\d+) rows appended"},
        {"k": "last", "label": "New Teacher Feedback", "re": r"\[Write . Teacher_Feedback\][^\n]*?Done . (\d+) rows appended"},
        {"k": "last", "label": "Source sessions fetched", "re": r"Total raw sessions:\s*(\d+)"},
        {"k": "last", "label": "Suspended students found", "re": r"Done . (\d+) suspended student"},
    ],
    "pyAssignmentSubmissions": [
        {"k": "last", "label": "Submissions extracted", "re": r"\[Submissions\] Total:\s*(\d+) submission"},
        {"k": "last", "label": "New submissions", "re": r"Appended:\s*(\d+)", "zero": "keep"},
        {"k": "last", "label": "Updated submissions", "re": r"Updated:\s*(\d+)"},
        {"k": "last", "label": "Unchanged", "re": r"Skipped \(unchanged\):\s*(\d+)"},
        {"k": "last", "label": "New assignments detected", "re": r"\[Trigger\]\s*New:\s*(\d+)"},
    ],
    "pyZohoSignatureStatusRefresh": [
        {"k": "last", "label": "Records fetched", "re": r"Records Fetched\s*:\s*(\d+)"},
        {"k": "last", "label": "Inserted (new)", "re": r"Records Inserted \(new\)\s*:\s*(\d+)"},
        {"k": "last", "label": "Updated", "re": r"Records Updated \(changed\)\s*:\s*(\d+)"},
        {"k": "last", "label": "Unchanged / skipped", "re": r"Records Unchanged / Skipped\s*:\s*(\d+)"},
        {"k": "last", "label": "Failed", "re": r"Records Failed\s*:\s*(\d+)"},
    ],
    "pyAttendaceFeedbackReport": [
        {"k": "count", "label": "Reports generated", "re": r"\[Drive\] . Uploaded", "zero": "keep"},
        {"k": "flag", "label": "E-mailed", "re": r"\[Email\] . Report sent to", "yes": "Yes", "no": "No"},
        {"k": "last", "label": "Sessions covered", "re": r"Done\.\s*Sessions:\s*(\d+)"},
        {"k": "last", "label": "Students covered", "re": r"Students:\s*(\d+)\s*\|\s*Absent"},
        {"k": "last", "label": "Absent", "re": r"Absent:\s*(\d+)"},
    ],
    "pyAssignmentSubmissionsReport": [
        {"k": "note_if", "re": r"\[Trigger\] Report generation skipped", "note": "No new submissions — report skipped (nothing to report)"},
        {"k": "last", "label": "New submissions", "re": r"New submissions\s*:\s*(\d+)"},
        {"k": "last", "label": "Updated submissions", "re": r"Updated submissions\s*:\s*(\d+)"},
        {"k": "count", "label": "Reports generated", "re": r"\[Drive\] . Uploaded"},
        {"k": "flag", "label": "E-mailed", "re": r"\[Email\] . Report sent to", "yes": "Yes", "no": "No"},
    ],
    "pyAssignmentSubmissionEmailReminder": [
        {"k": "last", "label": "Students to notify", "re": r"Students to notify\s*:\s*(\d+)", "zero": "keep"},
        {"k": "last", "label": "Reminders sent", "re": r"Emails sent successfully\s*:\s*(\d+)", "zero": "keep"},
        {"k": "last", "label": "Reminders failed", "re": r"Emails failed\s*:\s*(\d+)"},
        {"k": "last", "label": "Pending assignments", "re": r"Total pending assignments\s*:\s*(\d+)"},
    ],
    "pyBatchPlanner": [
        {"k": "last", "label": "Students analysed", "re": r"Computed progress for (\d+) students"},
        {"k": "last", "label": "Batch-plan rows", "re": r"Planned (\d+) technology"},
        {"k": "flag", "label": "Report uploaded", "re": r"\[Drive\] Uploaded", "yes": "Yes", "no": "No"},
    ],
    "pyStudentProfileReport": [
        {"k": "last", "label": "Students processed", "re": r"(\d+) student\(s\) processed"},
        {"k": "last", "label": "Failed", "re": r"student\(s\) processed,\s*(\d+) failed"},
        {"k": "flag", "label": "Master report e-mailed", "re": r"\[Email\] . Master report sent", "yes": "Yes", "no": "No"},
    ],
    "pyStudentAdditionalNote": [
        {"k": "last", "label": "Responses processed", "re": r"Total combined responses across \d+ source\(s\):\s*(\d+)"},
        {"k": "last", "label": "Matched to students", "re": r"Matched:\s*(\d+)"},
        {"k": "last", "label": "No response", "re": r"No response:\s*(\d+)"},
        {"k": "note_if", "re": r"No updates to write", "note": "No changes to write this run"},
    ],
    "pyAdmissionFormalitiesReport": [
        {"k": "last", "label": "Students processed", "re": r"Total Students Processed\s*:\s*(\d+)"},
        {"k": "last", "label": "Inserted", "re": r"Records Inserted\s*:\s*(\d+)"},
        {"k": "last", "label": "Updated", "re": r"Records Updated\s*:\s*(\d+)"},
        {"k": "last", "label": "Unchanged", "re": r"Records Skipped \(No Changes\)\s*:\s*(\d+)"},
        {"k": "last", "label": "Failed", "re": r"Records Failed\s*:\s*(\d+)"},
    ],
    # ── Sales ───────────────────────────────────────────────────────────────
    "pyInteraktUsers": [
        {"k": "last", "label": "Users fetched & upserted", "re": r"Summary:\s*(\d+) users fetched"},
        {"k": "last", "label": "New", "re": r"Appended:\s*(\d+)"},
        {"k": "last", "label": "Updated", "re": r"Updated:\s*(\d+)"},
        {"k": "last", "label": "Unchanged", "re": r"Skipped \(unchanged\):\s*(\d+)"},
        {"k": "last", "label": "Repeat-enquiry updated", "re": r"Repeat-enquiry refresh:\s*(\d+) contact"},
    ],
    "pyExotelInboxScrape": [
        {"k": "last", "label": "Call rows scraped", "re": r"Wrote (\d+) unique call rows"},
        {"k": "last", "label": "Notes captured", "re": r"Wrote (\d+) user-notes"},
    ],
    "pyExotelCallDetails": [
        {"k": "last", "label": "Calls fetched", "re": r"Fetched (\d+) call records from Exotel"},
        {"k": "last", "label": "Notes applied", "re": r"Applied Inbox Notes to (\d+) of \d+ calls"},
        {"k": "last", "label": "New", "re": r"Appended:\s*(\d+)"},
        {"k": "last", "label": "Updated", "re": r"Updated:\s*(\d+)"},
        {"k": "last", "label": "Unchanged", "re": r"Skipped \(unchanged\):\s*(\d+)"},
    ],
    "pyConsolidateLeadsLoad": [
        {"k": "last", "label": "Unique master leads", "re": r"Unique Master Leads\s*:\s*(\d+)"},
        {"k": "last", "label": "New leads inserted", "re": r"New Leads Inserted\s*:\s*(\d+)"},
        {"k": "last", "label": "Duplicates merged", "re": r"Duplicate Leads Merged\s*:\s*(\d+)"},
        {"k": "last", "label": "Total records processed", "re": r"Total Records Processed\s*:\s*(\d+)"},
        {"k": "last", "label": "Rows removed", "re": r"Leads Removed\s*:\s*(\d+)"},
    ],
    "pyConsolidatedLeadPerformanceReport": [
        {"k": "count", "label": "Reports generated", "re": r"report \|.*active leads:", "zero": "keep"},
        {"k": "last", "label": "Master rows read", "re": r"Master rows loaded:\s*(\d+)"},
        {"k": "flag", "label": "E-mailed", "re": r"\[email\] sent to", "yes": "Yes", "no": "No"},
    ],
    "pyLeadFollowUpAnalysisReport": [
        {"k": "count", "label": "Reports generated", "re": r"report \|.*active leads:", "zero": "keep"},
        {"k": "last", "label": "Active leads scored", "re": r"Active leads scored:\s*(\d+)"},
        {"k": "flag", "label": "E-mailed", "re": r"\[email\] sent to", "yes": "Yes", "no": "No"},
    ],
}


def _pretty(stem: str) -> str:
    s = re.sub(r"^py", "", stem)
    s = re.sub(r"(?<!^)(?=[A-Z])", " ", s)
    return s.strip()


def title_for(stem: str) -> str:
    return TITLES.get(stem, _pretty(stem))


def summarize(stem: str, text: str) -> dict:
    """Return {title, kpis:[(label,value)], note} for a script from its stdout."""
    text = text or ""
    kpis: list = []
    note = None
    for rule in _R.get(stem, []):
        k = rule["k"]
        if k == "note_if":
            if re.search(rule["re"], text):
                note = rule["note"]
            continue
        if k == "flag":
            found = re.search(rule["re"], text) is not None
            kpis.append((rule["label"], rule["yes"] if found else rule["no"]))
            continue
        matches = list(re.finditer(rule["re"], text))
        if not matches and k != "count":
            continue
        if k == "last":
            val = int(matches[-1].group(1))
        elif k == "first":
            val = int(matches[0].group(1))
        elif k == "sum":
            val = sum(int(m.group(1)) for m in matches)
        elif k == "count":
            val = len(matches)
        else:
            continue
        if val == 0 and rule.get("zero") != "keep":
            continue
        kpis.append((rule["label"], val))
    return {"title": title_for(stem), "kpis": kpis, "note": note}


_QUOTA_RE = re.compile(r"\b429\b|Quota exceeded|rate limit|RESOURCE_EXHAUSTED", re.I)
_ERR_RE = re.compile(r"([A-Za-z_][\w.]*(?:Error|Exception)):\s*([^\n]+)")


def business_error(text: str) -> str:
    """A short, human-readable reason a script failed (full traceback stays in the log)."""
    text = text or ""
    if _QUOTA_RE.search(text):
        return "Google Sheets API rate limit (HTTP 429) reached."
    m = _ERR_RE.findall(text)
    if m:
        typ, msg = m[-1]
        msg = msg.strip()
        return f"{typ}: {msg[:180]}"
    return "Did not complete — see the log file for details."


def pipeline_status(results: list) -> str:
    """SUCCESS (all executed ok, nothing skipped) / FAILED (all executed failed) /
    PARTIAL (mixed, or something skipped)."""
    scripts = [r for layer in results for r in layer.get("scripts", [])]
    executed = [r for r in scripts if r["status"] != "SKIPPED"]
    skipped = [r for r in scripts if r["status"] == "SKIPPED"]
    failed = [r for r in executed if r["status"] != "SUCCESS"]
    if not failed and not skipped:
        return "SUCCESS"
    if executed and all(r["status"] != "SUCCESS" for r in executed):
        return "FAILED"
    return "PARTIAL"
