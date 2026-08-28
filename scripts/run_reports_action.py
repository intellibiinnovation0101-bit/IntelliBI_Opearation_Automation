#!/usr/bin/env python3
"""
Layer 2 — Operations Reports & Actions.

Runs the reports, reminders, warnings and other operational actions in PARALLEL
where safe. Each script declares which Layer-1 data-collection jobs it depends
on; when run as part of the full pipeline, a script whose dependency FAILED is
SKIPPED (logged) so it never publishes stale/incomplete data — while independent
scripts still run.

    pyAttendaceFeedbackReport        <- SessionAttendance, StudentPayment
    pyAssignmentSubmissionsReport    <- AssignmentSubmissions
    pyAssignmentSubmissionEmailReminder <- AssignmentSubmissions
    pyBatchPlanner                   <- (independent — reads local data_inputs/)
    pyStudentProfileReport           <- SessionAttendance, AssignmentSubmissions, StudentPayment
    pyAdmissionFormalitiesReport     <- StudentPayment, ZohoSignatureStatus
    pyStudentAdditionalNote          <- (independent)

Run standalone:   python scripts/run_reports_action.py   (no gating — assumes L1 already ran)
"""
import os
import sys

sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap            # noqa: E402
import paths                 # noqa: E402
import logging_utils         # noqa: E402
import common_utils          # noqa: E402
import config_loader         # noqa: E402

LAYER = "Layer 2 — Operations Reports & Actions"

# script stem -> (display label, [Layer-1 dependency stems])
JOBS = [
    ("pyAttendaceFeedbackReport", "Attendance & Feedback Report",
        ["pySessionAttendanceStudentTeacherFeedbacks", "pyStudentPaymentClassesStudentEnrolled"]),
    ("pyAssignmentSubmissionsReport", "Assignment Submissions Report",
        ["pyAssignmentSubmissions"]),
    ("pyAssignmentSubmissionEmailReminder", "Assignment Submission Email Reminder",
        ["pyAssignmentSubmissions"]),
    ("pyBatchPlanner", "Batch Planner", []),
    ("pyStudentProfileReport", "Student Profile Report",
        ["pySessionAttendanceStudentTeacherFeedbacks", "pyAssignmentSubmissions",
         "pyStudentPaymentClassesStudentEnrolled"]),
    ("pyAdmissionFormalitiesReport", "Admission Formalities Report",
        ["pyStudentPaymentClassesStudentEnrolled", "pyZohoSignatureStatusRefresh"]),
    ("pyStudentAdditionalNote", "Student Additional Note", []),
]


def _timeout():
    t = config_loader.get("pipeline.script_timeout_seconds", "")
    try:
        return int(t) if str(t).strip() else None
    except (TypeError, ValueError):
        return None


def _max_parallel(n):
    """Cap concurrent scripts (config pipeline.max_parallel) to stay under the
    shared Google API per-minute quota."""
    try:
        mp = int(config_loader.get("pipeline.max_parallel", 2) or 2)
    except (TypeError, ValueError):
        mp = 2
    return max(1, min(n, mp))


def _skipped(stem, label, reason):
    return {"name": stem, "label": label, "script": "", "status": "SKIPPED",
            "returncode": None, "started": "-", "ended": "-", "duration": "-",
            "counts": [reason], "log_file": "", "error": reason}


def run(logger=None, l1_status=None) -> dict:
    """l1_status: {layer1_stem: bool_success}. None => standalone, no gating."""
    log = logger or logging_utils.get_logger("run_reports_action")
    logging_utils.section(log, f"{LAYER}: evaluating dependencies")
    timeout = _timeout()
    gate = config_loader.get("pipeline.stop_on_failure", True)

    runnable, skipped = [], []
    for stem, label, deps in JOBS:
        if l1_status is not None and gate:
            failed_deps = [d for d in deps if not l1_status.get(d, False)]
            if failed_deps:
                log.error("SKIP %s — upstream failed: %s", label, ", ".join(failed_deps))
                skipped.append(_skipped(stem, label,
                               f"skipped: upstream failed ({', '.join(failed_deps)})"))
                continue
        runnable.append((stem, label))

    log.info("%s: running %d, skipping %d", LAYER, len(runnable), len(skipped))

    def one(stem, label):
        return common_utils.run_script(paths.LAYER2_DIR / f"{stem}.py",
                                       label=label, timeout=timeout)

    scripts = []
    if runnable:
        from concurrent.futures import ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=_max_parallel(len(runnable))) as ex:
            futures = [ex.submit(one, s, l) for s, l in runnable]
            scripts = [f.result() for f in futures]
    scripts += skipped

    order = {stem: i for i, (stem, _, _) in enumerate(JOBS)}
    scripts.sort(key=lambda r: order.get(r["name"], 99))

    ok = all(r["status"] == "SUCCESS" for r in scripts)
    return {"layer": LAYER, "ok": ok, "scripts": scripts}


if __name__ == "__main__":
    result = run()
    sys.exit(0 if result["ok"] else 1)
