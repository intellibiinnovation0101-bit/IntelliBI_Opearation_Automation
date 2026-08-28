#!/usr/bin/env python3
"""
Layer 1 — Operations Data Collection / Refresh.

Refreshes/prepares the source datasets the Operations reports depend on. All
jobs are independent and run in PARALLEL:

    pyStudentPaymentClassesStudentEnrolled   (IntelliBIStudentInfo)
    pySessionAttendanceStudentTeacherFeedbacks (IntellBIAttendance)
    pyAssignmentSubmissions                  (IntelliBIAssessmentSubmission + trigger)
    pyZohoSignatureStatusRefresh             (Zoho Sign - Signers)

Run standalone:   python scripts/run_data_collection.py
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

LAYER = "Layer 1 — Operations Data Collection"

# script stem -> display label
JOBS = [
    ("pyStudentPaymentClassesStudentEnrolled", "Student Info & Payments"),
    ("pySessionAttendanceStudentTeacherFeedbacks", "Session Attendance & Feedback"),
    ("pyAssignmentSubmissions", "Assignment Submissions"),
    ("pyZohoSignatureStatusRefresh", "Zoho Signature Status"),
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


def run(logger=None) -> dict:
    log = logger or logging_utils.get_logger("run_data_collection")
    logging_utils.section(log, f"{LAYER}: launching {len(JOBS)} jobs in parallel")
    timeout = _timeout()

    def one(stem, label):
        return common_utils.run_script(paths.LAYER1_DIR / f"{stem}.py",
                                       label=label, timeout=timeout)

    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=_max_parallel(len(JOBS))) as ex:
        futures = [ex.submit(one, stem, label) for stem, label in JOBS]
        scripts = [f.result() for f in futures]

    # keep declared order for a stable summary
    order = {stem: i for i, (stem, _) in enumerate(JOBS)}
    scripts.sort(key=lambda r: order.get(r["name"], 99))

    ok = all(r["status"] == "SUCCESS" for r in scripts)
    failed = [r["name"] for r in scripts if r["status"] != "SUCCESS"]
    if failed:
        log.error("%s: FAILED jobs -> %s", LAYER, ", ".join(failed))
    else:
        log.info("%s: all jobs SUCCESS", LAYER)
    return {"layer": LAYER, "ok": ok, "scripts": scripts}


if __name__ == "__main__":
    result = run()
    sys.exit(0 if result["ok"] else 1)
