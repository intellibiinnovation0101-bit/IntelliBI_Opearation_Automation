#!/usr/bin/env python3
"""
================================================================================
  IntelliBI Operations Automation — master pipeline launcher (scripts/run_all.py)
  ------------------------------------------------------------------------------
  Runs the two layers in dependency order with per-script failure gating:

        Layer 1  ops_data_collection   (4 jobs, all in PARALLEL)
              |
              v   verify success per job
        Layer 2  ops_reports_action    (reports/actions in PARALLEL where safe;
                                        a report whose Layer-1 dependency FAILED
                                        is SKIPPED, independent ones still run)

  * Captures each script's success/failure and record counts.
  * Logs failures clearly and names the failed job.
  * Prevents dependent Layer-2 reports from running on stale/incomplete data
    (per-script dependency map in run_reports_action.py; honours
    pipeline.stop_on_failure in config.yaml).
  * Never silently ignores a failed source refresh — it is surfaced in the log
    and the completion e-mail.
  * On completion (success OR failure) e-mails a summary with each layer's log
    to config.yaml email.log_recipients
    (default info@intellibiinnovationstechnologies.in).

  Run:   python scripts/run_all.py      (or run_all.bat)
  Exit code 0 = every executed job succeeded and nothing was skipped, else 1.
================================================================================
"""
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap            # noqa: E402
import paths                 # noqa: E402
import logging_utils         # noqa: E402
import common_utils          # noqa: E402
import config_loader         # noqa: E402
import exec_summary          # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import run_data_collection   # noqa: E402
import run_reports_action    # noqa: E402


def main() -> int:
    log = logging_utils.get_logger("run_all")
    started = datetime.now()
    logging_utils.section(log, "IntelliBI Operations Automation — FULL PIPELINE START")

    # ── Layer 1 (parallel) ───────────────────────────────────────────────────
    l1 = run_data_collection.run(logger=logging_utils.get_logger("run_data_collection"))
    l1_status = {r["name"]: (r["status"] == "SUCCESS") for r in l1["scripts"]}
    if not l1["ok"]:
        failed = [n for n, ok in l1_status.items() if not ok]
        log.error("Layer 1 completed with FAILURES: %s. Dependent Layer-2 reports "
                  "will be skipped; independent ones continue.", ", ".join(failed))

    # ── Layer 2 (gated on Layer 1 per-script dependencies) ───────────────────
    l2 = run_reports_action.run(
        logger=logging_utils.get_logger("run_reports_action"),
        l1_status=l1_status)

    results = [l1, l2]
    ended = datetime.now()

    executed = [r for layer in results for r in layer["scripts"]
                if r["status"] != "SKIPPED"]
    skipped = [r for layer in results for r in layer["scripts"]
               if r["status"] == "SKIPPED"]
    overall_ok = all(r["status"] == "SUCCESS" for r in executed) and not skipped

    logging_utils.section(
        log, f"FULL PIPELINE {'SUCCESS' if overall_ok else 'COMPLETED WITH ISSUES'} in "
        f"{common_utils.fmt_duration((ended - started).total_seconds())}"
        + (f"  |  skipped: {len(skipped)}" if skipped else ""))

    # ── completion e-mail ────────────────────────────────────────────────────
    if config_loader.get("email.send_log_summary", True):
        recipients = config_loader.get(
            "email.log_recipients", ["info@intellibiinnovationstechnologies.in"])
        if isinstance(recipients, str):
            recipients = [recipients]
        attach = []
        if config_loader.get("email.attach_layer_logs", True):
            for layer in results:
                for r in layer["scripts"]:
                    lf = r.get("log_file")
                    if lf and os.path.exists(lf) and lf not in attach:
                        attach.append(lf)
        subject = (f"IntelliBI Operations Automation — {exec_summary.pipeline_status(results)}"
                   f" — {ended.strftime('%d-%b-%Y %H:%M')}")
        html = common_utils.build_summary_html(
            results, overall_ok,
            started.strftime("%Y-%m-%d %H:%M:%S"),
            ended.strftime("%Y-%m-%d %H:%M:%S"))
        common_utils.send_summary_email(subject, html, recipients, logger=log,
                                        attach_logs=attach)
    else:
        log.info("email.send_log_summary is false — no summary e-mail sent.")

    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
