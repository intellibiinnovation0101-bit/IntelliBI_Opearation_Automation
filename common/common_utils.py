"""
================================================================================
  IntelliBI Sales Automation — shared runner helpers  (common/common_utils.py)
  ------------------------------------------------------------------------------
  Utilities used by the layer runners and the run_all orchestrator:

    run_script(...)          launch one pipeline script as a subprocess, stream
                             its output to that script's log file, time it,
                             capture success/failure and record counts.
    fmt_duration(seconds)    "1m 12s" style formatting.
    build_summary_html(...)  assemble the completion e-mail body from per-layer
                             results (status, timings, record counts, errors).
    send_summary_email(...)  send that summary via the already-configured Gmail
                             SMTP account (credentials/email_config.py).
================================================================================
"""
from __future__ import annotations

import os
import random
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import paths
import logging_utils
import config_loader
import exec_summary

# Lines worth surfacing as "record counts" in the summary e-mail.
_COUNT_KEYWORDS = re.compile(
    r"\b(lead|leads|record|records|row|rows|user|users|call|calls|"
    r"upsert|insert|inserted|updated|update|skip|skipped|removed|duplicate|"
    r"duplicates|merged|scored|processed|wrote|written|uploaded|fetched|"
    r"enriched|masked|emailed|sent|total|unique)\b",
    re.IGNORECASE,
)
_HAS_NUMBER = re.compile(r"\d")


def fmt_duration(seconds: float) -> str:
    seconds = int(round(seconds))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}h {m}m {s}s"
    if m:
        return f"{m}m {s}s"
    return f"{s}s"


def extract_counts(text: str, limit: int = 20) -> list[str]:
    """Heuristically pull 'record count' lines out of a script's output."""
    out = []
    seen = set()
    for raw in text.splitlines():
        line = raw.strip()
        if not line or not _HAS_NUMBER.search(line):
            continue
        if not _COUNT_KEYWORDS.search(line):
            continue
        # strip a leading log prefix "2026-.. | INFO | name | "
        line = re.sub(r"^\d{4}-\d\d-\d\d[ T][\d:]+\s*\|.*?\|.*?\|\s*", "", line)
        line = line.strip(" -=|")
        if len(line) > 200 or line in seen:
            continue
        seen.add(line)
        out.append(line)
        if len(out) >= limit:
            break
    return out


# Signatures of a Google API rate-limit / quota failure (worth retrying).
_QUOTA_RE = re.compile(
    r"\b429\b|Quota exceeded|rateLimitExceeded|userRateLimitExceeded|"
    r"RESOURCE_EXHAUSTED|Rate Limit Exceeded|quotaExceeded", re.IGNORECASE)


def _run_script_once(script_path, log, env, timeout) -> dict:
    started_dt = datetime.now()
    t0 = time.time()
    captured: list[str] = []
    error = None
    rc = None
    try:
        proc = subprocess.Popen(
            [sys.executable, "-u", str(script_path)],
            cwd=str(paths.PROJECT_ROOT),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )
        for line in proc.stdout:  # stream -> this layer's log + capture
            line = line.rstrip("\n")
            captured.append(line)
            log.info(line)
        rc = proc.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        error = f"timed out after {timeout}s"
        rc = -1
    except Exception as e:
        error = f"{type(e).__name__}: {e}"
        rc = -1
    text = "\n".join(captured)
    return {
        "status": "SUCCESS" if rc == 0 else "FAILED",
        "returncode": rc,
        "started_dt": started_dt,
        "ended_dt": datetime.now(),
        "duration_s": time.time() - t0,
        "counts": extract_counts(text),
        "error": error,
        "is_quota": rc != 0 and bool(_QUOTA_RE.search(text)),
        "text": text,
    }


def run_script(script_path, logger=None, extra_env=None, label=None,
               timeout=None, retries=None, retry_wait=None) -> dict:
    """Run one pipeline script as a subprocess with the bootstrapped env.

    Returns a result dict: name, label, status ('SUCCESS'/'FAILED'), returncode,
    started, ended, duration_s, counts (list[str]), log_file, error, attempts.
    Never raises for a non-zero exit — it records the failure and returns.

    If the script fails specifically with a Google API quota / HTTP-429 error, it
    is retried up to `retries` times after `retry_wait` seconds (+ jitter). This
    absorbs the shared-service-account rate limit hit when run_all launches
    several reports in parallel (the per-minute quota resets, so a short wait
    lets the retry succeed). Non-quota failures are NOT retried.
    """
    script_path = Path(script_path)
    name = script_path.stem
    label = label or name
    log = logger or logging_utils.get_logger(name)
    log_file = logging_utils.log_file_for(name)

    if retries is None:
        try:
            retries = int(config_loader.get("pipeline.quota_retries", 2) or 0)
        except (TypeError, ValueError):
            retries = 2
    if retry_wait is None:
        try:
            retry_wait = int(config_loader.get("pipeline.quota_retry_wait_seconds", 70) or 70)
        except (TypeError, ValueError):
            retry_wait = 70

    env = dict(os.environ)
    if extra_env:
        env.update({k: str(v) for k, v in extra_env.items()})

    attempt = 0
    while True:
        tag = label + (f"  (retry {attempt}/{retries})" if attempt else "")
        logging_utils.section(log, f"START  {tag}  ({script_path.name})")
        res = _run_script_once(script_path, log, env, timeout)
        logging_utils.section(
            log, f"END    {tag}  [{res['status']}]  rc={res['returncode']}  "
                 f"in {fmt_duration(res['duration_s'])}")
        if res["status"] == "SUCCESS" or attempt >= retries or not res["is_quota"]:
            break
        attempt += 1
        wait = retry_wait + random.randint(0, 15)   # jitter de-syncs parallel retries
        log.warning("[%s] Google API quota / 429 hit — waiting %ds, then retry "
                    "%d/%d.", label, wait, attempt, retries)
        time.sleep(wait)

    return {
        "name": name,
        "label": label,
        "script": str(script_path),
        "status": res["status"],
        "returncode": res["returncode"],
        "started": res["started_dt"].strftime("%Y-%m-%d %H:%M:%S"),
        "ended": res["ended_dt"].strftime("%Y-%m-%d %H:%M:%S"),
        "duration_s": res["duration_s"],
        "duration": fmt_duration(res["duration_s"]),
        "counts": res["counts"],
        "log_file": str(log_file),
        "error": res["error"],
        "attempts": attempt + 1,
        # business-readable summary for the completion e-mail (logs keep full detail)
        "summary": exec_summary.summarize(name, res.get("text", "")),
        "error_brief": (exec_summary.business_error(res.get("text", ""))
                        if res["status"] != "SUCCESS" else None),
    }


# ── completion e-mail ────────────────────────────────────────────────────────
def _esc(x):
    return (str(x).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _pipeline_name() -> str:
    # e.g. "IntelliBI_Operations_Automation" -> "Operations Automation"
    n = paths.PROJECT_ROOT.name.replace("IntelliBI_", "").replace("_", " ")
    return n or "Pipeline"


_STATUS_COLOR = {"SUCCESS": "#1a7f37", "PARTIAL": "#b45309", "FAILED": "#b42318"}


def build_summary_html(layer_results, overall_ok, started, ended) -> str:
    """A concise, business-readable completion e-mail: pipeline status, per-script
    KPIs (what came in / changed / was generated or sent), and an Action-Required
    block only when something failed or was skipped. Technical detail stays in the
    log files (attached)."""
    status = exec_summary.pipeline_status(layer_results)
    color = _STATUS_COLOR.get(status, "#b42318")
    pipeline = _pipeline_name()

    scripts = [r for layer in layer_results for r in layer.get("scripts", [])]
    executed = [r for r in scripts if r["status"] != "SKIPPED"]
    ok = [r for r in executed if r["status"] == "SUCCESS"]
    failed = [r for r in executed if r["status"] != "SUCCESS"]
    skipped = [r for r in scripts if r["status"] == "SKIPPED"]

    dur = "—"
    try:
        d = (datetime.strptime(ended, "%Y-%m-%d %H:%M:%S")
             - datetime.strptime(started, "%Y-%m-%d %H:%M:%S")).total_seconds()
        dur = fmt_duration(d)
    except Exception:
        pass

    px = "padding:5px 12px;"
    out = [f'<div style="font-family:Arial,Helvetica,sans-serif;color:#111;max-width:680px">']

    # header band
    out.append(
        f'<div style="background:{color};color:#fff;padding:14px 18px;border-radius:6px 6px 0 0">'
        f'<div style="font-size:18px;font-weight:700">{_esc(pipeline)} &mdash; {status}</div></div>')
    out.append(
        '<table style="border-collapse:collapse;width:100%;font-size:13px;'
        'background:#f8fafc;border:1px solid #e5e7eb;border-top:none">'
        f'<tr><td style="{px}color:#555">Pipeline</td><td style="{px}font-weight:600">{_esc(pipeline)}</td>'
        f'<td style="{px}color:#555">Status</td><td style="{px}font-weight:700;color:{color}">{status}</td></tr>'
        f'<tr><td style="{px}color:#555">Started</td><td style="{px}">{_esc(started)}</td>'
        f'<td style="{px}color:#555">Completed</td><td style="{px}">{_esc(ended)}</td></tr>'
        f'<tr><td style="{px}color:#555">Duration</td><td style="{px}" colspan="3">{_esc(dur)}</td></tr>'
        f'<tr><td style="{px}color:#555">Result</td><td style="{px}" colspan="3">'
        f'{len(ok)} succeeded &bull; {len(failed)} failed &bull; {len(skipped)} skipped</td></tr>'
        '</table>')

    # per-script execution summary
    out.append('<h3 style="margin:18px 0 8px;font-size:15px">Execution Summary</h3>')
    pill = ("display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;"
            "font-weight:700;color:#fff;")
    for layer in layer_results:
        for r in layer.get("scripts", []):
            s = r.get("summary") or {"title": r.get("label", r["name"]), "kpis": [], "note": None}
            st = r["status"]
            pc = {"SUCCESS": "#1a7f37", "SKIPPED": "#6b7280"}.get(st, "#b42318")
            head = (f'<div style="margin:10px 0 2px">'
                    f'<span style="font-weight:700;font-size:14px">{_esc(s["title"])}</span> '
                    f'<span style="{pill}background:{pc}">{st}</span></div>')
            body = ""
            kpis = s.get("kpis") or []
            if kpis:
                cells = "".join(
                    f'<tr><td style="padding:2px 14px 2px 0;color:#444">{_esc(lbl)}</td>'
                    f'<td style="padding:2px 0;font-weight:700;text-align:right">{_esc(val)}</td></tr>'
                    for lbl, val in kpis)
                body += f'<table style="border-collapse:collapse;font-size:13px;margin:2px 0 0 2px">{cells}</table>'
            if s.get("note"):
                body += f'<div style="font-size:12px;color:#555;font-style:italic;margin-top:2px">{_esc(s["note"])}</div>'
            if st == "SKIPPED":
                body += f'<div style="font-size:12px;color:#6b7280">Skipped &mdash; {_esc(r.get("error") or "dependency not met")}</div>'
            elif st != "SUCCESS":
                body += f'<div style="font-size:12px;color:#b42318">{_esc(r.get("error_brief") or "Failed")}</div>'
            if not kpis and st == "SUCCESS" and not s.get("note"):
                body += '<div style="font-size:12px;color:#555">Completed.</div>'
            out.append(head + body)

    # action required
    if failed or skipped:
        out.append('<h3 style="margin:20px 0 6px;font-size:15px;color:#b42318">Action Required</h3>')
        th = "padding:6px 10px;border:1px solid #e5e7eb;text-align:left;background:#fef2f2"
        td = "padding:6px 10px;border:1px solid #e5e7eb;vertical-align:top"
        rows = []
        for r in failed:
            retry = f'retried {r.get("attempts",1)-1}x' if r.get("attempts", 1) > 1 else "no retry"
            rows.append(
                f'<tr><td style="{td}">{_esc((r.get("summary") or {}).get("title", r["name"]))}</td>'
                f'<td style="{td};color:#b42318;font-weight:600">FAILED</td>'
                f'<td style="{td}">{_esc(r.get("error_brief") or "See log")}</td>'
                f'<td style="{td}">{_esc(retry)}</td></tr>')
        for r in skipped:
            rows.append(
                f'<tr><td style="{td}">{_esc((r.get("summary") or {}).get("title", r["name"]))}</td>'
                f'<td style="{td};color:#6b7280;font-weight:600">SKIPPED</td>'
                f'<td style="{td}">{_esc(r.get("error") or "an upstream dependency failed — not run on stale data")}</td>'
                f'<td style="{td}">&mdash;</td></tr>')
        out.append(
            f'<table style="border-collapse:collapse;width:100%;font-size:12px">'
            f'<tr><th style="{th}">Process</th><th style="{th}">Status</th>'
            f'<th style="{th}">Details</th><th style="{th}">Retry</th></tr>'
            + "".join(rows) + '</table>')

    out.append(
        '<p style="margin:18px 0 0;color:#777;font-size:11px">Full technical logs '
        '(API calls, warnings, tracebacks, retries) are in the project '
        '<code>logs/</code> folder and attached to this e-mail.</p></div>')
    return "".join(out)


def send_summary_email(subject: str, html_body: str, recipients, logger=None,
                       attach_logs=None) -> bool:
    """Send the completion summary via the configured Gmail SMTP account."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    log = logger or logging_utils.get_logger("run_all")
    if not recipients:
        log.info("[email] no recipients configured; skipping summary email")
        return False
    try:
        import email_config as ecfg  # credentials/email_config.py (on sys.path)
        sender = ecfg.GMAIL_SENDER
        app_pass = ecfg.GMAIL_APP_PASS
    except Exception as e:
        log.error("[email] cannot load email_config: %s", e)
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText("This report is best viewed as HTML.", "plain"))
    alt.attach(MIMEText(html_body, "html"))
    msg.attach(alt)

    for f in (attach_logs or []):
        try:
            with open(f, "rb") as fh:
                part = MIMEApplication(fh.read(), Name=os.path.basename(f))
            part["Content-Disposition"] = f'attachment; filename="{os.path.basename(f)}"'
            msg.attach(part)
        except Exception as e:
            log.warning("[email] could not attach %s: %s", f, e)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as smtp:
            smtp.login(sender, app_pass)
            smtp.sendmail(sender, list(recipients), msg.as_string())
        log.info("[email] summary sent to %s", ", ".join(recipients))
        return True
    except Exception as e:
        log.error("[email] send failed: %s", e)
        return False
