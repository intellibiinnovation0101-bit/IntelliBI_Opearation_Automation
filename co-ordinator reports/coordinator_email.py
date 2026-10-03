"""
================================================================================
  IntelliBI Operations Automation
  COORDINATOR REPORT E-MAILS  (co-ordinator reports / coordinator_email.py)
  ------------------------------------------------------------------------------
  E-mail for the two Coordinator reports:
      pyCoordinatorTaskListReport.py   (task lists)
      pyCoordinatorTaskPerformanceReport.py        (completion & timeliness)

  Modelled on the IntelliBI report e-mail already in production
  (Sales: pyConsolidatedLeadPerformanceReport.send_email / build_email_body):
    * one e-mail per report, subject "<Type> <Report name> - <period>"
    * Gmail SMTP with the sender + app password from credentials/email_config.py
      (GMAIL_SENDER / GMAIL_APP_PASS — the Operations account), recipients
      validated (a missing comma never fuses two addresses), temporary SMTP
      errors retried through common/api_retry.py, never raises
    * the same card layout: navy header with the reporting period, "Hello Team",
      KPI cards, goal bars, a named "Open …" button + text link, signature and
      "Automated report · Generated …" footer.
    * star=True (the two Coordinator scripts' STAR_EMAIL_IN_GMAIL) stars the sent
      e-mail in the sending Gmail account through common/gmail_star.py — after a
      successful send only, best-effort, never affecting the send result.
  The e-mails carry the Google Sheet LINK (no attachment): the Coordinator works
  in the live sheet, and an attached copy would collect follow-ups the
  performance report never sees.
================================================================================
"""
from __future__ import annotations

import re
import html
from datetime import datetime

DEFAULT_SENDER = "info@intellibiinnovationstechnologies.in"
_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# =============================================================================
#  SENDING
# =============================================================================
def valid_recipients(recipients):
    """Split accidental joins (comma / semicolon / space), drop malformed
    addresses with a warning, de-duplicate case-insensitively."""
    seen, out = set(), []
    for raw in recipients or []:
        for part in re.split(r"[,\s;]+", str(raw).strip()):
            addr = part.strip()
            if not addr:
                continue
            if not _EMAIL_RE.match(addr):
                print(f"[Email] WARNING skipping malformed recipient {addr!r} "
                      f"— check EMAIL_RECIPIENTS for a missing comma")
                continue
            if addr.lower() not in seen:
                seen.add(addr.lower())
                out.append(addr)
    return out


def send(subject: str, html_body: str, recipients, sender: str = DEFAULT_SENDER,
         star: bool = False) -> bool:
    """Send one report e-mail. True when sent, False otherwise (never raises).
    The app password in credentials/email_config.py belongs to GMAIL_SENDER, so
    the requested sender must be that account. star=True: once sent, the e-mail
    is Starred (★) in that Gmail account (common/gmail_star.py, best-effort)."""
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    recipients = valid_recipients(recipients)
    if not recipients:
        print("[Email] no valid recipients — nothing sent")
        return False
    try:
        import email_config as ec                            # credentials/email_config.py
        account, app_pass = str(ec.GMAIL_SENDER).strip(), ec.GMAIL_APP_PASS
    except Exception as exc:                                 # noqa: BLE001
        print(f"[Email] FAILED — could not load credentials/email_config.py: {exc}")
        return False
    if account.lower() != str(sender).strip().lower():
        print(f"[Email] FAILED — sender {sender} is not the account configured in "
              f"credentials/email_config.py ({account or 'none'}).")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = account
    msg["To"] = ", ".join(recipients)
    gs = _gmail_star() if star else None
    if gs is not None:                                       # lets the sent copy be found & starred
        msg["Message-ID"] = gs.new_message_id(account)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    def _send():
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as srv:
            srv.login(account, app_pass)
            srv.sendmail(account, recipients, msg.as_string())

    try:
        try:
            import api_retry                                 # common/api_retry.py
            api_retry.call_with_retry(_send, "SMTP: send to " + ", ".join(recipients), log=print)
        except ImportError:                                  # pragma: no cover
            _send()
        print(f"[Email] ✓ sent to {', '.join(recipients)}  ({subject})")
    except Exception as exc:                                 # noqa: BLE001
        print(f"[Email] FAILED ({subject}): {exc}")
        return False
    if gs is not None:                                       # never changes the result
        gs.star_sent_message(account, app_pass, msg["Message-ID"], subject)
    return True


def _gmail_star():
    """common/gmail_star.py, or None (with a warning) if it cannot be loaded."""
    try:
        import gmail_star
        return gmail_star
    except Exception as exc:                                 # noqa: BLE001
        print(f"[Email] ★ starring unavailable — common/gmail_star.py: {exc}")
        return None


# =============================================================================
#  BODY  (same visual language as the Sales lead-performance e-mail)
# =============================================================================
NAVY, HEADER, BTN = "1B355E", "#1B355E", "#2B547E"
GREEN, AMBER, RED, MUTED = "1B5E20", "BF360C", "B71C1C", "5b6b86"
_E = html.escape


def gen_stamp():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d-%b-%Y %I:%M %p IST")
    except Exception:                                         # pragma: no cover
        return datetime.now().strftime("%d-%b-%Y %I:%M %p")


def band_color(pct, on_track=75.0, watch=45.0):
    if pct is None:
        return MUTED
    return GREEN if pct >= on_track else (AMBER if pct >= watch else RED)


def _cards(items):
    return "".join(
        "<td style='padding:6px;vertical-align:top'>"
        "<div style='background:#f4f8fd;border:1px solid #e2e8f0;border-radius:8px;"
        "padding:14px 8px;text-align:center'>"
        f"<div style='font-size:23px;font-weight:700;color:#{color}'>{v}</div>"
        f"<div style='font-size:11.5px;color:#5b6b86;margin-top:4px;line-height:1.35'>{lbl}</div>"
        "</div></td>"
        for lbl, v, color in items)


def card_block(items, per_row):
    if not items:
        return ""
    out = []
    for i in range(0, len(items), per_row):
        chunk = list(items[i:i + per_row])
        cells = _cards(chunk) + "<td style='padding:6px'></td>" * (per_row - len(chunk))
        out.append("<table role=\"presentation\" width=\"100%\" style=\"border-collapse:"
                   "separate;table-layout:fixed;margin:0 -6px 4px\"><tr>" + cells + "</tr></table>")
    return "".join(out)


def section(title):
    return ("<div style='font-size:11px;font-weight:700;letter-spacing:.06em;"
            "text-transform:uppercase;color:#5b6b86;margin:18px 0 8px;padding-bottom:5px;"
            f"border-bottom:1px solid #e2e8f0'>{title}</div>")


def bar(name, value_label, pct, hexc, goal_pct, note):
    pct = max(0, min(100, pct or 0))
    goal = (f"<div style='position:absolute;top:-2px;bottom:-2px;left:{goal_pct:.0f}%;"
            "width:2px;background:#33415580'></div>") if goal_pct else ""
    return (
        "<div style='margin:11px 0'>"
        "<table role='presentation' width='100%' style='border-collapse:collapse'><tr>"
        f"<td style='font-size:13px;color:#1a2a48'>{name}</td>"
        f"<td style='font-size:13px;font-weight:700;color:#{hexc};text-align:right'>{value_label}</td>"
        "</tr></table>"
        "<div style='height:12px;border-radius:999px;background:#eef1f6;position:relative;"
        "overflow:hidden;margin-top:5px'>"
        f"<div style='height:100%;border-radius:999px;background:#{hexc};width:{pct:.0f}%'></div>"
        f"{goal}</div>"
        f"<div style='font-size:10.5px;color:#5b6b86;margin-top:3px'>{note}</div></div>")


def table(headers, rows, align=None):
    """Compact bordered table (used for the task-group breakdown)."""
    align = align or ["left"] + ["center"] * (len(headers) - 1)
    th = "".join(f"<th style='background:{HEADER};color:#fff;padding:7px 8px;font-size:11.5px;"
                 f"text-align:{a}'>{h}</th>" for h, a in zip(headers, align))
    trs = []
    for i, r in enumerate(rows):
        bg = "#ffffff" if i % 2 == 0 else "#f6f8fb"
        tds = "".join(f"<td style='padding:6px 8px;font-size:12.5px;border-bottom:1px solid #e8edf5;"
                      f"text-align:{a};background:{bg}'>{v}</td>" for v, a in zip(r, align))
        trs.append(f"<tr>{tds}</tr>")
    return ("<table role='presentation' width='100%' style='border-collapse:collapse;"
            f"border:1px solid #e2e8f0'><tr>{th}</tr>{''.join(trs)}</table>")


def page(title, period, intro, sections_html, url, link_name, closing=""):
    """The report e-mail shell (header band, greeting, content, button, signature)."""
    if url:
        cta = (f'<p style="text-align:center;margin:22px 0 24px">'
               f'<a href="{_E(url)}" style="background:{BTN};color:#ffffff;text-decoration:none;'
               f'padding:13px 30px;border-radius:6px;font-weight:600;font-size:15px;'
               f'display:inline-block">Open {_E(link_name)} &nbsp;&rsaquo;</a></p>'
               f'<p style="margin:0;color:#5b6b86;font-size:13px">Or open it here: '
               f'<a href="{_E(url)}" style="color:{BTN};font-weight:600;text-decoration:none">'
               f'{_E(link_name)}</a></p>')
    else:
        cta = ('<p style="margin:20px 0 0;color:#5b6b86;font-size:13px">'
               '(This run did not upload the report to Google Drive.)</p>')
    return f"""<html><body style="margin:0;padding:24px;background:#eef2f8;
  font-family:'Segoe UI',Roboto,Arial,sans-serif;color:#1a2a48">
  <div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:10px;
    overflow:hidden;border:1px solid #e2e8f0">
    <div style="background:{HEADER};padding:22px 28px;color:#ffffff">
      <div style="font-size:20px;font-weight:700">IntelliBI &nbsp;&middot;&nbsp; {_E(title)}</div>
      <div style="font-size:13px;opacity:.85;margin-top:4px">Reporting Period: {_E(period)}</div>
    </div>
    <div style="padding:22px 28px">
      <p style="margin:0 0 14px">Hello Team,</p>
      <p style="margin:0 0 4px;line-height:1.5">{intro}</p>
      {sections_html}
      {closing}
      {cta}
      <p style="margin:26px 0 0;line-height:1.5">
        Thanks &amp; Regards,<br><b>IntelliBI Automation Team</b></p>
    </div>
    <div style="background:#f4f8fd;padding:12px 28px;font-size:12px;color:#8494ad;
      border-top:1px solid #e8edf5">
      Automated report &middot; Generated {gen_stamp()}</div>
  </div>
</body></html>"""


def _fmt_pct(p):
    return "—" if p is None else f"{p:.0f}%"


def _pct1(p):
    """Completion % exactly as the Dashboard scorecard shows it (0.0"%")."""
    return None if p is None else round(p, 1)


def _fmt_pct1(p):
    return "—" if p is None else f"{_pct1(p):.1f}%"


def goal_rows(groups, benchmark=95.0, group_labels=None):
    """Performance-vs-Goals rows for the task groups that HAVE tasks in the
    period: [(display label, registry name, completed, tasks, pct, met)].
    pct is the scorecard's own Completion % rounded like the Dashboard, and
    `met` is judged on that shown value (so 94.96 → 95.0% → met)."""
    labels = group_labels or {}
    rows = []
    for name, g in groups:
        if not g["tasks"]:
            continue                       # group not applicable to this period
        pct = _pct1(g["completion_pct"])
        rows.append((labels.get(name, name), name, g["completed"], g["tasks"], pct,
                     pct is not None and pct >= benchmark))
    return rows


def overall_goal_row(overall, benchmark=95.0):
    """The "Overall Completion %" row from the period's own summarise() — the
    SAME figure as the Dashboard's "All task groups" Completion % (shown and
    judged at 0.1%, like the task-group rows). None when the period has no tasks."""
    if not overall or not overall.get("tasks"):
        return None
    pct = _pct1(overall["completion_pct"])
    return ("Overall Completion %", "All task groups", overall["completed"], overall["tasks"], pct,
            pct is not None and pct >= benchmark)


def goals_section(groups, benchmark=95.0, group_labels=None, overall=None):
    """CLPR-style "Performance vs Goals": "Overall Completion %" first (when the
    period's summary is given), then one bar per applicable task group, each
    Completion % against the benchmark marker. Green = benchmark met, red = below."""
    rows = goal_rows(groups, benchmark, group_labels)
    first = overall_goal_row(overall, benchmark)
    if first:
        rows = [first] + rows
    html = section("Performance vs Goals")
    if not rows:
        return html + ("<p style='margin:8px 0;color:#5b6b86;font-size:13px'>"
                       "No tasks were generated in this period.</p>")
    html += (f"<p style='margin:0 0 2px;color:#5b6b86;font-size:12px'>Benchmark: "
             f"<b style='color:#1a2a48'>{benchmark:g}% Task Completion</b> &nbsp;&middot;&nbsp; "
             f"<span style='color:#{GREEN}'>&#9632;</span> met &nbsp; "
             f"<span style='color:#{RED}'>&#9632;</span> below</p>")
    for disp, name, done, n, pct, met in rows:
        note = f"Goal: {benchmark:g}% Task Completion"
        if disp != name:
            note += f" &nbsp;&middot;&nbsp; {_E(name)}"
        html += bar(_E(disp), f"{done} / {n} &middot; {pct:.1f}%", pct,
                    GREEN if met else RED, benchmark, note)
    return html


def performance_html(kind, label, s, groups, url, median_label, on_track=75.0, watch=45.0,
                     benchmark=95.0, group_labels=None):
    """Task Performance e-mail. s = summarise() of the period; groups =
    [(group name, summarise() dict)] — the Dashboard scorecard rows."""
    c_hex = band_color(s["completion_pct"], on_track, watch)
    t_hex = band_color(s["timely_pct"], on_track, watch)
    st_hex = {"On track": GREEN, "Watch": AMBER, "Behind": RED}.get(s["status"], MUTED)
    volume = [("Tasks Generated", s["tasks"], NAVY),
              ("Completed", s["completed"], GREEN if s["completed"] else NAVY),
              ("Pending", s["pending"], RED if s["pending"] else GREEN)]
    quality = [("Completion %", _fmt_pct(s["completion_pct"]), c_hex),
               ("Timely<br>Completion %", _fmt_pct(s["timely_pct"]), t_hex),
               ("Median Time<br>to Complete", median_label, NAVY)]
    body = (section(f"Overall &nbsp;&middot;&nbsp; <span style='color:#{st_hex}'>{_E(s['status'])}</span>")
            + section("Task Volume") + card_block(volume, 3)
            + section("Completion &amp; Timeliness") + card_block(quality, 3)
            + goals_section(groups, benchmark, group_labels, overall=s))
    rows = []
    for name, g in groups:
        if not g["tasks"]:
            continue
        gh = {"On track": GREEN, "Watch": AMBER, "Behind": RED}.get(g["status"], MUTED)
        rows.append([_E(name), g["tasks"], g["completed"],
                     f"<b style='color:#{RED if g['pending'] else NAVY}'>{g['pending']}</b>",
                     f"<b style='color:#{band_color(g['completion_pct'], on_track, watch)}'>"
                     f"{_fmt_pct1(g['completion_pct'])}</b>",
                     f"<b style='color:#{gh}'>{_E(g['status'])}</b>"])
    if rows:
        body += section("Task Groups") + table(
            ["Task Group", "Tasks", "Completed", "Pending", "Completion %", "Status"], rows)
    intro = (f"Please find the <b>{_E(kind)}</b> Coordinator task-performance report for "
             f"<b>{_E(label)}</b>. Here is a quick snapshot:")
    closing = ('<p style="margin:20px 0 0;line-height:1.5">The full report covers the task-group '
               'scorecard, the progress trend and the task register (every task, including '
               'every pending one).</p>')
    return page(f"{kind} Coordinator Task Performance Report", label, intro, body, url,
                f"{kind} Coordinator Task Performance Report", closing)


def batch_coordinator_html(kind, label, url, task_counts=None, note=None, today_label=None):
    """Batch Coordinator e-mail. task_counts = {task tab: rows needing action}
    for the Daily task list (None for the Weekly / Monthly / Manual roll-ups).
    today_label = the run day ("DD-Mon-YYYY"); a Daily report for another day
    (DAILY_DATE pinned) is then not worded as "Today's"."""
    is_today = today_label is None or today_label == label
    body = ""
    if task_counts:
        total = sum(task_counts.values())
        items = [(_E(k), v, RED if v else GREEN) for k, v in task_counts.items()]
        body += (section("Today's Task List" if is_today else f"Task List — {label}") + card_block([("Total Tasks", total, NAVY)], 1)
                 + card_block(items, 3))
    if note:
        body += f"<p style='margin:12px 0 0;color:#5b6b86;font-size:12.5px'>{_E(note)}</p>"
    if kind == "Daily":
        intro = (("Today's" if is_today else "The") +
                 f" Coordinator task list for <b>{_E(label)}</b> is ready. "
                 "Please work through each tab and record <b>Action Taken</b>, "
                 "<b>Follow-Up Comment</b> and <b>Follow-Up Done?</b> directly in the Google Sheet "
                 "— the evening Task Performance report reads them from there.")
        link = "Daily Coordinator Task List"
    else:
        intro = (f"Please find the <b>{_E(kind)}</b> Batch Coordinator follow-up summary for "
                 f"<b>{_E(label)}</b> (one row per learner / instructor over the period).")
        link = f"{kind} Batch Coordinator Report"
    return page(f"{kind} Batch Coordinator Report", label, intro, body, url, link)
