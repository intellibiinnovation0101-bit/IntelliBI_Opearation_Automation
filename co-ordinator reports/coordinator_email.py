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


_IMPACT_HEX = {"On target": GREEN, "Near target": AMBER, "Below target": RED}
_QUAD_HEX = {"ok": GREEN, "medium": AMBER, "high": RED, "info": "0D47A1", "muted": MUTED}


def _chip(text, hexc):
    return (f"<span style='display:inline-block;padding:2px 8px;border-radius:999px;"
            f"background:#{hexc}14;border:1px solid #{hexc}55;color:#{hexc};font-size:11px;"
            f"font-weight:700'>{_E(text)}</span>")


def _flow_cell(label, value, hexc, first=False):
    """One step of the Task Group -> Effort -> Completion -> Outcome strip."""
    arrow = ("" if first else
             "<td style='width:14px;text-align:center;color:#b6c2d4;font-size:14px;"
             "padding:0 2px'>&rarr;</td>")
    return (arrow + "<td style='text-align:center;padding:6px 4px;background:#f7f9fc;"
            "border:1px solid #e8edf5;border-radius:6px'>"
            f"<div style='font-size:16px;font-weight:700;color:#{hexc};line-height:1.2'>{value}</div>"
            f"<div style='font-size:10px;color:#5b6b86;margin-top:2px;text-transform:uppercase;"
            f"letter-spacing:.04em'>{label}</div></td>")


def _sub_label(text):
    return ("<div style='font-size:10px;font-weight:700;letter-spacing:.08em;color:#8494ad;"
            f"text-transform:uppercase;margin:12px 0 -4px'>{text}</div>")


def _group_card(title, subtitle, chip_html, accent_hex, body_html, number=None):
    """A bordered, separated block for one task group (e-mail-safe table)."""
    num = (f"<span style='display:inline-block;min-width:20px;height:20px;line-height:20px;"
           f"border-radius:10px;background:#{accent_hex};color:#ffffff;font-size:11px;"
           f"font-weight:700;text-align:center;margin-right:8px'>{number}</span>"
           if number is not None else "")
    sub = (f"<div style='font-size:11px;color:#8494ad;margin-top:2px;padding-left:"
           f"{'28px' if number is not None else '0'}'>{subtitle}</div>" if subtitle else "")
    return ("<table role='presentation' width='100%' style='border-collapse:separate;"
            "border-spacing:0;margin:14px 0 0;border:1px solid #dde4ee;"
            f"border-left:4px solid #{accent_hex};border-radius:8px;background:#ffffff'>"
            "<tr><td style='padding:10px 14px;background:#f4f7fb;border-bottom:1px solid #e8edf5;"
            "border-top-right-radius:8px'>"
            "<table role='presentation' width='100%' style='border-collapse:collapse'><tr>"
            f"<td style='font-size:14.5px;font-weight:700;color:#1a2a48'>{num}{title}{sub}</td>"
            f"<td style='text-align:right;vertical-align:top;white-space:nowrap'>{chip_html}</td>"
            "</tr></table></td></tr>"
            f"<tr><td style='padding:6px 14px 12px'>{body_html}</td></tr></table>")


_FR_HEX = {"ok": GREEN, "medium": AMBER, "high": RED, "info": "0D47A1", "muted": MUTED}
PREV_BAR_HEX = "8FA3BF"           # the previous period's bar: neutral slate (current = verdict colour)


def _change_html(v, band=1.0, unit=""):
    """▲ +3.2 pts (green) / ▼ 21.0 pts (red) / ● 0.4 pts (blue) / — ."""
    if v is None:
        return f"<span style='color:#{MUTED}'>&mdash;</span>"
    if v >= band:
        sym, hexc = "&#9650;", GREEN
    elif v <= -band:
        sym, hexc = "&#9660;", RED
    else:
        sym, hexc = "&#9679;", "0D47A1"
    return (f"<span style='display:inline-block;padding:1px 6px;border-radius:999px;"
            f"background:#{hexc}14;border:1px solid #{hexc}55;color:#{hexc};font-weight:700;"
            f"font-size:11.5px;white-space:nowrap'>{sym}&nbsp;{abs(v):.1f}"
            + (f"&nbsp;{unit}" if unit else "") + "</span>")


def _cmp_value(p, av, hexc):
    """An outcome / effort value of the comparison, or its status when there is
    no figure (Not available / Nothing to measure …); provisional / in-progress
    figures carry that tag."""
    word = (av or ("Final",))[0]
    if p is None:
        return f"<span style='color:#{MUTED};font-style:italic;font-size:11.5px'>{_E(word)}</span>"
    tag = {"Provisional": " <span style='font-size:10px;color:#BF360C'>prov.</span>",
           "In progress": " <span style='font-size:10px;color:#0D47A1'>to date</span>"}.get(word, "")
    return f"<b style='color:#{hexc}'>{p:.1f}%</b>{tag}"


def compare_card_body(r, cl, benchmark=95.0, band=1.0, impact_hex=None, near_band=15.0):
    """Previous → current for one follow-up group (e-mail): a 3-row comparison
    table (tasks done, effort %, outcome %), paired outcome bars against the
    target, the Follow-up → Result verdict and what each outcome measures.
    r = a scorecard row with r['compare']; cl = the view's compare labels."""
    c = r["compare"]
    ih = impact_hex or {}
    pt, ct = _E(cl["prev_tag"]), _E(cl["cur_tag"])
    eff_hex = lambda p: MUTED if p is None else (GREEN if round(p, 1) >= benchmark else RED)
    target = r.get("target")

    def out_hex(p):
        if p is None or target is None:
            return MUTED
        return ih.get("On target" if p >= target else ("Near target" if p >= target - near_band else
                                                       "Below target"), MUTED)

    def tasks_txt(s):
        return (f"<b>{s['completed']} / {s['tasks']}</b>" if s and s.get("tasks") else
                f"<span style='color:#{MUTED};font-style:italic;font-size:11.5px'>No tasks</span>")
    obs = (f"Next-day observation &middot; {ct}'s outcome window starts after {pt.lower()}'s task list"
           if cl.get("observation") == "next_day" else
           "Same-period comparison &middot; each period's outcome covers its own sessions / deadlines")
    th = ("<th style='width:22%;padding:6px 4px;font-size:11px;color:#ffffff;background:{bg};text-align:center;"
          "font-weight:700;line-height:1.25'>{t}<br><span style='font-weight:400;font-size:10px;"
          "opacity:.9'>{d}</span></th>")
    td = ("<td style='padding:7px 4px;font-size:12.5px;border-bottom:1px solid #e8edf5;"
          "text-align:{a};background:{bg}'>{v}</td>")
    head = ("<tr><th style='width:34%;padding:6px 6px;font-size:11px;color:#ffffff;background:#5b6b86;"
            "text-align:left'>&nbsp;</th>"
            + th.format(bg="#5b6b86", t=pt, d=_E(cl["prev_dates"]))
            + th.format(bg=HEADER, t=ct, d=_E(cl["cur_dates"]) + (
                f"<br>as of {_E(cl['as_of'])}" if cl.get("as_of") else ""))
            + th.format(bg="#7A4A12", t="Change", d=f"{ct} &minus; {pt}<br>(points)") + "</tr>")
    body_rows = [
        ("Tasks Completed", tasks_txt(c["prev_s"]), tasks_txt(r["s"]), ""),
        ("Effort &middot; Task Completion %",
         _cmp_value(c["prev_comp"], None, eff_hex(c["prev_comp"])) if c["prev_comp"] is not None
         else f"<span style='color:#{MUTED}'>&mdash;</span>",
         _cmp_value(c["cur_comp"], None, eff_hex(c["cur_comp"])) if c["cur_comp"] is not None
         else f"<span style='color:#{MUTED}'>&mdash;</span>",
         _change_html(c["d_completion"], band)),
        (f"Outcome &middot; {_E(r['measure'])}",
         _cmp_value(c["prev_actual"], c["prev_av"], out_hex(c["prev_actual"])),
         _cmp_value(c["cur_actual"], c["cur_av"], out_hex(c["cur_actual"])),
         _change_html(c["change"], band)),
    ]
    trs = ""
    for i, (lbl, a, b, d) in enumerate(body_rows):
        bg = "#ffffff" if i % 2 == 0 else "#f6f8fb"
        trs += ("<tr>" + td.format(a="left", bg=bg, v=f"<span style='color:#1a2a48'>{lbl}</span>")
                + td.format(a="center", bg=bg, v=a) + td.format(a="center", bg=bg, v=b)
                + td.format(a="center", bg=bg, v=d) + "</tr>")
    html = (f"<div style='font-size:10.5px;color:#8494ad;margin:8px 0 4px'>{obs}</div>"
            "<table role='presentation' width='100%' style='border-collapse:collapse;"
            f"border:1px solid #e2e8f0;table-layout:fixed'>{head}{trs}</table>")
    # paired outcome bars: previous (slate) above current (target-coloured)
    html += _sub_label(f"Actual Outcome &middot; {pt} vs {ct}")
    for tag, p, av, win, hexc in ((pt, c["prev_actual"], c["prev_av"], c["prev_window"], PREV_BAR_HEX),
                                  (ct, c["cur_actual"], c["cur_av"], c["cur_window"],
                                   out_hex(c["cur_actual"]))):
        note = (f"Target: {target:g}% &nbsp;&middot;&nbsp; " if target is not None else "") + _E(win or "")
        if p is None:
            html += (f"<div style='font-size:11.5px;color:#5b6b86;margin:10px 0 4px'>{tag}: "
                     f"{_E(av[0])}" + (f" &mdash; {_E(av[2])}" if av[2] else "") + "</div>")
        else:
            label = f"{p:.1f}%" + (f" ({_E(av[0].lower())})" if av[0] in ("Provisional", "In progress")
                                   else "")
            html += bar(f"{tag} &middot; {_E(r['measure'])}", label, p, hexc, target, note)
    fr_hex = _FR_HEX.get(c["verdict_level"], MUTED)
    html += ("<div style='margin:10px 0 0;padding:8px 10px;background:#f7f9fc;border:1px solid #e8edf5;"
             "border-radius:6px;font-size:12px;line-height:1.5'>"
             f"<b style='color:#1a2a48'>Follow-up &rarr; Result:</b> {_chip(c['verdict'], fr_hex)}"
             + (f" &nbsp;<span style='color:#5b6b86'>{_E(c['why'])}</span>" if c.get("why") else "")
             + "</div>")
    return html


def effort_outcome_section(rows, benchmark=95.0, group_labels=None, headline=None,
                           attention=None, quad_levels=None, overall=None, compare=None,
                           compare_band=1.0):
    """The upgraded "Performance vs Goals": ONE card per task group, read top to
    bottom as Task Group -> Effort -> Completion -> Actual Outcome:
      * header: the group, its registry name when the e-mail label differs, and
        the Effort -> Outcome verdict chip (accent colour = the verdict);
      * a 4-step strip: Tasks Generated -> Completed -> Task Completion % ->
        Actual Performance %;
      * an EFFORT bar (Task Completion % vs the completion benchmark) and an
        OUTCOME bar (Actual Performance % vs the area's target).
    "Overall Completion %" comes first in its own card. Presentation only: every
    value, threshold and colour is the one computed before (same bars, same
    rules). rows = pyCoordinatorTaskPerformanceReport.scorecard_rows."""
    labels = group_labels or {}
    quad_levels = quad_levels or {}
    html = section("Performance vs Goals &nbsp;&middot;&nbsp; Effort &rarr; Outcome")
    if headline:
        html += (f"<div style='background:#f4f8fd;border-left:4px solid {HEADER};padding:10px 12px;"
                 f"font-size:13.5px;line-height:1.45;margin:4px 0 10px'><b>{_E(headline)}</b></div>")
    html += (f"<p style='margin:0 0 2px;color:#5b6b86;font-size:12px;line-height:1.5'>Each task group "
             f"reads <b style='color:#1a2a48'>Effort &rarr; Completion &rarr; Actual Outcome</b>. "
             f"Effort bar: Task Completion % (goal {benchmark:g}%) &nbsp;&middot;&nbsp; Outcome bar: "
             f"Actual Performance % (goal = the area's target) &nbsp;&middot;&nbsp; "
             f"<span style='color:#{GREEN}'>&#9632;</span> met "
             f"<span style='color:#{AMBER}'>&#9632;</span> near <span style='color:#{RED}'>&#9632;</span> below</p>")
    if compare and any(r.get("compare") for r in rows):
        names = ", ".join(_E(labels.get(r["name"], r["name"])) for r in rows if r.get("compare"))
        pt, ct = _E(compare["prev_tag"]), _E(compare["cur_tag"])
        why = (f"their outcome window starts before the day's list is worked, so the result that "
               f"follows {pt.lower()}'s follow-ups is {ct.lower()}'s figure"
               if compare.get("observation") == "next_day" else
               "each period's outcome is measured over that period's own sessions / deadlines")
        html += (f"<p style='margin:6px 0 0;color:#5b6b86;font-size:12px;line-height:1.5'>"
                 f"<b style='color:#1a2a48'>{names}</b> are shown {pt} &rarr; {ct}: {why}. "
                 f"Changes are observations, not proof that the follow-ups caused them.</p>")
    shown = 0
    first = overall_goal_row(overall, benchmark)
    if first:                                   # "Overall Completion %" first, as before
        disp, name, done, n, p, met = first
        body = bar(_E(disp), f"{done} / {n} &middot; {p:.1f}%", p, GREEN if met else RED,
                   benchmark, f"Goal: {benchmark:g}% Task Completion &nbsp;&middot;&nbsp; {_E(name)}")
        html += _group_card("All Task Groups", "", "", NAVY, body)
    for r in rows:
        s = r["s"]
        if not s["tasks"] and r["actual"] is None:
            continue
        shown += 1
        disp = labels.get(r["name"], r["name"])
        q_hex = _QUAD_HEX.get(quad_levels.get(r["quadrant"], "muted"), MUTED)
        chip = _chip(r["quadrant"], q_hex) if r["quadrant"] and r["quadrant"] != "—" else ""
        # flow strip: Tasks Generated -> Completed -> Task Completion % -> Actual %
        p = _pct1(s["completion_pct"]) if s["tasks"] else None
        a = round(r["actual"], 1) if r["actual"] is not None else None
        eff_hex = (GREEN if p >= benchmark else RED) if p is not None else MUTED
        act_hex = _IMPACT_HEX.get(r["impact"], MUTED) if a is not None else MUTED
        flow = ("<table role='presentation' width='100%' style='border-collapse:separate;"
                "border-spacing:0;margin:8px 0 2px;table-layout:fixed'><tr>"
                + _flow_cell("Tasks Generated", s["tasks"], NAVY, first=True)
                + _flow_cell("Completed", s["completed"], NAVY)
                + _flow_cell("Task Completion", f"{p:.1f}%" if p is not None else "&mdash;", eff_hex)
                + _flow_cell("Actual Outcome", f"{a:.1f}%" if a is not None else "&mdash;", act_hex)
                + "</tr></table>")
        if compare and r.get("compare"):
            # follow-up group: previous → current comparison instead of the single-period strip
            vc = _FR_HEX.get(r["compare"]["verdict_level"], MUTED)
            q_line = (f"<div style='font-size:11px;color:#5b6b86;margin:8px 0 0'>Effort &rarr; Outcome "
                      f"({_E(compare['cur_tag'])}): {chip}</div>" if chip else "")
            html += _group_card(_E(disp), _E(r["name"]) if disp != r["name"] else "",
                                _chip(r["compare"]["verdict"], vc), vc,
                                compare_card_body(r, compare, benchmark, compare_band, _IMPACT_HEX)
                                + q_line, number=shown)
            continue
        body = flow + _sub_label("Effort &middot; Completion")
        if s["tasks"]:
            body += bar("Effort &middot; Task Completion %", f"{s['completed']} / {s['tasks']} &middot; {p:.1f}%",
                        p, GREEN if p >= benchmark else RED, benchmark,
                        f"Goal: {benchmark:g}% of tasks completed")
        else:
            body += ("<div style='font-size:11.5px;color:#5b6b86;margin:10px 0 4px'>Effort: no Coordinator "
                     "task in this period</div>")
        body += _sub_label("Actual Outcome")
        if a is not None:
            t = r["target"]
            body += bar(f"Outcome &middot; {_E(r['measure'])}", f"{a:.1f}%", a, act_hex, t,
                        f"Target: {t:g}% &nbsp;&middot;&nbsp; {_E(r['basis'])}"
                        + (f" &nbsp;&middot;&nbsp; {r['d_actual']:+.1f} pts vs previous period"
                           if r.get("d_actual") is not None else ""))
        else:
            note = (r["o"] or {}).get("note") if r.get("o") else ""
            body += ("<div style='font-size:11.5px;color:#5b6b86;margin:10px 0 4px'>Outcome: not measured"
                     + (f" — {_E(note)}" if note else "") + "</div>")
        html += _group_card(_E(disp), _E(r["name"]) if disp != r["name"] else "", chip, q_hex, body,
                            number=shown)
    if not shown:
        html += ("<p style='margin:8px 0;color:#5b6b86;font-size:13px'>"
                 "No tasks were generated and no outcome was measured in this period.</p>")
    att = [a for a in (attention or []) if a[1] in ("high", "medium")]
    if att:
        html += section("Needs Management Attention")
        html += table(["Area", "Why", "Action"],
                      [[f"<b style='color:#{_QUAD_HEX.get(lev, MUTED)}'>{_E(area)}</b>", _E(why), _E(act)]
                       for area, lev, why, act in att], align=["left", "left", "left"])
    return html


def performance_html(kind, label, s, groups, url, median_label, on_track=75.0, watch=45.0,
                     benchmark=95.0, group_labels=None, outcome_rows=None, headline=None,
                     attention=None, compare=None, compare_band=1.0):
    """Task Performance e-mail. s = summarise() of the period; groups =
    [(group name, summarise() dict)] — the Dashboard scorecard rows.
    outcome_rows (the Effort → Outcome scorecard rows) switch "Performance vs
    Goals" to the Effort → Outcome view; without them the completion-only
    goals are shown (effort-only runs)."""
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
            + (effort_outcome_section(outcome_rows, benchmark, group_labels, headline, attention,
                                      _quad_levels(), overall=s, compare=compare,
                                      compare_band=compare_band)
               if outcome_rows else goals_section(groups, benchmark, group_labels, overall=s)))
    rows = []
    act = {r["name"]: r for r in (outcome_rows or [])}
    for name, g in groups:
        if not g["tasks"]:
            continue
        gh = {"On track": GREEN, "Watch": AMBER, "Behind": RED}.get(g["status"], MUTED)
        if outcome_rows:
            r = act.get(name) or {}
            a = r.get("actual")
            rows.append([_E(name), g["tasks"], g["completed"],
                         f"<b style='color:#{band_color(g['completion_pct'], on_track, watch)}'>"
                         f"{_fmt_pct1(g['completion_pct'])}</b>",
                         (f"<b style='color:#{_IMPACT_HEX.get(r.get('impact'), MUTED)}'>{a:.1f}%</b>"
                          if a is not None else "—"),
                         _E(r.get("impact", "—"))])
            continue
        rows.append([_E(name), g["tasks"], g["completed"],
                     f"<b style='color:#{RED if g['pending'] else NAVY}'>{g['pending']}</b>",
                     f"<b style='color:#{band_color(g['completion_pct'], on_track, watch)}'>"
                     f"{_fmt_pct1(g['completion_pct'])}</b>",
                     f"<b style='color:#{gh}'>{_E(g['status'])}</b>"])
    if rows:
        body += section("Task Groups") + table(
            (["Task Group", "Tasks Generated", "Tasks Completed", "Task Completion %",
              "Actual Performance %", "Performance / Impact"] if outcome_rows else
             ["Task Group", "Tasks", "Completed", "Pending", "Completion %", "Status"]), rows)
    intro = (f"Please find the <b>{_E(kind)}</b> Coordinator task-performance report for "
             f"<b>{_E(label)}</b>. Here is a quick snapshot:")
    closing = ('<p style="margin:20px 0 0;line-height:1.5">The full report covers the Effort &rarr; '
               'Outcome scorecard, the effort-vs-outcome and progress trends, the task register '
               '(every task, including every pending one) and how each figure is measured.</p>'
               if outcome_rows else
               '<p style="margin:20px 0 0;line-height:1.5">The full report covers the task-group '
               'scorecard, the progress trend and the task register (every task, including '
               'every pending one).</p>')
    return page(f"{kind} Coordinator Task Performance Report", label, intro, body, url,
                f"{kind} Coordinator Task Performance Report", closing)


def _quad_levels():
    try:
        import coordinator_outcomes as CO
        return CO.QUAD_LEVEL
    except Exception:                                         # pragma: no cover
        return {}


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
