"""
================================================================================
  IntelliBI — Gmail "Starred" for report e-mails   (common/gmail_star.py)
  ------------------------------------------------------------------------------
  Shared by the report scripts that opt in (the SAME file is kept in the Sales
  and the Operations projects):
      Operations  co-ordinator reports/pyCoordinatorTaskPerformanceReport.py
                  co-ordinator reports/pyBatchCoordinatorDailyAttendanceReport.py
                  (both through co-ordinator reports/coordinator_email.send)
      Sales       sales_reports/pyConsolidatedLeadPerformanceReport.py
                  sales_reports/pyLeadFollowUpAnalysisReport.py

  HOW IT WORKS
    1. Before the e-mail is sent, the script stamps it with its own unique
       Message-ID (new_message_id) — a standard header Gmail would otherwise
       add itself; subject, body, recipients and attachments are untouched.
    2. After Gmail SMTP has accepted it, star_sent_message() logs in to the
       SAME Gmail account over IMAP (imap.gmail.com:993, the same address and
       app password as the SMTP send — nothing new to configure), finds that
       exact message by its Message-ID in "All Mail" and sets the IMAP \\Flagged
       flag, which IS Gmail's Starred (★). Because the sender's own address
       (info@…) is also a recipient, the starred message is the one in its Inbox.

  SCOPE / LIMITS
    * A star belongs to ONE mailbox. The sending account can star its own copy;
      no e-mail header can make a message arrive starred in somebody else's
      Gmail. External recipients (e.g. @gmail.com addresses) star it with a
      one-time Gmail filter — see docs/GMAIL_STARRED_REPORTS.md.
    * Best-effort and isolated: a starring problem (IMAP disabled, network,
      message not visible yet) is printed as a warning and NEVER changes the
      send result, the retries, the exit code or anything else in the run.

  CHECK FROM THE COMMAND LINE (read-only)
      python common/gmail_star.py --check            # last 2 days of report mails
      python common/gmail_star.py --check --days 7
================================================================================
"""
from __future__ import annotations

import imaplib
import re
import socket
import time
from email.utils import make_msgid

IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993
IMAP_TIMEOUT = 15                       # seconds per network operation
# how long to wait for the just-sent message to become visible in the mailbox
FIND_DELAYS = (1, 2, 3, 4, 5, 5, 5, 5)  # seconds between look-ups (≈30 s in total)
MAX_SESSIONS = 2                        # reconnect once on a dropped IMAP session
# accounts whose IMAP was unreachable / refused in THIS run: later e-mails of the
# same run skip starring at once instead of waiting for the same failure again
_GIVE_UP: dict = {}

# Subjects of the four report e-mails (used only by --check)
REPORT_SUBJECTS = ("Coordinator Task Performance Report", "Batch Coordinator Report",
                   "Lead Report", "Lead Follow-Up Analysis Report")


def new_message_id(sender: str) -> str:
    """A unique RFC 5322 Message-ID on the sender's domain, set on the e-mail
    BEFORE it is sent so the sent copy can be found again exactly."""
    domain = sender.split("@", 1)[1] if "@" in (sender or "") else "intellibi.local"
    return make_msgid(idstring="intellibi-report", domain=domain)


def _all_mail_folder(imap) -> str:
    """The mailbox holding every message ('[Gmail]/All Mail', or its localised
    name) — found by its \\All special-use flag, never by a hard-coded name."""
    typ, data = imap.list()
    if typ == "OK":
        for raw in data or []:
            line = raw.decode("utf-8", "replace") if isinstance(raw, bytes) else str(raw)
            if "\\All" in line:
                m = re.search(r'"([^"]+)"\s*$', line) or re.search(r"(\S+)\s*$", line)
                if m:
                    return m.group(1)
    return "[Gmail]/All Mail"


def _quote(s: str) -> str:
    """IMAP quoted string (RFC 3501): backslash and double quote escaped."""
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _search(imap, msgid: str, subject: str | None):
    """UIDs of the message: exact Message-ID first; the subject (sent today) only
    as a fallback, newest match only."""
    bare = msgid.strip().strip("<>")
    typ, data = imap.uid("SEARCH", "X-GM-RAW", _quote(f"rfc822msgid:{bare}"))
    uids = (data[0] or b"").split() if typ == "OK" and data else []
    if uids:
        return uids
    typ, data = imap.uid("SEARCH", None, "HEADER", "Message-ID", _quote(msgid.strip()))
    uids = (data[0] or b"").split() if typ == "OK" and data else []
    if uids or not subject:
        return uids
    words = re.findall(r"[A-Za-z0-9]+", subject)[:10]          # ASCII words: safe in X-GM-RAW
    if not words:
        return []
    typ, data = imap.uid("SEARCH", "X-GM-RAW",
                         _quote(f'in:sent newer_than:1d subject:({" ".join(words)})'))
    uids = (data[0] or b"").split() if typ == "OK" and data else []
    return uids[-1:]                                          # the newest one only


def star_sent_message(sender: str, app_pass: str, message_id: str, subject: str | None = None,
                      log=print, delays=FIND_DELAYS) -> bool:
    """Star (★) the just-sent message in the sender's Gmail mailbox. Returns
    True when starred, False otherwise. NEVER raises."""
    if not (sender and app_pass and message_id):
        log("  [Email] ★ not starred — no sender / Message-ID available.")
        return False
    if sender.lower() in _GIVE_UP:
        log(f"  [Email] ★ not starred — {_GIVE_UP[sender.lower()]} (see the first warning of this run).")
        return False
    last_err = None
    for _session in range(MAX_SESSIONS):
        imap = None
        try:
            imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=IMAP_TIMEOUT)
            imap.login(sender, app_pass)
            imap.select(_quote(_all_mail_folder(imap)), readonly=False)
            for wait in (0,) + tuple(delays):
                if wait:
                    time.sleep(wait)
                    imap.noop()                                 # refresh the mailbox view
                uids = _search(imap, message_id, subject)
                if uids:
                    for uid in uids:
                        imap.uid("STORE", uid, "+FLAGS", "(\\Flagged)")
                    log(f"  [Email] ★ starred in Gmail ({sender})")
                    return True
            log(f"  [Email] ★ not starred — the sent message did not appear in {sender} "
                f"within {sum(delays)} s (the e-mail itself was sent).")
            return False
        except imaplib.IMAP4.error as exc:                    # login refused, IMAP disabled …
            msg = str(exc)
            hint = (" — check that IMAP is enabled for this Gmail account "
                    "(Gmail ▸ Settings ▸ Forwarding and POP/IMAP) and that the app password is valid"
                    if "LOGIN" in msg.upper() or "AUTH" in msg.upper() or "IMAP" in msg.upper() else "")
            log(f"  [Email] ★ not starred — Gmail IMAP: {msg}{hint} (the e-mail itself was sent).")
            _GIVE_UP[sender.lower()] = "Gmail IMAP login refused"
            return False
        except (OSError, socket.timeout, imaplib.IMAP4.abort) as exc:   # network: retry once
            last_err = exc
        except Exception as exc:                                # noqa: BLE001 — never break a run
            log(f"  [Email] ★ not starred — {type(exc).__name__}: {exc} (the e-mail itself was sent).")
            return False
        finally:
            if imap is not None:
                try:
                    imap.logout()
                except Exception:                               # noqa: BLE001
                    pass
    log(f"  [Email] ★ not starred — Gmail IMAP unreachable ({last_err}); the e-mail itself was sent.")
    _GIVE_UP[sender.lower()] = "Gmail IMAP unreachable"
    return False


def check_recent(sender: str, app_pass: str, days: int = 2, log=print) -> list:
    """Read-only: list the recent report e-mails in the sender's mailbox with
    their Starred state. Returns [(date, subject, starred)]."""
    out = []
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, timeout=IMAP_TIMEOUT)
    try:
        imap.login(sender, app_pass)
        imap.select(_quote(_all_mail_folder(imap)), readonly=True)
        subj = " OR ".join(f'"{s}"' for s in REPORT_SUBJECTS)
        typ, data = imap.uid("SEARCH", "X-GM-RAW",
                             _quote(f"from:{sender} newer_than:{int(days)}d subject:({subj})"))
        uids = (data[0] or b"").split() if typ == "OK" and data else []
        for uid in uids[-50:]:
            typ, d = imap.uid("FETCH", uid, "(FLAGS BODY.PEEK[HEADER.FIELDS (SUBJECT DATE)])")
            if typ != "OK" or not d or not isinstance(d[0], tuple):
                continue
            flags = d[0][0].decode("utf-8", "replace")
            hdr = d[0][1].decode("utf-8", "replace")
            from email.header import decode_header, make_header
            sm = re.search(r"^Subject:\s*(.*?)\r?\n(?!\s)", hdr, re.S | re.M | re.I)
            dm = re.search(r"^Date:\s*(.*)$", hdr, re.M | re.I)
            subject = str(make_header(decode_header(sm.group(1).replace("\r\n", "")))) if sm else "?"
            out.append(((dm.group(1).strip() if dm else "?"), subject, "\\Flagged" in flags))
    finally:
        try:
            imap.logout()
        except Exception:                                       # noqa: BLE001
            pass
    for dt, s, st in out:
        log(f"  {'★' if st else '☆'}  {dt[:31]:<31}  {s}")
    if not out:
        log("  (no report e-mails found in that period)")
    return out


if __name__ == "__main__":                                       # pragma: no cover — manual check
    import argparse
    import os
    import sys
    ap = argparse.ArgumentParser(description="Show which recent IntelliBI report e-mails are Starred.")
    ap.add_argument("--check", action="store_true", help="list recent report e-mails and their star")
    ap.add_argument("--days", type=int, default=2)
    a = ap.parse_args()
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, os.path.join(root, "credentials"))
    import email_config as _ec                                   # the account the reports are sent from
    print(f"Report e-mails in {_ec.GMAIL_SENDER} (last {a.days} day(s)):  ★ = starred")
    check_recent(_ec.GMAIL_SENDER, _ec.GMAIL_APP_PASS, a.days)
