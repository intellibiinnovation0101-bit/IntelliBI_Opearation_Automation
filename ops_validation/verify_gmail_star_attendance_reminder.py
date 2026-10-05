"""
Verification of Gmail Starred (★) for the report e-mails of
  ops_reports_action/pyAttendaceFeedbackReport.py          (send_report)
  ops_reports_action/pyAssignmentSubmissionEmailReminder.py (send_staff_summary_email)
(run from the project root:  python ops_validation\\verify_gmail_star_attendance_reminder.py).

No network: Gmail SMTP and IMAP are in memory, ONE MAILBOX PER ACCOUNT (each
recipient gets their own copy; IMAP \\Flagged = Gmail's STARRED label). Checks:
  * the report e-mail is starred in the info@ mailbox only, exactly that message
  * other recipients receive it normally (unstarred, never logged in to)
  * subject / recipients / body / attachments identical with starring on or off
  * failed send → no star attempt; star failure → e-mail result unchanged
  * student reminder e-mails (not sent to info@) are never starred
  * STAR_EMAIL_IN_GMAIL = False → no IMAP at all
Learner data is synthetic.
"""
import email
import imaplib
import io
import os
import re
import smtplib
import sys
import types
from contextlib import redirect_stdout
from datetime import date

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
for p in ("common", "ops_reports_action"):
    sys.path.insert(0, os.path.join(ROOT, p))
INFO, HR = "info@intellibiinnovationstechnologies.in", "intellibihropsb2ch@gmail.com"
ec = types.ModuleType("email_config")
ec.GMAIL_SENDER, ec.GMAIL_APP_PASS = INFO, "test-only"
sys.modules["email_config"] = ec                       # never the real credentials

import pandas as pd                                     # noqa: E402
import gmail_star as GS                                 # noqa: E402
import pyAttendaceFeedbackReport as AR                  # noqa: E402
import pyAssignmentSubmissionEmailReminder as ASG       # noqa: E402

GS.FIND_DELAYS = (0, 0)
GS.time.sleep = lambda s: None
FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


BOXES, LOGINS, SENT = {}, [], []
STATE = {"smtp_fail": False, "imap_fail": False}


def reset():
    BOXES.clear(); LOGINS.clear(); SENT.clear(); GS._GIVE_UP.clear()


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
        SENT.append((sender, list(rcpts), raw))
        m = email.message_from_string(raw)
        subj = str(email.header.make_header(email.header.decode_header(m["Subject"])))
        for a in {sender.lower(), *[r.lower() for r in rcpts]}:   # each mailbox: own copy
            box = BOXES.setdefault(a, [])
            box.append({"uid": str(100 + len(box)).encode(), "msgid": m["Message-ID"] or "",
                        "subject": subj, "flags": set(), "to": list(rcpts)})


class FakeIMAP:
    def __init__(self, *a, **k):
        self.user = None

    def login(self, user, pw):
        LOGINS.append(user.lower())
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
            return "OK", [b" ".join(x["uid"] for x in box
                                    if m and x["msgid"].strip("<>") == m.group(1))]
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


def starred(addr):
    return [x["subject"] for x in BOXES.get(addr, []) if "\\Flagged" in x["flags"]]


def payloads(raw):
    m = email.message_from_string(raw)
    return [(p.get_content_type(), p.get_filename(), p.get_payload(decode=True))
            for p in m.walk() if not p.is_multipart()]


# =============================================================================
print("\n== 1. Attendance & Feedback report e-mail ==")
check("STAR_EMAIL_IN_GMAIL on by default", AR.STAR_EMAIL_IN_GMAIL, True)
reset()
out = io.StringIO()
with redirect_stdout(out):
    AR.send_report("daily", "04-Oct-2026", "IntelliBI_Daily_Attendance.xlsx", b"PK-test-bytes",
                   12, 120, 9, 92.5, att_rows=[], report_url="https://docs.google.com/x")
subj = "IntelliBI Daily Attendance & Feedback Report — 04-Oct-2026"
check("sent to the existing recipients", SENT[0][1], AR.REPORT_TO + AR.REPORT_CC)
check("starred in the info@ mailbox", starred(INFO), [subj])
check("other recipient receives it, not starred", ([x["subject"] for x in BOXES[HR]], starred(HR)), ([subj], []))
check("only info@ opened for starring", sorted(set(LOGINS)), [INFO])
check("logged", "[Email] ★ starred in Gmail (info@intellibiinnovationstechnologies.in)" in out.getvalue(), True)
with_star = payloads(SENT[0][2])
reset(); AR.STAR_EMAIL_IN_GMAIL = False
with redirect_stdout(io.StringIO()):
    AR.send_report("daily", "04-Oct-2026", "IntelliBI_Daily_Attendance.xlsx", b"PK-test-bytes",
                   12, 120, 9, 92.5, att_rows=[], report_url="https://docs.google.com/x")
AR.STAR_EMAIL_IN_GMAIL = True
_strip = lambda ps: [(t, f, re.sub(rb"Generated [^<]*", b"", b)) for t, f, b in ps]
check("STAR off → no IMAP; subject, body and attachment identical with it on",
      (LOGINS, _strip(payloads(SENT[0][2])) == _strip(with_star)), ([], True))
reset(); STATE["smtp_fail"] = True
try:
    with redirect_stdout(io.StringIO()):
        AR.send_report("daily", "04-Oct-2026", "f.xlsx", b"x", 1, 1, 0, 100.0)
    raised = False
except smtplib.SMTPAuthenticationError:
    raised = True
STATE["smtp_fail"] = False
check("send failure behaves as before (raises) and nothing is starred", (raised, LOGINS), (True, []))
reset(); STATE["imap_fail"] = True
out = io.StringIO()
with redirect_stdout(out):
    AR.send_report("weekly", "28-Sep-2026 to 04-Oct-2026", "f.xlsx", b"x", 1, 1, 0, 100.0)
STATE["imap_fail"] = False
check("star failure: e-mail still sent, one warning, no exception",
      (len(SENT), "[Email] ★ not starred" in out.getvalue(), starred(INFO)), (1, True, []))

# =============================================================================
print("\n== 2. Assignment reminder run: staff summary e-mail ==")
check("STAR_EMAIL_IN_GMAIL on by default", ASG.STAR_EMAIL_IN_GMAIL, True)
TODAY = date(2026, 10, 5)
df = pd.DataFrame([{"student_name": "Learner One", "student_email": "learner.one@example.com",
                    "student_phone": "+919000000000", "class_name": "SQL",
                    "class_subject": "01-Sep-2026 To Current Date", "assessment_title": "SQL Joins",
                    "assessment_id": "A1", "maximum_marks": "10",
                    "submission_deadline": "06/10/2026 23:45:00 IST",
                    "submission_start_date": "01/10/2026 10:00:00 IST",
                    "submission_status": "Not Submitted"}])
cons = ASG.consolidate_by_student(ASG._dedup_records(ASG.find_pending_reminders(df, TODAY)))
check("synthetic reminder built", len(cons), 1)
reset()
with redirect_stdout(io.StringIO()):
    ok_student = ASG.send_reminder_email(cons[0])
check("student reminder sent but NOT starred (info@ is not a recipient)",
      (ok_student, LOGINS, starred(INFO), BOXES["learner.one@example.com"][0]["flags"]),
      (True, [], [], set()))
reset()
out = io.StringIO()
with redirect_stdout(out):
    ok = ASG.send_staff_summary_email(cons, cons, [], [], TODAY,
                                      extra_pdfs=[{"filename": "Email_Status.pdf", "pdf_bytes": b"%PDF-1.4",
                                                  "course": "(all courses)",
                                                  "assignment": "Email Status Report",
                                                  "student_count": 1,
                                                  "reminder_counts": {}}])
ssubj = "[IntelliBI] Reminder Run — 05-Oct-2026 • 1/1 sent"
check("staff summary sent to the existing recipients", (ok, SENT[0][1]), (True, ASG.STAFF_SUMMARY_RECIPIENTS))
check("staff summary starred in the info@ mailbox", starred(INFO), [ssubj])
check("other recipient receives it, not starred", ([x["subject"] for x in BOXES[HR]], starred(HR)), ([ssubj], []))
check("attachment unchanged", [f for _t, f, _b in payloads(SENT[0][2]) if f], ["Email_Status.pdf"])
check("only info@ opened for starring", sorted(set(LOGINS)), [INFO])
reset(); STATE["smtp_fail"] = True
with redirect_stdout(io.StringIO()):
    ok = ASG.send_staff_summary_email(cons, cons, [], [], TODAY)
STATE["smtp_fail"] = False
check("send failure → False as before, no star attempt", (ok, LOGINS), (False, []))
reset(); STATE["imap_fail"] = True
out = io.StringIO()
with redirect_stdout(out):
    ok = ASG.send_staff_summary_email(cons, cons, [], [], TODAY)
STATE["imap_fail"] = False
check("star failure → still True, one warning", (ok, "[Email] ★ not starred" in out.getvalue()), (True, True))
reset(); ASG.STAR_EMAIL_IN_GMAIL = False
with redirect_stdout(io.StringIO()):
    ASG.send_staff_summary_email(cons, cons, [], [], TODAY)
ASG.STAR_EMAIL_IN_GMAIL = True
check("STAR off → sent, no IMAP", (len(SENT), LOGINS), (1, []))
reset()
with redirect_stdout(io.StringIO()):
    ASG.send_staff_summary_email(cons, cons, [], [], TODAY, dry_run=True)
check("dry run → nothing sent, nothing starred", (SENT, LOGINS), ([], []))

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
