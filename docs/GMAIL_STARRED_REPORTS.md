# Starred (★) report e-mails in Gmail

These report e-mails are starred automatically once they are sent:

| Project | Script | Subject |
|---|---|---|
| Operations | `co-ordinator reports/pyCoordinatorTaskPerformanceReport.py` | `<Type> Coordinator Task Performance Report - <period>` |
| Operations | `co-ordinator reports/pyCoordinatorTaskListReport.py` | `<Type> Batch Coordinator Report - <period>` |
| Sales | `sales_reports/pyConsolidatedLeadPerformanceReport.py` | `<Type> Lead Report - <period>` |
| Sales | `sales_reports/pyLeadFollowUpAnalysisReport.py` | `<Type> Lead Follow-Up Analysis Report - <period>` |

## How it works
- The shared module is `common/gmail_star.py`. The same file is kept in both projects.
- Each script has `STAR_EMAIL_IN_GMAIL = True`, next to `SEND_EMAIL`. Set it to `False` to stop starring; sending is unaffected.
- **Before sending:** the e-mail gets its own unique `Message-ID` header, a standard header Gmail would otherwise add itself. Subject, body, sender, recipients and attachments are unchanged.
- **After Gmail has accepted it:** the script logs in to the **same Gmail account** over IMAP (`imap.gmail.com`), using the same address and app password as the send, from `credentials/email_config.py`. It finds that exact message in All Mail and sets IMAP `\Flagged`, which is Gmail's **Starred**. The sender (`info@intellibiinnovationstechnologies.in`) is also a recipient, so the message in its **Inbox** shows ★.
- **Failures never block a run:** starring is best-effort. If IMAP is disabled, refused or unreachable, one warning `[Email] ★ not starred — …` is printed. The e-mail, the run summary and the exit code are unchanged, and the rest of that run skips starring instead of waiting again.
- Other e-mails are never starred, including the morning and evening batch summaries.

## Other recipients' mailboxes
A star belongs to **one mailbox**. The sending account can star its own copy. No e-mail header can make a message arrive starred in somebody else's Gmail, and the other recipients are personal `@gmail.com` accounts. Each of them can star these reports automatically with a one-time Gmail filter:

1. In Gmail, open the search-options arrow in the search bar.
2. **From:** `info@intellibiinnovationstechnologies.in`
3. **Subject:** `"Coordinator Task Performance Report" OR "Batch Coordinator Report" OR "Lead Report" OR "Lead Follow-Up Analysis Report"`
4. Click **Create filter**, tick **Star it** (optionally also **Never send it to Spam**), and click **Create filter**.

## One-time requirement
IMAP must be allowed for the sending account. For a personal Gmail account, IMAP is always on. For a Google Workspace account, the administrator can switch it off: Admin console ▸ Apps ▸ Google Workspace ▸ Gmail ▸ End User Access ▸ POP and IMAP access. The existing app password covers IMAP; no new credential is needed.

## Checking it
- `python ops_validation\verify_gmail_star.py` (Operations) and `python sales_validation\verify_gmail_star.py` (Sales) test everything offline against an in-memory Gmail.
- Live, read-only: `python common\gmail_star.py --check` lists the last 2 days of report e-mails in the sending mailbox with ★ / ☆. Add `--days 7` for a longer period.
