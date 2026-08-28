"""
Template for credentials/email_config.py  (copy this file, remove '.example').

Central Gmail SMTP credentials used by the report scripts and the run_all
summary e-mail. Rotate the app password at myaccount.google.com/apppasswords
(2-Step Verification must be ON).  This file is git-ignored.
"""

# Primary account — sends report + pipeline-summary e-mails.
GMAIL_SENDER   = "info@intellibiinnovationstechnologies.in"
GMAIL_APP_PASS = "xxxx xxxx xxxx xxxx"   # 16-char Gmail app password

# Secondary account (only if a report uses it).
GMAIL_SENDER_DIGITAL   = "intellibidigital@gmail.com"
GMAIL_APP_PASS_DIGITAL = "xxxx xxxx xxxx xxxx"
