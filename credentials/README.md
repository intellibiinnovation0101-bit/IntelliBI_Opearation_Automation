# credentials/ — secrets & session state (never commit)

Everything here is git-ignored (see the project `.gitignore`). Copy the real
files here **once per machine**; the pipeline reads them from here via
`common/paths.py`, so no source-code paths ever change.

## Required

| File | Purpose | How to get it |
|------|---------|---------------|
| `service_account.json` | Google service-account key (Sheets v4 + Drive v3). Account `intellibi-data-pipeline@intellibi-mis.iam.gserviceaccount.com`. | Google Cloud console. Every target sheet/folder must be shared with this account (Editor). |
| `email_config.py` | Gmail SMTP senders + app passwords (primary Operations account + secondary "digital" account). | Copy `email_config.example.py` → `email_config.py`; fill the 16-char app passwords (`myaccount.google.com/apppasswords`). |
| `wise_config.py` | Wise API credentials (`HEADERS`) for the Student Info / Attendance / Assignment collectors. | Copy `wise_config.example.py` → `wise_config.py`; paste the Wise API key/user id. |

## Layer-1 collector credentials

| File | Used by |
|------|---------|
| `zoho_credentials.json` | `pyZohoSignatureStatusRefresh.py` — `{data_center, client_id, client_secret, refresh_token}` from api-console.zoho.in (scope `ZohoSign.documents.READ`). |
| `zoho_sign_spreadsheet.json` | cached target-spreadsheet id for the Zoho signers sheet. |
| `.assignment_report_trigger.json` | runtime trigger written by `pyAssignmentSubmissions.py`, read by `pyAssignmentSubmissionsReport.py` (auto-created; do not commit). |

