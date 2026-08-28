# pyZohoSignatureStatusRefresh.py — Zoho Sign Signature Status Refresh

**Layer 1 — Operations Data Collection** · `ops_data_collection/`

## Purpose
Pulls the people who have **signed** Zoho Sign documents and syncs them into a Google Sheet ("Zoho Sign - Signers") in a Drive folder. This feeds the Admission Formalities report.

## Input sources
- **Zoho Sign** India data centre (`https://sign.zoho.in`) via OAuth refresh token in `credentials/zoho_credentials.json`:
  ```json
  {"data_center":"in","client_id":"...","client_secret":"...","refresh_token":"..."}
  ```
  (scope `ZohoSign.documents.READ`; generate at api-console.zoho.in).

## Output
- Google Sheet **"Zoho Sign - Signers"**, tab **Signers**, in a Drive folder (id cached in `credentials/zoho_sign_spreadsheet.json`). Upsert-by-key (no duplicate signers; status changes update in place).

## Main business logic
- Reuses the project plumbing exactly like the Interakt sync: `utils.get_sheets_service`, `interakt_common.get_or_create_spreadsheet`, `utils.upsert_rows` (auto header + column re-align).
- Re-running never duplicates a signer; a status moving to "completed" is updated in place.

## Important fields / columns
`recipient_name`, `recipient_email`, `recipient_status`, `sent_date`, `signed_date`, `expiry_date`, `document_name`.

## Dependencies
- Project: `common/utils.py`, `common/interakt_common.py`, `credentials/service_account.json`, `credentials/zoho_credentials.json`.
- Packages: `requests`, `google-api-python-client`, `google-auth`.

## Execution flow
Refresh Zoho OAuth token → list signed documents/recipients → get/create signers sheet → upsert. Logs to `logs/zoho_sign_sync.log`.

## Email / report behaviour
None — data-refresh job.

## Configuration used
`credentials/zoho_credentials.json` (tokens), `google.service_account_file`.

## Known issues
- The Zoho refresh token must be valid; regenerate via `get_zoho_refresh_token.py` (kept in the original repo) if it is revoked. See `README_zoho_sign.md` in the original project.

## Future improvements
- Bundle a small token-refresh helper into the Operations project; externalise the Drive folder id into `config.yaml`.

## Change history
- 2026-08-24 — Moved into Operations project; credential/state/log paths made project-root-relative. No business-logic change.

## Email Summary Metrics

The pipeline completion e-mail shows these business KPIs for this script (derived from the script's own run output — no business logic changed). Full technical detail stays in `logs/pyZohoSignatureStatusRefresh.log`.

- **Records fetched** — volume handled this run (context, not a change count)
- **Inserted (new)** — count for the current run
- **Updated** — count for the current run
- **Unchanged / skipped** — already up to date / no change this run
- **Failed** — records/items that errored this run

Zero-valued KPIs are omitted to keep the e-mail concise. The script's line in the e-mail is marked **SUCCESS / FAILED / SKIPPED**; on failure an **Action Required** row shows a short business reason + retry status.
