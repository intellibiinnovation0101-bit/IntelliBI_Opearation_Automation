"""
================================================================================
  IntelliBI Operations Automation — Google Sheets helpers (common/google_sheet_utils.py)
  ------------------------------------------------------------------------------
  A single, documented entry point for the project's shared Google Sheets /
  service-account plumbing. It re-exports the helpers implemented in
  ``common/utils.py`` — the module the existing pipeline scripts already import —
  so new code can depend on a clearly-named facade without duplicating logic:

        from google_sheet_utils import get_sheets_service, upsert_rows

  The authoritative implementation stays in ``utils.py`` (unchanged) to preserve
  every existing behaviour; this module only gives it a discoverable name in
  ``common/``. Existing scripts keep importing ``utils`` directly.
================================================================================
"""
from __future__ import annotations

import utils  # common/ is on sys.path via _bootstrap
from utils import *  # noqa: F401,F403  (re-export the shared helpers)

# Commonly used helpers, surfaced explicitly for editor discoverability.
get_sheets_service = utils.get_sheets_service
upsert_rows = getattr(utils, "upsert_rows", None)
overwrite_rows = getattr(utils, "overwrite_rows", None)
clean_sheet = getattr(utils, "clean_sheet", None)
sort_sheet_by_column = getattr(utils, "sort_sheet_by_column", None)
upsert_scd2 = getattr(utils, "upsert_scd2", None)
col_letter = getattr(utils, "col_letter", None)


def service():
    """Convenience: return an authenticated service-account Sheets client."""
    return utils.get_sheets_service()
