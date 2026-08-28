"""
================================================================================
  IntelliBI Operations Automation — runtime bootstrap  (common/_bootstrap.py)
  ------------------------------------------------------------------------------
  Imported as the very first project import by every entry script and runner.
  Makes the moved scripts run from any folder on any machine by:

    1. Putting  common/  and  credentials/  on sys.path so the shared modules
       (utils, interakt_common, exotel_common, email_config, wise_config, ...)
       import by name exactly as when everything lived in one flat folder.
    2. Seeding environment-variable defaults (only when unset) for every
       override the scripts honour — service account, batch-planner base dir,
       export/temp locations — all anchored to PROJECT_ROOT.
    3. Loading config/config.yaml and exporting operator settings to the
       environment.
    4. Creating the writable working folders.

  Import side-effect only:   `import _bootstrap`
================================================================================
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# Snapshot operator-set env vars BEFORE seeding defaults, so config.yaml never
# overrides a deliberate command-line/OS override.
_ORIGINAL_ENV_KEYS = set(os.environ)

_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_COMMON = _PROJECT_ROOT / "common"
_CREDENTIALS = _PROJECT_ROOT / "credentials"
_CONFIG = _PROJECT_ROOT / "config"

for _p in (str(_COMMON), str(_CREDENTIALS), str(_PROJECT_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import paths  # noqa: E402  (common/ now on sys.path)

paths.ensure_dirs()


def _setdefault(key, value):
    if value is not None and not os.environ.get(key):
        os.environ[key] = str(value)


# ── path defaults the scripts already understand ─────────────────────────────
_setdefault("GOOGLE_SERVICE_ACCOUNT_FILE", paths.CREDENTIALS_DIR / "service_account.json")
_setdefault("GOOGLE_APPLICATION_CREDENTIALS", os.environ["GOOGLE_SERVICE_ACCOUNT_FILE"])

# Batch Planner reads its source workbooks from BP_BASE_DIR
_setdefault("BP_BASE_DIR", paths.DATA_INPUTS_DIR)

# scratch stays inside the portable project
_setdefault("TMPDIR", paths.TEMP_DIR)
_setdefault("TEMP", paths.TEMP_DIR)
_setdefault("TMP", paths.TEMP_DIR)

# ── optional YAML config ─────────────────────────────────────────────────────
try:
    import config_loader
    config_loader.apply_to_environment(_ORIGINAL_ENV_KEYS)
except Exception as _e:  # never let optional config break a run
    sys.stderr.write(f"[bootstrap] config.yaml not applied: {_e}\n")

PROJECT_ROOT = str(_PROJECT_ROOT)
