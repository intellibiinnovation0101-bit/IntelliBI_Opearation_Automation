"""
================================================================================
  IntelliBI Operations Automation — configuration loader (common/config_loader.py)
  ------------------------------------------------------------------------------
  ONE centralized configuration mechanism. Reads config/config.yaml and:
    * exposes it as a plain dict via  load()  /  get("a.b.c", default)
    * maps operator settings onto the environment variables the existing
      scripts already understand (so no business logic changed), via
      apply_to_environment().

  Precedence (highest first):
    1. A variable already set in the OS environment (operator ad-hoc run)
    2. A value in config/config.yaml
    3. The project-root path defaults from common/_bootstrap.py

  YAML is optional: if PyYAML or the file is missing, the pipeline runs on the
  path defaults and nothing raises.
================================================================================
"""
from __future__ import annotations

import os
from pathlib import Path

import paths  # common/ on sys.path via _bootstrap

CONFIG_YAML = paths.CONFIG_DIR / "config.yaml"

# config.yaml dotted-key  ->  environment variable the scripts read
_KEY_TO_ENV = {
    "google.service_account_file":          "GOOGLE_SERVICE_ACCOUNT_FILE",
    "batch_planner.base_dir":               "BP_BASE_DIR",
}

# env vars whose value names a file/folder -> resolve inside the project
_PATH_VALUED_ENV = {
    "GOOGLE_SERVICE_ACCOUNT_FILE": paths.CREDENTIALS_DIR,
    "BP_BASE_DIR": paths.PROJECT_ROOT,
}

_cache = None


def load() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    data = {}
    if CONFIG_YAML.exists():
        try:
            import yaml
            with open(CONFIG_YAML, "r", encoding="utf-8") as fh:
                data = yaml.safe_load(fh) or {}
        except Exception as e:
            import sys
            sys.stderr.write(f"[config_loader] could not read config.yaml: {e}\n")
            data = {}
    _cache = data if isinstance(data, dict) else {}
    return _cache


def get(dotted: str, default=None):
    node = load()
    for part in dotted.split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return default
    return node


def _resolve_pathish(env_key: str, value: str) -> str:
    base = _PATH_VALUED_ENV.get(env_key)
    if base is None:
        return value
    p = Path(str(value)).expanduser()
    return str(p if p.is_absolute() else (base / str(value)))


def apply_to_environment(original_keys=None) -> None:
    original_keys = original_keys or set()
    cfg = load()
    if not cfg:
        return
    for dotted, env_key in _KEY_TO_ENV.items():
        if env_key in original_keys:
            continue
        val = get(dotted, None)
        if val is None or val == "":
            continue
        os.environ[env_key] = _resolve_pathish(env_key, val)
