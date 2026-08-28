"""
================================================================================
  IntelliBI Sales Automation — centralized logging  (common/logging_utils.py)
  ------------------------------------------------------------------------------
  Every script and runner logs through get_logger(); output goes to BOTH the
  console and a file under  logs/  whose location is derived from PROJECT_ROOT
  (never hard-coded).  Log lines carry timestamp, level, and logger name so a
  returning operator can reconstruct exactly what each layer did.

  Helpers:
    get_logger(name)        -> a configured logging.Logger (console + file)
    log_file_for(name)      -> the Path of that logger's file
    read_tail(path, n)      -> last n lines of a log file (for the email body)
    section(logger, title)  -> log a visible banner line
================================================================================
"""
from __future__ import annotations

import logging
import os
from pathlib import Path

import paths  # common/ on sys.path via _bootstrap

_LEVEL = os.environ.get("SALES_LOG_LEVEL", "INFO").upper()
_FMT = "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s"
_DATEFMT = "%Y-%m-%d %H:%M:%S"

# One file per logger name, appended across runs (rotate by hand / Task Scheduler).
_configured: dict[str, logging.Logger] = {}


def log_file_for(name: str) -> Path:
    paths.LOGS_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c if (c.isalnum() or c in "._-") else "_" for c in name)
    return paths.LOGS_DIR / f"{safe}.log"


def get_logger(name: str) -> logging.Logger:
    if name in _configured:
        return _configured[name]
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, _LEVEL, logging.INFO))
    logger.propagate = False

    fmt = logging.Formatter(_FMT, datefmt=_DATEFMT)

    fh = logging.FileHandler(log_file_for(name), encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    _configured[name] = logger
    return logger


def section(logger: logging.Logger, title: str) -> None:
    bar = "=" * 72
    logger.info(bar)
    logger.info(title)
    logger.info(bar)


def read_tail(path, n: int = 60) -> str:
    """Return the last n lines of a text file (safe if missing)."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            lines = fh.readlines()
        return "".join(lines[-n:]).rstrip()
    except Exception as e:
        return f"(could not read {path}: {e})"
