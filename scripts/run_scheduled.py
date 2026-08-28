#!/usr/bin/env python3
"""
================================================================================
  Scheduled-run wrapper  (scripts/run_scheduled.py)
  ------------------------------------------------------------------------------
  The entry point Windows Task Scheduler fires. It wraps the project's
  scripts/run_all.py with two guarantees:

    * OVERLAP PROTECTION — an OS advisory file lock ensures a new trigger never
      starts while a previous run of the SAME project is still in progress. If
      the lock is held, this trigger logs and exits without launching a second
      instance.

    * ONCE-PER-DAY SUCCESS GATE (--once-per-day) — used by Operations. Before
      running, it checks a per-day success marker. If today already succeeded,
      the trigger exits immediately (so the 11:00 / 12:00 fallback windows do
      NOT run a second time). It records success only when run_all exits 0, so a
      failed early window is retried by the next fallback window.

  run_all.py itself sends the detailed summary e-mail on every real execution,
  so you receive an e-mail per actual run (success or failure); skipped triggers
  are logged but do not e-mail.

  Usage:
      python scripts/run_scheduled.py --label sales
      python scripts/run_scheduled.py --label ops --once-per-day

  Exit code mirrors run_all (0 = success). A skipped trigger exits 0.
================================================================================
"""
import os
import sys

# ── portable bootstrap: common/ on path, env + config seeded ─────────────────
sys.path.insert(0, os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "common"))
import _bootstrap            # noqa: E402
import paths                 # noqa: E402
import logging_utils         # noqa: E402

import argparse              # noqa: E402
import subprocess           # noqa: E402
from datetime import datetime  # noqa: E402


class FileLock:
    """Cross-platform non-blocking advisory lock on a file handle.

    The lock is held by the OS for the lifetime of this process and released
    automatically if the process dies — so a crashed run never leaves a stale
    lock that blocks tomorrow's schedule.
    """

    def __init__(self, path):
        self.path = str(path)
        self.fh = None

    def acquire(self) -> bool:
        self.fh = open(self.path, "a+")
        try:
            self.fh.seek(0)
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except OSError:
            self.fh.close()
            self.fh = None
            return False

    def release(self):
        if not self.fh:
            return
        try:
            self.fh.seek(0)
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        finally:
            self.fh.close()
            self.fh = None


def _sched_dir():
    d = paths.CACHE_DIR / "scheduler"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _marker_path(label):
    return _sched_dir() / f"{label}_last_success.txt"


def _already_succeeded_today(label) -> bool:
    p = _marker_path(label)
    if not p.exists():
        return False
    try:
        return p.read_text(encoding="utf-8").strip() == datetime.now().strftime("%Y-%m-%d")
    except Exception:
        return False


def _mark_success_today(label):
    try:
        _marker_path(label).write_text(datetime.now().strftime("%Y-%m-%d"), encoding="utf-8")
    except Exception:
        pass


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--label", required=True, help="short id for the lock/marker (e.g. sales, ops)")
    ap.add_argument("--once-per-day", action="store_true",
                    help="skip if the pipeline already succeeded today (retry-window mode)")
    args = ap.parse_args()

    log = logging_utils.get_logger("run_scheduled")
    today = datetime.now().strftime("%Y-%m-%d")
    logging_utils.section(log, f"SCHEDULED TRIGGER  label={args.label}  "
                               f"once_per_day={args.once_per_day}  {today}")

    # 1. once-per-day success gate (Operations fallback windows)
    if args.once_per_day and _already_succeeded_today(args.label):
        log.info("[%s] already completed successfully today (%s) — skipping this "
                 "fallback trigger.", args.label, today)
        return 0

    # 2. overlap protection
    lock = FileLock(_sched_dir() / f"{args.label}.lock")
    if not lock.acquire():
        log.warning("[%s] a previous run is still in progress — skipping this "
                    "trigger (no overlap).", args.label)
        return 0

    try:
        run_all = paths.SCRIPTS_DIR / "run_all.py"
        log.info("[%s] launching %s", args.label, run_all)
        rc = subprocess.call([sys.executable, str(run_all)], cwd=str(paths.PROJECT_ROOT))
        if rc == 0:
            log.info("[%s] pipeline SUCCEEDED.", args.label)
            if args.once_per_day:
                _mark_success_today(args.label)
                log.info("[%s] marked %s as complete — later windows will skip.",
                         args.label, today)
        else:
            msg = f"[{args.label}] pipeline FAILED (rc={rc})."
            if args.once_per_day:
                msg += " The next fallback window will retry."
            log.error(msg)
        return rc
    finally:
        lock.release()


if __name__ == "__main__":
    sys.exit(main())
