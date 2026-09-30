"""
================================================================================
  IntelliBI — transient-failure retry for Google API / SMTP calls
  (common/api_retry.py)
  ------------------------------------------------------------------------------
  WHY THIS EXISTS
    The production report logs show the report process dying on errors that
    Google itself classifies as transient — e.g.

        HttpError 500 ... drive/v3/files ... "Internal Error"
        HttpError 503 ... sheets ... "The service is currently unavailable."

    One such response, with no retry, killed the whole run (rc=1) even though
    every other step of the report was fine and the very next scheduled run
    succeeded. This module gives every network call a bounded, back-off retry
    so a momentary Google / network / SMTP hiccup no longer fails a report.

  WHAT IS RETRIED (and only this)
    * HTTP 429 / 500 / 502 / 503 / 504 from a Google API
    * HTTP 403 whose reason is a rate limit (rateLimitExceeded / userRateLimit…)
    * connection-level failures: timeout, connection reset / aborted, broken
      pipe, remote disconnected, incomplete read, SSL EOF, DNS lookup failure
    * SMTP: server disconnected / connect / HELO errors and 4xx (temporary)
      response codes such as 421 / 450 / 451 / 452 / 454

  WHAT IS NEVER RETRIED (fails fast, exactly as before)
    * 400 / 401 / 403 (permission) / 404 — a wrong id, revoked access, a missing
      sheet, an expired credential; retrying cannot fix these
    * storageQuotaExceeded
    * SMTP 5xx (e.g. 535 bad credentials)
    * any non-network exception (a bug, a KeyError, bad data)

  HOW
    call_with_retry(fn, what)      run fn(); on a transient error wait
                                   (10s, 20s, 40s, 80s + jitter) and try again,
                                   up to `attempts` calls in total, then re-raise
                                   the LAST error. Every retry is logged.
    execute(request, what)         call_with_retry(request.execute, what) for a
                                   googleapiclient request object.
    run_steps_with_retry(...)      run a list of named, ORDERED delivery steps
                                   (upload → share → e-mail …) remembering which
                                   steps already completed, so a transient
                                   failure re-runs ONLY the steps still pending.
                                   That is what keeps a retry from producing a
                                   duplicate Drive file or a duplicate e-mail.

  Tunables (config.yaml → pipeline:)
    api_retry_attempts        total calls per operation      (default 5)
    api_retry_base_wait_sec   first wait, doubles each retry (default 10, cap 120)
    delivery_retry_attempts   passes per report delivery     (default 3)
    delivery_retry_wait_sec   wait between passes            (default 90, doubles)
================================================================================
"""
from __future__ import annotations

import random
import socket
import ssl
import time

try:                                   # config.yaml tunables (optional)
    import config_loader as _cfg
except Exception:                       # pragma: no cover - standalone use
    _cfg = None


def _cfg_int(key, default):
    if _cfg is None:
        return default
    try:
        v = _cfg.get(key, default)
        return int(v) if str(v).strip() else default
    except (TypeError, ValueError):
        return default


TRANSIENT_HTTP_STATUS = {429, 500, 502, 503, 504}
_TRANSIENT_TEXT = (
    "internal error", "backenderror", "currently unavailable", "service unavailable",
    "ratelimitexceeded", "userratelimitexceeded", "quota exceeded", "resource_exhausted",
    "timed out", "timeout", "connection reset", "connection aborted", "broken pipe",
    "remote end closed", "remotedisconnected", "incompleteread", "eof occurred",
    "temporarily unavailable", "try again later", "bad gateway", "gateway time",
    "name or service not known", "getaddrinfo failed", "nodename nor servname",
    "max retries exceeded", "unable to find the server",
)
_NEVER_TEXT = ("storagequota", "storage quota")
_SMTP_TEMP_CODES = {421, 450, 451, 452, 454}


def http_status(exc) -> int | None:
    """HTTP status carried by a googleapiclient HttpError (or similar), else None."""
    resp = getattr(exc, "resp", None)
    st = getattr(resp, "status", None)
    if st is None:
        st = getattr(exc, "status_code", None)
    try:
        return int(st) if st is not None else None
    except (TypeError, ValueError):
        return None


def is_transient(exc) -> bool:
    """True when a retry has a real chance of succeeding (see module doc)."""
    text = str(exc).lower()
    if any(t in text for t in _NEVER_TEXT):
        return False

    # --- SMTP -------------------------------------------------------------
    try:
        import smtplib
        if isinstance(exc, (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError,
                            smtplib.SMTPHeloError)):
            return True
        if isinstance(exc, smtplib.SMTPResponseException):
            return int(getattr(exc, "smtp_code", 0) or 0) in _SMTP_TEMP_CODES
        if isinstance(exc, smtplib.SMTPException):
            return False
    except Exception:            # pragma: no cover
        pass

    # --- Google API HTTP errors -------------------------------------------
    st = http_status(exc)
    if st is not None:
        if st in TRANSIENT_HTTP_STATUS:
            return True
        if st == 403:            # only the rate-limit flavour of 403 is transient
            return "ratelimit" in text.replace(" ", "")
        return False

    # --- connection-level problems ----------------------------------------
    if isinstance(exc, (socket.timeout, TimeoutError, ConnectionError, ssl.SSLError)):
        return True
    try:
        import http.client as _hc
        if isinstance(exc, (_hc.RemoteDisconnected, _hc.IncompleteRead,
                            _hc.BadStatusLine)):
            return True
    except Exception:            # pragma: no cover
        pass
    try:
        import httplib2
        if isinstance(exc, (httplib2.ServerNotFoundError, httplib2.HttpLib2Error)):
            return True
    except Exception:
        pass
    if isinstance(exc, OSError) and getattr(exc, "errno", None) in (
            104, 110, 111, 10053, 10054, 10060, 10061):
        return True

    # last resort: recognised transient wording from any wrapper exception
    return any(t in text for t in _TRANSIENT_TEXT)


def short_error(exc, limit: int = 220) -> str:
    """One-line, log-friendly description of an exception."""
    msg = " ".join(str(exc).split())
    st = http_status(exc)
    head = f"{type(exc).__name__}" + (f" HTTP {st}" if st else "")
    out = f"{head}: {msg}" if msg else head
    return out if len(out) <= limit else out[:limit - 1] + "…"


def _sleep(seconds: float):
    time.sleep(seconds)


def call_with_retry(fn, what: str = "call", attempts: int | None = None,
                    base_wait: float | None = None, max_wait: float = 120.0,
                    log=print, on_retry=None):
    """Run fn() and retry it on TRANSIENT errors with exponential back-off.

    attempts   total number of calls (default pipeline.api_retry_attempts = 5)
    base_wait  first wait in seconds; doubles each retry, capped at max_wait
               (default pipeline.api_retry_base_wait_sec = 10)
    on_retry   optional callback(attempt_no, exc) run BEFORE each retry — used by
               upload_report_to_drive to check whether the failed call actually
               succeeded server-side (so a retry never duplicates a file). If it
               returns a non-None value, that value is returned instead of
               retrying.
    Non-transient errors are re-raised immediately, untouched.
    """
    if attempts is None:
        attempts = max(1, _cfg_int("pipeline.api_retry_attempts", 5))
    if base_wait is None:
        base_wait = max(1, _cfg_int("pipeline.api_retry_base_wait_sec", 10))
    last = None
    for n in range(1, attempts + 1):
        try:
            return fn()
        except Exception as exc:                  # noqa: BLE001 — classified below
            last = exc
            if not is_transient(exc) or n >= attempts:
                raise
            if on_retry is not None:
                recovered = on_retry(n, exc)
                if recovered is not None:
                    log(f"  [retry] {what}: earlier attempt had actually succeeded — reusing it.")
                    return recovered
            wait = min(max_wait, base_wait * (2 ** (n - 1))) + random.uniform(0, 3)
            log(f"  [retry] {what}: transient error ({short_error(exc)}) — "
                f"retry {n}/{attempts - 1} in {wait:.0f}s")
            _sleep(wait)
    raise last                                    # pragma: no cover (defensive)


def execute(request, what: str = "Google API call", **kw):
    """call_with_retry() for a googleapiclient request object."""
    return call_with_retry(request.execute, what, **kw)


class DeliveryFailed(RuntimeError):
    """A delivery step failed for good (after every retry). `.step` names it."""

    def __init__(self, step: str, exc: Exception):
        super().__init__(f"step '{step}' failed: {short_error(exc)}")
        self.step = step
        self.cause = exc


def run_steps_with_retry(label: str, steps, state: dict | None = None,
                         attempts: int | None = None, base_wait: float | None = None,
                         log=print) -> dict:
    """Execute ordered delivery steps [(name, fn), …] with pass-level retry.

    * Each fn() is called with no arguments and its return value stored in
      state[name]. A step already present in `state` is SKIPPED — that is the
      idempotency guarantee: after a transient failure only the steps that did
      not complete are run again, so nothing is uploaded or e-mailed twice.
    * On a transient error the whole pass is retried after a wait (default
      pipeline.delivery_retry_wait_sec = 90s, doubling), up to
      pipeline.delivery_retry_attempts passes (default 3).
    * A NON-transient error, or exhausted retries, raises DeliveryFailed naming
      the step and the root error, so the caller can log it and move on to the
      next report instead of aborting the whole process.
    Returns the (completed) state dict.
    """
    state = state if state is not None else {}
    if attempts is None:
        attempts = max(1, _cfg_int("pipeline.delivery_retry_attempts", 3))
    if base_wait is None:
        base_wait = max(1, _cfg_int("pipeline.delivery_retry_wait_sec", 90))
    for n in range(1, attempts + 1):
        try:
            for name, fn in steps:
                if name in state:
                    continue
                state[name] = fn()
            return state
        except DeliveryFailed:
            raise
        except Exception as exc:                  # noqa: BLE001
            failed_step = next((nm for nm, _ in steps if nm not in state), "?")
            if not is_transient(exc) or n >= attempts:
                raise DeliveryFailed(failed_step, exc) from exc
            wait = min(600, base_wait * (2 ** (n - 1))) + random.uniform(0, 5)
            done = [nm for nm, _ in steps if nm in state]
            log(f"  [retry] {label}: step '{failed_step}' hit a transient error "
                f"({short_error(exc)}); completed so far: {done or 'none'} — "
                f"retrying the remaining steps in {wait:.0f}s (pass {n}/{attempts - 1})")
            _sleep(wait)
    return state                                  # pragma: no cover (defensive)
