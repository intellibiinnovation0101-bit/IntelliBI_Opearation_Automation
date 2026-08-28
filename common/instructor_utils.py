"""
================================================================================
  IntelliBI Operations Automation — instructor mapping (common/instructor_utils.py)
  ------------------------------------------------------------------------------
  Reusable helper for the class -> instructor mapping used by the Student Info
  and Session Attendance pipelines. The mapping is fetched LIVE from the Wise
  API (there is no static instructor JSON), so this helper centralizes that
  fetch for any future consumer.

  The existing pipeline scripts keep their own inline `_fetch_class_instructor_map`
  implementations unchanged (business logic is preserved); this module offers the
  same capability as a shared, documented utility for new code to build on.

  Usage:
      from wise_config import HEADERS            # credentials/wise_config.py
      from instructor_utils import fetch_class_instructor_map
      m = fetch_class_instructor_map(HEADERS, institute_id, base_url)
      instructor = m.get(str(class_id), "")
================================================================================
"""
from __future__ import annotations

from typing import Dict


def _teacher_names(session: dict) -> str:
    """Extract a display instructor name from a Wise 'session'/'class' record."""
    for key in ("teacherName", "instructorName", "teacher_name", "instructor"):
        v = session.get(key)
        if v:
            return str(v).strip()
    teachers = session.get("teachers") or session.get("instructors") or []
    names = []
    for t in teachers:
        if isinstance(t, dict):
            n = t.get("name") or t.get("fullName") or t.get("teacherName")
            if n:
                names.append(str(n).strip())
        elif t:
            names.append(str(t).strip())
    return ", ".join(dict.fromkeys(names))


def fetch_class_instructor_map(headers: dict, institute_id: str,
                               base_url: str = "https://api.wise.live",
                               timeout: int = 60) -> Dict[str, str]:
    """Return {class_id: instructor_name} for an institute via the Wise API.

    Never raises — returns whatever it could build (or {} on total failure) so
    instructor loading never blocks a pipeline. Kept deliberately generic; the
    exact endpoint/pagination live in the calling scripts that need specifics.
    """
    import requests
    out: Dict[str, str] = {}
    try:
        url = f"{base_url}/institutes/{institute_id}/sessions"
        resp = requests.get(url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        payload = resp.json()
        rows = payload.get("sessions") or payload.get("data") or payload or []
        for s in rows:
            if not isinstance(s, dict):
                continue
            cid = s.get("classId") or s.get("class_id") or s.get("id")
            if cid is None:
                continue
            name = _teacher_names(s)
            if name:
                out[str(cid)] = name
    except Exception:
        return out
    return out
