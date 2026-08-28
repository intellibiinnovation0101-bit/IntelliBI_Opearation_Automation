"""
Template for credentials/wise_config.py  (copy this file, remove '.example').

Central Wise API credentials used by the Student Info, Session Attendance and
Assignment Submissions collectors. Rotate the API key in the Wise dashboard
(Settings -> Developer Settings -> API Credentials). This file is git-ignored.
"""
from base64 import b64encode

API_KEY   = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"   # 32-char hex from Wise
USER_ID   = "xxxxxxxxxxxxxxxxxxxxxxxx"
NAMESPACE = "learnwithintellibi"                 # <NAMESPACE>.wise.live

_BASIC = b64encode(f"{USER_ID}:{API_KEY}".encode()).decode()

HEADERS = {
    "Authorization":    f"Basic {_BASIC}",
    "user-agent":       f"VendorIntegrations/{NAMESPACE}",
    "x-api-key":        API_KEY,
    "x-wise-namespace": NAMESPACE,
    "Content-Type":     "application/json",
    "Accept":           "application/json",
}
