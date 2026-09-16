"""
email_sender.py — Sends renewal and offboarding emails via Microsoft Graph API.

Why Microsoft Graph and not SMTP:
  IBM uses Microsoft 365 / Exchange Online. Graph API is the modern,
  supported method for sending email from an IBM Outlook account.
  It requires a one-time Azure App Registration (instructions in DEPLOYMENT.md).

Confirmed email spec (docs/email-integration-spec.md):
  Subject : [TRAM ID] [TalentID/Serial] [PO Number] [Contractor Name] [Client Name] RENEWAL
  Body    : "Can you please help with this renewal?"
  Attach  : the populated .xlsm file
  Routing : determined by Sector → SECTOR_EMAIL_MAP in config.py

Auth flow:
  1. App uses Client Credentials (app-only) OR Delegated (user signs in).
  2. For a shared backend service, Client Credentials is recommended.
  3. Requires Azure AD app with Mail.Send permission granted by IBM IT admin.

Environment variables needed (add to .env and Railway):
  MS_TENANT_ID      — Azure AD tenant ID (from IBM Azure portal)
  MS_CLIENT_ID      — Azure App Registration client ID
  MS_CLIENT_SECRET  — Azure App Registration client secret
  MS_SENDER_EMAIL   — IBM email address to send from (e.g. Emanuel.Chaves@ibm.com)
"""

import base64
import os
import json
from datetime import datetime

import httpx

from config import SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT

# ── Microsoft Graph settings ──────────────────────────────────────────────────
MS_TENANT_ID     = os.getenv("MS_TENANT_ID", "")
MS_CLIENT_ID     = os.getenv("MS_CLIENT_ID", "")
MS_CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET", "")
MS_SENDER_EMAIL  = os.getenv("MS_SENDER_EMAIL", "")

GRAPH_TOKEN_URL  = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
GRAPH_SEND_URL   = "https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
GRAPH_SCOPE      = "https://graph.microsoft.com/.default"


# ── Public entry points ───────────────────────────────────────────────────────

def send_renewal_email(
    contractor: dict,
    tram_id_new: str,
    excel_filename: str,
    excel_bytes: bytes,
) -> dict:
    """
    Send the renewal email to the correct CSP team.

    Subject format (confirmed):
      [TRAM ID] [TalentID/Serial] [PO Number] [Contractor Name] [Client Name] RENEWAL

    Body (confirmed): "Can you please help with this renewal?"
    Attachment: the populated .xlsm file

    Returns:
      {"status": "sent", "to": str, "subject": str, "message_id": str}
      or raises an exception on failure.
    """
    to_email = _get_sector_email(contractor.get("sector", ""))
    subject  = _build_renewal_subject(contractor, tram_id_new)
    body     = "Can you please help with this renewal?"

    return _send(
        to_email=to_email,
        subject=subject,
        body=body,
        attachment_filename=excel_filename,
        attachment_bytes=excel_bytes,
    )


def send_offboarding_email(
    contractor: dict,
    offboard_data: dict,
) -> dict:
    """
    Send the offboarding notification email to the correct CSP team.
    Body is an HTML table with one data row and the 10 confirmed columns.
    """
    to_email = _get_sector_email(contractor.get("sector", ""))
    subject  = _build_offboarding_subject(contractor)
    body     = _build_offboarding_body(contractor, offboard_data)

    return _send(
        to_email=to_email,
        subject=subject,
        body=body,
        body_type="HTML",
        attachment_filename=None,
        attachment_bytes=None,
    )


def is_configured() -> bool:
    """Return True if all Microsoft Graph credentials are set."""
    return all([MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER_EMAIL])


# ── Core send logic ───────────────────────────────────────────────────────────

def _send(
    to_email: str,
    subject: str,
    body: str,
    attachment_filename: str | None,
    attachment_bytes: bytes | None,
    body_type: str = "Text",
) -> dict:
    """Build and send the email via Microsoft Graph."""

    if not is_configured():
        raise RuntimeError(
            "Microsoft Graph credentials not configured. "
            "Set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER_EMAIL "
            "in Railway environment variables."
        )

    token    = _get_access_token()
    payload  = _build_payload(to_email, subject, body, attachment_filename, attachment_bytes, body_type)
    send_url = GRAPH_SEND_URL.format(sender=MS_SENDER_EMAIL)

    resp = httpx.post(
        send_url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/json",
        },
        content=json.dumps(payload),
        timeout=30,
    )

    if resp.status_code not in (200, 202):
        raise RuntimeError(
            f"Graph API returned {resp.status_code}: {resp.text}"
        )

    return {
        "status":     "sent",
        "to":         to_email,
        "subject":    subject,
        "sent_at":    datetime.utcnow().isoformat() + "Z",
        "message_id": resp.headers.get("request-id", "unknown"),
    }


def _get_access_token() -> str:
    """Obtain a Microsoft Graph access token via client credentials flow."""
    token_url = GRAPH_TOKEN_URL.format(tenant=MS_TENANT_ID)

    resp = httpx.post(
        token_url,
        data={
            "grant_type":    "client_credentials",
            "client_id":     MS_CLIENT_ID,
            "client_secret": MS_CLIENT_SECRET,
            "scope":         GRAPH_SCOPE,
        },
        timeout=15,
    )

    if resp.status_code != 200:
        raise RuntimeError(
            f"Failed to obtain Microsoft Graph token: {resp.status_code} {resp.text}"
        )

    return resp.json()["access_token"]


def _build_payload(
    to_email: str,
    subject: str,
    body: str,
    attachment_filename: str | None,
    attachment_bytes: bytes | None,
    body_type: str = "Text",
) -> dict:
    """Build the Microsoft Graph sendMail request payload."""
    payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body,
            },
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ],
        },
        "saveToSentItems": True,
    }

    if attachment_filename and attachment_bytes:
        payload["message"]["attachments"] = [
            {
                "@odata.type":  "#microsoft.graph.fileAttachment",
                "name":         attachment_filename,
                "contentType":  "application/vnd.ms-excel.sheet.macroEnabled.12",
                "contentBytes": base64.b64encode(attachment_bytes).decode("utf-8"),
            }
        ]

    return payload


# ── Subject / body builders ───────────────────────────────────────────────────

def _build_renewal_subject(contractor: dict, tram_id_new: str) -> str:
    """
    Confirmed format:
    [TRAM ID] [TalentID/Serial] [PO Number] [Contractor Name] [Client Name] RENEWAL
    """
    parts = [
        tram_id_new                          or contractor.get("tramId", "–"),
        contractor.get("serial", "–"),
        contractor.get("poNumber", "–"),
        contractor.get("name", "–"),
        contractor.get("client", "–"),
        "RENEWAL",
    ]
    return " ".join(str(p) for p in parts)


def _build_offboarding_subject(contractor: dict) -> str:
    """Same format as renewal but OFFBOARDING suffix."""
    parts = [
        contractor.get("tramId", "–"),
        contractor.get("serial", "–"),
        contractor.get("poNumber", "–"),
        contractor.get("name", "–"),
        contractor.get("client", "–"),
        "OFFBOARDING",
    ]
    return " ".join(str(p) for p in parts)


def _build_offboarding_body(contractor: dict, offboard_data: dict) -> str:
    """
    Build the offboarding email body as a single-row HTML table.

    Columns (confirmed order):
      Sector | Manager Email (PM) | Contractor Name | PO Number |
      Serial Number | Last day (mm/dd/yyyy) | Reason For Termination |
      Laptop (yes/no) | Laptop returned (yes/no) | Comments

    Rules:
      - null / empty values  → "–"
      - Boolean (yes/no str) → "Yes" / "No"
      - Last day date        → MM/DD/YYYY
    """
    def _v(val) -> str:
        return str(val).strip() if val and str(val).strip() else "–"

    def _bool_display(val) -> str:
        if not val:
            return "–"
        return "Yes" if str(val).strip().lower() in ("yes", "true", "1") else "No"

    def _format_date(val) -> str:
        """Convert YYYY-MM-DD (or any common format) to MM/DD/YYYY."""
        raw = str(val).strip() if val else ""
        if not raw:
            return "–"
        # Try ISO format first
        try:
            from datetime import datetime as _dt
            return _dt.strptime(raw, "%Y-%m-%d").strftime("%m/%d/%Y")
        except ValueError:
            pass
        # Already in MM/DD/YYYY or unrecognised — return as-is
        return raw

    CELL  = 'style="border:1px solid #000;padding:8px"'
    HCELL = 'style="border:1px solid #000;padding:8px"'

    columns = [
        ("Sector",                    _v(contractor.get("sector"))),
        ("Manager Email (PM)",        _v(offboard_data.get("manager_email"))),
        ("Contractor Name",           _v(contractor.get("name"))),
        ("PO Number",                 _v(contractor.get("poNumber"))),
        ("Serial Number",             _v(contractor.get("serial"))),
        ("Last day (mm/dd/yyyy)",     _format_date(offboard_data.get("last_day"))),
        ("Reason For Termination",    _v(offboard_data.get("reason"))),
        ("Laptop (yes/no)",           _bool_display(offboard_data.get("laptop"))),
        ("Laptop returned (yes/no)",  _bool_display(offboard_data.get("laptop_returned"))),
        ("Comments",                  _v(offboard_data.get("comments"))),
    ]

    headers = "".join(f"<th {HCELL}>{h}</th>" for h, _ in columns)
    cells   = "".join(f"<td {CELL}>{v}</td>"   for _, v in columns)

    return (
        f'<table style="border-collapse:collapse;width:100%">'
        f"<thead><tr>{headers}</tr></thead>"
        f"<tbody><tr>{cells}</tr></tbody>"
        f"</table>"
    )


def _get_sector_email(sector: str) -> str:
    """Return the CSP email for a given sector string."""
    key = (sector or "").strip().lower()
    return SECTOR_EMAIL_MAP.get(key, SECTOR_EMAIL_DEFAULT)
