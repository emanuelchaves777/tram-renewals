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
    Body is an HTML table with all confirmed offboarding fields.
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
    Build the offboarding email body as an HTML table matching the confirmed
    field order: Sector | Manager Email (PM) | Contractor Name | PO Number |
    Serial Number | Last day | Reason For Termination | Laptop | Laptop returned | Comments
    """
    fields = [
        ("Sector",                   contractor.get("sector", "–")),
        ("Manager Email (PM)",       offboard_data.get("manager_email", "–")),
        ("Contractor Name",          contractor.get("name", "–")),
        ("PO Number",                contractor.get("poNumber", "–")),
        ("Serial Number",            contractor.get("serial", "–")),
        ("Last day (mm/dd/yyyy)",    offboard_data.get("last_day", "–")),
        ("Reason For Termination",   offboard_data.get("reason", "–")),
        ("Laptop (yes/no)",          offboard_data.get("laptop", "–")),
        ("Laptop returned (yes/no)", offboard_data.get("laptop_returned", "–")),
        ("Comments",                 offboard_data.get("comments", "–")),
    ]

    # ── HTML version (renders as a table in Outlook) ──────────────────────────
    header_cells = "".join(
        f'<th style="background:#1d4ed8;color:#fff;padding:8px 12px;'
        f'text-align:left;white-space:nowrap;font-size:13px">{h}</th>'
        for h, _ in fields
    )
    value_cells = "".join(
        f'<td style="padding:8px 12px;border-bottom:1px solid #e5e7eb;'
        f'font-size:13px;white-space:nowrap">{v or "–"}</td>'
        for _, v in fields
    )

    html = f"""
<p style="font-family:Segoe UI,Arial,sans-serif;font-size:14px">
  Can you please help with this offboarding?
</p>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;
              margin-top:12px;border:1px solid #e5e7eb">
  <thead><tr>{header_cells}</tr></thead>
  <tbody><tr>{value_cells}</tr></tbody>
</table>
"""
    return html


def _get_sector_email(sector: str) -> str:
    """Return the CSP email for a given sector string."""
    key = (sector or "").strip().lower()
    return SECTOR_EMAIL_MAP.get(key, SECTOR_EMAIL_DEFAULT)
