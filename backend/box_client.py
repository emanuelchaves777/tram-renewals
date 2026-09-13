"""
box_client.py — Box API connection.

Two modes:
  1. Developer token  (BOX_DEV_TOKEN set)   — for local testing only
  2. OAuth 2.0        (client id + secret)  — for production

Usage:
    from box_client import get_client
    client = get_client()
    items = client.folder(FOLDER_ID).get_items()
"""
import boxsdk
from config import BOX_CLIENT_ID, BOX_CLIENT_SECRET, BOX_DEV_TOKEN


def get_client() -> boxsdk.Client:
    """Return an authenticated Box client."""
    if BOX_DEV_TOKEN:
        # ── Local dev: developer token (expires every 60 min) ─────────────────
        auth = boxsdk.OAuth2(
            client_id=BOX_CLIENT_ID,
            client_secret=BOX_CLIENT_SECRET,
            access_token=BOX_DEV_TOKEN,
        )
    else:
        # ── Production: OAuth2 with stored tokens ─────────────────────────────
        # For a shared service account the tokens are stored in a file.
        # In production on Railway/Code Engine use environment variables instead.
        auth = boxsdk.OAuth2(
            client_id=BOX_CLIENT_ID,
            client_secret=BOX_CLIENT_SECRET,
        )

    return boxsdk.Client(auth)
