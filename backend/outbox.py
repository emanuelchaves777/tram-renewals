"""
outbox.py — Persistent email outbox.

Used when Microsoft Graph is not configured: emails (with the Excel attached) are
queued here and a small bridge script on an IBMer PC (outlook_mail_bridge.ps1)
polls /api/outbox/next and sends them through the local Outlook.

Storage (under DATA_DIR/outbox/): outbox.json (metadata) + <id>.bin (attachment bytes).
"""
import base64
import json
import threading
import uuid
from datetime import datetime, timedelta
from pathlib import Path

_dir: Path | None = None
_items: dict[str, dict] = {}
_lock = threading.Lock()

CLAIM_TIMEOUT = timedelta(minutes=5)   # a claimed-but-never-confirmed mail becomes claimable again


def init(data_dir: Path) -> None:
    global _dir, _items
    _dir = data_dir / "outbox"
    _dir.mkdir(parents=True, exist_ok=True)
    idx = _dir / "outbox.json"
    if idx.exists():
        try:
            _items = json.loads(idx.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"[outbox] Failed to load index: {exc}")
            _items = {}


def _save() -> None:
    (_dir / "outbox.json").write_text(json.dumps(_items, indent=2), encoding="utf-8")


def _now() -> datetime:
    return datetime.utcnow()


def enqueue(to_email, subject, body, body_type, attachment_filename, attachment_bytes) -> dict:
    if _dir is None:
        raise RuntimeError("Outbox not initialised.")
    mid = uuid.uuid4().hex[:12]
    with _lock:
        if attachment_filename and attachment_bytes:
            (_dir / f"{mid}.bin").write_bytes(attachment_bytes)
        _items[mid] = {
            "id": mid, "to": to_email, "subject": subject, "body": body,
            "body_type": body_type, "attachment_filename": attachment_filename or "",
            "has_attachment": bool(attachment_filename and attachment_bytes),
            "status": "pending", "created_at": _now().isoformat() + "Z",
            "claimed_at": None, "sent_at": None, "error": None,
        }
        _save()
    return {"status": "queued", "to": to_email, "subject": subject,
            "outbox_id": mid, "queued_at": _items[mid]["created_at"]}


def claim_next() -> dict | None:
    """Return the oldest sendable mail (with attachment as base64) and mark it claimed."""
    with _lock:
        now = _now()
        for item in sorted(_items.values(), key=lambda i: i["created_at"]):
            stale = (item["status"] == "claimed" and item["claimed_at"]
                     and now - datetime.fromisoformat(item["claimed_at"]) > CLAIM_TIMEOUT)
            if item["status"] == "pending" or stale:
                item["status"] = "claimed"
                item["claimed_at"] = now.isoformat()
                _save()
                out = dict(item)
                out["attachment_b64"] = ""
                if item["has_attachment"]:
                    out["attachment_b64"] = base64.b64encode(
                        (_dir / f"{item['id']}.bin").read_bytes()).decode("ascii")
                return out
    return None


def mark_sent(mid: str) -> bool:
    with _lock:
        item = _items.get(mid)
        if not item:
            return False
        item["status"], item["sent_at"], item["error"] = "sent", _now().isoformat() + "Z", None
        _save()
        f = _dir / f"{mid}.bin"
        if f.exists():
            f.unlink()   # attachment no longer needed once delivered
        return True


def mark_failed(mid: str, error: str) -> bool:
    with _lock:
        item = _items.get(mid)
        if not item:
            return False
        item["status"], item["error"] = "pending", (error or "")[:500]   # retry on next poll
        _save()
        return True


def list_items() -> list[dict]:
    with _lock:
        return sorted(_items.values(), key=lambda i: i["created_at"], reverse=True)[:100]
