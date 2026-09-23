"""
ledger.py — Persistent state ledger for app-processed actions.

Problem it solves:
  When a new CMO report is uploaded, it replaces the in-memory contractor data.
  Any state set by the app (offboarding submitted, renewal submitted) would be lost
  unless we persist it separately and re-apply it on every CMO upload.

How it works:
  - A JSON file on disk (/data/ledger.json) records every action taken through the app.
  - Each entry is keyed by contractor serial number.
  - On every CMO upload, _apply_ledger() is called to merge ledger state into the
    freshly parsed contractor records.
  - The ledger is append-only: new CMO data never removes an existing ledger entry.
    Only explicit "clear" actions (future feature) can remove an entry.

Ledger entry schema:
  {
    "serial":         str,           # contractor TalentID / Serial
    "offboard":       str | None,    # e.g. "Yes", "Submitted"
    "offboard_data":  dict | None,   # last_day, reason, laptop, etc.
    "renewal":        str | None,    # e.g. "Yes", "Submitted"
    "renewal_data":   dict | None,   # tram_id_new, end_date, etc.
    "updated_at":     str,           # ISO UTC timestamp
    "updated_by":     str,           # action type e.g. "OFFBOARDING_SUBMITTED"
  }
"""
import json
from datetime import datetime
from pathlib import Path

_LEDGER_PATH: Path | None = None  # set by init()
_ledger: dict[str, dict] = {}     # keyed by serial


def init(data_dir: Path) -> None:
    """Call once at startup with the data directory."""
    global _LEDGER_PATH, _ledger
    _LEDGER_PATH = data_dir / "ledger.json"
    _ledger = _load()


def _load() -> dict:
    if _LEDGER_PATH and _LEDGER_PATH.exists():
        try:
            return json.loads(_LEDGER_PATH.read_text(encoding="utf-8"))
        except Exception as exc:
            print(f"[ledger] Failed to load ledger: {exc}")
    return {}


def _save() -> None:
    if _LEDGER_PATH:
        try:
            _LEDGER_PATH.write_text(
                json.dumps(_ledger, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception as exc:
            print(f"[ledger] Failed to save ledger: {exc}")


def record_offboarding(serial: str, offboard_data: dict) -> None:
    """Called when an offboarding is submitted through the app."""
    entry = _ledger.get(serial, {"serial": serial})
    entry["offboard"]      = "Submitted"
    entry["offboard_data"] = offboard_data
    entry["updated_at"]    = datetime.utcnow().isoformat() + "Z"
    entry["updated_by"]    = "OFFBOARDING_SUBMITTED"
    _ledger[serial] = entry
    _save()


def record_renewal(serial: str, renewal_data: dict) -> None:
    """Called when a renewal is submitted through the app."""
    entry = _ledger.get(serial, {"serial": serial})
    entry["renewal"]      = "Submitted"
    entry["renewal_data"] = renewal_data
    entry["updated_at"]   = datetime.utcnow().isoformat() + "Z"
    entry["updated_by"]   = "RENEWAL_SUBMITTED"
    _ledger[serial] = entry
    _save()


def apply_to_contractors(contractors: list[dict]) -> list[dict]:
    """
    Merge ledger state into a freshly parsed contractor list.
    Ledger wins over CMO for offboard/renewal status fields.
    Contractors in the ledger but NOT in the CMO are re-injected
    as ghost records so they stay visible in the dashboard.
    """
    if not _ledger:
        return contractors

    # Build a lookup by serial for fast merge
    by_serial: dict[str, dict] = {c["serial"]: c for c in contractors if c.get("serial")}

    for serial, entry in _ledger.items():
        if serial in by_serial:
            c = by_serial[serial]
            if entry.get("offboard"):
                c["offboard"] = entry["offboard"]
            if entry.get("renewal"):
                c["renewal"] = entry["renewal"]
            # Store full ledger data on the record for detail panel display
            c["_ledger"] = entry
        else:
            # Contractor no longer in the CMO file — re-inject as a ghost record
            # so historical offboardings/renewals remain visible
            ghost = {
                "serial":      serial,
                "name":        entry.get("offboard_data", {}).get("contractor_name", serial),
                "offboard":    entry.get("offboard", "–"),
                "renewal":     entry.get("renewal", "–"),
                "_ledger":     entry,
                "_ghost":      True,   # flag so UI can style it differently
                # Minimal required fields to avoid None errors
                "client": "–", "project": "–", "vendor": "–", "contact": "–",
                "contactEmail": "–", "band": "–", "jrsTram": "–",
                "jrsStatus": "invalid", "workLocation": "–", "tramId": None,
                "endDate": None, "daysToExpiry": None, "bucket": "unknown", "dq": False,
                "sector": "–", "poNumber": "–", "cnum": None,
            }
            contractors.append(ghost)

    return contractors


def get_all() -> dict:
    """Return the full ledger (for debug/status endpoint)."""
    return dict(_ledger)


def delete_entry(serial: str) -> bool:
    """Remove a single ledger entry by serial. Returns True if it existed."""
    if serial in _ledger:
        del _ledger[serial]
        _save()
        return True
    return False


def clear_all() -> int:
    """Wipe the entire ledger. Returns number of entries removed."""
    global _ledger
    count = len(_ledger)
    _ledger = {}
    _save()
    return count
