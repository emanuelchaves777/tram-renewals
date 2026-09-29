"""
tram_jobs.py — Persistent job queue for TRAM automation via Power Automate Desktop.

Flow:
  1. PM submits renewal in the app → backend creates a "pending" job here.
  2. Power Automate Desktop polls GET /api/tram-jobs/next → picks up the job.
  3. PAD fills TRAM, reads the new TRAM ID from the confirmation page.
  4. PAD calls POST /api/tram-jobs/{job_id}/complete with { tram_id_new: "..." }.
  5. Backend wakes up, generates Excel, sends email.
  6. Frontend polls GET /api/tram-jobs/{job_id} until status = "done" | "failed".

Jobs are persisted to /data/tram_jobs.json so they survive restarts.
"""
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

# ── Storage ────────────────────────────────────────────────────────────────────
_JOBS_FILE: Optional[Path] = None   # set by init()
_jobs: dict[str, dict] = {}         # job_id → job dict


def init(data_dir: Path) -> None:
    """Called once at startup to set the storage path and load persisted jobs."""
    global _JOBS_FILE
    _JOBS_FILE = data_dir / "tram_jobs.json"
    _load()


def _load() -> None:
    if _JOBS_FILE and _JOBS_FILE.exists():
        try:
            raw = json.loads(_JOBS_FILE.read_text(encoding="utf-8"))
            _jobs.clear()
            _jobs.update(raw)
        except Exception:
            pass


def _save() -> None:
    if _JOBS_FILE:
        try:
            _JOBS_FILE.write_text(
                json.dumps(_jobs, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
        except Exception:
            pass


# ── Public API ─────────────────────────────────────────────────────────────────

def create_job(contractor: dict, renewal_fields: dict) -> dict:
    """
    Create a new TRAM automation job.

    Returns the job dict (include job_id to give to the frontend).
    """
    job_id = str(uuid.uuid4())
    job = {
        "job_id":          job_id,
        "status":          "pending",    # pending | claimed | done | failed
        "created_at":      _now(),
        "updated_at":      _now(),
        # Contractor identifiers
        "contractor_id":   contractor.get("serial", ""),
        "contractor_name": contractor.get("name", ""),
        "old_tram_id":     contractor.get("tramId", ""),
        "client":          contractor.get("client", ""),
        "vendor":          contractor.get("vendor", ""),
        "jrs":             contractor.get("jrsTram", ""),
        "band":            contractor.get("band", ""),
        "sector":          contractor.get("sector") or contractor.get("marketSector", ""),
        "work_location":   contractor.get("workLocation", ""),
        "current_end_date": contractor.get("endDate", ""),
        # Fields PM filled in the app — Power Automate will use these to fill TRAM
        "new_start_date":  renewal_fields.get("new_start_date", ""),
        "new_end_date":    renewal_fields.get("new_end_date", ""),
        "rate_cap":        renewal_fields.get("rate_cap", ""),
        "bill_rate":       renewal_fields.get("bill_rate", ""),
        "gp_pct":          renewal_fields.get("gp_pct", ""),
        "confirmed_jrs":   renewal_fields.get("confirmed_jrs", ""),
        "confirmed_band":  renewal_fields.get("confirmed_band", ""),
        "biz_just_1":      renewal_fields.get("biz_just_1", ""),
        "biz_just_2":      renewal_fields.get("biz_just_2", ""),
        "biz_just_3":      renewal_fields.get("biz_just_3", ""),
        # Result — filled in by Power Automate when done
        "tram_id_new":     None,
        "error":           None,
    }
    _jobs[job_id] = job
    _save()
    return job


def get_job(job_id: str) -> Optional[dict]:
    return _jobs.get(job_id)


def get_next_pending() -> Optional[dict]:
    """Return the oldest pending job, or None if queue is empty."""
    pending = [j for j in _jobs.values() if j["status"] == "pending"]
    if not pending:
        return None
    return min(pending, key=lambda j: j["created_at"])


def claim_job(job_id: str) -> bool:
    """Mark a job as claimed by Power Automate (prevents double-pickup)."""
    job = _jobs.get(job_id)
    if not job or job["status"] != "pending":
        return False
    job["status"] = "claimed"
    job["updated_at"] = _now()
    _save()
    return True


def complete_job(job_id: str, tram_id_new: str) -> Optional[dict]:
    """Mark a job as done with the new TRAM ID returned from TRAM."""
    job = _jobs.get(job_id)
    if not job:
        return None
    job["status"]      = "done"
    job["tram_id_new"] = tram_id_new.strip()
    job["updated_at"]  = _now()
    _save()
    return job


def fail_job(job_id: str, error: str) -> Optional[dict]:
    """Mark a job as failed with an error message."""
    job = _jobs.get(job_id)
    if not job:
        return None
    job["status"]     = "failed"
    job["error"]      = error
    job["updated_at"] = _now()
    _save()
    return job


def list_jobs(limit: int = 50) -> list[dict]:
    """Return most recent jobs, newest first."""
    all_jobs = sorted(_jobs.values(), key=lambda j: j["created_at"], reverse=True)
    return all_jobs[:limit]


# ── Helpers ────────────────────────────────────────────────────────────────────

def _now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
