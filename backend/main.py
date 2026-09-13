"""
main.py — FastAPI application entry point.

Endpoints:
  GET  /api/health              — liveness check
  GET  /api/contractors         — full contractor list (from cache)
  POST /api/refresh             — re-ingest from Box
  POST /api/refresh-taxonomy    — reload taxonomy from Box
  GET  /api/jrs/{jrs_value}     — validate a single JRS string
  POST /api/audit               — append an audit record
  GET  /api/audit               — retrieve audit log

Run locally:
  cd backend
  uvicorn main:app --reload --port 8000
"""
from datetime import datetime
from typing import Any

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from config import CORS_ORIGIN, SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT
from box_client import get_client
import ingestion
import jrs_validation

# ── App setup ─────────────────────────────────────────────────────────────────

app = FastAPI(
    title="TRAM Subcontractor Renewal & Offboarding API",
    version="1.0.0",
    docs_url="/api/docs",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[CORS_ORIGIN, "http://localhost:3000", "http://127.0.0.1:5500"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── In-memory state (replaced on refresh) ─────────────────────────────────────
_ingestion_result: dict | None = None
_audit_log: list[dict] = []


# ── Pydantic models ───────────────────────────────────────────────────────────

class AuditRecord(BaseModel):
    user:       str
    action:     str
    contractor: str | None = None
    detail:     str | None = None
    outcome:    str = "success"


# ── Health ────────────────────────────────────────────────────────────────────

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "data_loaded": _ingestion_result is not None,
        "taxonomy_loaded": jrs_validation.get_taxonomy_cache() is not None,
        "timestamp": datetime.utcnow().isoformat() + "Z",
    }


# ── Contractors ───────────────────────────────────────────────────────────────

@app.get("/api/contractors")
def get_contractors():
    """Return the full ingested contractor list plus metadata."""
    if _ingestion_result is None:
        raise HTTPException(
            status_code=503,
            detail="Data not yet loaded. Call POST /api/refresh first.",
        )
    return {
        "meta": {
            "filename":       _ingestion_result["filename"],
            "report_date":    _ingestion_result["report_date"],
            "ingested_at":    _ingestion_result["ingested_at"],
            "staleness_days": _ingestion_result["staleness_days"],
            "is_stale":       _ingestion_result["is_stale"],
            "missing_headers": _ingestion_result["missing_headers"],
            "total_accepted": len(_ingestion_result["contractors"]),
            "total_dq":       len(_ingestion_result["dq_exceptions"]),
        },
        "contractors":   _ingestion_result["contractors"],
        "dq_exceptions": _ingestion_result["dq_exceptions"],
    }


# ── Refresh from Box ──────────────────────────────────────────────────────────

@app.post("/api/refresh")
def refresh_data(background_tasks: BackgroundTasks):
    """
    Re-ingest contractor data from Box.
    Runs in the foreground for simplicity (file is small enough).
    """
    global _ingestion_result

    try:
        client = get_client()
        result = ingestion.ingest_from_box(client)
        _ingestion_result = result

        # Run JRS validation on all contractors
        _run_jrs_batch()

        _append_audit(
            user="system",
            action="INGESTION_COMPLETED",
            detail=(
                f"File: {result['filename']} · "
                f"{len(result['contractors'])} rows accepted · "
                f"{len(result['dq_exceptions'])} rejected · "
                f"{len(result['missing_headers'])} header warnings"
            ),
            outcome="success",
        )

        return {
            "status": "ok",
            "filename": result["filename"],
            "accepted": len(result["contractors"]),
            "dq":       len(result["dq_exceptions"]),
        }

    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ingestion failed: {str(e)}")


@app.post("/api/refresh-taxonomy")
def refresh_taxonomy():
    """Reload the JRS taxonomy from Box."""
    try:
        client = get_client()
        jrs_validation.load_taxonomy(client)

        if _ingestion_result:
            _run_jrs_batch()

        _append_audit(
            user="system",
            action="TAXONOMY_REFRESHED",
            detail="Taxonomy reloaded from Box.",
            outcome="success",
        )

        return {"status": "ok", "message": "Taxonomy loaded successfully."}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Taxonomy refresh failed: {str(e)}")


# ── JRS Validation ────────────────────────────────────────────────────────────

@app.get("/api/jrs/{jrs_value}")
def validate_jrs(jrs_value: str):
    """Validate a single JRS value against the loaded taxonomy."""
    result = jrs_validation.validate_jrs(jrs_value)
    return result


# ── Sector email routing ──────────────────────────────────────────────────────

@app.get("/api/sector-email/{sector}")
def get_sector_email(sector: str):
    """Return the CSP email address for a given sector."""
    key = sector.strip().lower()
    email = SECTOR_EMAIL_MAP.get(key, SECTOR_EMAIL_DEFAULT)
    return {"sector": sector, "email": email}


# ── Audit log ─────────────────────────────────────────────────────────────────

@app.post("/api/audit")
def post_audit(record: AuditRecord):
    """Append a user-triggered audit event."""
    _append_audit(
        user=record.user,
        action=record.action,
        contractor=record.contractor,
        detail=record.detail,
        outcome=record.outcome,
    )
    return {"status": "recorded"}


@app.get("/api/audit")
def get_audit(limit: int = 50):
    """Return the most recent audit events (newest first)."""
    return {"events": list(reversed(_audit_log[-limit:]))}


# ── Internal helpers ──────────────────────────────────────────────────────────

def _append_audit(
    user: str,
    action: str,
    contractor: str | None = None,
    detail: str | None = None,
    outcome: str = "success",
):
    _audit_log.append({
        "timestamp":  datetime.utcnow().isoformat() + "Z",
        "user":       user,
        "action":     action,
        "contractor": contractor or "—",
        "detail":     detail or "",
        "outcome":    outcome,
    })


def _run_jrs_batch():
    """Validate JRS for every contractor in the current ingestion result."""
    if not _ingestion_result:
        return
    for c in _ingestion_result["contractors"]:
        jrs_val = c.get("jrsTram")
        if jrs_val:
            result = jrs_validation.validate_jrs(jrs_val)
            c["jrsStatus"]  = result["outcome"]
            c["jrsResult"]  = result
            # Map outcome to the badge status the frontend uses
            if result["outcome"] == "active":
                c["jrsStatus"] = "active"
            elif result["outcome"] in ("one_replacement", "multi_replacement", "deleted"):
                c["jrsStatus"] = "warn"
            else:
                c["jrsStatus"] = "invalid"
        else:
            c["jrsStatus"] = "invalid"
            c["jrsResult"] = {
                "outcome": "not_found",
                "input": "",
                "replacements": [],
                "message": "No JRS value in source record.",
            }
