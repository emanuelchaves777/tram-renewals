"""
main.py — FastAPI application entry point.

Endpoints:
  GET  /api/health              — liveness check
  GET  /api/contractors         — full contractor list (from cache)
  POST /api/refresh             — re-ingest from Box
  POST /api/refresh-taxonomy    — reload taxonomy from Box
  GET  /api/jrs/{jrs_value}     — validate a single JRS string
  POST /api/submit-renewal      — ONE-SHOT: validate JRS + generate Excel + send email
  POST /api/submit-offboarding  — ONE-SHOT: generate offboarding email
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
from fastapi.responses import Response
from pydantic import BaseModel

from config import CORS_ORIGIN, SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT, BOX_FOLDER_ID
from box_client import get_client
import ingestion
import jrs_validation
import excel_generation
import email_sender

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


class ExcelRequest(BaseModel):
    contractor_id: str          # serial / TalentID to look up in ingestion cache
    pm_checklist:  dict         # all PM checklist fields from Step 2 of renewal workflow


class SubmitRenewalRequest(BaseModel):
    """
    Single-shot renewal submission.
    The PM fills in the review form; the backend does everything else automatically:
      1. JRS validation
      2. Excel generation (downloads template from Box, populates all fields)
      3. Email send via Microsoft Graph
      4. Audit write
    """
    contractor_id:   str
    # Fields the PM reviewed / corrected in the UI
    new_end_date:    str
    new_start_date:  str
    confirmed_jrs:   str
    confirmed_band:  str
    work_location:   str
    rate_cap:        str
    bill_rate:       str
    gp_pct:          str
    contract_type:   str
    niche_skills:    str
    biz_just_1:      str
    biz_just_2:      str = ""
    biz_just_3:      str = ""
    biz_just_4:      str = ""
    us_citizenship:  str = ""
    security_access: str = ""
    pen_testing:     str = ""
    requires_laptop: str = "No"
    laptop_os:       str = ""
    laptop_address:  str = ""
    contractor_phone:str = ""
    # TRAM ID the PM enters after submitting in TRAM (may be empty on first pass)
    tram_id_new:     str = ""


class SubmitOffboardingRequest(BaseModel):
    """Single-shot offboarding — generates and sends the email automatically."""
    contractor_id:    str
    last_day:         str
    reason:           str
    laptop:           str
    laptop_returned:  str = ""
    comments:         str = ""
    manager_email:    str = ""


class SendEmailRequest(BaseModel):
    contractor_id:      str
    tram_id_new:        str
    excel_filename:     str
    excel_bytes_base64: str     # base64-encoded .xlsm file
    email_type:         str = "renewal"   # "renewal" or "offboarding"
    offboard_data:      dict = {}


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


# ── ONE-SHOT Renewal Submission ───────────────────────────────────────────────

@app.post("/api/submit-renewal")
def submit_renewal(req: SubmitRenewalRequest):
    """
    Fully automated renewal pipeline triggered by a single PM click.

    Steps executed automatically:
      1. Resolve contractor record from cache
      2. Validate JRS against M389 taxonomy
      3. Download Excel template from Box
      4. Populate the Excel with all PM-reviewed fields
      5. Send the renewal email via Microsoft Graph (Excel attached)
      6. Write audit record

    Returns a single JSON response with all results so the frontend
    can display the confirmation screen in one shot.
    """
    # ── 1. Resolve contractor ─────────────────────────────────────────────────
    contractor = _resolve_contractor(req.contractor_id)

    # ── 2. JRS validation ─────────────────────────────────────────────────────
    jrs_to_use = req.confirmed_jrs or contractor.get("jrsTram", "")
    jrs_result = jrs_validation.validate_jrs(jrs_to_use)
    if jrs_result["outcome"] in ("multi_replacement", "not_found"):
        raise HTTPException(
            status_code=422,
            detail=(
                f"JRS '{jrs_to_use}' cannot be auto-resolved: {jrs_result['outcome']}. "
                "Please correct the JRS field before submitting."
            ),
        )
    # If there is exactly one replacement, apply it automatically
    final_jrs = jrs_to_use
    if jrs_result["outcome"] == "one_replacement" and jrs_result.get("replacements"):
        final_jrs = jrs_result["replacements"][0]

    # ── 3. Build PM checklist dict for Excel generation ───────────────────────
    biz_just_parts = [p for p in [
        req.biz_just_1, req.biz_just_2, req.biz_just_3, req.biz_just_4
    ] if p.strip()]
    biz_just = "\n".join(f"{i+1}. {p}" for i, p in enumerate(biz_just_parts))

    pm_checklist = {
        "manager_email":    contractor.get("pmIntranetId", ""),
        "start_date":       req.new_start_date,
        "end_date":         req.new_end_date,
        "tram_id_new":      req.tram_id_new or contractor.get("tramId", ""),
        "niche_skills":     req.niche_skills,
        "biz_just_1":       req.biz_just_1,
        "biz_just_2":       req.biz_just_2,
        "biz_just_3":       req.biz_just_3,
        "biz_just_4":       req.biz_just_4,
        "us_citizenship":   req.us_citizenship,
        "security_access":  req.security_access,
        "pen_testing":      req.pen_testing,
        "requires_laptop":  req.requires_laptop,
        "laptop_os":        req.laptop_os,
        "laptop_address":   req.laptop_address,
        "contractor_phone": req.contractor_phone,
        "comments":         biz_just,
        "jrs":              final_jrs,
        "band":             req.confirmed_band,
        "location":         req.work_location,
        "rateCap":          req.rate_cap,
        "billRate":         req.bill_rate,
        "gp":               req.gp_pct,
        "contractType":     req.contract_type,
    }

    # Patch contractor with PM-confirmed values for Excel
    contractor_for_excel = {
        **contractor,
        "jrsTram": final_jrs,
        "band":    req.confirmed_band,
        "workLocation": req.work_location,
        "endDate": req.new_end_date,
    }

    # ── 4. Download template + generate Excel ─────────────────────────────────
    try:
        box_client = get_client()
        template_bytes = excel_generation.download_template(box_client, BOX_FOLDER_ID)
    except Exception as e:
        raise HTTPException(
            status_code=503,
            detail=f"Could not download Excel template from Box: {e}",
        )

    excel_result = excel_generation.generate_renewal_excel(
        contractor=contractor_for_excel,
        pm_checklist=pm_checklist,
        template_bytes=template_bytes,
    )

    # ── 5. Send email ─────────────────────────────────────────────────────────
    email_result = None
    email_error  = None
    tram_id_for_subject = req.tram_id_new or contractor.get("tramId", "–")

    if email_sender.is_configured():
        try:
            email_result = email_sender.send_renewal_email(
                contractor=contractor,
                tram_id_new=tram_id_for_subject,
                excel_filename=excel_result["filename"],
                excel_bytes=excel_result["file_bytes"],
            )
        except Exception as e:
            email_error = str(e)
    else:
        email_error = (
            "Microsoft Graph credentials not configured. "
            "Set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER_EMAIL "
            "in Railway environment variables."
        )

    # ── 6. Audit ──────────────────────────────────────────────────────────────
    _append_audit(
        user="current_user",
        action="RENEWAL_SUBMITTED",
        contractor=req.contractor_id,
        detail=(
            f"JRS: {final_jrs} · "
            f"End: {req.new_end_date} · "
            f"Excel: {excel_result['filename']} · "
            f"Email: {'sent' if email_result else 'FAILED — ' + (email_error or '')}"
        ),
        outcome="success" if email_result else "warning",
    )

    return {
        "status":           "ok" if email_result else "partial",
        "jrs_result":       jrs_result,
        "final_jrs":        final_jrs,
        "excel_filename":   excel_result["filename"],
        "excel_sha256":     excel_result["sha256"],
        "fields_written":   excel_result["fields_written"],
        "missing_fields":   excel_result["missing_mandatory"],
        "email":            email_result,
        "email_error":      email_error,
        "submitted_at":     datetime.utcnow().isoformat() + "Z",
        "tram_id_used":     tram_id_for_subject,
        # Convenience: mailto fallback if Graph not configured
        "mailto_fallback":  _build_mailto(contractor, tram_id_for_subject, excel_result["filename"]),
    }


# ── ONE-SHOT Offboarding Submission ───────────────────────────────────────────

@app.post("/api/submit-offboarding")
def submit_offboarding(req: SubmitOffboardingRequest):
    """
    Fully automated offboarding pipeline triggered by a single PM click.
    Sends the offboarding email to the correct CSP team automatically.
    """
    contractor = _resolve_contractor(req.contractor_id)

    offboard_data = {
        "manager_email":    req.manager_email or contractor.get("pmIntranetId", ""),
        "last_day":         req.last_day,
        "reason":           req.reason,
        "laptop":           req.laptop,
        "laptop_returned":  req.laptop_returned,
        "comments":         req.comments,
    }

    email_result = None
    email_error  = None

    if email_sender.is_configured():
        try:
            email_result = email_sender.send_offboarding_email(
                contractor=contractor,
                offboard_data=offboard_data,
            )
        except Exception as e:
            email_error = str(e)
    else:
        email_error = (
            "Microsoft Graph credentials not configured. "
            "Set MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET, MS_SENDER_EMAIL "
            "in Railway environment variables."
        )

    _append_audit(
        user="current_user",
        action="OFFBOARDING_SUBMITTED",
        contractor=req.contractor_id,
        detail=(
            f"Last day: {req.last_day} · "
            f"Reason: {req.reason} · "
            f"Email: {'sent' if email_result else 'FAILED — ' + (email_error or '')}"
        ),
        outcome="success" if email_result else "warning",
    )

    return {
        "status":       "ok" if email_result else "partial",
        "email":        email_result,
        "email_error":  email_error,
        "submitted_at": datetime.utcnow().isoformat() + "Z",
        "mailto_fallback": _build_offboard_mailto(contractor, offboard_data),
    }


# ── Legacy: kept for backward compatibility ───────────────────────────────────

@app.post("/api/generate-excel")
def generate_excel(req: ExcelRequest):
    """Legacy endpoint — prefer /api/submit-renewal for new code."""
    contractor = _resolve_contractor(req.contractor_id)
    try:
        client = get_client()
        template_bytes = excel_generation.download_template(client, BOX_FOLDER_ID)
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"Could not download template from Box: {e}")

    result = excel_generation.generate_renewal_excel(
        contractor=contractor,
        pm_checklist=req.pm_checklist,
        template_bytes=template_bytes,
    )
    return Response(
        content=result["file_bytes"],
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        headers={
            "Content-Disposition": f'attachment; filename="{result["filename"]}"',
            "X-Filename":          result["filename"],
            "X-SHA256":            result["sha256"],
            "X-Missing-Fields":    ",".join(result["missing_mandatory"]),
        },
    )


@app.post("/api/send-email")
def send_email(req: SendEmailRequest):
    """Legacy endpoint — prefer /api/submit-renewal for new code."""
    import base64
    contractor = _resolve_contractor(req.contractor_id)
    if not email_sender.is_configured():
        raise HTTPException(status_code=503, detail="Microsoft Graph credentials not configured.")
    try:
        excel_bytes = base64.b64decode(req.excel_bytes_base64)
        if req.email_type == "offboarding":
            result = email_sender.send_offboarding_email(contractor=contractor, offboard_data=req.offboard_data)
        else:
            result = email_sender.send_renewal_email(
                contractor=contractor, tram_id_new=req.tram_id_new,
                excel_filename=req.excel_filename, excel_bytes=excel_bytes,
            )
        _append_audit(user="current_user", action="EMAIL_SENT", contractor=req.contractor_id,
                      detail=f"To: {result['to']} · Subject: {result['subject']}", outcome="success")
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Email send failed: {e}")


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


def _resolve_contractor(contractor_id: str) -> dict:
    """
    Look up a contractor by serial/TalentID in the in-memory cache.
    Falls back to a minimal stub so the pipeline never hard-blocks on a cache miss.
    """
    if _ingestion_result:
        all_records = (
            _ingestion_result.get("contractors", []) +
            _ingestion_result.get("dq_exceptions", [])
        )
        match = next((c for c in all_records if c.get("serial") == contractor_id), None)
        if match:
            return match
    # Stub — lets the pipeline continue; Excel/email will have partial data
    return {"serial": contractor_id, "name": contractor_id}


def _build_mailto(contractor: dict, tram_id: str, excel_filename: str) -> str:
    """Build a mailto: fallback link for when Graph email is not configured."""
    from urllib.parse import quote
    sector = contractor.get("sector", "")
    from config import SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT
    to = SECTOR_EMAIL_MAP.get(sector.strip().lower(), SECTOR_EMAIL_DEFAULT)
    subject = quote(
        f"{tram_id} {contractor.get('serial','–')} {contractor.get('poNumber','–')} "
        f"{contractor.get('name','–')} {contractor.get('client','–')} RENEWAL"
    )
    body = quote(f"Can you please help with this renewal?\n\nAttachment: {excel_filename}")
    return f"mailto:{to}?subject={subject}&body={body}"


def _build_offboard_mailto(contractor: dict, offboard_data: dict) -> str:
    """Build a mailto: fallback link for offboarding when Graph is not configured."""
    from urllib.parse import quote
    sector = contractor.get("sector", "")
    from config import SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT
    to = SECTOR_EMAIL_MAP.get(sector.strip().lower(), SECTOR_EMAIL_DEFAULT)
    subject = quote(
        f"{contractor.get('tramId','–')} {contractor.get('serial','–')} "
        f"{contractor.get('poNumber','–')} {contractor.get('name','–')} "
        f"{contractor.get('client','–')} OFFBOARDING"
    )
    lines = [
        "Please process the following offboarding request:",
        "",
        f"Contractor Name: {contractor.get('name','–')}",
        f"Serial Number: {contractor.get('serial','–')}",
        f"PO Number: {contractor.get('poNumber','–')}",
        f"Last Day: {offboard_data.get('last_day','–')}",
        f"Reason: {offboard_data.get('reason','–')}",
        f"Laptop: {offboard_data.get('laptop','–')}",
        f"Laptop Returned: {offboard_data.get('laptop_returned','–')}",
        f"Comments: {offboard_data.get('comments','–')}",
    ]
    body = quote("\n".join(lines))
    return f"mailto:{to}?subject={subject}&body={body}"
