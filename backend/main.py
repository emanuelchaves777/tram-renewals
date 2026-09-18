"""
main.py — FastAPI application entry point.

NO BOX DEPENDENCY. All files are uploaded directly by the user.

Endpoints:
  GET  /api/health                — liveness check
  POST /api/upload/report         — upload the contractor .xlsb report
  POST /api/upload/taxonomy       — upload the JRS taxonomy .xlsx
  POST /api/upload/template       — upload the Excel renewal template .xlsm
  GET  /api/contractors           — full contractor list (from cache)
  GET  /api/jrs/{jrs_value}       — validate a single JRS string
  POST /api/submit-renewal        — ONE-SHOT: validate JRS + generate Excel + send email
  POST /api/submit-offboarding    — ONE-SHOT: send offboarding email
  POST /api/audit                 — append an audit record
  GET  /api/audit                 — retrieve audit log

Persistence:
  Uploaded files are written to DATA_DIR (default /data, override via DATA_DIR env var).
  On startup the app auto-loads any files already on disk so uploads survive restarts.

Run locally:
  cd backend
  uvicorn main:app --reload --port 8000
"""
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pydantic import BaseModel

from config import CORS_ORIGIN, SECTOR_EMAIL_MAP, SECTOR_EMAIL_DEFAULT
import ingestion
import jrs_validation
import excel_generation
import email_sender
import ledger as ledger_module
try:
    import seed_data as _seed
    _HAS_SEED = True
except ImportError:
    _HAS_SEED = False

# ── Persistent storage directory ──────────────────────────────────────────────
# Railway provides a writable filesystem; we store the three uploaded files here
# so they survive container restarts without needing re-upload.
DATA_DIR = Path(os.getenv("DATA_DIR", "/data"))
DATA_DIR.mkdir(parents=True, exist_ok=True)

_REPORT_PATH   = DATA_DIR / "report.bin"      # raw bytes + ext meta
_TAXONOMY_PATH = DATA_DIR / "taxonomy.xlsx"
_TEMPLATE_PATH = DATA_DIR / "template.bin"    # raw bytes + ext meta

# ── Upload PIN ────────────────────────────────────────────────────────────────
# Set UPLOAD_PIN env var on Railway. If unset, uploads are open (dev mode).
UPLOAD_PIN = os.getenv("UPLOAD_PIN", "")


def _verify_pin(pin: str) -> None:
    """Raise 403 if PIN is configured and provided value doesn't match."""
    if UPLOAD_PIN and pin != UPLOAD_PIN:
        raise HTTPException(status_code=403, detail="Invalid upload PIN.")

# ── App setup ──────────────────────────────────────────────────────────────────

app = FastAPI(
    title="TRAM Subcontractor Renewal & Offboarding API",
    version="2.0.0",
    docs_url="/api/docs",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[CORS_ORIGIN, "http://localhost:3000", "http://127.0.0.1:5500",
                   "null", "*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── In-memory state ────────────────────────────────────────────────────────────
_ingestion_result: dict | None = None
_template_bytes:   bytes | None = None   # the .xlsm template
_audit_log:        list[dict] = []


# ── Startup: auto-load persisted files from disk ──────────────────────────────

def _persist_report(file_bytes: bytes, filename: str, ext: str) -> None:
    """Write report bytes + metadata to disk so it survives restarts."""
    import json
    _REPORT_PATH.write_bytes(file_bytes)
    (_REPORT_PATH.parent / "report_meta.json").write_text(
        json.dumps({"filename": filename, "ext": ext}), encoding="utf-8"
    )


def _persist_template(file_bytes: bytes, filename: str) -> None:
    import json
    _TEMPLATE_PATH.write_bytes(file_bytes)
    (_TEMPLATE_PATH.parent / "template_meta.json").write_text(
        json.dumps({"filename": filename}), encoding="utf-8"
    )


def _boot_load() -> None:
    """Called once at startup. Loads any previously persisted files from disk.
    If a file is missing from disk but seed_data.py is available, the seed is
    used as a fallback so fresh deploys work without manual re-uploads.
    """
    import json
    global _ingestion_result, _template_bytes

    # Always initialise ledger first so it's ready before report applies it
    ledger_module.init(DATA_DIR)
    print(f"[boot] Ledger loaded: {len(ledger_module.get_all())} entries")

    # Report
    meta_path = _REPORT_PATH.parent / "report_meta.json"
    if _REPORT_PATH.exists() and meta_path.exists():
        try:
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            file_bytes = _REPORT_PATH.read_bytes()
            result = ingestion.ingest_from_bytes(file_bytes, meta["filename"], meta["ext"])
            # Apply ledger overrides before caching
            result["contractors"] = ledger_module.apply_to_contractors(result["contractors"])
            _ingestion_result = result
            print(f"[boot] Loaded report from disk: {meta['filename']} — {len(result['contractors'])} contractors")
        except Exception as exc:
            print(f"[boot] Failed to load report from disk: {exc}")

    # Taxonomy — disk first, seed fallback
    if _TAXONOMY_PATH.exists():
        try:
            tax_bytes = _TAXONOMY_PATH.read_bytes()
            jrs_validation.load_taxonomy_from_bytes(tax_bytes)
            print("[boot] Loaded taxonomy from disk")
            if _ingestion_result:
                _run_jrs_batch()
        except Exception as exc:
            print(f"[boot] Failed to load taxonomy from disk: {exc}")
    elif _HAS_SEED:
        try:
            import base64 as _b64
            tax_bytes = _b64.b64decode(_seed.SEED_FILES["taxonomy"]["b64"])
            jrs_validation.load_taxonomy_from_bytes(tax_bytes)
            # Persist to disk so future restarts use the faster disk path
            _TAXONOMY_PATH.write_bytes(tax_bytes)
            print("[boot] Loaded taxonomy from seed (written to disk for future restarts)")
            if _ingestion_result:
                _run_jrs_batch()
        except Exception as exc:
            print(f"[boot] Failed to load taxonomy from seed: {exc}")

    # Template — disk first, seed fallback
    meta_path = _TEMPLATE_PATH.parent / "template_meta.json"
    if _TEMPLATE_PATH.exists() and meta_path.exists():
        try:
            _template_bytes = _TEMPLATE_PATH.read_bytes()
            print("[boot] Loaded template from disk")
        except Exception as exc:
            print(f"[boot] Failed to load template from disk: {exc}")
    elif _HAS_SEED:
        try:
            import base64 as _b64
            tpl_bytes = _b64.b64decode(_seed.SEED_FILES["template"]["b64"])
            _template_bytes = tpl_bytes
            tpl_name = _seed.SEED_FILES["template"]["filename"]
            # Persist to disk
            _TEMPLATE_PATH.write_bytes(tpl_bytes)
            (_TEMPLATE_PATH.parent / "template_meta.json").write_text(
                json.dumps({"filename": tpl_name}), encoding="utf-8"
            )
            print(f"[boot] Loaded template from seed: {tpl_name} (written to disk)")
        except Exception as exc:
            print(f"[boot] Failed to load template from seed: {exc}")


# Run at module load time (FastAPI startup)
_boot_load()


# ── Pydantic models ────────────────────────────────────────────────────────────

class AuditRecord(BaseModel):
    user:       str
    action:     str
    contractor: str | None = None
    detail:     str | None = None
    outcome:    str = "success"


class SubmitRenewalRequest(BaseModel):
    contractor_id:    str
    new_end_date:     str
    new_start_date:   str
    confirmed_jrs:    str
    confirmed_band:   str
    work_location:    str
    rate_cap:         str
    bill_rate:        str
    gp_pct:           str
    contract_type:    str
    niche_skills:     str
    biz_just_1:       str
    biz_just_2:       str = ""
    biz_just_3:       str = ""
    biz_just_4:       str = ""
    us_citizenship:   str = ""
    security_access:  str = ""
    pen_testing:      str = ""
    requires_laptop:  str = "No"
    laptop_os:        str = ""
    laptop_address:   str = ""
    contractor_phone: str = ""
    tram_id_new:      str = ""


class SubmitOffboardingRequest(BaseModel):
    contractor_id:   str
    last_day:        str
    reason:          str
    laptop:          str
    laptop_returned: str = ""
    comments:        str = ""
    manager_email:   str = ""


# ── Health ─────────────────────────────────────────────────────────────────────

@app.get("/api/health")
def health():
    return {
        "status":          "ok",
        "data_loaded":     _ingestion_result is not None,
        "taxonomy_loaded": jrs_validation.get_taxonomy_cache() is not None,
        "template_loaded": _template_bytes is not None,
        "timestamp":       datetime.utcnow().isoformat() + "Z",
    }


@app.get("/api/status")
def get_status():
    """
    Returns the current load state for all three files.
    Used by the Setup panel to show what is already loaded (especially after restarts).
    """
    import json

    report_info = None
    if _ingestion_result:
        report_info = {
            "filename":  _ingestion_result["filename"],
            "report_date": _ingestion_result["report_date"],
            "ingested_at": _ingestion_result["ingested_at"],
            "accepted":  len(_ingestion_result["contractors"]),
            "dq":        len(_ingestion_result["dq_exceptions"]),
            "persisted": _REPORT_PATH.exists(),
        }

    taxonomy_info = None
    if jrs_validation.get_taxonomy_cache() is not None:
        taxonomy_info = {
            "loaded":    True,
            "persisted": _TAXONOMY_PATH.exists(),
        }

    template_info = None
    if _template_bytes is not None:
        meta_path = _TEMPLATE_PATH.parent / "template_meta.json"
        tname = None
        if meta_path.exists():
            try:
                tname = json.loads(meta_path.read_text(encoding="utf-8")).get("filename")
            except Exception:
                pass
        template_info = {
            "filename":  tname,
            "size":      len(_template_bytes),
            "persisted": _TEMPLATE_PATH.exists(),
        }

    return {
        "report":   report_info,
        "taxonomy": taxonomy_info,
        "template": template_info,
        "all_ready": all([report_info, taxonomy_info, template_info]),
    }


# ── File Upload endpoints ──────────────────────────────────────────────────────

@app.post("/api/upload/report")
async def upload_report(file: UploadFile = File(...), pin: str = ""):
    """
    Upload the contractor management .xlsb report.
    Accepts: .xlsb or .xlsx
    PIN-protected: set UPLOAD_PIN env var on Railway.
    """
    _verify_pin(pin)
    global _ingestion_result

    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided.")

    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsb", "xlsx"):
        raise HTTPException(
            status_code=400,
            detail=f"Expected .xlsb or .xlsx file, got .{ext}"
        )

    file_bytes = await file.read()

    try:
        result = ingestion.ingest_from_bytes(file_bytes, file.filename, ext)

        # Apply ledger overrides so app-processed statuses survive the new upload
        result["contractors"] = ledger_module.apply_to_contractors(result["contractors"])
        _ingestion_result = result

        # Persist raw bytes to disk so the file survives restarts
        _persist_report(file_bytes, file.filename, ext)

        # Run JRS validation batch
        _run_jrs_batch()

        ledger_count = len(ledger_module.get_all())
        _append_audit(
            user="current_user",
            action="REPORT_UPLOADED",
            detail=(
                f"File: {file.filename} · "
                f"{len(result['contractors'])} rows accepted · "
                f"{len(result['dq_exceptions'])} rejected · "
                f"{ledger_count} ledger overrides applied"
            ),
            outcome="success",
        )

        return {
            "status":   "ok",
            "filename": result["filename"],
            "accepted": len(result["contractors"]),
            "dq":       len(result["dq_exceptions"]),
            "missing_headers": result["missing_headers"],
            "ledger_applied": ledger_count,
        }

    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to parse report: {e}")


@app.post("/api/upload/taxonomy")
async def upload_taxonomy(file: UploadFile = File(...), pin: str = ""):
    """
    Upload the JRS taxonomy .xlsx file.
    PIN-protected: set UPLOAD_PIN env var on Railway.
    """
    _verify_pin(pin)
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Expected a .xlsx taxonomy file.")

    file_bytes = await file.read()

    try:
        jrs_validation.load_taxonomy_from_bytes(file_bytes)

        # Persist to disk so it survives restarts
        _TAXONOMY_PATH.write_bytes(file_bytes)

        # Re-run JRS batch if report is already loaded
        if _ingestion_result:
            _run_jrs_batch()

        _append_audit(
            user="current_user",
            action="TAXONOMY_UPLOADED",
            detail=f"File: {file.filename}",
            outcome="success",
        )

        return {"status": "ok", "message": "Taxonomy loaded successfully."}

    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to parse taxonomy: {e}")


@app.post("/api/upload/template")
async def upload_template(file: UploadFile = File(...), pin: str = ""):
    """
    Upload the Excel renewal template (.xlsm or .xlsx).
    Stored in memory for use by the Excel generation step.
    Persisted to disk so it survives restarts.
    PIN-protected: set UPLOAD_PIN env var on Railway.
    """
    _verify_pin(pin)
    global _template_bytes

    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided.")

    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in ("xlsm", "xlsx"):
        raise HTTPException(
            status_code=400,
            detail=f"Expected .xlsm or .xlsx template, got .{ext}"
        )

    _template_bytes = await file.read()

    # Persist to disk
    _persist_template(_template_bytes, file.filename)

    _append_audit(
        user="current_user",
        action="TEMPLATE_UPLOADED",
        detail=f"File: {file.filename} · {len(_template_bytes)} bytes",
        outcome="success",
    )

    return {"status": "ok", "filename": file.filename, "size": len(_template_bytes)}


# ── Contractors ────────────────────────────────────────────────────────────────

@app.get("/api/contractors")
def get_contractors():
    if _ingestion_result is None:
        raise HTTPException(
            status_code=503,
            detail="No report loaded. Upload the contractor report via POST /api/upload/report.",
        )
    return {
        "meta": {
            "filename":        _ingestion_result["filename"],
            "report_date":     _ingestion_result["report_date"],
            "ingested_at":     _ingestion_result["ingested_at"],
            "staleness_days":  _ingestion_result["staleness_days"],
            "is_stale":        _ingestion_result["is_stale"],
            "missing_headers": _ingestion_result["missing_headers"],
            "total_accepted":  len(_ingestion_result["contractors"]),
            "total_dq":        len(_ingestion_result["dq_exceptions"]),
        },
        "contractors":   _ingestion_result["contractors"],
        "dq_exceptions": _ingestion_result["dq_exceptions"],
    }


# ── JRS Validation ─────────────────────────────────────────────────────────────

@app.get("/api/jrs/{jrs_value}")
def validate_jrs(jrs_value: str):
    return jrs_validation.validate_jrs(jrs_value)


# ── Sector email routing ───────────────────────────────────────────────────────

@app.get("/api/sector-email/{sector}")
def get_sector_email(sector: str):
    key   = sector.strip().lower()
    email = SECTOR_EMAIL_MAP.get(key, SECTOR_EMAIL_DEFAULT)
    return {"sector": sector, "email": email}


# ── ONE-SHOT Renewal Submission ────────────────────────────────────────────────

@app.post("/api/submit-renewal")
def submit_renewal(req: SubmitRenewalRequest):
    """
    Fully automated renewal pipeline triggered by a single PM click:
      1. Resolve contractor from cache
      2. Validate confirmed_jrs (the PM-reviewed value from the form)
      3. Populate Excel template
      4. Send renewal email via Microsoft Graph
      5. Write audit record

    JRS policy (per jrs-validation-rules.md):
      - confirmed_jrs is the value the PM has already reviewed and accepted in the UI.
      - We validate it to record the outcome; we do NOT silently substitute.
      - If the outcome is still unresolvable (multi_replacement / not_found / deleted),
        the UI should have blocked submission — we 422 here as a safety net.
    """
    contractor = _resolve_contractor(req.contractor_id)

    # Use the PM-confirmed JRS value (already reviewed/edited in the form)
    final_jrs  = req.confirmed_jrs or contractor.get("jrsTram", "")
    # Validate it to capture the outcome for the audit trail
    jrs_result = jrs_validation.validate_jrs(final_jrs)

    # Safety net: block if still unresolvable (UI should have caught this first)
    if jrs_result["outcome"] in ("multi_replacement", "not_found", "deleted"):
        raise HTTPException(
            status_code=422,
            detail=(
                f"JRS '{final_jrs}' cannot be used for submission: outcome={jrs_result['outcome']}. "
                "Escalate to the CSP team before submitting."
            ),
        )

    # Build pm_checklist dict
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
    }

    contractor_for_excel = {
        **contractor,
        "jrsTram":      final_jrs,
        "band":         req.confirmed_band,
        "workLocation": req.work_location,
        "endDate":      req.new_end_date,
    }

    # Excel generation — use uploaded template if available
    if not _template_bytes:
        raise HTTPException(
            status_code=503,
            detail="Excel template not uploaded. Use the Setup panel to upload the template file.",
        )

    excel_result = excel_generation.generate_renewal_excel(
        contractor=contractor_for_excel,
        pm_checklist=pm_checklist,
        template_bytes=_template_bytes,
    )

    # Email send
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
        email_error = "Microsoft Graph credentials not configured — use Outlook fallback link."

    # Persist to ledger — survives next CMO upload
    ledger_module.record_renewal(
        serial=req.contractor_id,
        renewal_data={
            "tram_id_new":   tram_id_for_subject,
            "new_end_date":  req.new_end_date,
            "confirmed_jrs": final_jrs,
            "excel":         excel_result["filename"],
        },
    )

    _append_audit(
        user="current_user",
        action="RENEWAL_SUBMITTED",
        contractor=req.contractor_id,
        detail=(
            f"JRS: {final_jrs} · End: {req.new_end_date} · "
            f"Excel: {excel_result['filename']} · "
            f"Email: {'sent' if email_result else 'fallback — ' + (email_error or '')}"
        ),
        outcome="success" if email_result else "warning",
    )

    return {
        "status":         "ok" if email_result else "partial",
        "jrs_result":     jrs_result,
        "final_jrs":      final_jrs,
        "excel_filename": excel_result["filename"],
        "excel_sha256":   excel_result["sha256"],
        "fields_written": excel_result["fields_written"],
        "missing_fields": excel_result["missing_mandatory"],
        "email":          email_result,
        "email_error":    email_error,
        "submitted_at":   datetime.utcnow().isoformat() + "Z",
        "tram_id_used":   tram_id_for_subject,
        "mailto_fallback": _build_mailto(contractor, tram_id_for_subject, excel_result["filename"]),
    }


# ── ONE-SHOT Offboarding Submission ───────────────────────────────────────────

@app.post("/api/submit-offboarding")
def submit_offboarding(req: SubmitOffboardingRequest):
    contractor = _resolve_contractor(req.contractor_id)

    offboard_data = {
        "manager_email":   req.manager_email or contractor.get("pmIntranetId", ""),
        "last_day":        req.last_day,
        "reason":          req.reason,
        "laptop":          req.laptop,
        "laptop_returned": req.laptop_returned,
        "comments":        req.comments,
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
        email_error = "Microsoft Graph credentials not configured — use Outlook fallback link."

    # Persist to ledger — survives next CMO upload
    ledger_module.record_offboarding(
        serial=req.contractor_id,
        offboard_data={
            **offboard_data,
            "contractor_name": contractor.get("name", req.contractor_id),
        },
    )

    _append_audit(
        user="current_user",
        action="OFFBOARDING_SUBMITTED",
        contractor=req.contractor_id,
        detail=(
            f"Last day: {req.last_day} · Reason: {req.reason} · "
            f"Email: {'sent' if email_result else 'fallback'}"
        ),
        outcome="success" if email_result else "warning",
    )

    return {
        "status":          "ok" if email_result else "partial",
        "email":           email_result,
        "email_error":     email_error,
        "submitted_at":    datetime.utcnow().isoformat() + "Z",
        "mailto_fallback": _build_offboard_mailto(contractor, offboard_data),
    }


# ── Ledger inspection ─────────────────────────────────────────────────────────

@app.get("/api/ledger")
def get_ledger():
    """Return the full state ledger. Useful for admin inspection."""
    return {"entries": ledger_module.get_all(), "count": len(ledger_module.get_all())}


# ── Audit log ──────────────────────────────────────────────────────────────────

@app.post("/api/audit")
def post_audit(record: AuditRecord):
    _append_audit(
        user=record.user, action=record.action,
        contractor=record.contractor, detail=record.detail, outcome=record.outcome,
    )
    return {"status": "recorded"}


@app.get("/api/audit")
def get_audit(limit: int = 50):
    return {"events": list(reversed(_audit_log[-limit:]))}


# ── Internal helpers ───────────────────────────────────────────────────────────

def _resolve_contractor(contractor_id: str) -> dict:
    if _ingestion_result:
        all_records = (
            _ingestion_result.get("contractors", []) +
            _ingestion_result.get("dq_exceptions", [])
        )
        match = next((c for c in all_records if c.get("serial") == contractor_id), None)
        if match:
            return match
        # Report is loaded but this serial wasn't found in it
        raise HTTPException(
            status_code=404,
            detail=(
                f"Contractor '{contractor_id}' not found in the loaded report. "
                "The report may have been re-uploaded without this record."
            ),
        )
    # No report loaded at all (server restarted / cold start)
    raise HTTPException(
        status_code=503,
        detail=(
            "No contractor report is loaded. The server may have restarted and lost its "
            "in-memory state. Please re-upload the report via the Setup panel and try again."
        ),
    )


def _append_audit(user, action, contractor=None, detail=None, outcome="success"):
    _audit_log.append({
        "timestamp":  datetime.utcnow().isoformat() + "Z",
        "user":       user,
        "action":     action,
        "contractor": contractor or "—",
        "detail":     detail or "",
        "outcome":    outcome,
    })


def _run_jrs_batch():
    if not _ingestion_result:
        return
    for c in _ingestion_result["contractors"]:
        jrs_val = c.get("jrsTram")
        if jrs_val:
            result = jrs_validation.validate_jrs(jrs_val)
            c["jrsResult"] = result
            if result["outcome"] == "active":
                c["jrsStatus"] = "active"
            elif result["outcome"] in ("one_replacement", "multi_replacement", "deleted"):
                c["jrsStatus"] = "warn"
            else:
                c["jrsStatus"] = "invalid"
        else:
            c["jrsStatus"] = "invalid"
            c["jrsResult"] = {
                "outcome": "not_found", "input": "",
                "replacements": [], "message": "No JRS value in source record.",
            }


def _build_mailto(contractor, tram_id, excel_filename):
    from urllib.parse import quote
    sector = contractor.get("sector", "")
    to = SECTOR_EMAIL_MAP.get(sector.strip().lower(), SECTOR_EMAIL_DEFAULT)
    subject = quote(
        f"{tram_id} {contractor.get('serial','–')} {contractor.get('poNumber','–')} "
        f"{contractor.get('name','–')} {contractor.get('client','–')} RENEWAL"
    )
    body = quote(f"Can you please help with this renewal?\n\nAttachment: {excel_filename}")
    return f"mailto:{to}?subject={subject}&body={body}"


def _build_offboard_mailto(contractor, offboard_data):
    from urllib.parse import quote
    from datetime import datetime as _dt

    def _v(val):
        return str(val).strip() if val and str(val).strip() else "–"

    def _bool_display(val):
        if not val:
            return "–"
        return "Yes" if str(val).strip().lower() in ("yes", "true", "1") else "No"

    def _format_date(val):
        raw = str(val).strip() if val else ""
        if not raw:
            return "–"
        try:
            return _dt.strptime(raw, "%Y-%m-%d").strftime("%m/%d/%Y")
        except ValueError:
            return raw

    sector = contractor.get("sector", "")
    to = SECTOR_EMAIL_MAP.get(sector.strip().lower(), SECTOR_EMAIL_DEFAULT)
    subject = quote(
        f"{_v(contractor.get('tramId'))} {_v(contractor.get('serial'))} "
        f"{_v(contractor.get('poNumber'))} {_v(contractor.get('name'))} "
        f"{_v(contractor.get('client'))} OFFBOARDING"
    )
    # mailto bodies are plain text — mirror the HTML table columns as labelled lines
    lines = [
        f"Sector:                   {_v(contractor.get('sector'))}",
        f"Manager Email (PM):       {_v(offboard_data.get('manager_email'))}",
        f"Contractor Name:          {_v(contractor.get('name'))}",
        f"PO Number:                {_v(contractor.get('poNumber'))}",
        f"Serial Number:            {_v(contractor.get('serial'))}",
        f"Last day (mm/dd/yyyy):    {_format_date(offboard_data.get('last_day'))}",
        f"Reason For Termination:   {_v(offboard_data.get('reason'))}",
        f"Laptop (yes/no):          {_bool_display(offboard_data.get('laptop'))}",
        f"Laptop returned (yes/no): {_bool_display(offboard_data.get('laptop_returned'))}",
        f"Comments:                 {_v(offboard_data.get('comments'))}",
    ]
    body = quote("\n".join(lines))
    return f"mailto:{to}?subject={subject}&body={body}"
