"""
ingestion.py — Parses the contractor management report from uploaded bytes.

No Box dependency. Files are uploaded directly via the /api/upload/report endpoint.

Key design rules:
  - Always resolve columns by HEADER NAME, never by position.
  - Use the earlier of OOBT PO Expected End Date and TRAM Request ID End Date.
  - Quarantine rows with missing mandatory fields into a DQ list.
  - Never raise an exception for a single bad row — isolate it and continue.
"""
import io
import tempfile
import os
from datetime import date, datetime
from typing import Any

import pyxlsb
import openpyxl

from config import (
    REPORT_TAB_NAME,
    MANDATORY_HEADERS,
    HEADER_ALIASES,
    STALENESS_WARNING_DAYS,
)


# ── Public entry point ────────────────────────────────────────────────────────

def ingest_from_bytes(file_bytes: bytes, filename: str, ext: str) -> dict:
    """
    Parse a contractor report from raw bytes.
    Accepts .xlsb (binary Excel) or .xlsx (standard Excel).

    Returns:
        {
          "filename":       str,
          "report_date":    str,
          "ingested_at":    str  (ISO UTC),
          "staleness_days": int,
          "is_stale":       bool,
          "missing_headers": [str],
          "contractors":    [dict],
          "dq_exceptions":  [dict],
        }
    """
    if ext == "xlsb":
        rows, missing_headers = _parse_xlsb(file_bytes, REPORT_TAB_NAME)
    else:
        rows, missing_headers = _parse_xlsx(file_bytes, REPORT_TAB_NAME)

    contractors, dq_exceptions = _build_contractor_records(rows, filename)
    report_date    = _extract_date_from_filename(filename)
    ingested_at    = datetime.utcnow().isoformat() + "Z"
    staleness_days = _compute_staleness(report_date)

    return {
        "filename":        filename,
        "report_date":     report_date,
        "ingested_at":     ingested_at,
        "staleness_days":  staleness_days,
        "is_stale":        staleness_days > STALENESS_WARNING_DAYS,
        "missing_headers": missing_headers,
        "contractors":     contractors,
        "dq_exceptions":   dq_exceptions,
    }


# ── .xlsb parsing ─────────────────────────────────────────────────────────────

def _parse_xlsb(file_bytes: bytes, tab_name: str) -> tuple[list[dict], list[str]]:
    """
    Parse the .xlsb binary workbook.
    Returns (rows, missing_mandatory_headers).

    Each row is a plain dict keyed by CANONICAL header name.
    """
    with tempfile.NamedTemporaryFile(suffix=".xlsb", delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    rows = []
    headers = []          # canonical names in column order
    missing_headers = []

    try:
        with pyxlsb.open_workbook(tmp_path) as wb:
            if tab_name not in wb.sheets:
                raise ValueError(
                    f"Tab '{tab_name}' not found. Available sheets: {wb.sheets}"
                )

            with wb.get_sheet(tab_name) as sheet:
                for row_idx, row in enumerate(sheet.rows()):
                    values = [cell.v for cell in row]

                    if row_idx == 0:
                        # ── Header row: resolve canonical names ───────────────
                        headers = _resolve_headers(values)
                        # Check mandatory headers are present
                        resolved_set = {h for h in headers if h}
                        missing_headers = [
                            m for m in MANDATORY_HEADERS if m not in resolved_set
                        ]
                        continue

                    # ── Data rows ─────────────────────────────────────────────
                    if not any(v for v in values if v is not None):
                        continue   # skip fully empty rows

                    record = {}
                    for col_idx, canonical in enumerate(headers):
                        if canonical and col_idx < len(values):
                            record[canonical] = _clean_value(values[col_idx])

                    rows.append(record)
    finally:
        os.unlink(tmp_path)

    return rows, missing_headers


def _parse_xlsx(file_bytes: bytes, tab_name: str) -> tuple[list[dict], list[str]]:
    """
    Parse a standard .xlsx workbook (openpyxl).
    Returns (rows, missing_mandatory_headers).
    Falls back to first sheet if tab_name not found.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)

    # Find the right sheet — exact match first, then case-insensitive, then first sheet
    ws = None
    if tab_name in wb.sheetnames:
        ws = wb[tab_name]
    else:
        ci = next((s for s in wb.sheetnames if s.lower() == tab_name.lower()), None)
        ws = wb[ci] if ci else wb[wb.sheetnames[0]]

    rows        = []
    headers     = []
    missing_headers = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
        values = list(row)
        if row_idx == 0:
            headers = _resolve_headers(values)
            resolved_set = {h for h in headers if h}
            missing_headers = [m for m in MANDATORY_HEADERS if m not in resolved_set]
            continue
        if not any(v for v in values if v is not None):
            continue
        record = {}
        for col_idx, canonical in enumerate(headers):
            if canonical and col_idx < len(values):
                record[canonical] = _clean_xlsx_value(values[col_idx])
        rows.append(record)

    wb.close()
    return rows, missing_headers


def _clean_xlsx_value(v: Any) -> Any:
    """Normalise a cell value from openpyxl — handles dates natively."""
    if v is None:
        return None
    if isinstance(v, str):
        stripped = v.strip()
        return stripped if stripped else None
    if isinstance(v, (date, datetime)):
        if isinstance(v, datetime):
            return v.date().isoformat()
        return v.isoformat()
    return v


def _resolve_headers(raw_values: list) -> list[str]:
    """
    Map raw header cell values to canonical names via HEADER_ALIASES.
    Unknown headers are kept as-is (lowercased). Empty cells become "".
    """
    resolved = []
    for v in raw_values:
        if v is None:
            resolved.append("")
            continue
        key = str(v).strip().lower()
        canonical = HEADER_ALIASES.get(key, str(v).strip())
        resolved.append(canonical)
    return resolved


def _clean_value(v: Any) -> Any:
    """Normalise a cell value: strip strings, convert xlsb date serials."""
    if v is None:
        return None
    if isinstance(v, str):
        stripped = v.strip()
        return stripped if stripped else None
    if isinstance(v, float) and v > 40000:
        # xlsb stores dates as float serial numbers (days since 1900-01-00)
        try:
            return _xlsb_date_to_iso(v)
        except Exception:
            return v
    return v


def _xlsb_date_to_iso(serial: float) -> str:
    """Convert an Excel/xlsb date serial to an ISO date string YYYY-MM-DD."""
    # Excel epoch: December 30, 1899
    from datetime import timedelta
    epoch = date(1899, 12, 30)
    delta = timedelta(days=int(serial))
    return (epoch + delta).isoformat()


# ── Contractor record builder ─────────────────────────────────────────────────

def _build_contractor_records(
    rows: list[dict], source_filename: str
) -> tuple[list[dict], list[dict]]:
    """
    Convert raw parsed rows into the canonical contractor shape the API serves.
    Rows missing mandatory fields go to dq_exceptions.
    """
    contractors = []
    dq_exceptions = []

    for row_num, row in enumerate(rows, start=2):  # row 1 = header
        issues = _validate_row(row)

        record = {
            # Identity
            "serial":             row.get("Serial Number"),
            "cnum":               row.get("TalentID/CNUM"),
            "name":               row.get("Contractor Full Name"),
            "email":              row.get("Contractor Intranet Address"),
            # Assignment
            "geography":          row.get("Geography"),
            "market":             row.get("Market/GMT"),
            "marketSector":       row.get("Market/Sector"),
            "country":            row.get("Country"),
            "hrLob":              row.get("HR LOB"),
            "sector":             row.get("Sector"),
            "client":             row.get("Client Name-PO"),
            "project":            row.get("Project Name-PO"),
            "vendor":             row.get("Vendor"),
            "poNumber":           row.get("PO Number"),
            # TRAM
            "tramId":             row.get("TRAM Request ID"),
            "tramRequester":      row.get("TRAM Requester"),
            "contractorAssignee": row.get("Contractor Assignee"),
            # People
            "contact":               row.get("Project Contact (TRAM)"),
            "contactEmail":          row.get("PM Intranet ID"),
            "pmNotesId":             row.get("PM Notes ID"),
            "pmIntranetId":          row.get("PM Intranet ID"),
            # BP Manager Intranet ID (CMO column "CL") — used as the primary PM identity filter
            "bpManagerIntranetId":   row.get("BP Manager Intranet ID"),
            # Role
            "band":               row.get("Actual Band"),
            "jrsTram":            row.get("JR/S (TRAM)"),
            "skillDescription":   row.get("Skill Description"),
            "workLocation":       row.get("Work Location"),
            "csaId":              row.get("CSA ID"),
            # Dates
            "endDatePO":          row.get("OOBT PO Expected End Date"),
            "endDateTRAM":        row.get("TRAM Request ID End Date"),
            "endDate":            _earlier_date(
                                      row.get("OOBT PO Expected End Date"),
                                      row.get("TRAM Request ID End Date"),
                                  ),
            # Computed
            "bucket":             _date_bucket(
                                      _earlier_date(
                                          row.get("OOBT PO Expected End Date"),
                                          row.get("TRAM Request ID End Date"),
                                      )
                                  ),
            "daysToExpiry":       _days_to(
                                      _earlier_date(
                                          row.get("OOBT PO Expected End Date"),
                                          row.get("TRAM Request ID End Date"),
                                      )
                                  ),
            # Workflow state — read from report columns if present, else default
            "renewal":    row.get("Renewal In Progress") or "–",
            "offboard":   row.get("Offboard In Progress") or "–",
            # Data quality
            "dq":         len(issues) > 0,
            "dqIssues":   issues,
            "sourceFile": source_filename,
            "sourceRow":  row_num,
            # JRS status — set to "pending" here; updated by jrs_validation module
            "jrsStatus":  "pending",
            "scenario":   "jrs_active",   # overwritten by JRS validation
        }

        if any(i["severity"] == "blocking" for i in issues):
            dq_exceptions.append(record)
        else:
            contractors.append(record)

    return contractors, dq_exceptions


def _validate_row(row: dict) -> list[dict]:
    """Return a list of DQ issue dicts for this row."""
    issues = []

    if not row.get("Serial Number"):
        issues.append({"field": "Serial Number", "severity": "blocking",
                        "message": "Missing contractor identifier (Serial Number / TalentID)."})

    if not row.get("Contractor Full Name"):
        issues.append({"field": "Contractor Full Name", "severity": "blocking",
                        "message": "Missing contractor name."})

    end_po   = row.get("OOBT PO Expected End Date")
    end_tram = row.get("TRAM Request ID End Date")
    if not end_po and not end_tram:
        issues.append({"field": "End Date", "severity": "warning",
                        "message": "Neither OOBT PO Expected End Date nor TRAM Request ID End Date is populated."})
    elif end_po and not _is_valid_date(end_po):
        issues.append({"field": "OOBT PO Expected End Date", "severity": "warning",
                        "message": f"Unparseable date value: '{end_po}'."})

    if not row.get("TRAM Request ID"):
        issues.append({"field": "TRAM Request ID", "severity": "warning",
                        "message": "Missing TRAM Request ID — renewal workflow cannot proceed."})

    return issues


# ── Date helpers ──────────────────────────────────────────────────────────────

def _is_valid_date(value: Any) -> bool:
    if not value:
        return False
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                datetime.strptime(value, fmt)
                return True
            except ValueError:
                continue
        return False
    return True


def _parse_date(value: Any) -> date | None:
    if not value:
        return None
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def _earlier_date(a: Any, b: Any) -> str | None:
    """Return the earlier of two date values as ISO string, or whichever is available."""
    da = _parse_date(a)
    db = _parse_date(b)
    if da and db:
        return min(da, db).isoformat()
    if da:
        return da.isoformat()
    if db:
        return db.isoformat()
    return None


def _days_to(iso_date: str | None) -> int | None:
    if not iso_date:
        return None
    try:
        d = date.fromisoformat(iso_date)
        return (d - date.today()).days
    except ValueError:
        return None


def _date_bucket(iso_date: str | None) -> str:
    days = _days_to(iso_date)
    if days is None:
        return "unknown"
    if days < 0:
        return "past-due"
    if days <= 7:
        return "0-7"
    if days <= 14:
        return "8-14"
    if days <= 30:
        return "15-30"
    if days <= 60:
        return "31-60"
    return "60+"


def _extract_date_from_filename(filename: str) -> str:
    """
    Extract 'DD.MM' from filename like:
    'Contractor Management Outlook report IBM Consulting NA 09.07.xlsb'
    Returns the date string or 'unknown'.
    """
    import re
    match = re.search(r"(\d{2}\.\d{2})\.xlsb$", filename)
    return match.group(1) if match else "unknown"


def _compute_staleness(report_date_str: str) -> int:
    """
    Compute how many days old the report is.
    report_date_str is 'DD.MM' — assumes current year.
    """
    if report_date_str == "unknown":
        return 0
    try:
        today = date.today()
        day, month = map(int, report_date_str.split("."))
        report_date = date(today.year, month, day)
        delta = (today - report_date).days
        return max(0, delta)
    except Exception:
        return 0
