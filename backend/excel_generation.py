"""
excel_generation.py — Generates the populated renewal .xlsm file.

Template is uploaded by the PM via the Setup panel (no Box required).

Spec:
  Tab to populate: Onboardings
  Data row       : Row 2 (headers in row 1)
  Output naming  : {Subcontractor Name}_Renewal_{YYYYMMDD}_{TalentID}.xlsm
  Returns        : bytes — sent directly in the API response
"""

import hashlib
import io
import re
from datetime import date, datetime
from typing import Any

import openpyxl

# ── Template identity ─────────────────────────────────────────────────────────
TEMPLATE_FILENAME = "Request Form Template_Contractor Team.xlsm"
TEMPLATE_TAB      = "Onboardings"
DATA_ROW          = 2   # row 1 = headers, row 2 = first data row

# ── Column map: column index (1-based) → template header → data source key ───
# Matches confirmed headers from OQ-18 exactly.
COLUMN_MAP = [
    # (col_index, header,                                          data_key)
    (1,  "Contractor's Hiring Manager (U.S. Manager Email)",      "manager_email"),
    (2,  "Fulfilment Specialist Email",                           "fulfilment_email"),   # ⏳ Pending CSP
    (3,  "Contractor's Full Name",                                "contractor_name"),
    (4,  "Start Date",                                            "start_date"),
    (5,  "End Date",                                              "end_date"),
    (6,  "Account/Client",                                        "client"),
    (7,  "Supplier's Name",                                       "vendor"),
    (8,  "Supplier's Contact",                                    "supplier_contact"),   # ⏳ Pending CSP
    (9,  "Open Seat #",                                           "tram_id_new"),
    (10, "Job Posting ID",                                        "job_posting_id"),     # ⏳ Pending
    (11, "TRAM",                                                  "tram_id_new"),
    (12, "Does the contractor have US citizenship?",              "us_citizenship"),
    (13, "Brief scope of work",                                   "scope_of_work"),
    (14, "Will this candidate be working on/access to security technologies?", "security_access"),
    (15, "Will the candidate be involved in penetration testing?","pen_testing"),
    (16, "Requires Laptop?",                                      "requires_laptop"),
    (17, "Windows, iOS",                                          "laptop_os"),
    (18, "Laptop Shipping Address",                               "laptop_address"),
    (19, "Contractor's Phone Number",                             "contractor_phone"),
    (20, "Comments",                                              "comments"),
]

# ── Mandatory fields (block generation if empty) ──────────────────────────────
MANDATORY_KEYS = {
    "manager_email", "contractor_name", "start_date", "end_date",
    "client", "vendor", "tram_id_new", "us_citizenship",
    "scope_of_work", "security_access", "pen_testing",
    "requires_laptop", "contractor_phone", "comments",
}


# ── Public entry point ────────────────────────────────────────────────────────

def generate_renewal_excel(
    contractor: dict,
    pm_checklist: dict,
    template_bytes: bytes,
) -> dict:
    """
    Populate the renewal template and return the result.

    Args:
        contractor:    Contractor record from ingestion (canonical field names).
        pm_checklist:  PM checklist values collected in Step 2 of renewal workflow.
        template_bytes: Raw bytes of the .xlsm template file.

    Returns:
        {
          "filename":   str,    # e.g. S_Deshmukh_Renewal_20250714_CVWMBY.xlsm
          "file_bytes": bytes,  # populated file ready for download / email attachment
          "sha256":     str,    # SHA-256 checksum for audit log
          "fields_written": int,
          "missing_mandatory": [str],  # empty if all mandatory fields present
        }
    """
    data = _build_data_dict(contractor, pm_checklist)
    missing = _check_mandatory(data)

    filename = _build_filename(contractor)
    populated_bytes = _populate_template(template_bytes, data)
    sha256 = hashlib.sha256(populated_bytes).hexdigest()

    fields_written = sum(1 for _, _, key in COLUMN_MAP if data.get(key))

    return {
        "filename":          filename,
        "file_bytes":        populated_bytes,
        "sha256":            sha256,
        "fields_written":    fields_written,
        "missing_mandatory": missing,
    }



# ── Data assembly ─────────────────────────────────────────────────────────────

def _build_data_dict(contractor: dict, pm_checklist: dict) -> dict:
    """
    Merge contractor record and PM checklist into the flat data dict
    keyed by COLUMN_MAP data_key values.
    """
    name_parts   = (contractor.get("name") or "").split()
    scope        = contractor.get("skillDescription") or ""
    niche_skills = pm_checklist.get("niche_skills", "")
    if niche_skills:
        scope = scope + "\n\nNiche skills: " + niche_skills

    biz_just = "\n".join([
        f"1. {pm_checklist.get('biz_just_1', '')}",
        f"2. {pm_checklist.get('biz_just_2', '')}",
        f"3. {pm_checklist.get('biz_just_3', '')}",
        f"4. {pm_checklist.get('biz_just_4', '')}",
    ]).strip()

    requires_laptop = pm_checklist.get("requires_laptop", "").lower()

    return {
        "manager_email":    pm_checklist.get("manager_email")     or contractor.get("pmIntranetId", ""),
        "fulfilment_email": pm_checklist.get("fulfilment_email", ""),   # ⏳ Pending CSP
        "contractor_name":  contractor.get("name", ""),
        "start_date":       pm_checklist.get("start_date", ""),
        "end_date":         pm_checklist.get("end_date")           or contractor.get("endDate", ""),
        "client":           contractor.get("client", ""),
        "vendor":           contractor.get("vendor", ""),
        "supplier_contact": pm_checklist.get("supplier_contact", ""),   # ⏳ Pending CSP
        "tram_id_new":      pm_checklist.get("tram_id_new")        or contractor.get("tramId", ""),
        "job_posting_id":   pm_checklist.get("job_posting_id", ""),     # ⏳ Pending
        "us_citizenship":   pm_checklist.get("us_citizenship", ""),
        "scope_of_work":    scope,
        "security_access":  pm_checklist.get("security_access", ""),
        "pen_testing":      pm_checklist.get("pen_testing", ""),
        "requires_laptop":  pm_checklist.get("requires_laptop", ""),
        "laptop_os":        pm_checklist.get("laptop_os", "") if requires_laptop == "yes" else "",
        "laptop_address":   pm_checklist.get("laptop_address", "") if requires_laptop == "yes" else "",
        "contractor_phone": pm_checklist.get("contractor_phone", ""),
        "comments":         biz_just or pm_checklist.get("comments", ""),
    }


def _check_mandatory(data: dict) -> list[str]:
    """Return list of mandatory field keys that are missing or empty."""
    return [key for key in MANDATORY_KEYS if not data.get(key)]


# ── Template population ───────────────────────────────────────────────────────

def _populate_template(template_bytes: bytes, data: dict) -> bytes:
    """
    Open the template .xlsm, write data into row 2 of the Onboardings tab,
    and return the populated file as bytes.

    Note: openpyxl preserves .xlsm format when keep_vba=True.
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(template_bytes),
        keep_vba=True,    # preserves macros in .xlsm
    )

    if TEMPLATE_TAB not in wb.sheetnames:
        # Fallback: try case-insensitive match
        match = next(
            (s for s in wb.sheetnames if s.lower() == TEMPLATE_TAB.lower()),
            None
        )
        if not match:
            raise ValueError(
                f"Tab '{TEMPLATE_TAB}' not found in template. "
                f"Available tabs: {wb.sheetnames}"
            )
        ws = wb[match]
    else:
        ws = wb[TEMPLATE_TAB]

    # Resolve headers from row 1 to confirm column positions
    header_to_col = {}
    for cell in ws[1]:
        if cell.value:
            header_to_col[str(cell.value).strip()] = cell.column

    # Write data into row 2
    for col_idx, header, key in COLUMN_MAP:
        value = data.get(key, "")
        if value is None:
            value = ""

        # Try to find the column by header name first (resilient to column moves)
        col = header_to_col.get(header, col_idx)
        ws.cell(row=DATA_ROW, column=col, value=value)

    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ── Filename builder ──────────────────────────────────────────────────────────

def _build_filename(contractor: dict) -> str:
    """
    Build the output filename per confirmed naming convention:
    {Subcontractor Name}_Renewal_{YYYYMMDD}_{TalentID}.xlsm

    Example: S_Deshmukh_Renewal_20250714_CVWMBY.xlsm
    """
    name  = contractor.get("name", "Unknown")
    parts = name.strip().split()

    if len(parts) >= 2:
        # "S. Deshmukh" → "S_Deshmukh"  |  "Sofia Deshmukh" → "S_Deshmukh"
        first = re.sub(r"[^A-Za-z]", "", parts[0])[:1].upper()
        last  = re.sub(r"[^A-Za-z]", "", parts[-1])
        name_token = f"{first}_{last}"
    else:
        name_token = re.sub(r"[^A-Za-z0-9]", "_", name)

    today      = date.today().strftime("%Y%m%d")
    talent_id  = re.sub(r"[^A-Za-z0-9]", "", contractor.get("serial") or "UNKNOWN")

    return f"{name_token}_Renewal_{today}_{talent_id}.xlsm"
