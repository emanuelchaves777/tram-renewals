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


# ── Client name normalisation ─────────────────────────────────────────────────
# Mirror of the _CLIENT_ALIAS map in app.html so the backend stores canonical names.
_CLIENT_ALIAS: dict = {
    # T-Mobile
    "t-mobile":                               "T-Mobile",
    "t-mo":                                   "T-Mobile",
    "tmobile":                                "T-Mobile",
    "t mobile":                               "T-Mobile",
    "t-mobile usa":                           "T-Mobile",
    "t-mobile usa inc":                       "T-Mobile",
    "t- mobile usa inc":                      "T-Mobile",
    "t-mobile us inc":                        "T-Mobile",
    "t-mobile us":                            "T-Mobile",
    "t-mobile inc":                           "T-Mobile",
    # AT&T
    "at&t":                                   "AT&T",
    "at&t cdo":                               "AT&T",
    "at&t inc.":                              "AT&T",
    "at&t services inc":                      "AT&T",
    "a t & t services inc (p)":              "AT&T",
    "att":                                    "AT&T",
    "at & t":                                 "AT&T",
    # Abbott
    "abbott":                                 "Abbott",
    "abbott laboratories":                    "Abbott",
    # AECC
    "aecc":                                   "AECC",
    "aecc ams":                               "AECC",
    "aecc sf remediation projec":             "AECC",
    "arkansas (aecc)":                        "AECC",
    "arkansas (aecc) ams":                    "AECC",
    # American Electric Power
    "aep":                                    "American Electric Power",
    "aep energy services inc":                "American Electric Power",
    "aep service corp.":                      "American Electric Power",
    "american electric power":                "American Electric Power",
    "american electric power (aep)":          "American Electric Power",
    "american electric power co inc":         "American Electric Power",
    "american electric power company, inc.":  "American Electric Power",
    "american electric power service":        "American Electric Power",
    "american electric power service corp":   "American Electric Power",
    # American Express
    "amex":                                   "American Express",
    "american express":                       "American Express",
    "american express co (p)":                "American Express",
    "american express company":               "American Express",
    "american express travel":                "American Express",
    "american express travel rdt vc":         "American Express",
    "american express travel related":        "American Express",
    "merican express company":                "American Express",
    # AmerisourceBergen / Cencora
    "amerisourcebergen corp":                 "Cencora",
    "amerisourcebergen corporation":          "Cencora",
    "amerisourcebergen services":             "Cencora",
    "cencora":                                "Cencora",
    "cencora inc.":                           "Cencora",
    "cencora, inc":                           "Cencora",
    # Altria
    "altria client services inc":             "Altria",
    "altria client services llc":             "Altria",
    "altria group, inc.":                     "Altria",
    # Anthem / Elevance Health
    "anthem":                                 "Elevance Health",
    "anthem inc":                             "Elevance Health",
    "anthem inc.":                            "Elevance Health",
    "elevance":                               "Elevance Health",
    "elevance health":                        "Elevance Health",
    "elevance health inc":                    "Elevance Health",
    "elevance health inc.":                   "Elevance Health",
    "elevance health, inc.":                  "Elevance Health",
    # Ahold Delhaize
    "ahold":                                  "Ahold Delhaize",
    "ahold delhaize":                         "Ahold Delhaize",
    "koninklijke ahold delhaize n.v.":        "Ahold Delhaize",
    "koninklijke ahold n.v.":                 "Ahold Delhaize",
    # Barclays
    "barclays":                               "Barclays",
    "barclays bank plc":                      "Barclays",
    "barclays plc":                           "Barclays",
    "barclays services corp":                 "Barclays",
    # Bank of America
    "bank of america":                        "Bank of America",
    "bank of america corporation":            "Bank of America",
    "bank of america national":               "Bank of America",
    "bofa":                                   "Bank of America",
    "boa":                                    "Bank of America",
    # Bank of Nova Scotia
    "bank of nova scotia":                    "Bank of Nova Scotia",
    "bank of nova scotia, the":               "Bank of Nova Scotia",
    "the bank of nova scotia":                "Bank of Nova Scotia",
    "the bank of nova scotia use":            "Bank of Nova Scotia",
    # BCBS / Horizon
    "blue cross and blue shield of":          "Blue Cross Blue Shield",
    "blue cross and blue shield of massachusetts": "BCBS Massachusetts",
    "blue cross blue shield of massachusetts":"BCBS Massachusetts",
    "bcbsma":                                 "BCBS Massachusetts",
    "horizon - bcbsnj":                       "Horizon BCBS of NJ",
    "horizon bcbs":                           "Horizon BCBS of NJ",
    "horizon bcbs of nj":                     "Horizon BCBS of NJ",
    "horizon healthcare services inc":        "Horizon BCBS of NJ",
    # Boeing
    "boeing company":                         "Boeing",
    "the boeing co":                          "Boeing",
    # Bruce Power
    "bruce power":                            "Bruce Power",
    "bruce power inc":                        "Bruce Power",
    "bruce power inc.":                       "Bruce Power",
    # Caterpillar
    "caterpillar inc.":                       "Caterpillar",
    # CIBC
    "canadian imperial bank of":              "CIBC",
    "canadian imperial bank of commerce":     "CIBC",
    # Cigna
    "cigna":                                  "Cigna",
    "cigna corp":                             "Cigna",
    "cigna corporation":                      "Cigna",
    # Chubb
    "chubb corporation":                      "Chubb",
    "chubb ina holdings inc":                 "Chubb",
    "chubb insurance":                        "Chubb",
    # Cognitus
    "cognitus":                               "Cognitus",
    "cognitus acquisition":                   "Cognitus",
    "cognitus adquisition":                   "Cognitus",
    # Comcast
    "comcast":                                "Comcast",
    "comcast corporation":                    "Comcast",
    "comcast corp":                           "Comcast",
    # Costco
    "costco wholesale corp":                  "Costco",
    "costco wholesale corporation":           "Costco",
    # Discover / DFS
    "discover financial services":            "Discover Financial Services",
    "dfs services llc":                       "Discover Financial Services",
    # DirecTV
    "directv":                                "DirecTV",
    "directtb":                               "DirecTV",
    "directv llc":                            "DirecTV",
    "directv llc at&t cdo bi":                "DirecTV",
    "direct tv":                              "DirecTV",
    "direcTb":                                "DirecTV",
    # DND
    "department of national defence":         "DND",
    "dnd":                                    "DND",
    "dnd - cfhis":                            "DND",
    "dnd - dhrim":                            "DND",
    "dnd - dlps":                             "DND",
    "dnd - iss":                              "DND",
    "dnd  misl":                              "DND",
    "dnd - misl":                             "DND",
    "dnd - misl - sap ehsm functional analyst":"DND",
    "dnd - remit project":                    "DND",
    "dnd / dhrim":                            "DND",
    "dnd cfhis":                              "DND",
    "dnd dhrim":                              "DND",
    "dnd dhrim ta 16":                        "DND",
    "dnd dlps":                               "DND",
    "dnd drm":                                "DND",
    "dnd drmis iss":                          "DND",
    "dnd iss":                                "DND",
    "dnd misl":                               "DND",
    "dnd- misl":                              "DND",
    "dnd misl ta 09":                         "DND",
    "dnd/ cfhis":                             "DND",
    "dnd/ misl":                              "DND",
    "drmis iss":                              "DND",
    # eBay
    "ebay inc":                               "eBay",
    "ebay inc.":                              "eBay",
    # Enbridge
    "enbridge inc.":                          "Enbridge",
    # Ernst & Young
    "ernst & young":                          "Ernst & Young",
    "ernst & young llp":                      "Ernst & Young",
    # ESDC
    "esdc":                                   "ESDC",
    "esdc - cpp":                             "ESDC",
    "esdc - oas":                             "ESDC",
    "esdc - tdri":                            "ESDC",
    "esdc tdri":                              "ESDC",
    "esdc tdri ta 95":                        "ESDC",
    # Fiserv
    "fiserv":                                 "Fiserv",
    "fiserv inc":                             "Fiserv",
    "fiserv inc (p)":                         "Fiserv",
    "fiserv, inc.":                           "Fiserv",
    # Ford
    "ford motor co":                          "Ford",
    # General Motors
    "general motors company":                 "General Motors",
    "general motors llc":                     "General Motors",
    # Gilead
    "gilead sciences inc":                    "Gilead",
    "gilead sciences, inc.":                  "Gilead",
    # Google
    "google":                                 "Google",
    "google inc.":                            "Google",
    "google llc":                             "Google",
    # Hakkoda
    "hakkoda":                                "Hakkoda",
    "hakkoda acquisition":                    "Hakkoda",
    # Honda
    "honda":                                  "Honda",
    "honda motor co., ltd.":                  "Honda",
    # Humana
    "humana inc.":                            "Humana",
    # IBM (internal)
    "ibm":                                    "IBM",
    "ibm corporation":                        "IBM",
    "ibm consulting":                         "IBM",
    # IPG / Interpublic
    "ipg":                                    "IPG",
    "interpublic group of companies inc":     "IPG",
    "interpublic group of cos inc (p)":       "IPG",
    # Johnson & Johnson
    "johnson & johnson":                      "Johnson & Johnson",
    "johnson & johnson services inc":         "Johnson & Johnson",
    # JPMorgan Chase
    "jpmorgan":                               "JPMorgan Chase",
    "jp morgan":                              "JPMorgan Chase",
    "jpmorgan chase":                         "JPMorgan Chase",
    "jp morgan chase":                        "JPMorgan Chase",
    "j.p. morgan":                            "JPMorgan Chase",
    # Juniper Networks
    "juniper networks inc":                   "Juniper Networks",
    "juniper networks, inc.":                 "Juniper Networks",
    # Kaiser
    "kaiser":                                 "Kaiser",
    "kaiser foundation health plan, inc.":    "Kaiser",
    # Kraft Heinz
    "kraft foods inc.":                       "Kraft Heinz",
    "kraft heinz foods co":                   "Kraft Heinz",
    "the kraft heinz":                        "Kraft Heinz",
    "the kraft heinz company":                "Kraft Heinz",
    # Kroger
    "kroger":                                 "Kroger",
    "kroger co":                              "Kroger",
    "the kroger co":                          "Kroger",
    # Lockheed Martin
    "lockheed martin":                        "Lockheed Martin",
    "lockheed martin corporation":            "Lockheed Martin",
    # Marriott
    "marriott international inc":             "Marriott",
    "marriott international, inc.":           "Marriott",
    "marriott ownership resorts inc":         "Marriott",
    "marriott vacation worldwide":            "Marriott",
    # Medtronic
    "medtronic":                              "Medtronic",
    "medtronic inc":                          "Medtronic",
    "medtronic, inc":                         "Medtronic",
    # MetLife
    "metlife":                                "MetLife",
    # Micron
    "micron technology inc":                  "Micron",
    "micron technology, inc.":                "Micron",
    # Morgan Stanley
    "morgan stanley":                         "Morgan Stanley",
    "morgan stanley services group inc":      "Morgan Stanley",
    "morgan stanley smith barney llc":        "Morgan Stanley",
    # National Grid
    "national grid":                          "National Grid",
    "national grid plc":                      "National Grid",
    "national grid usa service company":      "National Grid",
    # Navistar
    "navistar inc (p)":                       "Navistar",
    "navistar international":                 "Navistar",
    "navistar international corporation":     "Navistar",
    # Nestle
    "nestle":                                 "Nestle",
    "nestle regional globe office north":     "Nestle",
    # Neudesic
    "neudesic":                               "Neudesic",
    "neudesic acquisition":                   "Neudesic",
    # New York Life
    "new york life insurance co":             "New York Life",
    "new york life insurance company":        "New York Life",
    # NextEra Energy
    "nextera energy inc":                     "NextEra Energy",
    "nextera energy, inc.":                   "NextEra Energy",
    # Nintendo
    "nintendo of america inc.":               "Nintendo",
    # Norfolk Southern
    "norfolk southern corporation":           "Norfolk Southern",
    # NYPD
    "nypd":                                   "NYPD",
    "nypd - cdw":                             "NYPD",
    "new york city police department":        "NYPD",
    # Accelalpha
    "accelalpha":                             "Accelalpha",
    "accel alpha":                            "Accelalpha",
    "accelalpha acquisition":                 "Accelalpha",
    # Omnicom
    "omnicom":                                "Omnicom",
    "omnicom group inc":                      "Omnicom",
    # Oncor
    "oncor":                                  "Oncor",
    "oncor electric":                         "Oncor",
    "oncor electric delivery":                "Oncor",
    "oncor electric delivery c":              "Oncor",
    "oncor electric delivery co llc":         "Oncor",
    # PayPal
    "paypal":                                 "PayPal",
    "paypal inc":                             "PayPal",
    "paypal inc.":                            "PayPal",
    # PepsiCo
    "pepsico inc":                            "PepsiCo",
    "pepsico, inc.":                          "PepsiCo",
    # Pfizer
    "pfizer":                                 "Pfizer",
    "pfizer inc":                             "Pfizer",
    "pfizer inc.":                            "Pfizer",
    # PNC
    "pnc":                                    "PNC",
    "pnc bank":                               "PNC",
    "pnc bank canada branch":                 "PNC",
    "pnc bank national association":          "PNC",
    "pnc financial services group inc":       "PNC",
    # PSPC
    "pspc":                                   "PSPC",
    "pspc - posr":                            "PSPC",
    "pspc - sigma":                           "PSPC",
    "pspc ams":                               "PSPC",
    "pspc- psdpt (phoenix)":                  "PSPC",
    "pspc rpa":                               "PSPC",
    "pspc sigma":                             "PSPC",
    # Prudential
    "prudential financial inc":               "Prudential",
    "prudential financial inc.":              "Prudential",
    "the prudential insurance company of":    "Prudential",
    # RBC
    "rbc - royal bank of canada":             "Royal Bank of Canada",
    "royal bank of canada":                   "Royal Bank of Canada",
    # RTX
    "rtx":                                    "RTX",
    "rtx corporation":                        "RTX",
    "rtx- e-hub & rtx - common build":        "RTX",
    # SSC / Shared Services Canada
    "shared services canada":                 "SSC",
    "ssc":                                    "SSC",
    "ssc fortinet":                           "SSC",
    # StanCorp / Standard Insurance
    "stancorp financial group, inc":          "StanCorp",
    "stancorp financial group, inc.":         "StanCorp",
    "standard insurance co":                  "StanCorp",
    # Staples
    "staples inc":                            "Staples",
    "staples, inc.":                          "Staples",
    # State Farm
    "state farm insurance co":                "State Farm",
    "state farm mutual automobile":           "State Farm",
    "state farm mutual automobile (p)":       "State Farm",
    "state farm mutual automobile insurance company": "State Farm",
    # Suncor
    "suncor energy inc":                      "Suncor",
    "suncor energy services inc":             "Suncor",
    # SunTrust (historical, now Truist)
    "suntrust banks, inc.":                   "Truist",
    # TD Bank
    "td bank financial group":                "TD Bank",
    "the toronto-dominion bank":              "TD Bank",
    "toronto-dominion bank, the":             "TD Bank",
    # Telus
    "telus canada":                           "Telus",
    "telus communications inc":               "Telus",
    "telus corporation":                      "Telus",
    # Toyota
    "toyota motor corporation":               "Toyota",
    "toyota motor credit corp":               "Toyota",
    "toyota motor north america inc":         "Toyota",
    "toyota motor sales usa inc (p)":         "Toyota",
    "toyota tsusho america inc":              "Toyota",
    # Truist
    "truist":                                 "Truist",
    "truist financial corp":                  "Truist",
    "truist financial corporation":           "Truist",
    "truist financial corporation digital marketing": "Truist",
    # UnitedHealth
    "united healthcare services inc":         "UnitedHealth",
    "united health":                          "UnitedHealth",
    "unitedhealthcare":                       "UnitedHealth",
    "unitedhealth":                           "UnitedHealth",
    "uhc":                                    "UnitedHealth",
    # UPS
    "united parcel service oasis supply corp":"UPS",
    "united parcel service, inc.":            "UPS",
    # USAA
    "usaa":                                   "USAA",
    "usaa - hogwarts team extension":         "USAA",
    "usaa - ibm hogwarts ii":                 "USAA",
    "usaa - ibm hogwarts team":               "USAA",
    "united services automobile":             "USAA",
    "united services automobile association": "USAA",
    "united services automobile usaa emm rtb":"USAA",
    # Verizon
    "verizon":                                "Verizon",
    "verizon communications":                 "Verizon",
    "verizon sourcing llc":                   "Verizon",
    # W. L. Gore
    "w l gore & associates inc":              "W. L. Gore & Associates",
    "w. l. gore & associates, inc.":          "W. L. Gore & Associates",
    # Wells Fargo
    "wells fargo":                            "Wells Fargo",
    "wells fargo & co":                       "Wells Fargo",
    "wells fargo & company":                  "Wells Fargo",
    "wells fargo bank national":              "Wells Fargo",
    # Xcel Energy
    "xcel energy":                            "Xcel Energy",
    "xcel energy inc.":                       "Xcel Energy",
    "xcel energy services inc":               "Xcel Energy",
    # Citigroup / Citi
    "citibank":                               "Citigroup",
    "citi":                                   "Citigroup",
    "citigroup":                              "Citigroup",
}


def _normalise_client(raw: Any) -> Any:
    """Return canonical client name, or the original value if no alias matches."""
    if not raw:
        return raw
    key = str(raw).strip().lower()
    # collapse multiple spaces
    import re as _re
    key = _re.sub(r"\s+", " ", key)
    return _CLIENT_ALIAS.get(key, str(raw).strip())


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
            "client":             _normalise_client(row.get("Client Name-PO")),
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
            # Workflow state — NEVER read from CMO columns (unreliable).
            # Only the app ledger sets these after a PM submits through the app.
            "renewal":  "–",
            "offboard": "–",
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
