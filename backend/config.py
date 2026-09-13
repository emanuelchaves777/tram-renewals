"""
config.py — Central configuration loaded from environment variables.
All settings live here. No magic strings scattered across the codebase.
"""
import os
from dotenv import load_dotenv

load_dotenv()

# ── Box ───────────────────────────────────────────────────────────────────────
BOX_CLIENT_ID        = os.getenv("BOX_CLIENT_ID", "")
BOX_CLIENT_SECRET    = os.getenv("BOX_CLIENT_SECRET", "")
BOX_FOLDER_ID        = os.getenv("BOX_FOLDER_ID", "")
BOX_TAXONOMY_FILE_ID = os.getenv("BOX_TAXONOMY_FILE_ID", "")
BOX_DEV_TOKEN        = os.getenv("BOX_DEV_TOKEN", "")   # local dev only

# ── Report file ───────────────────────────────────────────────────────────────
# Filename prefix used to find the latest report in the Box folder
REPORT_FILE_PREFIX   = "Contractor Management Outlook report IBM Consulting NA"
REPORT_TAB_NAME      = "Compiled"

# Taxonomy sheet / column names (confirmed from M389 file)
TAXONOMY_ACTIVE_TAB      = "Consulting Short List"
TAXONOMY_ACTIVE_COL      = "JR/S"
TAXONOMY_CHANGELOG_TAB   = "Change Log"
TAXONOMY_PREV_COL        = "Previous Value (as of prior publication)"
TAXONOMY_NEW_COL         = "Current Value (effective this publication)"
TAXONOMY_ACTION_COL      = "Action taken"

# ── Staleness ─────────────────────────────────────────────────────────────────
STALENESS_WARNING_DAYS = int(os.getenv("STALENESS_WARNING_DAYS", "3"))

# ── Security ──────────────────────────────────────────────────────────────────
APP_SECRET_KEY = os.getenv("APP_SECRET_KEY", "dev-secret-change-in-production")
CORS_ORIGIN    = os.getenv("CORS_ORIGIN", "https://emanuelchaves777.github.io")

# ── Sector → CSP email routing (confirmed) ────────────────────────────────────
SECTOR_EMAIL_MAP = {
    "public":             "pubcontr@cr.ibm.com",
    "government":         "pubcontr@cr.ibm.com",
    "education":          "pubcontr@cr.ibm.com",
    "healthcare":         "pubcontr@cr.ibm.com",
    "financial services": "fsscontr@cr.ibm.com",
    "banking":            "fsscontr@cr.ibm.com",
    "insurance":          "fsscontr@cr.ibm.com",
    "capital markets":    "fsscontr@cr.ibm.com",
    "industrial":         "indcontr@cr.ibm.com",
    "manufacturing":      "indcontr@cr.ibm.com",
    "automotive":         "indcontr@cr.ibm.com",
    "energy":             "indcontr@cr.ibm.com",
    "chemicals":          "indcontr@cr.ibm.com",
    "technology":         "commsctr@cr.ibm.com",
    "communications":     "commsctr@cr.ibm.com",
    "telecommunications": "commsctr@cr.ibm.com",
    "media":              "commsctr@cr.ibm.com",
}
SECTOR_EMAIL_DEFAULT = "gbstccon@cr.ibm.com"

# ── Mandatory report headers (canonical names) ────────────────────────────────
MANDATORY_HEADERS = [
    "Serial Number",
    "Contractor Full Name",
    "Client Name-PO",
    "Project Name-PO",
    "OOBT PO Expected End Date",
    "TRAM Request ID",
    "JR/S (TRAM)",
    "Actual Band",
    "Work Location",
    "Project Contact (TRAM)",
    "Sector",
    "Vendor",
    "PO Number",
]

# Header aliases — maps alternate column names in the report to canonical names.
# Add new aliases here when the report format changes; never change the canonical names.
HEADER_ALIASES = {
    "serial":                         "Serial Number",
    "talent id":                       "Serial Number",
    "talentid":                        "Serial Number",
    "talentid/serial":                 "Serial Number",
    "serial number":                   "Serial Number",
    "contractor full name":            "Contractor Full Name",
    "contractor name":                 "Contractor Full Name",
    "full name":                       "Contractor Full Name",
    "client name-po":                  "Client Name-PO",
    "client name":                     "Client Name-PO",
    "client":                          "Client Name-PO",
    "project name-po":                 "Project Name-PO",
    "project name":                    "Project Name-PO",
    "project":                         "Project Name-PO",
    "oobt po expected end date":       "OOBT PO Expected End Date",
    "po expected end date":            "OOBT PO Expected End Date",
    "expected end date":               "OOBT PO Expected End Date",
    "tram request id end date":        "TRAM Request ID End Date",
    "tram end date":                   "TRAM Request ID End Date",
    "tram request id":                 "TRAM Request ID",
    "tram id":                         "TRAM Request ID",
    "jr/s (tram)":                     "JR/S (TRAM)",
    "jr/s":                            "JR/S (TRAM)",
    "jrs":                             "JR/S (TRAM)",
    "actual band":                     "Actual Band",
    "band":                            "Actual Band",
    "work location":                   "Work Location",
    "project contact (tram)":          "Project Contact (TRAM)",
    "project contact":                 "Project Contact (TRAM)",
    "sector":                          "Sector",
    "vendor":                          "Vendor",
    "po number":                       "PO Number",
    "geography":                       "Geography",
    "market/gmt":                      "Market/GMT",
    "market":                          "Market/GMT",
    "country":                         "Country",
    "hr lob":                          "HR LOB",
    "tram requester":                  "TRAM Requester",
    "tram requester (submitter)":      "TRAM Requester",
    "contractor assignee":             "Contractor Assignee",
    "pm notes id":                     "PM Notes ID",
    "pm intranet id":                  "PM Intranet ID",
    "skill description":               "Skill Description",
    "contractor notes id":             "Contractor Notes ID",
    "contractor intranet address":     "Contractor Intranet Address",
    "csa id":                          "CSA ID",
    "talentid/cnum":                   "TalentID/CNUM",
    "cnum":                            "TalentID/CNUM",
}
