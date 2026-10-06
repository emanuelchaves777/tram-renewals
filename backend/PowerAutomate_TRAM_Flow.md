TRAM Renewal Automation — Power Automate Desktop Flow
======================================================
Version: 2.0
Author:  Emanuel Chaves Gamboa
Purpose: Polls the TRAM Renewals app backend for pending jobs, fills TRAM
         automatically using the real 18-step flow, reads back the new TRAM
         ID, and calls the completion endpoint so the backend generates the
         Excel and sends the email.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REAL TRAM URL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

https://tram.projectservices.dal.app.cirrus.ibm.com/

(Not tram.ibm.com — that was wrong in v1.0)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW TO SET UP
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Open Power Automate Desktop → My flows → Import → select TRAM_Renewal_Automation.robin
2. Open TRAM in Chrome, navigate manually to: My Requests → Create New Requests
   → Spend → so the full Spend form is on screen
3. For EVERY element marked ★RECORD★ in the script:
   - In PAD: click "Add UI element" on that action
   - In Chrome: click the exact TRAM element to capture it
4. Schedule via make.powerautomate.com (see SCHEDULING section below)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
INPUT VARIABLES  (set once in the flow before recording)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Variable Name        Value
─────────────────    ─────────────────────────────────────────────────────
API_BASE             https://tram-renewals-production.up.railway.app
UPLOAD_PIN           [your UPLOAD_PIN from Railway env vars]
TRAM_URL             https://tram.projectservices.dal.app.cirrus.ibm.com/

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
THE 21 UI ELEMENTS TO RECORD  (in order)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

#   appmask name                    Where / what to click in TRAM
──  ──────────────────────────────  ─────────────────────────────────────────
1   My Requests Link                "My Requests" nav link on the TRAM home page
2   Create New Requests Link        "Create New Requests" tile/link
3   Spend Tile                      "Spend" option on the Select Request Type page
4   Import Yes Radio                "Yes" radio next to "Import from Existing TRAM request or Prom Seat ID?"
5   Import TRAM ID Input            The text input that appears after clicking Yes
6   Import Data Button              The "→" / "Import Data" button that triggers the import
7   Import Confirmation OK Button   "OK" button inside the "Import Confirmation / Import Completed" dialog
8   Import Error OK Button          "OK" button inside the error dialog ("Start Date must be in the future")
9   Spend Type Dropdown             The "Spend Type*" dropdown (select "Renewal")
10  Batch Review Date Field         The "Batch Review Date" date input
11  Expected Start Date Field       The "Expected Start Date*" date input
12  Planned End Date Field          The "Planned End Date*" date input
13  Band Equivalent Dropdown        The "Band Equivalent*" dropdown
14  JRS Dropdown                    The "JR/S*" dropdown
15  Hiring Manager Field            The "Hiring/Bluepages Manager*" input/search field
16  Project Contact Field           The "Project Contact" input/search field
17  Cost Rate Field                 The "Contractor Cost Rate*" number field
18  Bill Rate Field                 The "Contractor Bill Rate*" number field
19  Resource Analytics HUB Section  The "Resource Analytics HUB Data" section header/expander
20  Refresh Hub Data Button         The "Refresh Hub Data" button inside that section
21  Internal Search Yes Radio       "Yes" radio for "Internal Search Performed"
22  Resources Available No Radio    "No" radio for "Are there resources available...?"
23  Business Justification Field    The "Business Justification" or "Comments" text area
24  TRAM Submit Button              The Submit button at the bottom of the form
25  New TRAM ID Element             On the confirmation page: the element showing the new TRAM request ID

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
REAL FLOW — 18 STEPS (confirmed)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

─── STEP 1: Navigate to Create Spend Request ─────────────────────────────────

  My Requests → Create New Requests → Spend

─── STEP 2: Import from existing TRAM ────────────────────────────────────────

  "Import from Existing TRAM request or Prom Seat ID?" → Yes
  Paste old TRAM ID
  Click "Import Data" (→ button)

─── STEP 3: Handle import dialog (TWO possible outcomes) ─────────────────────

  Outcome A — Error dialog:
    "Unable to complete operation due to error: Start Date must be in the future"
    → Click OK and CONTINUE (this is expected — form is still pre-filled)

  Outcome B — Success dialog:
    "Import Confirmation × Import Completed"
    → Click OK

  Both outcomes lead to the same pre-filled form. Continue in both cases.

─── STEP 4: Spend Type ───────────────────────────────────────────────────────

  Spend Type* → select "Renewal" from dropdown

─── STEP 5: Batch Review Date ────────────────────────────────────────────────

  Batch Review Date → the closest Friday from today's date
  (Backend calculates and sends this in the job payload as batch_review_date)

─── STEP 6: Dates ────────────────────────────────────────────────────────────

  Expected Start Date*  → new_start_date from job
  Planned End Date*     → new_end_date from job

─── STEP 7: Band Equivalent ─────────────────────────────────────────────────

  Band Equivalent* → confirmed_band from job (select from dropdown)

─── STEP 8: JR/S ────────────────────────────────────────────────────────────

  Read current JR/S value.
  If it matches confirmed_jrs from the job → leave it.
  If it differs → select the correct value from the dropdown.

─── STEP 9: Hiring/Bluepages Manager + Project Contact ──────────────────────

  Read both fields.
  If they stay the same → leave them (do nothing).
  If hiring_manager is provided in the job → update both fields.

─── STEP 10: Rates ──────────────────────────────────────────────────────────

  Contractor Cost Rate*   → rate_cap from job
  Contractor Bill Rate*   → bill_rate from job

  (Resource GP% is recalculated automatically by TRAM after you enter the rates.
   No need to set it explicitly.)

─── STEP 11: Resource Analytics HUB Data ────────────────────────────────────

  Click "Resource Analytics HUB Data" section
  Click "Refresh Hub Data"
  Wait ~5 seconds

─── STEP 12: Compliance ─────────────────────────────────────────────────────

  Internal Search Performed               → Yes
  Are there resources available...?       → No

─── STEP 13: Submit ─────────────────────────────────────────────────────────

  Click Submit button at the bottom of the form
  Wait for confirmation page (up to 30s)

─── STEP 14: Read new TRAM ID ───────────────────────────────────────────────

  From the confirmation page, capture the new TRAM request ID
  (e.g. NA202506XXXXXXXXX).

  Method A — direct element capture (preferred):
    Record the element showing the TRAM ID on the confirmation page.

  Method B — regex fallback (if element is hard to isolate):
    Get page text, match pattern:  NA\d{10,16}[A-Z]?

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
JOB PAYLOAD — fields PAD reads from the backend
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Field name            Used in step
──────────────────    ─────────────────────────────────────────────────
old_tram_id           Step 2 — paste into Import ID field
new_start_date        Step 6 — Expected Start Date (MM/DD/YYYY)
new_end_date          Step 6 — Planned End Date (MM/DD/YYYY)
batch_review_date     Step 5 — Batch Review Date (closest Friday, MM/DD/YYYY)
confirmed_band        Step 7 — Band Equivalent dropdown
confirmed_jrs         Step 8 — JR/S dropdown (only changed if different)
hiring_manager        Step 9 — Hiring/Bluepages Manager (only set if non-empty)
rate_cap              Step 10 — Contractor Cost Rate
bill_rate             Step 10 — Contractor Bill Rate
biz_just_1/2/3        Steps combined into Business Justification text area

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCHEDULING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Go to make.powerautomate.com
2. New flow → Scheduled → every 5 minutes
3. Add action: "Run a desktop flow"
4. Select: "TRAM Renewal Automation"
5. Connection: your IBMer machine (must be online, PAD running)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TEST THE API BEFORE BUILDING THE FLOW
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Run these in PowerShell to verify the endpoints work:

# Check queue (should return 204 when empty)
Invoke-RestMethod "https://tram-renewals-production.up.railway.app/api/tram-jobs/next?pin=YOUR_PIN"

# View all jobs
Invoke-RestMethod "https://tram-renewals-production.up.railway.app/api/tram-jobs?pin=YOUR_PIN" | ConvertTo-Json

# Manually complete a job (for testing)
$body = '{"tram_id_new":"NA202503048216S","pin":"YOUR_PIN"}'
Invoke-RestMethod -Uri "https://tram-renewals-production.up.railway.app/api/tram-jobs/JOB_ID/complete" -Method Post -Body $body -ContentType "application/json"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TIPS FOR RECORDING TRAM UI ELEMENTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- Open TRAM in Chrome FIRST, navigate to the Create Spend form manually
- In PAD: every web action has an "Add UI element" button — click it,
  then click the element in the TRAM page to capture it
- For date fields: try typing directly. If a datepicker opens, type the
  date then press Escape to close it, then Tab to confirm
- For dropdowns: use "Select option in drop-down list on web page"
- For the Hiring Manager / Project Contact fields: these are typically
  a type-ahead / people-picker. Record the input, type the intranet ID,
  then wait for the suggestion and press Enter or click the first result
- For the confirmation TRAM ID: right-click in Chrome → Inspect →
  note the element's ID or class, then record it in PAD
- For dialogs (Import Confirmation / error): the dialog must be OPEN
  when you click "Add UI element" — use TRAM ID NA202503048216S to
  trigger a real import so both dialogs appear during recording
