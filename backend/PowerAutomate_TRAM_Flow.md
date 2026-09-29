TRAM Renewal Automation — Power Automate Desktop Flow
======================================================
Version: 1.0
Author:  Emanuel Chaves Gamboa
Purpose: Polls the TRAM Renewals app backend for pending jobs, fills TRAM 9.6
         automatically, reads back the new TRAM ID, and calls the completion
         endpoint so the backend can generate the Excel and send the email.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
HOW TO SET UP
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. Open Power Automate Desktop → New flow → name it "TRAM Renewal Automation"
2. Add a "Set variable" action at the top for each INPUT variable below
3. Build the flow following the STEPS section
4. Schedule it to run every 5 minutes via a cloud flow

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
INPUT VARIABLES  (set these once in the flow)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Variable Name        Value
─────────────────    ─────────────────────────────────────────────────────
API_BASE             https://tram-renewals-production.up.railway.app
UPLOAD_PIN           [your UPLOAD_PIN from Railway env vars]
TRAM_URL             https://tram.ibm.com

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FLOW STEPS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

─── STEP 1: Poll for the next pending job ────────────────────────────────────

Action: Invoke web service
  URL:    %API_BASE%/api/tram-jobs/next?pin=%UPLOAD_PIN%
  Method: GET
  Store response body in: ResponseBody
  Store status code in:   StatusCode

If StatusCode = 204
  → No jobs pending. End flow.

Action: Parse JSON
  JSON:     %ResponseBody%
  Store in: JobData

Set variables from JobData:
  JOB_ID       = %JobData['job_id']%
  OLD_TRAM_ID  = %JobData['old_tram_id']%
  NEW_START    = %JobData['new_start_date']%
  NEW_END      = %JobData['new_end_date']%
  RATE_CAP     = %JobData['rate_cap']%
  BILL_RATE    = %JobData['bill_rate']%
  GP_PCT       = %JobData['gp_pct']%
  JRS          = %JobData['confirmed_jrs']%
  BAND         = %JobData['confirmed_band']%
  BIZ_JUST_1   = %JobData['biz_just_1']%
  BIZ_JUST_2   = %JobData['biz_just_2']%
  BIZ_JUST_3   = %JobData['biz_just_3']%

─── STEP 2: Open TRAM in Chrome ──────────────────────────────────────────────

Action: Launch new Chrome
  URL:      %TRAM_URL%
  Store in: Browser

Action: Wait for web page to load (timeout 30s)

─── STEP 3: Navigate to Create Spend Request ─────────────────────────────────

Action: Click link on web page
  UI element: Link "My Requests"     ← record this

Action: Wait for web page to load

Action: Click link on web page
  UI element: Link "Create New Requests"

Action: Wait for web page to load

Action: Click on web page
  UI element: "Spend" tile           ← record this

Action: Wait for web page to load

─── STEP 4: Set Import = Yes, paste old TRAM ID ─────────────────────────────

Action: Click radio button on web page
  UI element: Radio button "Yes" next to
              "Import from Existing TRAM request or Prom Seat ID?"
              ← record this radio button

Action: Wait for element to appear
  UI element: The text input that appears for the existing TRAM ID

Action: Populate text field on web page
  UI element: Existing TRAM ID input field  ← record this
  Text:       %OLD_TRAM_ID%

Action: Send keys  (Tab to trigger the import)
  Keys: {Tab}

Action: Wait for web page to load (timeout 20s)
  [Form will pre-fill from the imported request]

─── STEP 5: Update mandatory fields ──────────────────────────────────────────

Action: Clear text field + Populate
  UI element: "Expected Start Date"  ← record this
  Text:       %NEW_START%            (format: MM/DD/YYYY — verify TRAM's format)

Action: Clear text field + Populate
  UI element: "Planned End Date"     ← record this
  Text:       %NEW_END%

Action: Clear text field + Populate
  UI element: "Cost Rate"            ← record this
  Text:       %RATE_CAP%

Action: Clear text field + Populate
  UI element: "Bill Rate"            ← record this
  Text:       %BILL_RATE%

Action: Clear text field + Populate
  UI element: "Resource GP %"        ← record this
  Text:       %GP_PCT%

[Only if JRS changed from prior request:]
Action: Clear text field + Populate
  UI element: "JR/S" field           ← record this
  Text:       %JRS%

─── STEP 6: Fill compliance & justification ─────────────────────────────────

Action: Set radio button on web page
  UI element: "Internal Search Performed" → Yes  ← record

Action: Set radio button on web page
  UI element: "Are there resources available..." → No  ← record

Action: Populate text field on web page
  UI element: "Business Justification" / Comments  ← record
  Text:       %BIZ_JUST_1%
  [Append BIZ_JUST_2 and BIZ_JUST_3 with newlines if not empty]

─── STEP 7: Submit ───────────────────────────────────────────────────────────

Action: Click button on web page
  UI element: Submit button at bottom of form  ← record this

Action: Wait for web page to load (timeout 30s)

─── STEP 8: Extract new TRAM ID from confirmation page ──────────────────────

The confirmation page shows: "Submission completed" + the new TRAM ID.

Method A — Direct element (preferred):
  Action: Get details of element on web page
    UI element: The span/div containing the TRAM ID  ← record from real page
    Attribute:  Own Text
    Store in:   NEW_TRAM_ID

Method B — Regex fallback (if element is hard to isolate):
  Action: Get details of web page
    Browser:  %Browser%
    Get:      Web page text
    Store in: PageText

  Action: Get matches with regular expression
    Text:     %PageText%
    Pattern:  NA\d{10,16}[A-Z]?
    Store in: TramIdMatches

  NEW_TRAM_ID = %TramIdMatches[0]%

─── STEP 9: Send completion to backend ──────────────────────────────────────

Action: Invoke web service
  URL:          %API_BASE%/api/tram-jobs/%JOB_ID%/complete
  Method:       POST
  Content type: application/json
  Request body: {"tram_id_new": "%NEW_TRAM_ID%", "pin": "%UPLOAD_PIN%"}
  Store status: CompleteStatus

If CompleteStatus = 200
  → Done. Backend generates Excel + sends email. Log success.

─── STEP 10: Close browser ───────────────────────────────────────────────────

Action: Close web browser
  Browser: %Browser%

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ERROR HANDLING  (wrap every section in Try/Catch)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

In every Catch block:

Action: Invoke web service
  URL:          %API_BASE%/api/tram-jobs/%JOB_ID%/fail
  Method:       POST
  Content type: application/json
  Request body: {"error": "PAD error: %LastError%", "pin": "%UPLOAD_PIN%"}

This marks the job failed. The frontend shows an error + offers manual fallback.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCHEDULING (recommended)
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
$body = '{"tram_id_new":"TESTID123","pin":"YOUR_PIN"}'
Invoke-RestMethod -Uri ".../api/tram-jobs/JOB_ID/complete" -Method Post -Body $body -ContentType "application/json"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TIPS FOR RECORDING TRAM UI ELEMENTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- Open TRAM in Chrome FIRST, navigate to the Create Spend form manually
- In PAD: every web action has an "Add UI element" button — click it,
  then click the element in the TRAM page to capture it
- For date fields: try typing directly. If a datepicker appears, use
  "Send Keys" to type the date then press Escape to close the picker
- For dropdowns: use "Select option in drop-down list on web page"
- For the confirmation TRAM ID: right-click it in Chrome → Inspect →
  note the element's ID or class, then record it in PAD
