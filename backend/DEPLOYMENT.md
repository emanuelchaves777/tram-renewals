# TRAM Renewals Backend — Deployment Guide

## What Was Built

```
backend/
├── main.py           ← FastAPI app — all API endpoints
├── ingestion.py      ← Reads .xlsb from Box, parses Compiled tab
├── jrs_validation.py ← Validates JRS against M389 taxonomy
├── box_client.py     ← Box API authentication
├── config.py         ← All settings (env vars + constants)
├── requirements.txt  ← Python dependencies
├── Dockerfile        ← Container definition for deployment
└── .env.example      ← Copy to .env and fill in secrets
```

## API Endpoints

| Method | Path | What it does |
|---|---|---|
| `GET` | `/api/health` | Liveness check |
| `GET` | `/api/contractors` | Full contractor list + metadata |
| `POST` | `/api/refresh` | Re-read .xlsb from Box |
| `POST` | `/api/refresh-taxonomy` | Reload JRS taxonomy from Box |
| `GET` | `/api/jrs/{value}` | Validate one JRS string |
| `GET` | `/api/sector-email/{sector}` | CSP routing email for a sector |
| `POST` | `/api/audit` | Record an audit event |
| `GET` | `/api/audit` | Retrieve audit log |

---

## Step 1 — Run Locally (Test Before Deploying)

### Prerequisites
- Python 3.12 installed — download from python.org
- A Box developer token (see Step 2 below)

### Commands
```powershell
# Go into the backend folder
cd backend

# Create a virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Copy the example env file
copy .env.example .env
# Then open .env in Notepad and fill in your values

# Start the server
uvicorn main:app --reload --port 8000
```

Open your browser at: http://localhost:8000/api/health
You should see: `{"status":"ok",...}`

Interactive API docs: http://localhost:8000/api/docs

---

## Step 2 — Get a Box Developer Token (for local testing)

1. Go to https://developer.box.com
2. Sign in with your IBM Box account (Emanuel.Chaves@ibm.com)
3. Click **My Apps** → **Create New App**
4. Choose **Custom App** → **User Authentication (OAuth 2.0)**
5. Name it `TRAM Renewals API`
6. Once created, go to the **Configuration** tab
7. Copy **Client ID** and **Client Secret** into your `.env`
8. Scroll down to **Developer Token** → click **Generate Developer Token**
9. Copy it into `BOX_DEV_TOKEN=` in your `.env` (it lasts 60 minutes)

### Get your Box Folder ID
- Open Box in a browser and navigate to:
  `IBM \ 3. Consulting NA\`
- Look at the URL: `https://ibm.ent.box.com/folder/123456789`
- The number at the end is your `BOX_FOLDER_ID`

### Get your Taxonomy File ID
- In Box, open the M389 taxonomy .xlsx file
- URL will be: `https://ibm.ent.box.com/file/987654321`
- The number is your `BOX_TAXONOMY_FILE_ID`

---

## Step 3 — Deploy to Railway.app (Free, No Credit Card)

Railway is a free hosting platform that runs Docker containers.

### One-time setup
1. Go to https://railway.app and sign up with your GitHub account
2. Click **New Project** → **Deploy from GitHub repo**
3. Connect your GitHub account if asked
4. Select the `tram-renewals` repository
5. Railway auto-detects the `backend/Dockerfile`
6. Click **Add Variables** and enter all your `.env` values:
   - `BOX_CLIENT_ID`
   - `BOX_CLIENT_SECRET`
   - `BOX_FOLDER_ID`
   - `BOX_TAXONOMY_FILE_ID`
   - `APP_SECRET_KEY`
   - `CORS_ORIGIN` = `https://emanuelchaves777.github.io`
7. Click **Deploy**

### Your backend URL
After deploy, Railway gives you a URL like:
`https://tram-renewals-production.up.railway.app`

Test it: `https://tram-renewals-production.up.railway.app/api/health`

---

## Step 4 — Connect the Frontend to the Backend

Once you have the Railway URL, update `app.html`:

1. Open `app.html` in a text editor
2. Find the line near the top of the `<script>` section:
   ```javascript
   const API_BASE = '';  // will be set after backend is deployed
   ```
3. Change it to:
   ```javascript
   const API_BASE = 'https://tram-renewals-production.up.railway.app';
   ```
4. Upload the updated `app.html` to GitHub
5. GitHub Pages deploys automatically within 60 seconds

---

## Step 5 — First Run Sequence

After everything is deployed, do this once in order:

1. Open the app URL
2. Click **↻ Refresh Taxonomy** — loads the JRS taxonomy from Box
3. Click **↻ Refresh from Box** — loads the contractor report
4. The dashboard fills with real data

From then on, click **Refresh from Box** whenever a new report is posted in Box.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `404 No .xlsb file found` | Check `BOX_FOLDER_ID` is the Consulting NA folder, not a subfolder |
| `Box auth error` | Developer token expired (60 min limit) — generate a new one |
| `Tab 'Compiled' not found` | The report format changed — check available sheet names in Box |
| `Column not found in taxonomy` | Taxonomy file format changed — check tab names match M389 |
| CORS error in browser | Make sure `CORS_ORIGIN` in Railway env matches your exact GitHub Pages URL |
