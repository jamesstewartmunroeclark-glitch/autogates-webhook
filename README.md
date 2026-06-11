# Auto Gates Vic — Quote Webhook

Receives Web3Forms submissions, generates a Word quote doc, and emails it to James.

---

## How It Works

1. Customer submits quote form on website
2. Web3Forms sends the data to this server via webhook
3. Server generates a populated Word (.docx) quote
4. Word doc is emailed to you as an attachment

---

## Deploy to Railway (Free — takes ~10 minutes)

### Step 1 — Create a GitHub repo

1. Go to https://github.com and sign in (or create a free account)
2. Click **New repository**
3. Name it `autogates-webhook`, set to **Private**, click **Create**
4. Upload all files from this folder (drag and drop into the repo)

### Step 2 — Deploy on Railway

1. Go to https://railway.app and sign in with GitHub
2. Click **New Project → Deploy from GitHub repo**
3. Select your `autogates-webhook` repo
4. Railway will detect it's a Python app and deploy automatically
5. Once deployed, click your project → **Settings → Domains**
6. Click **Generate Domain** — you'll get a URL like:
   `https://autogates-webhook-production.up.railway.app`
7. Copy this URL — you'll need it for the next steps

### Step 3 — Add Environment Variables

In Railway, go to your project → **Variables** → add these:

| Variable                   | Value                                      |
|----------------------------|--------------------------------------------|
| SMTP_HOST                  | smtp.gmail.com                             |
| SMTP_PORT                  | 587                                        |
| SMTP_USER                  | your.gmail@gmail.com                       |
| SMTP_PASS                  | (your Gmail App Password — see below)      |
| TO_EMAIL                   | james@autogatevic.com.au                   |
| GOOGLE_SHEET_ID            | (your Google Sheet ID — see below)         |
| GOOGLE_CREDENTIALS_JSON    | (your service account JSON — see below)    |

#### Getting a Gmail App Password
1. Go to https://myaccount.google.com/security
2. Enable **2-Step Verification** if not already on
3. Go to **App Passwords** (search for it)
4. Select app: **Mail**, device: **Other** → type "Railway"
5. Copy the 16-character password → paste into SMTP_PASS

#### Setting up Google Sheets logging

**1. Create the spreadsheet:**
1. Go to https://sheets.google.com → create a new blank sheet
2. Name it **Auto Gates Vic — Estimates Log**
3. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/`**THIS_IS_THE_ID**`/edit`
4. Paste it into Railway as `GOOGLE_SHEET_ID`

**2. Create a Service Account (so the server can write to it):**
1. Go to https://console.cloud.google.com
2. Create a new project (or use existing)
3. Enable **Google Sheets API** and **Google Drive API**
4. Go to **IAM & Admin → Service Accounts → Create Service Account**
5. Name it `autogates-webhook`, click Create
6. Click the service account → **Keys → Add Key → JSON**
7. Download the JSON file
8. Open the JSON file, copy the entire contents
9. In Railway, paste the whole JSON as the value for `GOOGLE_CREDENTIALS_JSON`

**3. Share the sheet with the service account:**
1. Open your Google Sheet
2. Click **Share**
3. Paste the `client_email` value from the JSON (looks like `autogates-webhook@project.iam.gserviceaccount.com`)
4. Give it **Editor** access → Share

That's it — every submission will now add a row to your sheet automatically.

**Your sheet columns will be:**
`Reference | Date | Name | Phone | Email | Address | Gate Type | Width | Motor | Driveway | Estimate | Preferred Date | Notes | Status`

The **Status** column defaults to `Pending` — you update it manually to:
- `Pending` — new request, not yet contacted
- `Booked` — site measure scheduled
- `Won` — quote accepted
- `Lost` — didn't proceed

### Step 4 — Add Webhook to Web3Forms

Web3Forms doesn't natively support webhooks, so we use a redirect trick:

1. In your quote form HTML, find the Web3Forms fetch call
2. After a successful submission, add a secondary fetch to your Railway URL:

```javascript
// After the Web3Forms fetch succeeds, also call your webhook:
await fetch('https://YOUR-RAILWAY-URL.up.railway.app/webhook', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(formData)
});
```

That's it! Every form submission will now:
- Send data to Web3Forms (existing)
- Trigger your Railway server to generate + email the Word doc

---

## Testing

Visit your Railway URL in a browser — you should see:
```json
{"status": "Auto Gates Vic webhook running"}
```

To test a full submission, use a tool like https://reqbin.com to POST sample JSON to:
`https://YOUR-RAILWAY-URL.up.railway.app/webhook`

---

## Files

- `app.py` — Flask webhook receiver + email sender
- `docx_generator.py` — Word doc generator
- `requirements.txt` — Python dependencies
- `Procfile` — tells Railway how to run the app
