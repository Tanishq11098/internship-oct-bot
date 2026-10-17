# Production-Ready Job-Hunting & Internship Intelligence Bot

An automated Python intelligence bot engineered to discover, score, rank, and track high-conviction Finance, Investment Banking, Venture Capital, Private Equity, Consulting, and Strategy internships and fresher opportunities.

Runs autonomously via GitHub Actions every morning at **8:00 AM IST** (`02:30 UTC`), updates your live Google Spreadsheet (`seen_listings`, `all_listings`, `tracker`), generates a styled 50-day rolling digest (`internships.xlsx`), and delivers a daily HTML email summary directly to your inbox.

---

## Architecture Overview

```
Daily GitHub Actions Trigger (08:00 AM IST)
                │
                ▼
  Scrapers & APIs (HTML Scrapers + Rotated APIs + Apify)
                │
                ▼
  Cross-Run Deduplication & Noise Filtering (Title Gate)
                │
                ▼
  Pre-Scoring & Detail-Page Enrichment (Top 40 fetches)
                │
                ▼
  Quality Verification (Credential, Experience & Location Rules)
                │
                ▼
  Multi-Factor Scoring Engine (Max 100 Points)
                │
                ▼
  Live Google Sheets Sync (seen_listings, all_listings, tracker)
                │
                ▼
  Excel Digest Generation (internships.xlsx - 50-Day History)
                │
                ▼
  Email Notification (Gmail SMTP SSL) + GitHub Actions Artifact Upload
```

---

## Key Features

1. **Live Google Sheets Integration (Option A)**:
   - Your Google Sheet is the master live database.
   - **`seen_listings`**: Records normalized URLs and 16-character SHA-256 fingerprints to ensure zero duplicate notifications across runs.
   - **`all_listings`**: Clean master record of all qualifying roles without bloat.
   - **`tracker`**: Auto-populates high-conviction listings (score ≥ 60) with pre-configured dropdowns (`To Apply`, `Applied`, `Interview`, `Offer`, `Rejected`, `Ghosted`, `Withdrawn`).
2. **Professional OpenPyXL Excel Digest**:
   - Generates `internships.xlsx` containing the last 50 days of opportunities.
   - Includes Navy blue headers, clickable apply links, frozen panes, auto-filters, and auto-adjusted column widths.
   - Dedicated `Source Stats` sheet with per-scraper diagnostic telemetry.
3. **Resilient Multi-Source Engine**:
   - **HTML Scrapers**: Shine, TimesJobs (with SSL bypass for broken certs), Internshala (best-effort anti-bot handling).
   - **APIs**: Adzuna (India), JSearch (RapidAPI), Jooble, SerpAPI (Google Jobs).
   - **Apify Integration**: Supports custom actor IDs with safe JSON parameter substitution and 150-second timeouts.
4. **Scoring Engine (Max 100 Points)**:
   - **Tier-1 Employer Match**: 30 points (80+ global banks, Indian boutiques, VC/PE funds, consulting firms).
   - **Domain Fit**: Up to 30 points (VC/PE, IB, Equity Research, Founder's Office, Consulting, Asset Mgmt, etc.).
   - **Role Type**: PPO Internship (20), Graduate Program (18), Full-Time Fresher (14), Internship (12).
   - **Location**: Delhi-NCR (10) or Remote (8).
   - **Entry-Level Signal**: 6 points (Fresher, 0-1 yrs, 2027 batch).
   - **Deadline Urgency**: 4 points if application deadline is within 14 days.
5. **Zero-Result Alerting**:
   - Dispatches a detailed diagnostic alert email if all scrapers return 0 cards, isolating HTTP status, block page detection, and error messages without crashing.

---

## GitHub Secrets and Variables

Configure these in **Settings > Secrets and variables > Actions**:

### Required Repository Secrets

| Secret Name | Description |
|---|---|
| `EMAIL` | Your sender Gmail address. |
| `PASSWORD` | Your 16-character Gmail App Password (created in Google Account > Security > 2-Step Verification > App Passwords). |
| `GSHEET_ID` | The ID of your Google Sheet (from the URL between `/d/` and `/edit`). |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The full raw JSON credentials file of your Google Cloud Service Account. |

### Optional Repository Secrets

| Secret Name | Description |
|---|---|
| `EMAIL_TO` | Target recipient email address (defaults to `EMAIL` if not set). |
| `ADZUNA_APP_ID` | Adzuna API Application ID (India search). |
| `ADZUNA_APP_KEY` | Adzuna API Application Key. |
| `RAPIDAPI_KEY` | RapidAPI key for JSearch job aggregation. |
| `JOOBLE_KEY` | Jooble API Key. |
| `SERPAPI_KEY` | SerpAPI Key for Google Jobs extraction. |
| `APIFY_TOKEN` | Apify API personal token. |

### Repository Variables (Vars)

| Variable Name | Default | Description |
|---|---|---|
| `APIFY_ACTORS` | *(empty)* | Comma-separated Apify actor IDs (e.g. `scrapio/linkedin-search-jobs-scraper`). |
| `APIFY_INPUT_JSON` | `{"queries": ["{query}"], "maxItems": 50}` | JSON template with `{query}` and `{location}` placeholders. |
| `APIFY_MAX_ITEMS` | `100` | Maximum items to process per Apify actor per daily run. |
| `STRICT_LOCATION` | `true` | Restricts eligible roles strictly to Delhi NCR and Remote. |
| `API_QUERIES_PER_RUN` | `4` | Number of rotated search queries to execute on paid APIs per run. |

---

## Local Setup & Development

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/your-username/internship-bot.git
cd internship-bot
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure Local Environment
Create a local `.env` file (which is git-ignored):
```env
EMAIL=your-email@gmail.com
PASSWORD=your-app-password
GSHEET_ID=your-google-sheet-id
GOOGLE_SERVICE_ACCOUNT_JSON={"type": "service_account", ...}
STRICT_LOCATION=true
```

### 3. Run the Bot
```bash
python internship_bot.py
```

---

## Maintenance & Selector Updates

If an HTML source (Shine, TimesJobs, Internshala) changes its layout, you do **not** need to rewrite the bot. Simply update the selector mappings in `internship_bot.py`:

- **Shine**: Update `cards = soup.select(...)` in `scrape_shine()`.
- **TimesJobs**: Update `cards = soup.select("li.clearfix.job-bx")` in `scrape_timesjobs()`.
- **Internshala**: Update `cards = soup.select("div.individual_internship")` in `scrape_internshala()`.

TimesJobs uses `verify=False` to bypass known SSL certificate chain validation errors. If TimesJobs resolves their certificate issue in the future, remove `verify=False` in `scrape_timesjobs()`.

---

## Troubleshooting

- **Email fails to send**:
  Verify you are using an **App Password**, not your main Google account password. Standard Google passwords will fail with `smtplib.SMTPAuthenticationError`.
- **Artifact missing in GitHub Actions**:
  The workflow uploads `internships.xlsx` using `actions/upload-artifact@v4`. Check the run logs to verify the file was created.
- **Google Sheets 403 Forbidden**:
  Ensure you shared your Google Sheet with the Service Account email address as **Editor**.

---

All Rights Reserved.
