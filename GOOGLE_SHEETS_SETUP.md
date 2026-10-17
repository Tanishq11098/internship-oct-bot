# Google Sheets & Service Account Setup Guide

This guide walks you step-by-step through setting up Google Cloud, enabling the Google Sheets API, creating a Service Account, and configuring GitHub Secrets for the automated bot.

---

## Step 1: Create a Google Cloud Project

1. Open your browser and navigate to the [Google Cloud Console](https://console.cloud.google.com/).
2. Sign in with your Google account.
3. In the top navigation bar, click the project dropdown (next to "Google Cloud") and select **New Project**.
4. Enter a Project Name (e.g., `Job-Hunting-Bot`) and click **Create**.
5. Once created, make sure the new project is selected in the project dropdown.

---

## Step 2: Enable Google Sheets & Drive APIs

1. In the Google Cloud Console, open the left navigation menu (hamburger icon `☰`) and go to **APIs & Services** > **Library**.
2. In the search box, search for: `Google Sheets API`.
3. Click on **Google Sheets API** from the results and click **Enable**.
4. Return to **APIs & Services** > **Library**.
5. Search for: `Google Drive API`.
6. Click on **Google Drive API** and click **Enable**.

---

## Step 3: Create a Service Account

1. In the left navigation menu, go to **APIs & Services** > **Credentials**.
2. Click **+ CREATE CREDENTIALS** at the top of the page and select **Service Account**.
3. Fill in the Service Account details:
   - **Service account name**: `job-bot-service`
   - **Service account ID**: (automatically populated)
   - **Description**: Service account for automated internship bot persistence
4. Click **Create and Continue**.
5. (Optional) For Role, select **Editor** (or leave blank and grant permission directly in Google Sheets).
6. Click **Done**.

---

## Step 4: Generate and Download the JSON Key

1. Under the **Service Accounts** section in **Credentials**, find the service account you just created.
2. Note the email address (it looks like `job-bot-service@your-project-id.iam.gserviceaccount.com`).
3. Click on the service account email, then go to the **Keys** tab at the top.
4. Click **ADD KEY** > **Create new key**.
5. Select **JSON** as the key type and click **Create**.
6. A JSON file will automatically download to your computer. Keep this file secure and private! Never commit this file to GitHub or share it publicly.

---

## Step 5: Create and Share your Google Sheet

1. Go to [Google Sheets](https://sheets.google.com/) and create a **Blank spreadsheet**.
2. Name your spreadsheet (e.g., `Internship & Job Tracker`).
3. Click the **Share** button in the top right corner.
4. In the "Add people and groups" field, paste the **Service Account email address** copied from Step 4.
5. Set permissions to **Editor** and uncheck "Notify people" (service accounts do not have an inbox).
6. Click **Share** (or **Send**).

---

## Step 6: Extract your Google Sheet ID

Look at your browser's address bar while viewing your Google Sheet. The URL will look like:

```text
https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0
```

The long alphanumeric string between `/d/` and `/edit` is your **Sheet ID**:
`1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms`

Copy this ID for the next step.

---

## Step 7: Configure GitHub Secrets

1. Navigate to your repository on [GitHub](https://github.com/).
2. Go to **Settings** > **Secrets and variables** > **Actions**.
3. Under **Repository secrets**, click **New repository secret** and add the following:

| Secret Name | Value | Description |
|---|---|---|
| `GSHEET_ID` | `1BxiMVs0XRA5...` | The Sheet ID extracted in Step 6. |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | `{"type": "service_account", ...}` | The entire raw text content of the downloaded JSON file from Step 4. |

---

## Step 8: Tab Management & Tracker Operation

You do **not** need to create any sheets or column headers manually! On its first execution, the bot will automatically initialize three tabs:

1. **`seen_listings`**:
   - Maintains a deduplication registry of all visited job URLs and company-role cryptographic fingerprints.
   - Prevents re-fetching or re-processing expired or rejected roles across consecutive days.
2. **`all_listings`**:
   - Master archive of every kept listing matching your domain criteria.
   - Preserved permanently; never auto-pruned.
3. **`tracker`**:
   - Automated candidate tracker for listings scoring **60 or higher** (up to 10 additions per day).
   - Features pre-configured headers and interactive status dropdowns:
     - `To Apply`, `Applied`, `Interview`, `Offer`, `Rejected`, `Ghosted`, `Withdrawn`.

---

## Troubleshooting

- **`gspread.exceptions.SpreadsheetNotFound`**:
  Ensure the `GSHEET_ID` matches your sheet URL and that you shared the sheet with the exact service account email address as an **Editor**.
- **`google.auth.exceptions.DefaultCredentialsError`**:
  Verify that the `GOOGLE_SERVICE_ACCOUNT_JSON` secret contains valid JSON without missing braces or truncation.
