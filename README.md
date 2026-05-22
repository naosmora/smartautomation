# LA Ops Hub

A free internal platform for the Ops Hub teams — automated reports, prompt library, SOPs, useful links, and AMA. Hosted on GitHub Pages, automated with GitHub Actions.

## Deploy in 5 minutes

### Step 1 — Fork this repo
Click **Fork** (top right on GitHub). Name it `la-ops-hub`.

### Step 2 — Enable GitHub Pages
**Settings → Pages → Source: Deploy from branch → main / root → Save**

Your site is live at: `https://YOUR_USERNAME.github.io/la-ops-hub/`

### Step 3 — Add secrets for the weekly report
**Settings → Secrets and variables → Actions → New repository secret**

Add these secrets:
| Secret name | What it is |
|---|---|
| `JIRA_BASE_URL` | e.g. `https://yourcompany.atlassian.net` |
| `JIRA_EMAIL` | Your JIRA login email |
| `JIRA_API_TOKEN` | From https://id.atlassian.com/manage-profile/security/api-tokens |
| `JIRA_PROJECT_KEY` | e.g. `OPS` |
| `SHAREPOINT_TENANT_ID` | Azure AD tenant ID |
| `SHAREPOINT_CLIENT_ID` | App registration client ID |
| `SHAREPOINT_CLIENT_SECRET` | App registration secret |
| `SHAREPOINT_SITE_ID` | SharePoint site ID |

### Step 4 — Run the first report manually
**Actions → Weekly ops report → Run workflow**

The report updates every Monday at 9am automatically after that.

## File structure

```
la-ops-hub/
├── index.html              ← Homepage / hub
├── report.html             ← Weekly report (reads data/report.json)
├── prompts.html            ← Copilot prompt library + email generator
├── wiki.html               ← SOPs and team knowledge base
├── links.html              ← Useful links directory
├── ama.html                ← Ask me anything
├── assets/css/main.css     ← All styles
├── data/
│   └── report.json         ← Auto-generated each Monday (or edit manually)
├── scripts/
│   └── generate_report.py  ← Fetches JIRA + SharePoint, writes report.json
└── .github/workflows/
    └── weekly-report.yml   ← GitHub Actions schedule
```

## How to update content

Everything is just files — edit directly on GitHub (click the file, click the pencil).

| What to update | File to edit |
|---|---|
| Team members / SOPs | `wiki.html` |
| Useful links | `links.html` (edit the JS `links` object) |
| AMA questions & answers | `ama.html` (edit the `qas` array) |
| Prompt library | `prompts.html` (edit the `prompts` array) |
| Report data (manual) | `data/report.json` |

## How the auto-report works

Every Monday at 9am, GitHub Actions:
1. Spins up a free virtual machine
2. Runs `scripts/generate_report.py`
3. The script calls the JIRA API and SharePoint Graph API
4. Generates `data/report.json`
5. Commits and pushes the file
6. GitHub Pages publishes it instantly

** GitHub Actions free tier = 2,000 min/month. Each run ≈ 2 min. Weekly = ~8 min/month.

## SharePoint integration setup

1. Go to [Azure Portal](https://portal.azure.com) → App registrations → New registration
2. Name it `la-ops-hub`, single tenant
3. Go to API permissions → Add → Microsoft Graph → Application permissions:
   - `Sites.Read.All`
   - `Files.Read.All`
4. Grant admin consent
5. Go to Certificates & secrets → New client secret → copy the value
6. Add all values as GitHub secrets (see Step 3 above)

---
Powered by GitHub Pages - All Rights Reserved
