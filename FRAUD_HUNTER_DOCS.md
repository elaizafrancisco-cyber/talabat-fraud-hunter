# Fraud Hunter & Data Verifier - System Documentation

> **Powered by Talabat Platform - UAE Finance**

---

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [MCP Gateway Integrations](#mcp-gateway-integrations)
4. [Dynamic Variables & External Storage](#dynamic-variables--external-storage)
5. [Installation & Setup](#installation--setup)
6. [Dashboard Features](#dashboard-features)
7. [Potential Fraud Detection Logic](#potential-fraud-detection-logic)
8. [Data Tables & Formatting](#data-tables--formatting)
9. [Email Automation](#email-automation)
10. [Download & Export Options](#download--export-options)
11. [Security & Access Controls (RBAC)](#security--access-controls-rbac)
12. [Public Access Links](#public-access-links)
13. [Upload Data](#upload-data)
14. [Slack Integration](#slack-integration)
15. [Configuration Reference](#configuration-reference)

---

## Overview

The **Fraud Hunter** system is a web-based fraud analysis dashboard built for the Talabat UAE Finance team. It processes large-scale transaction data from Detailed Reports, Checkout data, AX365, and Branch Inquiry files to identify potential fraud patterns.

**Key Capabilities:**
- Process 3M+ transaction records via streaming
- Identify potential fraud based on configurable thresholds
- Interactive drill-down dashboard with Chart.js
- Automated email alerts with suspicious email summaries
- Slack channel sharing with file attachments
- PDF and Excel export
- Google Sheets integration via Drive sync
- Role-Based Access Control (RBAC)
- Public guest access links for stakeholders

**Tech Stack:** Node.js v20, Express.js, ExcelJS (streaming), Chart.js, html2pdf.js

---

## Architecture

```
FraudAnalysis/
  server.js              Express server (port 3000)
  fraud_engine.js        Core analysis engine
  .claude.json           MCP server configuration
  package.json           Dependencies
  upload_meta.json       Upload metadata tracking
  guest_tokens.json      Public access tokens
  dashboard_data.json    Cached analysis results
  public/
    index.html           Single-page application (dashboard)
    assets/
      talabat_logo.png   Corporate logo (from G: Drive)
```

**G: Drive Path Structure:**
```
G:\Shared drives\AR Team (Talabat UAE Finance) - Shared Drive\Claude\Fraud\
  Input/
    AX365/               AX365 reconciliation files
    Branch Inquiry/      Branch inquiry exports
    Checkout/            Checkout payment CSVs
    Detailed Reports/    Transaction detail Excel files
  Output/
    Fraud_Analysis_*.xlsx  Generated reports
    dashboard_data.json    Cached analysis results
  upload_meta.json         Upload metadata tracking
  guest_tokens.json        Public access tokens
  Talabat logo_TSquare.png Corporate logo source
```

**Data Flow:**
1. Upload data files to `G:\...\Fraud\Input\` subfolders (or use Google Shared Drive sync)
2. Run analysis engine (streaming processors read from Input)
3. Engine outputs Excel report + dashboard JSON to `G:\...\Fraud\Output\`
4. Dashboard renders from JSON with interactive charts

**Network Fix:** `dns.setDefaultResultOrder('ipv4first')` at Line 1 of `server.js` forces IPv4 to prevent `ENETUNREACH` on Gmail SMTP.

---

## MCP Gateway Integrations

Four MCP (Model Context Protocol) servers are configured in `.claude.json`:

| Service            | Endpoint URL                                                         |
|--------------------|----------------------------------------------------------------------|
| Google Workspace   | `https://talabatai.dhhmena.com/mcp/gateway/google-workspace/mcp`     |
| Slack              | `https://talabatai.dhhmena.com/mcp/gateway/slack/mcp`                |
| Looker             | `https://talabatai.dhhmena.com/mcp/gateway/looker/mcp`               |
| Tableau            | `https://talabatai.dhhmena.com/mcp/gateway/tableau/mcp`              |

All connections use **SSE (Server-Sent Events)** transport.

---

## Dynamic Variables & External Storage

| Node    | Type   | URL                                                                                        |
|---------|--------|--------------------------------------------------------------------------------------------|
| Input   | Source | `https://drive.google.com/drive/folders/13WyEDftGEHZjhbCRaXF298gRWbpuUJX8`               |
| Output  | Report | `https://drive.google.com/drive/folders/1ZWyk9K86N81G4Xe1iyB8Y6yvrbVz-6IX`               |

**Local Virtual Path (Google Drive for Desktop):** `G:\Shared drives\AR Team (Talabat UAE Finance) - Shared Drive\Claude\Fraud`

---

## Installation & Setup

### Cloud Deployment (Render.com — Free Tier)

**Permanent URL:** `https://talabat-fraud-hunter.onrender.com`

**GitHub Repository:** `https://github.com/elaizafrancisco-cyber/talabat-fraud-hunter`

The application is deployed on Render.com free tier with auto-HTTPS. It auto-deploys from the GitHub `master` branch on every push.

| Property          | Value                                            |
|-------------------|--------------------------------------------------|
| Platform          | Render.com (Free Web Service)                    |
| Runtime           | Node.js 20                                       |
| Build Command     | `npm install`                                    |
| Start Command     | `node server.js`                                 |
| Storage           | Ephemeral (data resets on redeploy/sleep)         |
| Auto-Sleep        | After 15 min of inactivity (spins up on request) |
| HTTPS             | Automatic via Render                             |

**Data on Cloud:** Files uploaded via the web UI are stored in ephemeral container storage. Data is not persisted across redeploys or sleep cycles — upload fresh data each session. This is by design to protect company data.

**Re-Deploy:** Push to `master` branch on GitHub. Render auto-detects and redeploys within 2-3 minutes.

### Local Development

```bash
# Prerequisites: Node.js v20+
cd FraudAnalysis
npm install

# Start server
node server.js
# Dashboard: http://localhost:3000
# Uses G: Drive paths when available, falls back to ./data/
```

**Dependencies:** `express`, `exceljs`, `multer`, `nodemailer`, `@slack/web-api`, `xlsx`, `open`

---

## Dashboard Features

### Interactive Drill-Down
- Click any chart element (Customer Group, Date, Restaurant, Branch) to filter ALL dashboard views
- Active filter bar shows current filter with a "Clear" button
- Filtered stats recalculate: total orders, fraud count, fraud amount, fraud %

### Celebration Popup
- Triggers automatically when `potentialFraudCount === 0`
- Displays confetti animation (80 particles, 8 Talabat brand colors)
- Shows "No Fraud Detected! Happy party poopers!" message
- Fires on every data load/refresh

### Statistics Cards
- **Total Orders** - from Detailed Report source
- **Potential Fraud** - count of flagged orders
- **Fraud Amount** - total AED value
- **Fraud %** - percentage of total orders

### Charts
- Potential Fraud by Customer Group (doughnut)
- Potential Fraud by Date (bar)
- Potential Fraud by Restaurant Top 10 (horizontal bar)
- Potential Fraud by Branch Top 10 (combo bar + line)

---

## Potential Fraud Detection Logic

**Sub-description:**
> Analyzed by: Total Sale >= AED 10,000 and/or Orders >= 1,000, email used multiple times across flagged orders, or duplicate Total Sale with same Restaurant ID and same Customer Email.

**Flagging Criteria (all records where `Potential Fraud = 'Yes'`):**
1. **High Amount:** Total Sale >= AED 10,000
2. **Multi-Email Usage:** Same customer email appears in 2+ flagged orders
3. **Duplicate Sale Pattern:** Same Total Sale amount + same Restaurant ID + same Customer Email

**Sorting:** Highest to Lowest Total Sale, then Restaurant ID (ascending), then Customer Email (ascending)

---

## Data Tables & Formatting

### Date Format
All dates use **DD-MMM-YYYY** format (e.g., `17-Jun-2026`). Time components are stripped from table displays.

### Financial Metrics
All AED values rounded to **2 decimal places** with comma grouping (e.g., `AED 1,254,320.50`).

### Potential Fraud Table - 17 Columns

| # | Column                  |
|---|-------------------------|
| 1 | Serial #                |
| 2 | Restaurant Id           |
| 3 | Branch Id               |
| 4 | Restaurant              |
| 5 | Branch                  |
| 6 | Order Id                |
| 7 | Date / Time             |
| 8 | Payment Method          |
| 9 | Status                  |
| 10| Service Type            |
| 11| SubTotal                |
| 12| Total Sale              |
| 13| Net Payment Per Order   |
| 14| Customer Group          |
| 15| Customer Account        |
| 16| Customer Email          |
| 17| Potential Fraud Reasons |

**Filters:** Restaurant (text), Restaurant ID (text), Date (text), Customer Email (text)

### Suspicious Emails Table

| Column        | Description                            |
|---------------|----------------------------------------|
| Rank          | Position by suspicion score            |
| Customer Email| Email address                          |
| Order Count   | Number of orders placed                |
| Total Amount  | Sum of all order amounts (AED)         |
| Max Burst     | Max consecutive orders in 60-min window|
| First Seen    | Earliest order date                    |
| Last Seen     | Most recent order date                 |

**Sorting:** Highest to Lowest Order Count, then Highest to Lowest Total Amount
**Search:** Partial match on Customer Email

---

## Email Automation

### Default Email Template
```
Dear Team,

Please verify the emails below that have suspicious transactions.

Regards,
UAE Talabat Finance
```

### Email Table Structure
The automated email includes a table with **2 columns only**:

| Rank | Customer Email |
|------|----------------|

### Button State Lifecycle
1. **Click:** Button disables, text changes to "Sending..."
2. **Success:** Text changes to "Sent!" for 3 seconds
3. **Reset:** Button returns to original state

### SMTP Configuration
Set in Settings page: Host, Port, User, Password

---

## Error Handling & Diagnostics

### Graceful Degradation
The analysis engine wraps each data source in independent try/catch blocks:

| Data Source      | Failure Type | Behavior                                    |
|------------------|-------------|----------------------------------------------|
| Detailed Reports | **Critical** | Analysis halts, error returned with fix tips |
| Excel Generation | **Critical** | Analysis halts, error returned with fix tips |
| Branch Inquiry   | Non-critical | Warning emitted, analysis continues          |
| AX365            | Non-critical | Warning emitted, analysis continues          |
| Checkout         | Non-critical | Warning emitted, analysis continues          |

### File Error Reports
When errors occur, the system returns a `fileErrors` array with:
- **step** — which processing stage failed
- **folder** — directory path where the file was expected
- **files** — list of files found in that folder
- **error** — the error message
- **fix** — suggested resolution

Warnings are displayed in the analysis log with amber highlighting and fix suggestions.

---

## Download & Export Options

Accessible via the **Download** dropdown button in the topbar:

| Format         | Description                                           |
|----------------|-------------------------------------------------------|
| Excel (.xlsx)  | Full fraud analysis workbook with all sheets           |
| PDF (.pdf)     | Captures the **currently active tab** (landscape A3)   |
| Google Sheets  | Exports to Google Drive Output folder                  |

### Universal PDF Download
The PDF download captures whichever tab is currently active (Dashboard, Potential Fraud, Suspicious Emails, etc.), not just the dashboard. The filename includes the tab name: `Fraud_Analysis_<TabName>.pdf`.

**Google Sheets Export:** If the Google Shared Drive (`G:\`) is mounted locally, the Excel file is copied to the Output folder automatically. Otherwise, the user is prompted to download and manually upload.

---

## Security & Access Controls

### Global Public Access Framework
All operational features are publicly accessible to anyone navigating to the dashboard URL. No login or role selection is required.

**Declassified Features (Full Public Access):**

| Feature            | Access Level | Operations Available                                    |
|--------------------|-------------|----------------------------------------------------------|
| Dashboard          | Public       | View stats, charts, interactive drill-down, filters      |
| Potential Fraud    | Public       | View/search/sort flagged orders, all 17 columns          |
| Suspicious Emails  | Public       | View/search/sort email rankings                          |
| Upload Data        | Public       | Upload files to all 4 data folders                       |
| Run Analysis       | Public       | Execute fraud analysis processing runs                   |
| Share & Export     | Public       | Slack sharing, Email distribution, Excel/PDF/GSheets     |

### Settings Tab Password Gate (Option B)
The **Settings** tab is visible in the sidebar for everyone but protected by a master password challenge.

**Behavior:**
1. User clicks "Settings" in the sidebar navigation
2. A browser prompt appears: *"Enter Administrative Master Password to access configuration settings:"*
3. **Correct password** (`543210`) — toast "Access Granted", Settings panel opens. Session remains unlocked for subsequent clicks until page reload.
4. **Wrong password or cancel** — toast "Access Denied: Invalid Password.", user is bounced back to the Dashboard tab.

**Protected Settings (behind password gate):**
- User name / role configuration
- Public access link generation
- Slack bot token and channel ID
- SMTP email server credentials (host, port, user, password)

### Public Access Links

**Generating a Guest Link:**
1. Unlock **Settings** with the master password
2. Find the **Public Access Link** section
3. Click **Generate** to create a unique token URL
4. Click **Copy** to copy to clipboard
5. Share the link with stakeholders

**Guest Access Behavior:**
- Guest users access via `/guest/{token}` URL
- Full public access to all operational features
- Settings tab still requires password authentication
- Tokens are stored in `guest_tokens.json` on G: Drive

---

## Upload Data

### Folder Structure
| Folder           | Content                        | Accepted Formats |
|------------------|--------------------------------|------------------|
| Detailed Reports | Transaction detail exports     | .xlsx, .csv      |
| Checkout         | Checkout payment data          | .xlsx, .csv      |
| AX365            | AX365 reconciliation files     | .xlsx, .csv      |
| Branch Inquiry   | Branch inquiry exports         | .xlsx, .csv      |

### Upload Methods
1. **Direct folder upload:** Click the "+ Upload" button on any folder tab
2. **Drag & drop:** Drag files into the upload zone (uploads to currently selected folder)
3. **Click to browse:** Click the upload zone to open file picker

### Upload Metadata
Each uploaded file records:
- **Date Uploaded:** Timestamp of upload
- **Date Modified:** Last modification time
- **Uploader Name:** Name from Settings > Your Name

Maximum file size: **500 MB per file**

---

## Slack Integration

### Setup
1. Go to **Settings > Slack Configuration**
2. Enter your Slack Bot Token (`xoxb-...`)
3. Set default Channel ID (default: `C085XTZTC05`)

### Sharing
1. Go to **Share & Export > Share to Slack**
2. Customize channel ID and message
3. Click **Send to Slack**
4. The Excel report file is automatically attached

---

## Configuration Reference

### API Endpoints

| Method | Endpoint                    | Description                         | Access    |
|--------|-----------------------------|-------------------------------------|-----------|
| POST   | `/api/upload/:folder`       | Upload files to a data folder       | Public    |
| GET    | `/api/files`                | List all uploaded files with metadata| Public    |
| POST   | `/api/run`                  | Run fraud analysis                  | Public    |
| GET    | `/api/results`              | Get latest analysis results JSON    | Public    |
| GET    | `/api/status`               | Get analysis run status/logs        | Public    |
| GET    | `/api/download/excel`       | Download analysis Excel file        | Public    |
| POST   | `/api/share/slack`          | Share report to Slack               | Public    |
| POST   | `/api/share/email`          | Send email with suspicious emails   | Public    |
| POST   | `/api/export/gsheets`       | Export to Google Drive              | Public    |
| POST   | `/api/public/generate`      | Generate public guest access link   | Public    |
| GET    | `/guest/:token`             | Guest access to dashboard           | Public    |
| GET    | `/api/guest/verify/:token`  | Verify guest token validity         | Public    |

> **Note:** All API endpoints are publicly accessible. The Settings UI panel is the only protected area, gated by a client-side password prompt (master key: `543210`).

### Brand Assets

**Logo:** Corporate Talabat logo (`Talabat logo_TSquare.png`) sourced from `G:\...\Fraud\` and served via `/assets/talabat_logo.png`. Displayed in the topbar at 36px height.

**Subtitle:** "Powered by Talabat Platform - UAE Finance" displayed below the logo in the topbar.

### Brand Colors

| Color     | Hex       | Usage                |
|-----------|-----------|----------------------|
| Primary   | `#FF5900` | Buttons, headers     |
| Accent    | `#411517` | Sidebar, titles      |
| Secondary | `#F4EDE3` | Table alternating    |
| Highlight | `#CFFF00` | Accents, progress    |

### Environment

| Setting    | Local Value                                                                                | Cloud Value                              |
|------------|--------------------------------------------------------------------------------------------|------------------------------------------|
| Port       | 3000                                                                                       | `$PORT` (set by Render)                  |
| Node.js    | v20.20.2                                                                                   | v20 (Render runtime)                     |
| Data Root  | `G:\Shared drives\...\Claude\Fraud` (auto-detected)                                        | `/opt/render/project/src/data`           |
| Input Dir  | `{DATA_ROOT}\Input`                                                                        | `{DATA_ROOT}/Input`                      |
| Output Dir | `{DATA_ROOT}\Output`                                                                       | `{DATA_ROOT}/Output`                     |
| DNS Fix    | `dns.setDefaultResultOrder('ipv4first')` — forces IPv4 routing for SMTP                    | Same                                     |

**Path Resolution Logic:** The server checks for G: Drive first. If available (local dev), uses it. Otherwise falls back to `DATA_ROOT` env var or `./data/` directory. This allows the same codebase to run locally and in the cloud.

---

*Document generated for internal distribution - Talabat UAE Finance*
*Last updated: 17-Jun-2026*
