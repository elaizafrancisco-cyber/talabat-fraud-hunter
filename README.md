# Talabat Fraud Hunter & Data Verifier

**UAE Finance — Fraud Analysis Platform**

---

## Dashboard Links

| Version | URL | Branch |
|---------|-----|--------|
| **Full Server App (Render)** | [https://talabat-fraud-hunter.onrender.com](https://talabat-fraud-hunter.onrender.com) | `master` |
| **Browser Version (GitHub Pages)** | [https://elaizafrancisco-cyber.github.io/talabat-fraud-hunter/](https://elaizafrancisco-cyber.github.io/talabat-fraud-hunter/) | `gh-pages` |

---

## Overview

The Fraud Hunter system analyzes Talabat UAE transaction data to detect potential fraud patterns. It cross-references Detailed Reports, Checkout payment data, Branch Inquiry, and AX365 records to flag suspicious orders and emails.

---

## Two Versions

### 1. Render — Full Server App (`master` branch)

The production server-based application with all features:

- **Server-Side Processing:** Node.js + ExcelJS streaming for large files (100MB+)
- **Settings Page:** Password-protected (master key: `543210`)
- **SMTP Email:** Send fraud alerts directly via configured SMTP server
- **Slack Integration:** Post analysis results to Slack channels via API
- **Google Sheets Export:** Direct integration with Google Drive
- **PDF & Excel Downloads:** Full report generation
- **File Storage:** Server-side file management in Google Drive

**Tech Stack:** Node.js v20, Express, ExcelJS, XLSX

### 2. GitHub Pages — Browser Version (`gh-pages` branch)

A stripped-down, client-side-only version designed for browser limitations:

- **No Settings Page** — removed entirely
- **No SMTP Email** — replaced with `mailto:` link (opens default email client)
- **No Slack API** — replaced with copy-to-clipboard for pasting in Slack
- **Browser-Side Excel** — SheetJS CDN for reading .xlsx and .csv files
- **In-Memory Processing** — all data loaded and processed in the browser
- **Optimized Engine** — chunked processing with `yieldToUI()` to prevent memory crashes

---

## Fraud Detection Rules

Orders are flagged as **Potential Fraud** when any of the following conditions are met:

| Rule | Threshold |
|------|-----------|
| Total Sale unusually high | >= AED 10,000 |
| Duplicate Total Sale under same Restaurant ID | > 1 occurrence |
| Duplicate Total Sale under same Branch ID | > 1 occurrence |
| Duplicate Total Sale under same Customer Group | > 1 occurrence |
| Duplicate Total Sale under same Customer Email | > 1 occurrence |
| Customer Email frequency | Used > 2 times across flagged orders |
| Checkout vs Total Sale mismatch | Difference < -200 (credit card/talabat credit) |
| Cash high-value orders | Payment Method = Cash |

### Potential Fraud Filter (Second Pass)

After initial flagging, a second filter identifies the most suspicious records:

- Total Sale >= AED 10,000
- Email used multiple times across flagged orders
- Duplicate Total Sale with same Restaurant ID AND same Customer Email

---

## Data Sources (Upload Folders)

| Folder | File Type | Description |
|--------|-----------|-------------|
| **Detailed Reports** | `.xlsx` / `.csv` | Main transaction data with order details |
| **Checkout** | `.csv` | Payment gateway data with customer emails |
| **AX365** | `.xlsx` | AX365 branch code reference data |
| **Branch Inquiry** | `.xlsx` | Customer account and group mapping by branch |

### Column Mapping — Detailed Reports

| Column Index (0-based) | Field |
|------------------------|-------|
| 11 | Restaurant Id |
| 12 | Restaurant |
| 13 | Branch Id |
| 14 | Branch |
| 15 | Order Id |
| 17 | Date / Time |
| 18 | Payment Method |
| 19 | Status |
| 20 | Service Type |
| 24 | SubTotal |
| 31 | Total Sale |
| 63 | Net Payment Per Order |

---

## Features

- **Dashboard** — Summary stats, 4 interactive charts (by Customer Group, Date, Restaurant, Branch) with click-to-drill-down filtering
- **Potential Fraud Table** — Sortable, searchable table of all flagged orders
- **Suspicious Emails** — Top 20 emails ranked by suspicion score (frequency + burst patterns)
- **Upload Data** — Drag & drop or browse for each data folder
- **Run Analysis** — Execute fraud engine with real-time progress log
- **Share & Export** — Excel download, PDF snapshot, email (mailto/SMTP), Slack (copy/API), Google Drive link

---

## Repository

- **GitHub:** [elaizafrancisco-cyber/talabat-fraud-hunter](https://github.com/elaizafrancisco-cyber/talabat-fraud-hunter)
- **Branches:**
  - `master` — Full server app (auto-deploys to Render)
  - `gh-pages` — Browser-only version (served by GitHub Pages)

---

## Google Drive Location

```
G:\Shared drives\AR Team (Talabat UAE Finance) - Shared Drive\Claude\Fraud\FraudAnalysis\
```

**Google Drive URL:** [Open in Drive](https://drive.google.com/drive/folders/1-oMUrBczTD1tv5mSyRFe00_pTlDjg1yH)

---

## Brand Colors

| Color | Hex | Usage |
|-------|-----|-------|
| Primary | `#FF5900` | Buttons, accents, chart highlights |
| Accent | `#411517` | Sidebar, headings |
| Secondary | `#F4EDE3` | Backgrounds, alternating rows |
| Highlight | `#CFFF00` | Success indicators, lime accents |

---

*Powered by Talabat Platform — UAE Finance*
