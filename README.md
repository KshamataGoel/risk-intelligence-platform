# Risk Intelligence Platform

A unified, AI-powered risk intelligence web application for CRO leadership teams. Provides natural language insights, automated reporting, document management, and advanced AML cluster analysis — all grounded in approved internal datasets.

---

## Modules

| Module | Description |
|--------|-------------|
| **Insights** | Ask natural language questions against Excel datasets. Returns AI-generated summaries, dynamic charts, and stat cards. |
| **Reports** | Generate structured 10-section executive intelligence reports and leadership email drafts from Vault documents and Insights data. |
| **Vault** | Browse, search, and summarise internal risk documents (Word files) including regulatory sources, risk analyses, and emails. |
| **Analysis** | K-Means cluster analysis, feature catalogue browser, scatter plot visualisation, and anomaly detection — run on demand via the Run Model button. |

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend | Node.js + Express |
| Frontend | React (via CDN + Babel standalone — no build step) |
| AI | Groq API (`llama-3.1-8b-instant`) |
| Excel parsing | SheetJS (`xlsx`) |
| Word parsing | `mammoth` |
| Charts | Chart.js 4 |
| HTTP client | `axios` |

---

## Project Structure

```
RiskIntelligencePlatform-v2/
├── server.js               # All backend routes and data processing
├── public/
│   └── index.html          # All frontend UI, React components, CSS
├── data/
│   ├── insights/           # Place Dataset.xlsx here
│   ├── vault/              # Place .docx risk documents here
│   └── analysis/           # Place Analysis.xlsx here
├── .env                    # API keys and config (created from .env.example)
├── .env.example            # Template for environment variables
├── start.bat               # Windows one-click launcher
└── package.json
```

---

## Quick Start

### Option 1 — Windows (double-click)
```
start.bat
```
This installs dependencies, creates `.env` if missing, and starts the server.

### Option 2 — Manual
```bash
# 1. Install dependencies (first time only)
npm install

# 2. Set up environment
copy .env.example .env
# Edit .env and add your GROQ_API_KEY

# 3. Start the server
node server.js

# 4. Open in browser
# http://localhost:3000
```

---

## Environment Variables

Create a `.env` file in the project root (or copy from `.env.example`):

```env
GROQ_API_KEY=your_groq_api_key_here
GROQ_MODEL=llama-3.1-8b-instant
PORT=3000
```

Get a free Groq API key at [console.groq.com](https://console.groq.com).

---

## Data Setup

### Insights — `data/insights/Dataset.xlsx`
Required sheets:

| Sheet | Purpose |
|-------|---------|
| `Parties` | Party type and count breakdown |
| `Risk_Ratings` | Customer risk tier distribution |
| `Top10_Accounts_Volume` | Top accounts by transaction volume |
| `Monthly_TXN_Trends` | Monthly transaction volumes |
| `Top10_Counterparty_Exposures` | Top counterparty exposure values |

### Vault — `data/vault/*.docx`
Place Word documents here. Supported document types:
- Risk Analysis reports
- Regulatory source documents
- Internal emails
- Any compliance or policy documents

Each document should include metadata at the top (Date, Author, Classification, Status).

### Analysis — `data/analysis/Analysis.xlsx`
Required sheets:

| Sheet | Purpose |
|-------|---------|
| `Feature_Catalogue` | AML feature definitions, categories, tags, and risk relevance |
| `Scatter_Plot_Data` | Party-level data with cluster labels for scatter visualisation |
| `Anomaly_Detection` | Pre-computed anomaly records with distance, type, and investigation status |

---

## Key Features

### Insights
- Natural language question routing to the correct dataset
- Context-aware chart highlighting (largest, smallest, or specifically mentioned segment)
- AI executive summary grounded strictly in dataset values — no hallucinated figures
- Suggested follow-up questions
- Data preview table

### Reports
- Select one or more source types (Risk Analysis, Regulatory, Email Intelligence, Insights)
- Date range filter with timezone-safe parsing
- 10-section AI report: Executive Summary, Key Risk Themes, Regulatory Developments, Customer & Transaction Insights, Email Intelligence, Required Actions, Deadlines, Ownership & Teams, Leadership Attention Required, Supporting Evidence
- Chart.js visualisations from Insights data embedded in report
- Generate Email — produces a leadership-ready email draft in a child window modal with Copy button
- Print-friendly layout

### Vault
- Full-text search across all documents
- Filter by section (Risk Analysis, Regulatory, Email Intelligence)
- Document summary modal — AI-generated structured brief (Executive Summary, Key Findings, Risk/Regulatory Impact, Required Actions, Deadlines, Owners, Leadership Attention)
- Author, date, classification metadata extraction from Word documents

### Analysis
1. Browse the **Feature Catalogue** — search, filter by category and risk relevance
2. Select 2+ features to enable the model
3. Click **Run Model** to generate:
   - K-Means scatter plot with cluster colour coding and axis selection
   - AI-generated cluster descriptions
   - Anomaly detection table with Cluster, Risk, Status, and Type filters
   - Insights Quick Summary — AI narrative combining cluster patterns and anomaly risk
4. **Refresh Data** button to reload the Excel without restarting the server

---

## Stopping / Restarting the Server

If port 3000 is already in use:

```powershell
# PowerShell — kill all node processes
Get-Process -Name "node" -ErrorAction SilentlyContinue | Stop-Process -Force
```

Then run `node server.js` again.

---

## Notes

- All AI responses are grounded in uploaded datasets only — no external web data is used
- Groq API token limits apply; document content is truncated to stay within limits
- The frontend uses React via CDN (no npm build required for the UI)
- Editing `public/index.html` takes effect on browser refresh; editing `server.js` requires a server restart
