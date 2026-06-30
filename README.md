# 🔍 Subcontractor Finder

AI-powered UK subcontractor discovery tool. Enter a trade package and UK region; the AI researches suitable subcontractors and returns a structured, exportable table cross-referenced against your D&B (DNBi) database.

Configured once by an admin — end users just search, no API keys or file uploads required.

---

## Features

- **Dual AI support** — works with both **ChatGPT (OpenAI, GPT-4o)** and **Gemini (Google, gemini-2.5-flash)**; switch between them with a toggle in the sidebar
- **Work Package–driven trade selection** — dropdown populated from your real WP list; free-text search box for trades not in the list
- **Flexible area targeting** — pick a broad UK region, or type a specific town/city/postcode for tighter results
- **Preferred Supplier prioritisation across multiple clusters** — upload any number of cluster preferred-supplier exports at once (or have the admin commit a folder of them); all are combined, each company tagged with its source cluster, and matches are flagged and sorted to the top when relevant to the chosen trade
- **D&B database lookup** — pre-loaded by the admin; Turnover and D&B Risk are auto-populated by matching company name / registration number
- **Turnover-based company sizing** — Company Size band and an estimated Number of Employees are derived automatically from turnover (D&B figure, or an AI-researched estimate when D&B has no record)
- **Interactive supplier map** — plots suggested suppliers on a UK map; bubble size reflects company turnover, preferred suppliers shown in a different colour
- **Smart sorting** — results ranked by Turnover → Company Size → Proximity to area, with preferred suppliers always prioritised
- **Editable results table** — edit any cell before exporting
- **Export to Excel**
- **Light, high-contrast UI** — readable on any screen
- **Friendly error messages** — quota/billing issues (HTTP 429), temporary overload (503), invalid keys, and outdated model names are explained in plain language

---

## One-time Admin Setup

This is the part you do **once**. After this, regular users only see the search boxes — no API key field, no file upload.

### Step 1 — Add your API key(s) as Streamlit secrets

In Streamlit Cloud: open your app → **Settings → Secrets** → paste:

```toml
OPENAI_API_KEY = "sk-your-openai-key-here"
GEMINI_API_KEY = "AIza-your-gemini-key-here"
```

You can configure one or both. The provider toggle in the sidebar shows "Ready to search" for whichever key is present, letting you (or users) switch providers if one runs out of quota.

Save. The app reads these automatically; they are never shown to end users.

(For local development, copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill in the real keys. That file is git-ignored so it won't be committed.)

### Step 2 — Commit the reference files into the repo

Place these in the root of the repository, alongside `app.py`:

```
your-repo/
├── app.py
├── requirements.txt
├── dnb_database.xlsx              ← D&B (DNBi) risk/turnover export
├── wp_list.xlsx                   ← Work Package list (trade/package descriptions)
├── preferred_suppliers/           ← folder — add as many cluster files as you like
│   ├── cluster_1.xlsx
│   ├── cluster_2.xlsx
│   └── cluster_3.xlsx
├── .gitignore
└── README.md
```

Every `.xlsx` file inside `preferred_suppliers/` is loaded and combined automatically — no limit on how many you add. Each company is tagged with which cluster file it came from (shown in the **Preferred Cluster** column in results). To add a new cluster later, just drop another file into that folder and push.

Push to GitHub — Streamlit Cloud redeploys automatically and loads everything at startup.

None of these files are strictly required for the app to run — any piece the admin hasn't provided yet simply falls back to a manual upload box, shown to whoever is searching, for that session only (the supplier uploader accepts multiple files at once too, so several clusters can be combined on the fly).

To update any of these later, just replace the file(s) in the repo and push. No code changes needed.

### Verifying setup

Once both are in place, the app sidebar will show **"✅ Ready to search"**. If either is missing, a warning is shown along with manual fallback fields so the app still works while you finish setup.

---

## Quick Start (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Deploy to Streamlit Cloud (from GitHub)

1. Push this folder to a GitHub repo (e.g. `subcontractor-finder`)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Fill in:
   - **Repository:** `your-github-username/subcontractor-finder`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**
5. Complete the **Admin Setup** steps above (Secrets + `dnb_database.xlsx`)
6. Share the generated URL with your team

---

## D&B (DNBi) Excel File Format

Built around the real DNBi export format:

| Column | Example |
|--------|---------|
| Company Name | `7 STEEL (UK) LIMITED` |
| Business Registration Number | `04661575` |
| Sales (Revenue) | `461,104,000` |
| Overall Business Risk | `low-moderate` |
| City - D&B | `CARDIFF` |
| Zip/Postal Code - D&B | `CF24 5NN` |
| Country or Region - D&B | `GB` |

Minor naming variations (e.g. "Registration No.", "Turnover", "Risk") are also auto-detected. Multiple sheets are supported and combined automatically.

Risk values recognised: `low`, `low-moderate`, `moderate`, `high`, `Severe`, `Undetermined`, `Out of Business`.

---

## Output Table Structure

| Column | Description |
|--------|-------------|
| Company Name | Full legal company name |
| Registration No. | UK Companies House number |
| Trade Scope | AI-generated scope description |
| Close to Area | Proximity score (1–10) |
| Location | Registered / main office address (from D&B where matched) |
| Turnover | D&B figure where available, otherwise an AI-researched estimate |
| Turnover Source | "D&B" or "Estimated" |
| Company Size | Band derived from turnover (Micro / Small / Medium / Large / Major) |
| Number of Employees | Rough estimate derived from turnover |
| D&B Risk | Risk rating from D&B |
| Preferred Supplier | "Yes" if found in any loaded Preferred Supplier list |
| Preferred Cluster | Which cluster file the preferred match came from |
| Website | Company website |
| Notes | Additional commentary |

Results are sorted by **Turnover (desc) → Company Size (desc) → Proximity (desc)**, with preferred suppliers always shown first.

---

## API Keys

Supports two providers:

- **OpenAI (ChatGPT / GPT-4o)** — get a key at [platform.openai.com/api-keys](https://platform.openai.com/api-keys). A **429 "insufficient_quota" error** means this key's billing/credit needs attention — it's an account issue, not an app bug.
- **Gemini (Google)** — get a key at [aistudio.google.com/apikey](https://aistudio.google.com/apikey). Uses `gemini-2.5-flash`.

If neither admin key works at a given moment (e.g. one hits quota), switch providers using the toggle at the top of the sidebar — no code change needed.

If the admin hasn't configured a key yet, the sidebar shows a manual input field as a temporary fallback (entered keys are session-only, never stored).

**Security note:** never paste real API keys into chat tools, emails, or shared documents. If a key has ever been shared outside of Streamlit's encrypted Secrets manager, revoke and regenerate it immediately from the provider's dashboard.

---

## Folder Structure

```
subcontractor-finder/
├── app.py                          ← main Streamlit application
├── requirements.txt                ← Python dependencies
├── dnb_database.xlsx               ← admin-provided D&B export (you add this)
├── wp_list.xlsx                    ← admin-provided Work Package list (you add this)
├── preferred_suppliers/            ← admin-provided cluster files (you add these)
│   ├── cluster_1.xlsx
│   └── cluster_2.xlsx
├── .streamlit/secrets.toml.example ← template for local secrets
├── .gitignore
└── README.md
```
