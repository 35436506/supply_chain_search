# 🔍 Subcontractor Finder

AI-powered UK subcontractor discovery tool. Enter a trade package and UK region; the AI researches suitable subcontractors and returns a structured, exportable table cross-referenced against your D&B (DNBi) database.

Configured once by an admin — end users just search, no API keys or file uploads required.

---

## Features

- **Dual AI support** — works with both **ChatGPT (OpenAI, GPT-4o)** and **Gemini (Google, gemini-2.5-flash)**; switch between them with a toggle in the sidebar
- **D&B database lookup** — pre-loaded by the admin; Turnover, D&B Risk, and Location are auto-populated by matching company name / registration number
- **Editable results table** — edit any cell before exporting
- **Export to Excel & CSV**
- **Proximity scoring** — each company scored 1–10 for how local they are to the target region
- **Filter & sort** — by D&B Risk, proximity score, or company name
- **Light, high-contrast UI** — readable on any screen
- **Friendly error messages** — quota/billing issues (HTTP 429), invalid keys, and outdated model names are explained in plain language, with a suggestion to switch providers

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

### Step 2 — Commit the D&B database into the repo

Rename your D&B (DNBi) export to exactly **`dnb_database.xlsx`** and place it in the root of the repository, alongside `app.py`:

```
your-repo/
├── app.py
├── requirements.txt
├── dnb_database.xlsx   ← your D&B export, committed once
├── .gitignore
└── README.md
```

Push to GitHub — Streamlit Cloud redeploys automatically and loads the file at startup.

To update the database later, just replace `dnb_database.xlsx` in the repo and push. No code changes needed.

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
| Contact | Phone and email |
| Turnover | From D&B database (£) |
| D&B Risk | Risk rating from D&B |
| Website | Company website |
| AI Notes | Additional AI commentary |

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
├── .streamlit/secrets.toml.example ← template for local secrets
├── .gitignore
└── README.md
```
