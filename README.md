# 🏗️ Subcontractor Finder

AI-powered UK subcontractor discovery tool. Enter a trade package and UK region; Claude or Gemini researches suitable subcontractors and returns a structured, exportable table matched against your D&B database.

---

## Features

- **Dual AI support** — works with both Anthropic (Claude) and Google (Gemini) APIs
- **D&B database lookup** — upload your own Excel file; Turnover, C/Line, and D&B Risk columns are auto-populated by matching company name / registration number
- **Editable results table** — edit any cell before exporting
- **Export to Excel & CSV** — formatted output matching Brighton tab structure
- **Proximity scoring** — each company is scored 1–10 for how local they are to the target region
- **Filter & sort** — by D&B Risk, proximity score, or company name

---

## Quick Start (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Deploy to Streamlit Cloud (from GitHub)

1. **Push this folder to a GitHub repo** (e.g. `lor-subcontractor-finder`)

2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**

3. Fill in:
   - **Repository:** `your-github-username/lor-subcontractor-finder`
   - **Branch:** `main`
   - **Main file path:** `app.py`

4. Click **Deploy** — that's it. Streamlit Cloud reads `requirements.txt` automatically.

5. Share the generated URL with your procurement team.

---

## D&B (DNBi) Excel File Format

The app is built around the **real DNBi export format**, with these exact columns:

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

## Constructionline (C/Line) Excel File

Constructionline (C/Line) data is **not accessible to the AI** — it isn't a platform Claude or Gemini can query directly. Instead, the app has a **separate uploader** in the sidebar for a C/Line export (e.g. supplier list, membership/accreditation report).

- **No file uploaded:** the C/Line column shows `Not available`
- **File uploaded, company not found in it:** shows `Not registered`
- **File uploaded, company found:** shows the membership/status value (e.g. `Gold`, `SSIP Verified`)

Expected columns (auto-detected, names flexible): `Company Name`, `Registration Number`, and a status/level column such as `Status`, `Membership Level`, or `Accreditation Level`.

As soon as you upload a Constructionline export, the table will automatically start populating that column on the next search — no code changes required.

---

## Output Table Structure

Matches the **Brighton tab** format from the master Subcontractors Excel:

| Column | Description |
|--------|-------------|
| Company Name | Full legal company name |
| Registration No. | UK Companies House number |
| Trade Scope | AI-generated scope description |
| Close to Area | Proximity score (1–10) |
| Location | Registered / main office address |
| Contact | Phone and email |
| Turnover | From D&B database (£) |
| C/Line | Credit line tier from D&B |
| D&B Risk | Risk rating from D&B |
| Website | Company website |
| AI Notes | Additional AI commentary |

---

## API Keys

Keys are entered in the sidebar and **never stored** — they exist only for the current browser session.

- **Claude:** Get a key at [console.anthropic.com](https://console.anthropic.com)
- **Gemini:** Get a key at [aistudio.google.com](https://aistudio.google.com)

---

## Folder Structure

```
lor-subcontractor-finder/
├── app.py              ← main Streamlit application
├── requirements.txt    ← Python dependencies
└── README.md           ← this file
```
