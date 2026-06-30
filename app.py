import streamlit as st
import pandas as pd
import json
import io
import re
import os
import time
from datetime import datetime
import requests

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Subcontractor Finder",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS — light background, dark readable text everywhere ─────────────
st.markdown("""
<style>
    /* Sidebar: light background, dark text */
    [data-testid="stSidebar"] { background: #f4f6f9; border-right: 1px solid #dde3ea; }
    [data-testid="stSidebar"] * { color: #1a1f2b !important; }
    [data-testid="stSidebar"] label { color: #3a4150 !important; font-size: 0.82rem; font-weight: 500; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #1a3a5c !important; }
    [data-testid="stSidebar"] .stCaption, [data-testid="stSidebar"] small { color: #5a6272 !important; }

    /* Header */
    .app-header {
        background: #ffffff;
        color: #1a1f2b;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        border: 1px solid #dde3ea;
        border-left: 5px solid #2563eb;
    }
    .app-header h1 { margin:0; font-size: 1.6rem; color: #1a1f2b; }
    .app-header p  { margin:0.3rem 0 0; color: #5a6272; font-size: 0.9rem; }

    /* Metric cards */
    .metric-row { display:flex; gap:1rem; margin-bottom:1.5rem; flex-wrap:wrap; }
    .metric-card {
        background: white;
        border: 1px solid #e0e8f0;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        flex: 1;
        min-width: 140px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    .metric-card .val { font-size: 2rem; font-weight: 700; color: #1a3a5c; }
    .metric-card .lbl { font-size: 0.75rem; color: #5a6272; text-transform: uppercase; letter-spacing:0.05em; margin-top:0.2rem; }

    /* D&B risk badges */
    .badge {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.75rem;
        font-weight: 600;
        color: white;
    }
    .badge-low { background:#2e7d32; }
    .badge-low-mod { background:#558b2f; }
    .badge-mod { background:#f57c00; }
    .badge-high { background:#c62828; }
    .badge-unknown { background:#546e7a; }

    /* Section headings */
    .section-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1a3a5c;
        border-bottom: 2px solid #2563eb;
        padding-bottom: 0.4rem;
        margin: 1.5rem 0 1rem;
    }

    /* AI response box */
    .ai-box {
        background: #f0f6ff;
        border: 1px solid #b0c8e8;
        border-left: 4px solid #1a3a5c;
        border-radius: 6px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
        font-size: 0.9rem;
        line-height: 1.6;
        white-space: pre-wrap;
        color: #1a1f2b;
    }

    .stDataFrame { border-radius: 8px; overflow: hidden; }

    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# ADMIN CONFIGURATION
# ════════════════════════════════════════════════════════════════════════════
# This app supports a one-time admin setup so end users never need to
# enter an API key or upload the D&B file themselves.
#
# Two ways to provide the OpenAI API key (checked in this order):
#   1. Streamlit secrets:  .streamlit/secrets.toml -> OPENAI_API_KEY = "sk-..."
#      (Set this in Streamlit Cloud: App settings -> Secrets)
#   2. Environment variable: OPENAI_API_KEY
#
# Two ways to provide the D&B database (checked in this order):
#   1. Commit a file named "dnb_database.xlsx" into this repo's root folder.
#      The app loads it automatically on startup — no upload needed.
#   2. If neither secret key nor committed file exists, the app falls back
#      to showing manual input boxes (so it still works before setup).
# ════════════════════════════════════════════════════════════════════════════

DEFAULT_DB_PATH = os.path.join(os.path.dirname(__file__), "dnb_database.xlsx")


def get_admin_api_key(provider: str) -> str:
    """Look for an admin-configured key in secrets or env vars, per provider."""
    secret_name = "OPENAI_API_KEY" if provider == "ChatGPT (OpenAI)" else "GEMINI_API_KEY"
    try:
        if secret_name in st.secrets:
            return st.secrets[secret_name]
    except Exception:
        pass
    return os.environ.get(secret_name, "")


def get_admin_db_df() -> pd.DataFrame | None:
    """Load the committed D&B database file if it exists in the repo."""
    if os.path.exists(DEFAULT_DB_PATH):
        try:
            return load_excel_any(DEFAULT_DB_PATH)
        except Exception as e:
            st.sidebar.error(f"Could not read bundled D&B file: {e}")
    return None


# ════════════════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════════════════

UK_REGIONS = [
    "London", "South East (Brighton, Guildford, Maidstone)",
    "South West (Bristol, Exeter, Plymouth)",
    "East of England (Cambridge, Norwich, Ipswich)",
    "East Midlands (Nottingham, Leicester, Derby)",
    "West Midlands (Birmingham, Coventry, Wolverhampton)",
    "Yorkshire & Humber (Leeds, Sheffield, Hull)",
    "North West (Manchester, Liverpool, Preston)",
    "North East (Newcastle, Sunderland, Durham)",
    "Scotland (Edinburgh, Glasgow, Aberdeen)",
    "Wales (Cardiff, Swansea, Newport)",
    "Northern Ireland (Belfast)",
]

TRADE_PACKAGES = [
    "Mechanical & Public Health (Drainage, Pipework, HVAC)",
    "Electrical & ELV",
    "Civil & Groundworks (Drainage, Earthworks, Piling)",
    "Structural Steelwork",
    "Concrete Frame & Formwork",
    "Cladding & Curtain Wall",
    "Roofing & Waterproofing",
    "Façade & External Works",
    "Fit-Out & Joinery",
    "Flooring (Screeds, Raised Access, Finishes)",
    "Partition & Drylining",
    "Painting & Decorating",
    "Fire Protection & Suppression",
    "Lifts & Escalators",
    "Demolition & Enabling Works",
    "Landscaping & External Works",
    "Temporary Works & Shoring",
    "Specialist Concrete (Post-tension, Precast)",
    "MEP Commissioning",
    "Building Management Systems (BMS)",
]

DB_RISK_COLOURS = {
    "low": "badge-low",
    "low-moderate": "badge-low-mod",
    "moderate": "badge-mod",
    "high": "badge-high",
    "severe": "badge-high",
    "undetermined": "badge-unknown",
    "out-of-business": "badge-high",
}


def risk_badge(risk_str: str) -> str:
    if not risk_str or str(risk_str).strip() == "":
        return '<span class="badge badge-unknown">Unknown</span>'
    r = str(risk_str).strip().lower().replace(" ", "-")
    css = DB_RISK_COLOURS.get(r, "badge-unknown")
    return f'<span class="badge {css}">{risk_str}</span>'


def load_excel_any(file_or_path) -> pd.DataFrame | None:
    """Generic Excel loader — combines all sheets. Accepts a path or uploaded file."""
    try:
        xl = pd.ExcelFile(file_or_path)
        frames = []
        for sh in xl.sheet_names:
            df = xl.parse(sh)
            df["_source_sheet"] = sh
            frames.append(df)
        return pd.concat(frames, ignore_index=True) if frames else None
    except Exception as e:
        st.sidebar.error(f"Could not read file: {e}")
        return None


def _normalise_cols(df: pd.DataFrame) -> dict:
    col_map = {}
    for col in df.columns:
        cl = str(col).lower()
        for ch in [" ", "_", "-", ".", "&", "(", ")", "/"]:
            cl = cl.replace(ch, "")
        col_map[cl] = col
    return col_map


def find_db_columns(db_df: pd.DataFrame) -> dict:
    """Detect D&B (DNBi) export columns — matches the real DNBi export format:
    Company Name | Business Registration Number | Sales (Revenue) | Overall Business Risk |
    City - D&B | Zip/Postal Code - D&B | Country or Region - D&B
    """
    cm = _normalise_cols(db_df)
    return {
        "name":    cm.get("companyname") or cm.get("name") or cm.get("company"),
        "reg":     cm.get("businessregistrationnumber") or cm.get("registrationno") or cm.get("regno")
                   or cm.get("companynumber") or cm.get("companynr"),
        "sales":   cm.get("salesrevenue") or cm.get("sales") or cm.get("turnover") or cm.get("revenue"),
        "risk":    cm.get("overallbusinessrisk") or cm.get("businessrisk") or cm.get("dbrisk") or cm.get("risk"),
        "city":    cm.get("citydb") or cm.get("city"),
        "zip":     cm.get("zippostalcodedb") or cm.get("zippostalcode") or cm.get("postcode") or cm.get("zip"),
        "country": cm.get("countryorregiondb") or cm.get("countryorregion") or cm.get("country"),
    }


def _match_row(df: pd.DataFrame, name_col, reg_col, company_name: str, reg_no: str = ""):
    """Find best matching row by company name (exact -> partial) or registration number."""
    if name_col is None or name_col not in df.columns:
        return None

    name_norm = str(company_name).lower().strip()
    mask = df[name_col].astype(str).str.lower().str.strip() == name_norm

    if not mask.any():
        mask = df[name_col].astype(str).str.lower().str.contains(
            re.escape(name_norm[:15]), na=False
        )

    if not mask.any() and reg_no and reg_col and reg_col in df.columns:
        mask = df[reg_col].astype(str).str.strip().str.upper() == str(reg_no).strip().upper()

    if not mask.any():
        return None

    return df[mask].iloc[0]


def lookup_db(db_df: pd.DataFrame, company_name: str, reg_no: str = "") -> dict:
    """Look up a company in the D&B (DNBi) database export."""
    if db_df is None or db_df.empty:
        return {}

    cols = find_db_columns(db_df)
    row = _match_row(db_df, cols["name"], cols["reg"], company_name, reg_no)
    if row is None:
        return {}

    def safe(key):
        col = cols.get(key)
        if col and col in db_df.columns:
            v = row[col]
            return "" if pd.isna(v) else str(v)
        return ""

    city = safe("city")
    zip_ = safe("zip")
    country = safe("country")
    location = ", ".join([p for p in [city, zip_, country] if p])

    return {
        "D&B Risk":     safe("risk"),
        "Turnover":     safe("sales"),
        "D&B Location": location,
        "_db_matched":  True,
    }


def _post_with_retry(url, headers, body, timeout=60, max_retries=3, backoff_base=2):
    """POST with retry on transient errors (503 overloaded, 429 with Retry-After,
    connection resets). Raises the last error if all retries are exhausted."""
    last_exc = None
    for attempt in range(max_retries):
        try:
            r = requests.post(url, headers=headers, json=body, timeout=timeout)
            if r.status_code in (503, 502, 504):
                last_exc = requests.HTTPError(response=r)
                if attempt < max_retries - 1:
                    time.sleep(backoff_base * (2 ** attempt))
                    continue
            r.raise_for_status()
            return r
        except requests.HTTPError:
            raise
        except requests.RequestException as e:
            last_exc = e
            if attempt < max_retries - 1:
                time.sleep(backoff_base * (2 ** attempt))
                continue
    if last_exc:
        raise last_exc


def call_openai(api_key: str, prompt: str) -> str:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    body = {
        "model": "gpt-4o",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 2500,
    }
    r = _post_with_retry("https://api.openai.com/v1/chat/completions", headers, body)
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"]


def call_gemini(api_key: str, prompt: str) -> str:
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
    body = {"contents": [{"parts": [{"text": prompt}]}]}
    r = _post_with_retry(url, headers, body)
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def call_ai(provider: str, api_key: str, prompt: str) -> str:
    if provider == "ChatGPT (OpenAI)":
        return call_openai(api_key, prompt)
    else:
        return call_gemini(api_key, prompt)


def friendly_api_error(provider: str, e: requests.HTTPError) -> str:
    status = e.response.status_code
    body = e.response.text[:300]
    if status == 503:
        other = "ChatGPT" if provider == "Gemini (Google)" else "Gemini"
        return (f"{provider} is temporarily overloaded (503) — Google/OpenAI's servers are at capacity right now. "
                f"This usually clears within a minute or two. Try **Find Subcontractors** again, "
                f"or switch to {other} using the provider toggle while you wait.")
    if status == 429:
        if provider == "ChatGPT (OpenAI)":
            return ("OpenAI quota exceeded (429) — this API key has no remaining credit or hit its rate limit. "
                    "Check billing at platform.openai.com/account/billing, or switch to Gemini using the provider selector.")
        else:
            return ("Gemini quota exceeded (429) — this API key hit its rate or usage limit. "
                    "Check quota at aistudio.google.com, or switch to ChatGPT using the provider selector.")
    if status == 401 or status == 403:
        return f"{provider} rejected this API key ({status}) — it may be invalid, revoked, or missing permissions."
    if status == 404:
        return f"{provider} model not found (404) — the model name in the code may be outdated. Details: {body}"
    return f"{provider} API error ({status}): {body}"


def build_ai_prompt(trade: str, region: str, extra_notes: str) -> str:
    return f"""You are a UK construction procurement specialist.

A procurement team member needs you to identify suitable subcontractors for the following:

**Trade / Package:** {trade}
**Project Region / Area:** {region}
**Additional Notes:** {extra_notes or 'None'}

Please provide:
1. A brief market overview for this trade in the given UK region (2-3 sentences).
2. A list of **8–12 recommended subcontractors** — ideally a mix of large, mid-tier, and specialist firms.

For EACH company return a JSON object inside a ```json block with this EXACT schema:
{{
  "companies": [
    {{
      "company_name": "...",
      "registration_no": "...",
      "trade_scope": "...",
      "location": "...",
      "contact": "...",
      "website": "...",
      "proximity_score": <integer 1-10 where 10 = based right in the area>,
      "notes": "..."
    }}
  ]
}}

Rules:
- Use real UK Companies House registration numbers where known, else leave blank.
- Trade scope should be 1-2 specific sentences about what they deliver relevant to this package.
- Location should be the company's registered / main office address.
- Proximity score: 10 = headquartered in or immediately adjacent to the target area; 1 = national but mobilisable.
- Be factual; do not invent contact details if unsure — leave blank.
- Respond with your market overview text FIRST, then the JSON block.
"""


def parse_ai_response(text: str):
    """Extract narrative + list of companies from AI response."""
    json_match = re.search(r"```json\s*(.*?)\s*```", text, re.DOTALL)
    narrative = text
    companies = []

    if json_match:
        narrative = text[: json_match.start()].strip()
        try:
            payload = json.loads(json_match.group(1))
            companies = payload.get("companies", [])
        except json.JSONDecodeError:
            pass

    return narrative, companies


def companies_to_df(companies: list, db_df, region: str) -> pd.DataFrame:
    rows = []
    for c in companies:
        name = c.get("company_name", "")
        reg  = c.get("registration_no", "")

        db_info = lookup_db(db_df, name, reg)
        location = db_info.get("D&B Location") or c.get("location", "")

        rows.append({
            "Company Name":      name,
            "Registration No.":  reg,
            "Trade Scope":       c.get("trade_scope", ""),
            "Close to Area":     c.get("proximity_score", ""),
            "Location":          location,
            "Contact":           c.get("contact", ""),
            "Turnover":          db_info.get("Turnover", ""),
            "D&B Risk":          db_info.get("D&B Risk", ""),
            "Website":           c.get("website", ""),
            "AI Notes":          c.get("notes", ""),
            "_db_matched":       bool(db_info.get("_db_matched")),
        })
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, trade: str, region: str) -> bytes:
    buf = io.BytesIO()
    export_df = df.drop(columns=[c for c in df.columns if c.startswith("_")], errors="ignore")
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Subcontractors")
        wb  = writer.book
        ws  = writer.sheets["Subcontractors"]

        hdr_fmt = wb.add_format({
            "bold": True, "bg_color": "#1a3a5c", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        cell_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
        money_fmt = wb.add_format({"border": 1, "valign": "top", "num_format": "£#,##0"})

        col_widths = [30, 16, 50, 12, 35, 30, 14, 18, 35, 45]
        for i, (col, w) in enumerate(zip(export_df.columns, col_widths)):
            ws.set_column(i, i, w, cell_fmt)
            ws.write(0, i, col, hdr_fmt)

        turn_idx = list(export_df.columns).index("Turnover") if "Turnover" in export_df.columns else -1
        for row_idx, row in export_df.iterrows():
            for col_idx, val in enumerate(row):
                if col_idx == turn_idx:
                    try:
                        ws.write_number(row_idx + 1, col_idx, int(float(str(val).replace(",", "").replace("£", "").strip())), money_fmt)
                        continue
                    except Exception:
                        pass
                ws.write(row_idx + 1, col_idx, str(val) if val else "", cell_fmt)

        ws.set_row(0, 30)

        meta = wb.add_worksheet("Search Info")
        meta.write("A1", "Trade Package")
        meta.write("B1", trade)
        meta.write("A2", "Region")
        meta.write("B2", region)
        meta.write("A3", "Exported")
        meta.write("B3", datetime.now().strftime("%Y-%m-%d %H:%M"))

    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# LOAD ADMIN CONFIG
# ════════════════════════════════════════════════════════════════════════════

admin_db_df = get_admin_db_df()

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔍 Search")

    ai_provider = st.radio("AI Provider", ["ChatGPT (OpenAI)", "Gemini (Google)"], horizontal=True)
    admin_api_key = get_admin_api_key(ai_provider)

    if admin_api_key and admin_db_df is not None:
        st.success("✅ Ready to search — no setup needed.")
        api_key = admin_api_key
        db_df = admin_db_df
        st.caption(f"D&B database loaded: {len(db_df):,} records")
    elif admin_api_key and admin_db_df is None:
        st.warning("⚠️ D&B database not found in repo. Add `dnb_database.xlsx` — see Admin Setup below.")
        api_key = admin_api_key
        db_df = None
    else:
        st.warning(f"⚠️ No admin key configured for {ai_provider}. Enter one below for this session, or see Admin Setup at the bottom of the page.")
        api_key = st.text_input(f"{ai_provider} API Key", type="password", placeholder="sk-..." if "OpenAI" in ai_provider else "AIza...")
        db_df = admin_db_df
        if db_df is None:
            db_file = st.file_uploader("D&B Excel export (temporary, this session only)", type=["xlsx", "xls"])
            db_df = load_excel_any(db_file) if db_file else None

    st.markdown("---")
    st.markdown("### Trade & Area")
    trade  = st.selectbox("Trade / Package", ["— Select —"] + TRADE_PACKAGES)
    region = st.selectbox("UK Region / Area", ["— Select —"] + UK_REGIONS)
    extra  = st.text_area("Additional Notes", placeholder="e.g. CHAS accredited, min turnover £10M, experience in healthcare...", height=100)

    search_btn = st.button("🔍 Find Subcontractors", use_container_width=True, type="primary")

    st.markdown("---")
    st.markdown("### Previous Searches")
    if "history" not in st.session_state:
        st.session_state.history = []
    if st.session_state.history:
        for h in reversed(st.session_state.history[-5:]):
            st.caption(f"• {h['trade'][:35]}… | {h['region'][:20]}")
    else:
        st.caption("No searches yet.")


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
  <h1>🔍 Subcontractor Finder</h1>
  <p>AI-powered UK subcontractor discovery, cross-referenced against your D&B database</p>
</div>
""", unsafe_allow_html=True)

with st.expander("ℹ️ How to use this tool", expanded=False):
    st.markdown("""
    1. **Choose an AI provider** (ChatGPT or Gemini) in the sidebar.
    2. **Select a Trade Package and UK Region**, add any extra notes (minimum turnover, accreditations, etc.).
    3. Click **Find Subcontractors** — AI researches suitable suppliers and the app cross-references them against the D&B database.
    4. **Review and edit** the table inline, then **export to Excel or CSV**.

    If you see a setup warning in the sidebar, the admin API key(s) and/or D&B database haven't been configured yet for this deployment — see **Admin Setup** at the bottom of this page.
    """)

if "result_df"   not in st.session_state: st.session_state.result_df   = None
if "narrative"   not in st.session_state: st.session_state.narrative   = ""
if "last_trade"  not in st.session_state: st.session_state.last_trade  = ""
if "last_region" not in st.session_state: st.session_state.last_region = ""

# ── Search execution ──────────────────────────────────────────────────────────
if search_btn:
    if trade == "— Select —" or region == "— Select —":
        st.warning("Please select both a Trade Package and a UK Region before searching.")
    elif not api_key:
        st.error(f"No {ai_provider} API key available. Ask your admin to configure one, or enter one in the sidebar.")
    else:
        with st.spinner(f"Searching for {trade} subcontractors in {region}…"):
            prompt = build_ai_prompt(trade, region, extra)
            try:
                raw = call_ai(ai_provider, api_key, prompt)
                narrative, companies = parse_ai_response(raw)
                df = companies_to_df(companies, db_df, region)

                st.session_state.result_df   = df
                st.session_state.narrative   = narrative
                st.session_state.last_trade  = trade
                st.session_state.last_region = region
                st.session_state.raw_ai      = raw

                st.session_state.history.append({"trade": trade, "region": region})

            except requests.HTTPError as e:
                st.error(friendly_api_error(ai_provider, e))
            except Exception as e:
                st.error(f"Error: {e}")

# ── Results display ───────────────────────────────────────────────────────────
if st.session_state.result_df is not None:
    df   = st.session_state.result_df
    nav  = st.session_state.narrative
    trde = st.session_state.last_trade
    rgn  = st.session_state.last_region

    total      = len(df)
    with_db    = int(df.get("_db_matched", pd.Series([False]*len(df))).sum()) if "_db_matched" in df.columns else 0
    low_risk   = df["D&B Risk"].str.lower().str.contains("low", na=False).sum()
    local_cnt  = (pd.to_numeric(df["Close to Area"], errors="coerce") >= 7).sum()

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="val">{total}</div><div class="lbl">Subcontractors Found</div></div>
      <div class="metric-card"><div class="val">{with_db}</div><div class="lbl">D&B Records Matched</div></div>
      <div class="metric-card"><div class="val">{low_risk}</div><div class="lbl">Low / Low-Mod Risk</div></div>
      <div class="metric-card"><div class="val">{local_cnt}</div><div class="lbl">Locally Based (Score ≥7)</div></div>
    </div>
    """, unsafe_allow_html=True)

    if nav:
        st.markdown('<div class="section-title">📊 Market Overview</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="ai-box">{nav}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">🔧 Filter & Export</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        risk_opts = ["All"] + sorted(df["D&B Risk"].dropna().unique().tolist())
        risk_filt = st.selectbox("D&B Risk", risk_opts)
    with col2:
        min_prox = st.slider("Min Proximity Score", 1, 10, 1)
    with col3:
        sort_by = st.selectbox("Sort by", ["Close to Area ↓", "Company Name ↑", "D&B Risk"])

    view = df.copy()
    if risk_filt != "All":
        view = view[view["D&B Risk"] == risk_filt]
    prox_numeric = pd.to_numeric(view["Close to Area"], errors="coerce").fillna(0)
    view = view[prox_numeric >= min_prox]

    if sort_by == "Close to Area ↓":
        view = view.copy()
        view["_sort"] = pd.to_numeric(view["Close to Area"], errors="coerce")
        view = view.sort_values("_sort", ascending=False).drop(columns=["_sort"])
    elif sort_by == "Company Name ↑":
        view = view.sort_values("Company Name")
    elif sort_by == "D&B Risk":
        order = {"Low": 0, "Low-Moderate": 1, "Moderate": 2, "High": 3}
        view["_sort"] = view["D&B Risk"].map(order).fillna(99)
        view = view.sort_values("_sort").drop(columns=["_sort"])

    display_cols = [c for c in view.columns if not c.startswith("_")]

    st.markdown(f"**{len(view)} companies shown** — you can edit cells directly:")
    edited = st.data_editor(
        view[display_cols].reset_index(drop=True),
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Website":       st.column_config.LinkColumn("Website"),
            "Close to Area": st.column_config.NumberColumn("Close to Area", min_value=1, max_value=10),
            "Trade Scope":   st.column_config.TextColumn("Trade Scope", width="large"),
            "AI Notes":      st.column_config.TextColumn("AI Notes",    width="large"),
            "Turnover":      st.column_config.NumberColumn("Turnover (£)", format="£%d"),
        },
        hide_index=True,
        height=500,
    )

    st.markdown('<div class="section-title">⬇️ Export</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)
    with col_a:
        xlsx_bytes = to_excel_bytes(edited, trde, rgn)
        fname = f"Subcontractors_{rgn[:15].replace(' ','_')}_{trde[:20].replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            "📥 Download Excel",
            data=xlsx_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_b:
        csv_bytes = edited.to_csv(index=False).encode()
        st.download_button(
            "📄 Download CSV",
            data=csv_bytes,
            file_name=fname.replace(".xlsx", ".csv"),
            mime="text/csv",
            use_container_width=True,
        )

    with st.expander("🤖 Raw AI Response", expanded=False):
        st.text(st.session_state.get("raw_ai", ""))

else:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#5a6272;">
        <div style="font-size:4rem;">🔍</div>
        <h3 style="color:#3a4150;">Ready to search</h3>
        <p>Select a Trade Package and UK Region in the sidebar, then click <strong>Find Subcontractors</strong>.</p>
    </div>
    """, unsafe_allow_html=True)

# ── Admin setup instructions (always visible at bottom, collapsed) ───────────
st.markdown("---")
with st.expander("🛠️ Admin Setup (one-time, do this so users never need to upload anything)", expanded=False):
    st.markdown("""
    **Step 1 — Add your API keys as Streamlit secrets**

    In Streamlit Cloud: open your app → **Settings → Secrets** → paste:
    ```toml
    OPENAI_API_KEY = "sk-your-openai-key-here"
    GEMINI_API_KEY = "AIza-your-gemini-key-here"
    ```
    You can add either one or both — the provider selector in the sidebar will show "Ready to search" for whichever key is configured. Save, and the app picks them up automatically; no user ever sees or enters them.

    **Step 2 — Commit the D&B database into the GitHub repo**

    Rename your D&B export to exactly `dnb_database.xlsx` and add it to the root of the repository (same folder as `app.py`). Push to GitHub — Streamlit Cloud redeploys automatically and loads the file on every app start.

    ```
    your-repo/
    ├── app.py
    ├── requirements.txt
    ├── dnb_database.xlsx   ← admin-provided, loaded automatically
    └── README.md
    ```

    To update the database later, just replace `dnb_database.xlsx` in the repo and push — no code changes needed.

    Once both are in place, the sidebar will show **"✅ Ready to search"** and end users will only see the AI provider toggle and the trade/area search boxes.

    **About 429 "quota exceeded" errors:** this means the API key itself has no remaining credit or hit its rate limit — it isn't a bug in the app. Check billing at platform.openai.com (OpenAI) or aistudio.google.com (Gemini), or simply switch providers using the toggle at the top of the sidebar.
    """)
