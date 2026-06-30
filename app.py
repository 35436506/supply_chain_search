import streamlit as st
import pandas as pd
import json
import io
import re
from datetime import datetime
import requests

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Subcontractor Finder | Laing O'Rourke",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Sidebar */
    [data-testid="stSidebar"] { background: #0d1b2a; }
    [data-testid="stSidebar"] * { color: #e8eaf0 !important; }
    [data-testid="stSidebar"] .stTextInput label,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stFileUploader label { color: #a8b8d0 !important; font-size:0.8rem; }
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color: #c8a84b !important; }

    /* Header */
    .lor-header {
        background: linear-gradient(135deg, #0d1b2a 0%, #1a3a5c 100%);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        border-left: 5px solid #c8a84b;
    }
    .lor-header h1 { margin:0; font-size: 1.6rem; color: white; }
    .lor-header p  { margin:0.3rem 0 0; color: #a8c8e8; font-size: 0.9rem; }

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
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    .metric-card .val { font-size: 2rem; font-weight: 700; color: #1a3a5c; }
    .metric-card .lbl { font-size: 0.75rem; color: #666; text-transform: uppercase; letter-spacing:0.05em; margin-top:0.2rem; }

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
        border-bottom: 2px solid #c8a84b;
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
    }

    /* Table styling */
    .stDataFrame { border-radius: 8px; overflow: hidden; }

    /* Hide default streamlit branding */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


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


def load_excel_any(uploaded_file) -> pd.DataFrame | None:
    """Generic Excel loader — combines all sheets."""
    try:
        xl = pd.ExcelFile(uploaded_file)
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


def find_cline_columns(cline_df: pd.DataFrame) -> dict:
    """Detect Constructionline export columns. C/Line exports typically include
    Company Name, Registration Number, and a status/level field (e.g. 'Gold',
    'Silver', 'SSIP Verified', membership level, or pass/fail status)."""
    cm = _normalise_cols(cline_df)
    return {
        "name":   cm.get("companyname") or cm.get("name") or cm.get("company") or cm.get("suppliername"),
        "reg":    cm.get("registrationnumber") or cm.get("companynumber") or cm.get("companynr") or cm.get("regno"),
        "status": cm.get("clinestatus") or cm.get("status") or cm.get("membershiplevel") or cm.get("level")
                  or cm.get("accreditationlevel") or cm.get("certificationlevel"),
        "expiry": cm.get("expirydate") or cm.get("expiry") or cm.get("renewaldate"),
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


def lookup_cline(cline_df: pd.DataFrame, company_name: str, reg_no: str = "") -> dict:
    """Look up a company in the uploaded Constructionline export.
    Returns 'Not available' if no C/Line file has been uploaded at all,
    or 'Not registered' if the file is uploaded but the company isn't in it."""
    if cline_df is None or cline_df.empty:
        return {"C/Line": "Not available"}

    cols = find_cline_columns(cline_df)
    row = _match_row(cline_df, cols["name"], cols["reg"], company_name, reg_no)
    if row is None:
        return {"C/Line": "Not registered"}

    status_col = cols.get("status")
    if status_col and status_col in cline_df.columns:
        v = row[status_col]
        val = "" if pd.isna(v) else str(v)
        return {"C/Line": val if val else "Listed"}
    return {"C/Line": "Listed"}


def call_claude(api_key: str, prompt: str) -> str:
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    body = {
        "model": "claude-opus-4-6",
        "max_tokens": 2000,
        "messages": [{"role": "user", "content": prompt}],
    }
    r = requests.post("https://api.anthropic.com/v1/messages", headers=headers, json=body, timeout=60)
    r.raise_for_status()
    data = r.json()
    return data["content"][0]["text"]


def call_gemini(api_key: str, prompt: str) -> str:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key={api_key}"
    body = {"contents": [{"parts": [{"text": prompt}]}]}
    r = requests.post(url, json=body, timeout=60)
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def build_ai_prompt(trade: str, region: str, extra_notes: str) -> str:
    return f"""You are a UK construction procurement specialist at Laing O'Rourke.

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
    # Split narrative and JSON
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


def companies_to_df(companies: list, db_df, cline_df, region: str) -> pd.DataFrame:
    rows = []
    for c in companies:
        name = c.get("company_name", "")
        reg  = c.get("registration_no", "")

        db_info    = lookup_db(db_df, name, reg)
        cline_info = lookup_cline(cline_df, name, reg)

        # Prefer D&B's own location data if matched, else fall back to AI's guess
        location = db_info.get("D&B Location") or c.get("location", "")

        rows.append({
            "Company Name":      name,
            "Registration No.":  reg,
            "Trade Scope":       c.get("trade_scope", ""),
            "Close to Area":     c.get("proximity_score", ""),
            "Location":          location,
            "Contact":           c.get("contact", ""),
            "Turnover":          db_info.get("Turnover", ""),
            "C/Line":            cline_info.get("C/Line", ""),
            "D&B Risk":          db_info.get("D&B Risk", ""),
            "Website":           c.get("website", ""),
            "AI Notes":          c.get("notes", ""),
            "_db_matched":       bool(db_info.get("_db_matched")),
        })
    return pd.DataFrame(rows)


def to_excel_bytes(df: pd.DataFrame, trade: str, region: str) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Subcontractors")
        wb  = writer.book
        ws  = writer.sheets["Subcontractors"]

        # Formats
        hdr_fmt = wb.add_format({
            "bold": True, "bg_color": "#1a3a5c", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        cell_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
        money_fmt = wb.add_format({"border": 1, "valign": "top", "num_format": "£#,##0"})

        col_widths = [30, 16, 50, 12, 35, 30, 14, 14, 18, 35, 45]
        for i, (col, w) in enumerate(zip(df.columns, col_widths)):
            ws.set_column(i, i, w, cell_fmt)
            ws.write(0, i, col, hdr_fmt)

        # Turnover as number where possible
        turn_idx = list(df.columns).index("Turnover") if "Turnover" in df.columns else -1
        for row_idx, row in df.iterrows():
            for col_idx, val in enumerate(row):
                if col_idx == turn_idx:
                    try:
                        ws.write_number(row_idx + 1, col_idx, int(float(str(val).replace(",", "").replace("£", "").strip())), money_fmt)
                        continue
                    except Exception:
                        pass
                ws.write(row_idx + 1, col_idx, str(val) if val else "", cell_fmt)

        ws.set_row(0, 30)

        # Meta sheet
        meta = wb.add_worksheet("Search Info")
        meta.write("A1", "Trade Package")
        meta.write("B1", trade)
        meta.write("A2", "Region")
        meta.write("B2", region)
        meta.write("A3", "Exported")
        meta.write("B3", datetime.now().strftime("%Y-%m-%d %H:%M"))

    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🔑 API Keys")
    claude_key  = st.text_input("Anthropic (Claude) Key", type="password", placeholder="sk-ant-...")
    gemini_key  = st.text_input("Google (Gemini) Key",    type="password", placeholder="AIza...")
    ai_provider = st.radio("Active AI", ["Claude", "Gemini"], horizontal=True)

    st.markdown("---")
    st.markdown("## 📁 Databases")

    st.caption("**D&B (DNBi)** — exported risk/turnover database. Required columns: Company Name, Business Registration Number, Sales (Revenue), Overall Business Risk, City - D&B, Zip/Postal Code - D&B.")
    db_file = st.file_uploader("D&B Excel export", type=["xlsx", "xls"], key="db_upload")
    db_df = load_excel_any(db_file) if db_file else None
    if db_df is not None:
        st.success(f"✅ D&B loaded — {len(db_df):,} records")
    else:
        st.info("⏳ No D&B file uploaded yet")

    st.caption("**Constructionline (C/Line)** — upload when available. Until then, the C/Line column will show 'Not available'.")
    cline_file = st.file_uploader("Constructionline Excel export", type=["xlsx", "xls"], key="cline_upload")
    cline_df = load_excel_any(cline_file) if cline_file else None
    if cline_df is not None:
        st.success(f"✅ C/Line loaded — {len(cline_df):,} records")
    else:
        st.warning("⏳ C/Line not uploaded — column will be blank until provided")

    st.markdown("---")
    st.markdown("## 🔍 Search Parameters")
    trade  = st.selectbox("Trade / Package", ["— Select —"] + TRADE_PACKAGES)
    region = st.selectbox("UK Region / Area", ["— Select —"] + UK_REGIONS)
    extra  = st.text_area("Additional Notes", placeholder="e.g. CHAS accredited, min turnover £10M, experience in healthcare...", height=100)

    search_btn = st.button("🔍 Find Subcontractors", use_container_width=True, type="primary")

    st.markdown("---")
    st.markdown("## 📋 Previous Results")
    if "history" not in st.session_state:
        st.session_state.history = []
    if st.session_state.history:
        for i, h in enumerate(reversed(st.session_state.history[-5:])):
            st.caption(f"• {h['trade'][:35]}… | {h['region'][:20]}")
    else:
        st.caption("No searches yet.")


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="lor-header">
  <h1>🏗️ Subcontractor Finder</h1>
  <p>AI-powered UK subcontractor discovery for Laing O'Rourke Procurement</p>
</div>
""", unsafe_allow_html=True)

# ── How-to guide ─────────────────────────────────────────────────────────────
with st.expander("ℹ️ How to use this tool", expanded=False):
    st.markdown("""
    1. **Add your API key** in the sidebar — Claude (Anthropic) or Gemini (Google).
    2. **Upload your D&B (DNBi) Excel export** — required columns: `Company Name`, `Business Registration Number`, `Sales (Revenue)`, `Overall Business Risk`, `City - D&B`, `Zip/Postal Code - D&B`. The app matches AI-suggested companies against this file by name (and registration number as a fallback) to populate **Turnover**, **D&B Risk**, and **Location**.
    3. **Upload a Constructionline (C/Line) export when you have one.** Until then, the **C/Line** column will simply show *"Not available"* — the app will automatically start filling it in as soon as a file is uploaded, no code changes needed.
    4. **Select a Trade Package and UK Region**, add any extra notes (minimum turnover, accreditations, etc.).
    5. Click **Find Subcontractors** — the AI researches suitable suppliers; the app cross-references D&B and C/Line and returns a structured table.
    6. **Review and edit** the table inline, then **export to Excel or CSV**.

    **D&B (DNBi) file format** — columns expected (case-insensitive, minor naming variants tolerated):
    `Company Name`, `Business Registration Number`, `Sales (Revenue)`, `Overall Business Risk`, `City - D&B`, `Zip/Postal Code - D&B`, `Country or Region - D&B`

    **Constructionline file format** — typically:
    `Company Name`, `Registration Number`, and a status/level column (e.g. `Status`, `Membership Level`, `Accreditation Level`) — column names are auto-detected.
    """)

# ── Session state init ────────────────────────────────────────────────────────
if "result_df"  not in st.session_state: st.session_state.result_df  = None
if "narrative"  not in st.session_state: st.session_state.narrative  = ""
if "last_trade" not in st.session_state: st.session_state.last_trade = ""
if "last_region"not in st.session_state: st.session_state.last_region= ""

# ── Search execution ──────────────────────────────────────────────────────────
if search_btn:
    if trade == "— Select —" or region == "— Select —":
        st.warning("Please select both a Trade Package and a UK Region before searching.")
    elif ai_provider == "Claude" and not claude_key:
        st.error("Please enter your Anthropic (Claude) API key in the sidebar.")
    elif ai_provider == "Gemini" and not gemini_key:
        st.error("Please enter your Google (Gemini) API key in the sidebar.")
    else:
        with st.spinner(f"Searching for {trade} subcontractors in {region}…"):
            prompt = build_ai_prompt(trade, region, extra)
            try:
                if ai_provider == "Claude":
                    raw = call_claude(claude_key, prompt)
                else:
                    raw = call_gemini(gemini_key, prompt)

                narrative, companies = parse_ai_response(raw)
                df = companies_to_df(companies, db_df, cline_df, region)

                st.session_state.result_df   = df
                st.session_state.narrative   = narrative
                st.session_state.last_trade  = trade
                st.session_state.last_region = region
                st.session_state.raw_ai      = raw

                st.session_state.history.append({"trade": trade, "region": region})

            except requests.HTTPError as e:
                st.error(f"API error ({e.response.status_code}): {e.response.text[:300]}")
            except Exception as e:
                st.error(f"Error: {e}")

# ── Results display ───────────────────────────────────────────────────────────
if st.session_state.result_df is not None:
    df   = st.session_state.result_df
    nav  = st.session_state.narrative
    trde = st.session_state.last_trade
    rgn  = st.session_state.last_region

    # Metric cards
    total      = len(df)
    with_db    = int(df.get("_db_matched", pd.Series([False]*len(df))).sum()) if "_db_matched" in df.columns else int(df["D&B Risk"].astype(bool).sum())
    low_risk   = df["D&B Risk"].str.lower().str.contains("low", na=False).sum()
    local_cnt  = (pd.to_numeric(df["Close to Area"], errors="coerce") >= 7).sum()
    cline_avail = cline_df is not None
    cline_matched = (df["C/Line"].astype(str) != "Not available").sum() if cline_avail else 0

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="val">{total}</div><div class="lbl">Subcontractors Found</div></div>
      <div class="metric-card"><div class="val">{with_db}</div><div class="lbl">D&B Records Matched</div></div>
      <div class="metric-card"><div class="val">{low_risk}</div><div class="lbl">Low / Low-Mod Risk</div></div>
      <div class="metric-card"><div class="val">{local_cnt}</div><div class="lbl">Locally Based (Score ≥7)</div></div>
    </div>
    """, unsafe_allow_html=True)

    if not cline_avail:
        st.info("ℹ️ Constructionline (C/Line) file not uploaded yet — that column will read **'Not available'** until you upload a C/Line export in the sidebar.")

    # AI narrative
    if nav:
        st.markdown('<div class="section-title">📊 Market Overview</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="ai-box">{nav}</div>', unsafe_allow_html=True)

    # Filters
    st.markdown('<div class="section-title">🔧 Filter & Export</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        risk_opts = ["All"] + sorted(df["D&B Risk"].dropna().unique().tolist())
        risk_filt = st.selectbox("D&B Risk", risk_opts)
    with col2:
        min_prox = st.slider("Min Proximity Score", 1, 10, 1)
    with col3:
        sort_by = st.selectbox("Sort by", ["Close to Area ↓", "Company Name ↑", "D&B Risk"])

    # Apply filters
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

    # Editable table
    st.markdown(f"**{len(view)} companies shown** — you can edit cells directly:")
    edited = st.data_editor(
        view.reset_index(drop=True),
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

    # Export
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

    # Raw AI toggle
    with st.expander("🤖 Raw AI Response", expanded=False):
        st.text(st.session_state.get("raw_ai", ""))

else:
    # Empty state
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#999;">
        <div style="font-size:4rem;">🔍</div>
        <h3 style="color:#ccc;">Ready to search</h3>
        <p>Select a Trade Package and UK Region in the sidebar, then click <strong>Find Subcontractors</strong>.</p>
    </div>
    """, unsafe_allow_html=True)
