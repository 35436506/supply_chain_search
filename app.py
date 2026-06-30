import streamlit as st
import pandas as pd
import json
import io
import re
import os
import time
from datetime import datetime
import requests
import pgeocode
import plotly.graph_objects as go

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
# enter an API key or upload files themselves.
#
# API keys (checked in this order):
#   1. Streamlit secrets: .streamlit/secrets.toml -> OPENAI_API_KEY / GEMINI_API_KEY
#   2. Environment variables
#
# Reference files (checked in this order):
#   1. Commit a file/folder with the matching name into this repo's root —
#      loaded automatically on startup, no upload needed.
#        - dnb_database.xlsx          -> D&B (DNBi) risk/turnover database
#        - wp_list.xlsx               -> Work Package list (trade/package descriptions)
#        - preferred_suppliers/*.xlsx -> ANY NUMBER of Preferred Supplier lists,
#                                         one per cluster. Drop as many .xlsx files
#                                         as you like into this folder; all are
#                                         combined automatically.
#   2. If a committed file/folder isn't found, the app falls back to a manual
#      upload box for that session only (multi-file uploader for suppliers).
# ════════════════════════════════════════════════════════════════════════════

DEFAULT_DB_PATH       = os.path.join(os.path.dirname(__file__), "dnb_database.xlsx")
DEFAULT_WP_PATH       = os.path.join(os.path.dirname(__file__), "wp_list.xlsx")
DEFAULT_SUPPLIER_DIR  = os.path.join(os.path.dirname(__file__), "preferred_suppliers")


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


def get_admin_wp_df() -> pd.DataFrame | None:
    """Load the committed Work Package list if it exists in the repo."""
    if os.path.exists(DEFAULT_WP_PATH):
        try:
            return load_wp_list(DEFAULT_WP_PATH)
        except Exception as e:
            st.sidebar.error(f"Could not read bundled WP list: {e}")
    return None


def get_admin_supplier_df() -> pd.DataFrame | None:
    """Load and combine every committed Preferred Supplier cluster file found
    in the preferred_suppliers/ folder. Each file becomes its own 'Cluster'
    column value (the filename, minus extension)."""
    if not os.path.isdir(DEFAULT_SUPPLIER_DIR):
        return None

    frames = []
    for fname in sorted(os.listdir(DEFAULT_SUPPLIER_DIR)):
        if not fname.lower().endswith((".xlsx", ".xls")):
            continue
        fpath = os.path.join(DEFAULT_SUPPLIER_DIR, fname)
        try:
            df = load_preferred_suppliers(fpath)
            if df is not None and not df.empty:
                df["Cluster"] = os.path.splitext(fname)[0]
                frames.append(df)
        except Exception as e:
            st.sidebar.error(f"Could not read {fname}: {e}")

    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


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

# Used only if no WP list has been uploaded/configured by the admin yet.
FALLBACK_TRADE_PACKAGES = [
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

COMPANY_SIZE_BANDS = [
    (0,            1_000_000,    "Micro (<£1M)"),
    (1_000_000,    10_000_000,   "Small (£1M–£10M)"),
    (10_000_000,   50_000_000,   "Medium (£10M–£50M)"),
    (50_000_000,   250_000_000,  "Large (£50M–£250M)"),
    (250_000_000,  float("inf"), "Major (£250M+)"),
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
        for ch in [" ", "_", "-", ".", "&", "(", ")", "/", "£", "$", "€"]:
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


def load_wp_list(file_or_path) -> pd.DataFrame | None:
    """Load a Work Package list. Expected columns roughly: WP Number,
    Correct Description, WP Description 1/2/3, Manager. Combines whichever
    description columns exist into a single clean 'Description' per WP."""
    try:
        xl = pd.ExcelFile(file_or_path)
        sheet = "Master WP" if "Master WP" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet)
    except Exception as e:
        st.sidebar.error(f"Could not read WP list: {e}")
        return None

    cm = _normalise_cols(df)
    num_col   = cm.get("wpnumber") or cm.get("number") or cm.get("wpno")
    correct_col = cm.get("correctdescription")
    desc_cols = [cm.get(f"wpdescription{i}") for i in (1, 2, 3) if cm.get(f"wpdescription{i}")]

    if num_col is None:
        return None

    rows = []
    for _, r in df.iterrows():
        wp_no = r.get(num_col, "")
        if pd.isna(wp_no):
            continue
        desc_parts = []
        if correct_col and pd.notna(r.get(correct_col, None)):
            desc_parts.append(str(r[correct_col]).strip())
        for dc in desc_cols:
            v = r.get(dc, None)
            if pd.notna(v) and str(v).strip():
                desc_parts.append(str(v).strip())
        desc = " / ".join(dict.fromkeys(desc_parts))
        if not desc:
            continue
        rows.append({"WP Number": str(wp_no).strip(), "Description": desc})

    return pd.DataFrame(rows) if rows else None


def load_preferred_suppliers(file_or_path) -> pd.DataFrame | None:
    """Load a Preferred Supplier list (Cluster export format). The
    'Preferred' sheet typically has its real header on row 2 (index 1) and
    the Package column is forward-filled (blank means same package as the
    row above). Falls back to header row 0 if row 1 doesn't look right,
    so files from different clusters with slightly different layouts
    still load."""
    try:
        xl = pd.ExcelFile(file_or_path)
    except Exception as e:
        st.sidebar.error(f"Could not read preferred supplier list: {e}")
        return None

    sheet = "Preferred" if "Preferred" in xl.sheet_names else xl.sheet_names[0]

    def _try_parse(header_row):
        df = xl.parse(sheet, header=header_row)
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df.columns = [str(c).strip() for c in df.columns]
        cm = _normalise_cols(df)
        return df, cm

    df, cm = _try_parse(1)
    if cm.get("companyname") is None:
        df, cm = _try_parse(0)

    pkg_col    = cm.get("package")
    name_col   = cm.get("companyname")
    status_col = cm.get("status")
    reg_col    = cm.get("registrationno")
    addr_col   = cm.get("tradingaddress")
    turn_col   = cm.get("turnover")
    risk_col   = cm.get("dnbirisk")
    cline_col  = cm.get("clinelevel")

    if name_col is None:
        return None

    if pkg_col:
        df[pkg_col] = df[pkg_col].ffill()

    keep = {
        "Package":          pkg_col,
        "Company Name":     name_col,
        "Status":           status_col,
        "Registration No.": reg_col,
        "Trading Address":  addr_col,
        "Turnover":         turn_col,
        "D&B Risk":         risk_col,
        "C/Line Level":     cline_col,
    }
    out = pd.DataFrame()
    for new_col, old_col in keep.items():
        out[new_col] = df[old_col] if old_col and old_col in df.columns else ""

    out = out[out["Company Name"].astype(str).str.strip() != ""]
    out = out[out["Company Name"].notna()]
    return out.reset_index(drop=True)


def find_preferred_match(supplier_df: pd.DataFrame, trade: str, company_name: str) -> dict:
    """Check whether a company appears in any loaded preferred supplier
    cluster list."""
    if supplier_df is None or supplier_df.empty:
        return {"is_preferred": False}

    name_norm = str(company_name).lower().strip()
    mask = supplier_df["Company Name"].astype(str).str.lower().str.strip() == name_norm
    if not mask.any():
        mask = supplier_df["Company Name"].astype(str).str.lower().str.contains(
            re.escape(name_norm[:15]), na=False
        )
    if not mask.any():
        return {"is_preferred": False}

    row = supplier_df[mask].iloc[0]
    return {
        "is_preferred": True,
        "status": str(row.get("Status", "")) if pd.notna(row.get("Status", "")) else "",
        "package": str(row.get("Package", "")) if pd.notna(row.get("Package", "")) else "",
        "cluster": str(row.get("Cluster", "")) if pd.notna(row.get("Cluster", "")) else "",
    }


_geocoder = None

def get_geocoder():
    global _geocoder
    if _geocoder is None:
        _geocoder = pgeocode.Nominatim("GB")
    return _geocoder


def geocode_location(location_text: str):
    """Best-effort UK geocoding from a free-text location/postcode string.
    Returns (lat, lon) or (None, None) if it can't be resolved."""
    if not location_text:
        return None, None
    geo = get_geocoder()

    pc_match = re.search(r"[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}", str(location_text).upper())
    if pc_match:
        result = geo.query_postal_code(pc_match.group(0).replace(" ", ""))
        if result is not None and pd.notna(result.latitude):
            return float(result.latitude), float(result.longitude)

    outward_match = re.search(r"\b[A-Z]{1,2}\d[A-Z\d]?\b", str(location_text).upper())
    if outward_match:
        result = geo.query_postal_code(outward_match.group(0))
        if result is not None and pd.notna(result.latitude):
            return float(result.latitude), float(result.longitude)

    return None, None


def turnover_to_size_band(turnover) -> str:
    """Map a turnover figure to a human-readable company size band."""
    try:
        val = float(str(turnover).replace(",", "").replace("£", "").strip())
    except (ValueError, TypeError):
        return ""
    for lo, hi, label in COMPANY_SIZE_BANDS:
        if lo <= val < hi:
            return label
    return ""


def turnover_to_employee_estimate(turnover) -> str:
    """Rough UK construction-sector rule of thumb (~£150k revenue per
    employee). Used only as a size indicator, not a precise figure."""
    try:
        val = float(str(turnover).replace(",", "").replace("£", "").strip())
    except (ValueError, TypeError):
        return ""
    if val <= 0:
        return ""
    est = val / 150_000
    if est < 1:
        return "<5"
    if est < 10:
        return f"~{round(est)}"
    if est < 100:
        return f"~{round(est / 10) * 10}"
    return f"~{round(est / 50) * 50}"


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


def build_ai_prompt(trade: str, region: str, specific_area: str, extra_notes: str, preferred_names: list) -> str:
    area_line = f"**Specific Area (priority focus):** {specific_area}\n" if specific_area else ""
    preferred_line = ""
    if preferred_names:
        sample = ", ".join(preferred_names[:30])
        preferred_line = (
            f"\n**Known preferred suppliers for this trade (prioritise / include these if genuinely suitable, "
            f"do not invent new ones with these exact names):** {sample}\n"
        )

    return f"""You are a UK construction procurement specialist.

A procurement team member needs you to identify suitable subcontractors for the following:

**Trade / Package:** {trade}
**Project Region / Area:** {region}
{area_line}**Additional Notes:** {extra_notes or 'None'}
{preferred_line}
Please provide:
1. A brief market overview for this trade in the given UK region (2-3 sentences).
2. A list of **8–12 recommended subcontractors** — ideally a mix of large, mid-tier, and specialist firms. If a specific area was given above, prioritise companies based in or very near it.

For EACH company return a JSON object inside a ```json block with this EXACT schema:
{{
  "companies": [
    {{
      "company_name": "...",
      "registration_no": "...",
      "trade_scope": "...",
      "location": "...",
      "website": "...",
      "proximity_score": <integer 1-10 where 10 = based right in the area>,
      "estimated_turnover": <number in GBP, or null if truly unknown>,
      "turnover_is_estimate": <true if this figure came from your own knowledge/web research rather than an official filing, false if it's a known/reported figure>,
      "notes": "..."
    }}
  ]
}}

Rules:
- Use real UK Companies House registration numbers where known, else leave blank.
- Trade scope should be 1-2 specific sentences about what they deliver relevant to this package.
- Location should be the company's registered / main office address, including a UK postcode if you know it (helps map placement).
- Proximity score: 10 = headquartered in or immediately adjacent to the target area; 1 = national but mobilisable.
- For estimated_turnover: if you don't have a precise reported figure, use your knowledge of the company (employee count, market position, public mentions, press coverage) to give a reasonable estimated annual turnover in GBP as an indicator of company size. Mark turnover_is_estimate as true in that case. Only leave estimated_turnover null if you genuinely have no basis at all for an estimate.
- Be factual and conservative with estimates; do not invent precise-looking figures with false confidence — round estimates to a sensible level (e.g. nearest £100k or £1M).
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


def companies_to_df(companies: list, db_df, supplier_df, trade: str, region: str) -> pd.DataFrame:
    rows = []
    for c in companies:
        name = c.get("company_name", "")
        reg  = c.get("registration_no", "")

        db_info = lookup_db(db_df, name, reg)
        location = db_info.get("D&B Location") or c.get("location", "")

        # Turnover: prefer D&B's official figure; fall back to the AI's estimate
        db_turnover = db_info.get("Turnover", "")
        ai_turnover = c.get("estimated_turnover", None)
        is_estimate = bool(c.get("turnover_is_estimate", False))

        if db_turnover:
            turnover = db_turnover
            turnover_source = "D&B"
        elif ai_turnover is not None:
            turnover = ai_turnover
            turnover_source = "Estimated"
            is_estimate = True
        else:
            turnover = ""
            turnover_source = ""

        pref = find_preferred_match(supplier_df, trade, name)
        lat, lon = geocode_location(location)

        rows.append({
            "Company Name":      name,
            "Registration No.":  reg,
            "Trade Scope":       c.get("trade_scope", ""),
            "Close to Area":     c.get("proximity_score", ""),
            "Location":          location,
            "Turnover":          turnover,
            "Turnover Source":   turnover_source,
            "Company Size":      turnover_to_size_band(turnover),
            "Number of Employees": turnover_to_employee_estimate(turnover),
            "D&B Risk":          db_info.get("D&B Risk", ""),
            "Preferred Supplier": "Yes" if pref.get("is_preferred") else "",
            "Preferred Cluster": pref.get("cluster", "") if pref.get("is_preferred") else "",
            "Website":           c.get("website", ""),
            "Notes":             c.get("notes", ""),
            "_db_matched":       bool(db_info.get("_db_matched")),
            "_turnover_is_estimate": is_estimate,
            "_lat": lat,
            "_lon": lon,
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

        col_widths = [30, 16, 50, 12, 35, 14, 14, 18, 18, 14, 14, 16, 35, 45]
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

admin_db_df       = get_admin_db_df()
admin_wp_df       = get_admin_wp_df()
admin_supplier_df = get_admin_supplier_df()

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

    # Work Package list — admin-provided, with session fallback
    wp_df = admin_wp_df
    if wp_df is None:
        with st.expander("📋 Upload Work Package list", expanded=False):
            st.caption("Upload your WP list to populate the Trade/Package dropdown with real package descriptions. Not configured by admin yet — see Admin Setup below.")
            wp_file = st.file_uploader("Work Package Excel file", type=["xlsx", "xls"], key="wp_upload")
            wp_df = load_wp_list(wp_file) if wp_file else None

    # Preferred Supplier lists — admin-provided (multiple cluster files), with
    # session fallback (multi-file upload, so several clusters can be combined)
    supplier_df = admin_supplier_df
    if supplier_df is None:
        with st.expander("⭐ Upload Preferred Supplier list(s)", expanded=False):
            st.caption("Upload one or more cluster preferred-supplier exports — you can select multiple files at once. All are combined and used to prioritise known suppliers in results. Not configured by admin yet — see Admin Setup below.")
            supplier_files = st.file_uploader(
                "Preferred Supplier Excel file(s)",
                type=["xlsx", "xls"],
                key="supplier_upload",
                accept_multiple_files=True,
            )
            if supplier_files:
                frames = []
                for f in supplier_files:
                    d = load_preferred_suppliers(f)
                    if d is not None and not d.empty:
                        d["Cluster"] = os.path.splitext(f.name)[0]
                        frames.append(d)
                supplier_df = pd.concat(frames, ignore_index=True) if frames else None
    if supplier_df is not None and not supplier_df.empty:
        n_clusters = supplier_df["Cluster"].nunique() if "Cluster" in supplier_df.columns else 1
        st.caption(f"⭐ {len(supplier_df):,} preferred suppliers loaded across {n_clusters} cluster file(s)")

    st.markdown("---")
    st.markdown("### Trade & Area")

    if wp_df is not None and not wp_df.empty:
        trade_options = ["— Select —"] + wp_df["Description"].dropna().unique().tolist()
    else:
        trade_options = ["— Select —"] + FALLBACK_TRADE_PACKAGES

    trade_dropdown = st.selectbox("Trade / Package (from WP list)", trade_options)
    trade_search = st.text_input("...or type a trade to search for", placeholder="e.g. acoustic ceilings, asphalt paving")
    trade = trade_search.strip() if trade_search.strip() else trade_dropdown

    region = st.selectbox("UK Region / Area", ["— Select —"] + UK_REGIONS)
    specific_area = st.text_input("Or a specific area (town, city, postcode)", placeholder="e.g. Reading, RG1, Maidstone")

    extra = st.text_area("Additional Notes", placeholder="e.g. CHAS accredited, min turnover £10M, experience in healthcare...", height=100)

    search_btn = st.button("🔍 Find Subcontractors", use_container_width=True, type="primary")


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
    2. **Pick a Trade/Package** from the dropdown (populated from your Work Package list) — or just type a trade in the search box if it's not listed.
    3. **Pick a UK Region**, or enter a specific town/city/postcode if you need a smaller area.
    4. Add any extra notes (minimum turnover, accreditations, etc.), then click **Find Subcontractors**.
    5. Results are cross-referenced against your D&B database and Preferred Supplier list, sorted by **Turnover → Company Size → Proximity**, with preferred suppliers prioritised to the top.
    6. **Review the map** of supplier locations, **edit the table** inline, then **export to Excel**.

    If you see a setup warning in the sidebar, the admin API key(s) and/or reference files haven't been configured yet for this deployment — see **Admin Setup** at the bottom of this page.
    """)

if "result_df"   not in st.session_state: st.session_state.result_df   = None
if "narrative"   not in st.session_state: st.session_state.narrative   = ""
if "last_trade"  not in st.session_state: st.session_state.last_trade  = ""
if "last_region" not in st.session_state: st.session_state.last_region = ""

# ── Search execution ──────────────────────────────────────────────────────────
if search_btn:
    if not trade or trade == "— Select —":
        st.warning("Please select or type a Trade Package before searching.")
    elif region == "— Select —" and not specific_area.strip():
        st.warning("Please select a UK Region or enter a specific area before searching.")
    elif not api_key:
        st.error(f"No {ai_provider} API key available. Ask your admin to configure one, or enter one in the sidebar.")
    else:
        effective_region = specific_area.strip() if specific_area.strip() else region
        with st.spinner(f"Searching for {trade} subcontractors in {effective_region}…"):
            preferred_names = []
            if supplier_df is not None and not supplier_df.empty:
                trade_lower = trade.lower()
                pkg_mask = supplier_df["Package"].astype(str).str.lower().str.contains(
                    re.escape(trade_lower[:20]), na=False
                )
                preferred_names = supplier_df.loc[pkg_mask, "Company Name"].dropna().unique().tolist()
                if not preferred_names:
                    preferred_names = supplier_df["Company Name"].dropna().unique().tolist()[:30]

            prompt = build_ai_prompt(trade, region if region != "— Select —" else "", specific_area.strip(), extra, preferred_names)
            try:
                raw = call_ai(ai_provider, api_key, prompt)
                narrative, companies = parse_ai_response(raw)
                df = companies_to_df(companies, db_df, supplier_df, trade, effective_region)

                st.session_state.result_df   = df
                st.session_state.narrative   = narrative
                st.session_state.last_trade  = trade
                st.session_state.last_region = effective_region
                st.session_state.raw_ai      = raw

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
    preferred_cnt = (df["Preferred Supplier"] == "Yes").sum()

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="val">{total}</div><div class="lbl">Subcontractors Found</div></div>
      <div class="metric-card"><div class="val">{preferred_cnt}</div><div class="lbl">Preferred Suppliers</div></div>
      <div class="metric-card"><div class="val">{with_db}</div><div class="lbl">D&B Records Matched</div></div>
      <div class="metric-card"><div class="val">{low_risk}</div><div class="lbl">Low / Low-Mod Risk</div></div>
      <div class="metric-card"><div class="val">{local_cnt}</div><div class="lbl">Locally Based (Score ≥7)</div></div>
    </div>
    """, unsafe_allow_html=True)

    if nav:
        st.markdown('<div class="section-title">📊 Market Overview</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="ai-box">{nav}</div>', unsafe_allow_html=True)

    # ── Sort: Turnover (desc) -> Company Size band (desc) -> Close to Area (desc) ──
    def _turnover_numeric(v):
        try:
            return float(str(v).replace(",", "").replace("£", "").strip())
        except (ValueError, TypeError):
            return None

    size_rank = {label: i for i, (_, _, label) in enumerate(COMPANY_SIZE_BANDS)}

    view = df.copy()
    view["_turn_sort"] = view["Turnover"].apply(_turnover_numeric)
    view["_size_sort"] = view["Company Size"].map(size_rank)
    view["_prox_sort"] = pd.to_numeric(view["Close to Area"], errors="coerce")
    view["_pref_sort"] = (view["Preferred Supplier"] == "Yes").astype(int)

    view = view.sort_values(
        by=["_pref_sort", "_turn_sort", "_size_sort", "_prox_sort"],
        ascending=[False, False, False, False],
        na_position="last",
    ).drop(columns=["_turn_sort", "_size_sort", "_prox_sort", "_pref_sort"])

    # ── Map of suggested suppliers ──────────────────────────────────────────────
    map_df = view.dropna(subset=["_lat", "_lon"]) if "_lat" in view.columns else pd.DataFrame()
    if not map_df.empty:
        st.markdown('<div class="section-title">🗺️ Supplier Locations</div>', unsafe_allow_html=True)

        def _bubble_size(turnover):
            t = _turnover_numeric(turnover)
            if t is None or t <= 0:
                return 14
            import math
            return max(14, min(60, 10 + 8 * math.log10(max(t, 1000))))

        sizes = map_df["Turnover"].apply(_bubble_size)
        colors = map_df["Preferred Supplier"].apply(lambda x: "#f57c00" if x == "Yes" else "#2563eb")
        hover_text = map_df.apply(
            lambda r: f"<b>{r['Company Name']}</b><br>Turnover: £{r['Turnover']}<br>{r['Location']}"
            if r["Turnover"] else f"<b>{r['Company Name']}</b><br>{r['Location']}",
            axis=1,
        )

        fig = go.Figure(go.Scattermapbox(
            lat=map_df["_lat"], lon=map_df["_lon"],
            mode="markers",
            marker=dict(size=sizes, color=colors, opacity=0.75),
            text=hover_text,
            hoverinfo="text",
        ))
        fig.update_layout(
            mapbox=dict(style="open-street-map", zoom=6, center=dict(lat=float(map_df["_lat"].mean()), lon=float(map_df["_lon"].mean()))),
            margin=dict(l=0, r=0, t=0, b=0),
            height=450,
        )
        st.plotly_chart(fig, use_container_width=True)
        st.caption("🔵 Standard supplier · 🟠 Preferred supplier · Bubble size reflects company turnover (bigger = larger company)")
    else:
        st.info("📍 No locations could be mapped — locations need a recognisable UK postcode or town/city to plot.")

    st.markdown('<div class="section-title">🔧 Filter & Export</div>', unsafe_allow_html=True)
    risk_opts = ["All"] + sorted(df["D&B Risk"].dropna().unique().tolist())
    risk_filt = st.selectbox("D&B Risk", risk_opts)

    if risk_filt != "All":
        view = view[view["D&B Risk"] == risk_filt]

    display_cols = [c for c in view.columns if not c.startswith("_")]

    st.markdown(f"**{len(view)} companies shown**, sorted by Turnover → Company Size → Proximity — you can edit cells directly:")
    edited = st.data_editor(
        view[display_cols].reset_index(drop=True),
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Website":             st.column_config.LinkColumn("Website"),
            "Close to Area":       st.column_config.NumberColumn("Close to Area", min_value=1, max_value=10),
            "Trade Scope":         st.column_config.TextColumn("Trade Scope", width="large"),
            "Notes":               st.column_config.TextColumn("Notes", width="large"),
            "Turnover":            st.column_config.NumberColumn("Turnover (£)", format="£%d"),
            "Preferred Supplier":  st.column_config.TextColumn("Preferred Supplier"),
        },
        hide_index=True,
        height=500,
    )

    st.markdown('<div class="section-title">⬇️ Export</div>', unsafe_allow_html=True)
    xlsx_bytes = to_excel_bytes(edited, trde, rgn)
    fname = f"Subcontractors_{rgn[:15].replace(' ','_')}_{trde[:20].replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        "📥 Download Excel",
        data=xlsx_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
    You can add either one or both — the provider selector in the sidebar will show "Ready to search" for whichever key is configured.

    **Step 2 — Commit the reference files into the GitHub repo**

    Place these files in the root of the repository, alongside `app.py`, using these exact names:

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
    └── README.md
    ```

    Every `.xlsx` file inside `preferred_suppliers/` is loaded and combined automatically — there's no limit on how many cluster files you add. Each company is tagged with which cluster file it came from (shown in the **Preferred Cluster** column in results). To add a new cluster later, just drop another file into that folder and push.

    None of these files are required for the app to run — any missing piece simply falls back to a manual upload box for that session (the supplier uploader accepts multiple files at once too).

    Once everything is in place, the sidebar will show **"✅ Ready to search"** and end users will only see the AI provider toggle and the trade/area search boxes.

    **About API errors:**
    - **429 "quota exceeded"** — the API key's billing/credit needs attention. Check platform.openai.com (OpenAI) or aistudio.google.com (Gemini), or switch providers using the toggle.
    - **503 "overloaded"** — the provider's servers are temporarily at capacity. The app retries automatically; if it still fails, try again shortly or switch providers.
    """)
