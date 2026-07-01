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
# API key (checked in this order):
#   1. Streamlit secrets: .streamlit/secrets.toml -> GEMINI_API_KEY
#   2. Environment variable: GEMINI_API_KEY
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


def get_admin_api_key() -> str:
    """Look for an admin-configured Gemini key in secrets or env vars."""
    try:
        if "GEMINI_API_KEY" in st.secrets:
            return st.secrets["GEMINI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "")


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
    """Find best matching row by registration number first (most reliable),
    then a normalised exact name match, then a conservative bidirectional
    partial match as a last resort. Avoids loose substring matches that can
    attach the wrong company's verified data to an AI-suggested name."""
    if name_col is None or name_col not in df.columns:
        return None

    # 1. Registration number — most reliable
    if reg_no and reg_col and reg_col in df.columns:
        reg_mask = df[reg_col].astype(str).str.strip().str.upper() == str(reg_no).strip().upper()
        if reg_mask.any():
            return df[reg_mask].iloc[0]

    # 2. Normalised exact name match
    target_norm = _normalise_company_name(company_name)
    name_norms = df[name_col].astype(str).apply(_normalise_company_name)
    exact_mask = name_norms == target_norm
    if exact_mask.any():
        return df[exact_mask].iloc[0]

    # 3. Conservative bidirectional partial match
    if len(target_norm) >= 6:
        partial_mask = name_norms.apply(lambda n: target_norm in n or n in target_norm)
        if partial_mask.any():
            return df[partial_mask].iloc[0]

    return None


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
        "reg":          safe("reg"),
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

    # Strip whitespace first, then deduplicate — the cluster export has
    # "Email" and "Email " (trailing space) which become duplicates after strip.
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated(keep="first")]
    cm = _normalise_cols(df)

    pkg_col       = cm.get("package")
    name_col      = cm.get("companyname")
    status_col    = cm.get("status")
    reg_col       = cm.get("registrationno")
    addr_col      = cm.get("tradingaddress")
    turn_col      = cm.get("turnover")
    risk_col      = cm.get("dnbirisk")
    cline_col     = cm.get("clinelevel")
    contact_col   = cm.get("keycontact")
    email_col     = cm.get("email")
    phone_col     = cm.get("phone")

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
        "Key Contact":      contact_col,
        "Email":            email_col,
        "Phone":            phone_col,
    }
    out = pd.DataFrame()
    for new_col, old_col in keep.items():
        out[new_col] = df[old_col] if old_col and old_col in df.columns else ""

    out = out[out["Company Name"].astype(str).str.strip() != ""]
    out = out[out["Company Name"].notna()]
    return out.reset_index(drop=True)


def _normalise_company_name(name: str) -> str:
    """Normalise a company name for safer matching: lowercase, strip common
    suffixes (Ltd, Limited, etc.) and punctuation, collapse whitespace."""
    n = str(name).lower().strip()
    n = re.sub(r"[.,&()]", " ", n)
    n = re.sub(r"\b(ltd|limited|llp|plc|group|inc|co)\b", "", n)
    n = re.sub(r"\s+", " ", n).strip()
    return n


def find_preferred_match(supplier_df: pd.DataFrame, trade: str, company_name: str, reg_no: str = "") -> dict:
    """Check whether a company appears in any loaded preferred supplier
    cluster list. Uses registration number as the most reliable signal when
    available, then a normalised exact-name match, then a conservative
    partial match only as a last resort — all to avoid attaching the wrong
    company's verified data to an AI-suggested name."""
    if supplier_df is None or supplier_df.empty:
        return {"is_preferred": False}

    row = None

    # 1. Registration number match — most reliable, avoids name-spelling drift
    if reg_no and "Registration No." in supplier_df.columns:
        reg_mask = supplier_df["Registration No."].astype(str).str.strip().str.upper() == str(reg_no).strip().upper()
        if reg_mask.any():
            row = supplier_df[reg_mask].iloc[0]

    # 2. Normalised exact name match (ignores Ltd/Limited/punctuation differences)
    if row is None:
        target_norm = _normalise_company_name(company_name)
        name_norms = supplier_df["Company Name"].astype(str).apply(_normalise_company_name)
        exact_mask = name_norms == target_norm
        if exact_mask.any():
            row = supplier_df[exact_mask].iloc[0]

    # 3. Conservative partial match — only if the AI name is a clear substring
    #    match in both directions (handles things like trading-name variants)
    if row is None:
        target_norm = _normalise_company_name(company_name)
        if len(target_norm) >= 6:
            name_norms = supplier_df["Company Name"].astype(str).apply(_normalise_company_name)
            partial_mask = name_norms.apply(lambda n: target_norm in n or n in target_norm)
            if partial_mask.any():
                row = supplier_df[partial_mask].iloc[0]

    if row is None:
        return {"is_preferred": False}

    def safe(col):
        v = row.get(col, "")
        return str(v) if pd.notna(v) else ""

    return {
        "is_preferred":    True,
        "status":          safe("Status"),
        "package":         safe("Package"),
        "cluster":         safe("Cluster"),
        "company_name":    safe("Company Name"),
        "registration_no": safe("Registration No."),
        "location":        safe("Trading Address"),
        "turnover":        safe("Turnover"),
        "db_risk":         safe("D&B Risk"),
        "cline_level":     safe("C/Line Level"),
        "key_contact":     safe("Key Contact"),
        "email":           safe("Email"),
        "phone":           safe("Phone"),
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


def call_gemini(api_key: str, prompt: str) -> str:
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
    body = {"contents": [{"parts": [{"text": prompt}]}]}
    r = _post_with_retry(url, headers, body)
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def friendly_api_error(e: requests.HTTPError) -> str:
    status = e.response.status_code
    body = e.response.text[:300]
    if status == 503:
        return ("Gemini is temporarily overloaded (503) — Google's servers are at capacity right now. "
                "This usually clears within a minute or two. The app already retried automatically; "
                "please try **Find Subcontractors** again shortly.")
    if status == 429:
        return ("Gemini quota exceeded (429) — this API key hit its rate or usage limit. "
                "Check quota and billing at aistudio.google.com, or ask your admin to update the key in Streamlit Secrets.")
    if status == 401 or status == 403:
        return f"Gemini rejected this API key ({status}) — it may be invalid, revoked, or missing permissions."
    if status == 404:
        return f"Gemini model not found (404) — the model name in the code may be outdated. Details: {body}"
    return f"Gemini API error ({status}): {body}"


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
      "proximity_score": <integer 1-10 where 10 = headquartered right in the target area, 1 = national/remote>,
      "estimated_turnover": <number in GBP, or null if truly unknown>,
      "turnover_is_estimate": <true if this figure came from your own knowledge/web research rather than an official filing, false if it's a known/reported figure>,
      "notes": "..."
    }}
  ]
}}

Rules:
- For registration_no: only provide a UK Companies House number if you are highly confident it is correct for this exact company. If you are not certain, leave it blank — a wrong number is worse than no number, since this app will flag and surface it to a procurement team.
- For website: only provide a URL if you are highly confident it is the company's real site. Leave blank if unsure.
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
        ai_name = c.get("company_name", "")
        ai_reg  = c.get("registration_no", "")

        db_info = lookup_db(db_df, ai_name, ai_reg)
        pref    = find_preferred_match(supplier_df, trade, ai_name, ai_reg)

        is_verified = bool(db_info.get("_db_matched")) or bool(pref.get("is_preferred"))

        # ── Company name ──────────────────────────────────────────────────────
        # Preferred supplier list is the most authoritative source for name.
        name = pref.get("company_name") or ai_name

        # ── Registration number ───────────────────────────────────────────────
        # ONLY show a registration number that came from a trusted source file.
        # Never show what the AI suggested — it is unreliable and misleads users.
        if pref.get("registration_no"):
            reg_display = pref["registration_no"]          # from preferred supplier list
        elif db_info.get("_db_matched") and db_info.get("reg"):
            reg_display = db_info["reg"]                   # from D&B record
        else:
            reg_display = ""                               # unknown — leave blank

        # ── Location ──────────────────────────────────────────────────────────
        location = (pref.get("location") or db_info.get("D&B Location")
                    or c.get("location", ""))

        # ── Risk ──────────────────────────────────────────────────────────────
        db_risk = pref.get("db_risk") or db_info.get("D&B Risk", "")

        # ── Turnover ──────────────────────────────────────────────────────────
        # Precedence: preferred-supplier list > D&B > AI estimate
        pref_turn  = pref.get("turnover", "")
        db_turn    = db_info.get("Turnover", "")
        ai_turn    = c.get("estimated_turnover", None)
        is_estimate = False

        if pref_turn:
            turnover = pref_turn
        elif db_turn:
            turnover = db_turn
        elif ai_turn is not None:
            turnover = ai_turn
            is_estimate = True
        else:
            turnover = ""

        lat, lon = geocode_location(location)

        rows.append({
            "Company Name":        name,
            "Registration No.":    reg_display,
            "Trade Scope":         c.get("trade_scope", ""),
            "Location":            location,
            "Turnover":            turnover,
            "Company Size":        turnover_to_size_band(turnover),
            "Number of Employees": turnover_to_employee_estimate(turnover),
            "D&B Risk":            db_risk,
            "C/Line Level":        pref.get("cline_level", "") if pref.get("is_preferred") else "",
            "Contact":             pref.get("key_contact", "") if pref.get("is_preferred") else "",
            "Email":               pref.get("email", "") if pref.get("is_preferred") else "",
            "Phone":               pref.get("phone", "") if pref.get("is_preferred") else "",
            "Notes":               c.get("notes", ""),
            "Preferred Supplier":  "Yes" if pref.get("is_preferred") else "",
            "Close to Area":       c.get("proximity_score", ""),  # raw score, ranked later
            "_verified":           is_verified,
            "_db_matched":         bool(db_info.get("_db_matched")),
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

        col_widths = [5, 30, 16, 50, 12, 35, 14, 14, 18, 18, 14, 14, 16, 14, 28, 32, 16, 35, 45]
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

    # API key and all reference files loaded silently from repo/secrets.
    # No status messages or upload widgets shown to users.
    api_key     = get_admin_api_key()
    db_df       = admin_db_df
    wp_df       = admin_wp_df
    supplier_df = admin_supplier_df

    st.markdown("### Trade & Area")

    if wp_df is not None and not wp_df.empty:
        trade_options = ["— Select —"] + wp_df["Description"].dropna().unique().tolist()
    else:
        trade_options = ["— Select —"] + FALLBACK_TRADE_PACKAGES

    trade_dropdown = st.selectbox("Trade / Package", trade_options)
    trade_search   = st.text_input("Or type a trade to search for", placeholder="e.g. acoustic ceilings, asphalt paving")
    trade = trade_search.strip() if trade_search.strip() else trade_dropdown

    region        = st.selectbox("UK Region / Area", ["— Select —"] + UK_REGIONS)
    specific_area = st.text_input("Or a specific area (town, city, postcode)", placeholder="e.g. Reading, RG1, Maidstone")

    extra = st.text_area("Additional Notes", placeholder="e.g. CHAS accredited, min turnover £10M, experience in healthcare...", height=100)

    search_btn = st.button("🔍 Find Subcontractors", use_container_width=True, type="primary")


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="app-header">
  <h1>🔍 Subcontractor Finder</h1>
  <p>UK subcontractor discovery, cross-referenced against D&B and Preferred Supplier data &nbsp;·&nbsp;
  <a href="https://find-and-update.company-information.service.gov.uk/" target="_blank" style="color:#93c5fd;">
  🏛️ Search Companies House</a></p>
</div>
""", unsafe_allow_html=True)

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
        st.error("No Gemini API key available. Ask your admin to configure one, or enter one in the sidebar.")
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
                raw = call_gemini(api_key, prompt)
                narrative, companies = parse_ai_response(raw)
                df = companies_to_df(companies, db_df, supplier_df, trade, effective_region)

                st.session_state.result_df   = df
                st.session_state.narrative   = narrative
                st.session_state.last_trade  = trade
                st.session_state.last_region = effective_region
                st.session_state.raw_ai      = raw

            except requests.HTTPError as e:
                st.error(friendly_api_error(e))
            except Exception as e:
                st.error(f"Error: {e}")

# ── Results display ───────────────────────────────────────────────────────────
if st.session_state.result_df is not None:
    df   = st.session_state.result_df
    nav  = st.session_state.narrative
    trde = st.session_state.last_trade
    rgn  = st.session_state.last_region

    total      = len(df)
    verified_cnt = int(df.get("_verified", pd.Series([False]*len(df))).sum()) if "_verified" in df.columns else 0
    with_db    = int(df.get("_db_matched", pd.Series([False]*len(df))).sum()) if "_db_matched" in df.columns else 0
    low_risk   = df["D&B Risk"].str.lower().str.contains("low", na=False).sum()
    local_cnt  = (pd.to_numeric(df["Close to Area"], errors="coerce") <= 3).sum()
    preferred_cnt = (df["Preferred Supplier"] == "Yes").sum()

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="val">{total}</div><div class="lbl">Subcontractors Found</div></div>
      <div class="metric-card"><div class="val">{verified_cnt}</div><div class="lbl">Verified (D&B or Preferred List)</div></div>
      <div class="metric-card"><div class="val">{preferred_cnt}</div><div class="lbl">Preferred Suppliers</div></div>
      <div class="metric-card"><div class="val">{low_risk}</div><div class="lbl">Low / Low-Mod Risk</div></div>
      <div class="metric-card"><div class="val">{local_cnt}</div><div class="lbl">Closest to Area (Rank ≤3)</div></div>
    </div>
    """, unsafe_allow_html=True)

    unverified_cnt = total - verified_cnt
    if unverified_cnt > 0:
        st.warning(
            f"⚠️ {unverified_cnt} of {total} companies were not found in your D&B or Preferred Supplier data — "
            f"their registration number is left blank. Please verify these companies on "
            f"[Companies House](https://find-and-update.company-information.service.gov.uk/) before contacting them."
        )

    if nav:
        st.markdown('<div class="section-title">📊 Market Overview</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="ai-box">{nav}</div>', unsafe_allow_html=True)

    # ── Sort: Preferred first -> Turnover (desc) -> Company Size band (desc) -> Close to Area (desc) ──
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
    ).drop(columns=["_turn_sort", "_size_sort", "_pref_sort"])

    # Convert raw AI proximity score (1–10) to a unique rank across the actual
    # result set: rank 1 = closest to target area, rank N = furthest.
    # method="first" ensures every company gets a distinct rank even when
    # the AI gave them the same proximity score.
    view["Close to Area"] = view["_prox_sort"].rank(
        ascending=False, method="first", na_option="bottom"
    ).astype("Int64")
    view = view.drop(columns=["_prox_sort"])

    view = view.reset_index(drop=True)
    view.insert(0, "#", view.index + 1)

    # ── Map of suggested suppliers ──────────────────────────────────────────────
    map_df = view.dropna(subset=["_lat", "_lon"]) if "_lat" in view.columns else pd.DataFrame()
    if not map_df.empty:
        st.markdown('<div class="section-title">🗺️ Supplier Locations</div>', unsafe_allow_html=True)
        st.caption("Each numbered marker matches the **#** column in the table below. Bubble size reflects company turnover. Markers sitting on the same town are nudged apart slightly so they don't overlap.")

        map_df = map_df.reset_index(drop=True)
        map_df["_label"] = map_df["#"].astype(str)

        # Spread out markers that geocoded to the exact same point (e.g. several
        # companies in the same town), so they don't render as one indistinguishable blob.
        import math
        coord_counts = {}
        jittered_lat, jittered_lon = [], []
        for _, r in map_df.iterrows():
            key = (round(r["_lat"], 3), round(r["_lon"], 3))
            n = coord_counts.get(key, 0)
            coord_counts[key] = n + 1
            if n == 0:
                jittered_lat.append(r["_lat"])
                jittered_lon.append(r["_lon"])
            else:
                angle = (n - 1) * (2 * math.pi / 8)
                radius = 0.035 * (1 + (n - 1) // 8)
                jittered_lat.append(r["_lat"] + radius * math.sin(angle))
                jittered_lon.append(r["_lon"] + radius * math.cos(angle))
        map_df["_lat_j"] = jittered_lat
        map_df["_lon_j"] = jittered_lon

        def _bubble_size(turnover):
            t = _turnover_numeric(turnover)
            if t is None or t <= 0:
                return 22
            return max(22, min(56, 16 + 7 * math.log10(max(t, 1000))))

        # A distinct, high-contrast colour per company, cycling through a
        # qualitative palette so adjacent markers are visually separable —
        # not just one blue blob and one orange blob.
        palette = [
            "#2563eb", "#dc2626", "#16a34a", "#9333ea", "#ea580c",
            "#0891b2", "#db2777", "#65a30d", "#7c3aed", "#0d9488",
            "#ca8a04", "#e11d48", "#4f46e5", "#059669", "#c2410c",
        ]
        marker_colors = [palette[i % len(palette)] for i in range(len(map_df))]
        # Preferred suppliers get a gold ring via a distinct border-like outer
        # marker drawn underneath (Scattermapbox has no native border width).
        is_preferred = (map_df["Preferred Supplier"] == "Yes").tolist()

        sizes = map_df["Turnover"].apply(_bubble_size).tolist()
        hover_text = map_df.apply(
            lambda r: (
                f"<b>#{r['_label']} {r['Company Name']}</b><br>"
                f"{'⭐ Preferred Supplier<br>' if r['Preferred Supplier'] == 'Yes' else ''}"
                f"Turnover: £{r['Turnover']:,}" if isinstance(r['Turnover'], (int, float)) and r['Turnover']
                else f"<b>#{r['_label']} {r['Company Name']}</b><br>{'⭐ Preferred Supplier<br>' if r['Preferred Supplier'] == 'Yes' else ''}{r['Location']}"
            ) + f"<br>{r['Location']}",
            axis=1,
        )

        fig = go.Figure()

        # Gold halo behind preferred-supplier markers so they stand out beyond just colour
        pref_idx = [i for i, p in enumerate(is_preferred) if p]
        if pref_idx:
            fig.add_trace(go.Scattermapbox(
                lat=[map_df["_lat_j"].iloc[i] for i in pref_idx],
                lon=[map_df["_lon_j"].iloc[i] for i in pref_idx],
                mode="markers",
                marker=dict(size=[sizes[i] + 14 for i in pref_idx], color="#fbbf24", opacity=0.55),
                hoverinfo="skip",
                showlegend=False,
            ))

        fig.add_trace(go.Scattermapbox(
            lat=map_df["_lat_j"], lon=map_df["_lon_j"],
            mode="markers+text",
            marker=dict(size=sizes, color=marker_colors, opacity=0.9),
            text=map_df["_label"],
            textfont=dict(size=11, color="white"),
            hovertext=hover_text,
            hoverinfo="text",
            showlegend=False,
        ))

        fig.update_layout(
            mapbox=dict(style="open-street-map", zoom=6, center=dict(lat=float(map_df["_lat"].mean()), lon=float(map_df["_lon"].mean()))),
            margin=dict(l=0, r=0, t=0, b=0),
            height=480,
        )
        st.plotly_chart(fig, use_container_width=True)
        st.caption("🟡 Gold halo = Preferred Supplier · Bubble size reflects turnover · Number on each marker matches the table's # column")
    else:
        st.info("📍 No locations could be mapped — locations need a recognisable UK postcode or town/city to plot.")

    display_cols = [c for c in view.columns if not c.startswith("_")]

    st.markdown(f"**{len(view)} companies shown**, sorted by Preferred Supplier → Turnover → Company Size → Proximity. Registration numbers are only shown when verified against D&B or your Preferred Supplier list — use [Companies House](https://find-and-update.company-information.service.gov.uk/) to look up any with a blank registration number.")
    edited = st.data_editor(
        view[display_cols].reset_index(drop=True),
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "#":                     st.column_config.NumberColumn("#", width="small"),
            "Close to Area":         st.column_config.NumberColumn("Close to Area (Rank)", help="1 = closest to target area, higher = further away"),
            "Trade Scope":           st.column_config.TextColumn("Trade Scope", width="large"),
            "Notes":                 st.column_config.TextColumn("Notes", width="large"),
            "Turnover":              st.column_config.NumberColumn("Turnover (£)", format="£%d"),
            "Preferred Supplier":    st.column_config.TextColumn("Preferred Supplier"),
            "C/Line Level":          st.column_config.TextColumn("C/Line Level"),
            "Contact":               st.column_config.TextColumn("Contact"),
            "Email":                 st.column_config.TextColumn("Email"),
            "Phone":                 st.column_config.TextColumn("Phone"),
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

else:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#5a6272;">
        <div style="font-size:4rem;">🔍</div>
        <h3 style="color:#3a4150;">Ready to search</h3>
        <p>Select a Trade Package and UK Region in the sidebar, then click <strong>Find Subcontractors</strong>.</p>
    </div>
    """, unsafe_allow_html=True)



