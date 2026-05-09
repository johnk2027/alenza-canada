
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import io
import re
import json
import sqlite3
import hashlib
import secrets
import os
import zipfile
import urllib.request
from urllib.parse import quote_plus, urlencode
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List

try:
    from PIL import Image
except Exception:
    Image = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except Exception:
    SimpleDocTemplate = None


# ============================================================
# ALENZA CAPITAL | ENTERPRISE CRE UNDERWRITING SUITE
# 950-target build: underwriting, OCR, offline AI, optional Ollama,
# rent roll, diligence, lender quotes, amortization, Excel/PDF/JSON.
# ============================================================

APP_VERSION = "Platform Build 3.2 Canada | 950 Target"

DEFAULTS = {
    "sponsor": "Client Name",
    "property_type": "Multifamily",
    "transaction_type": "Acquisition",
    "purchase_price": 12500000,
    "appraisal": 13750000,
    "existing_debt": 8500000,
    "gross_income": 1550000,
    "vacancy_loss": 70000,
    "operating_expenses": 420000,
    "noi": 1060322,
    "target_ltv": 0.75,
    "target_ltc": 0.80,
    "target_dscr": 1.25,
    "target_dy": 0.085,
    "rate": 0.0525,
    "amort": 25,
    "loan_term": 5,
    "fees": 0.02,
    "closing_costs": 50000,
    "capex_reserve": 0,
    "interest_reserve": 0,
    "debt_structure": "Amortizing",
    "credit_profile": "Market",
    "property_address": "100 King Street West, Toronto, ON"
}

PROFILE_PRESETS = {
    "Conservative": {"target_ltv": 0.65, "target_ltc": 0.70, "target_dscr": 1.35, "target_dy": 0.095},
    "Market": {"target_ltv": 0.75, "target_ltc": 0.80, "target_dscr": 1.25, "target_dy": 0.085},
    "Aggressive": {"target_ltv": 0.80, "target_ltc": 0.85, "target_dscr": 1.20, "target_dy": 0.075},
    "Custom": {}
}

PROPERTY_RISK_WEIGHTS = {
    "Multifamily": 1.00,
    "Industrial": 0.98,
    "Medical Office": 1.03,
    "Self-Storage": 1.04,
    "Mixed-Use": 1.07,
    "Retail": 1.10,
    "Office": 1.18,
    "Hospitality": 1.25,
    "Other": 1.12,
}


# ============================================================
# PAGE CONFIG + THEME
# ============================================================

st.set_page_config(
    page_title="Alenza Capital Underwriting Suite",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
:root {
    --bg-main: #05080F;
    --bg-panel: #0B1220;
    --border: #1E293B;
    --text-main: #F8FAFC;
    --text-muted: #94A3B8;
    --accent: #1D4ED8;
    --accent-soft: #2563EB;
}
.main { background-color: var(--bg-main); color: var(--text-main); font-family: "Helvetica Neue", Arial, sans-serif; }
section[data-testid="stSidebar"] { background-color: var(--bg-panel) !important; border-right: 1px solid var(--border); }
h1, h2, h3 { letter-spacing: -0.025em; }
[data-testid="stMetricValue"] { font-size: 26px !important; font-weight: 800 !important; color: var(--accent-soft) !important; }
[data-testid="stMetricLabel"] { font-size: 11px !important; text-transform: uppercase; letter-spacing: 1.4px; color: var(--text-muted) !important; }
div[data-testid="stMetric"] {
    background-color: var(--bg-panel);
    padding: 17px;
    border-radius: 6px;
    border: 1px solid var(--border);
    box-shadow: 0 3px 14px rgba(0,0,0,.35);
}
.stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--border); }
.stTabs [data-baseweb="tab"] {
    background-color: var(--bg-panel);
    border: 1px solid var(--border);
    border-bottom: none;
    border-radius: 4px 4px 0 0;
    padding: 9px 14px;
    color: var(--text-muted);
    font-weight: 700;
    letter-spacing: .04em;
    text-transform: uppercase;
    font-size: 10px;
}
.stTabs [aria-selected="true"] { background-color: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important; }
.stDownloadButton>button, .stButton>button {
    width: 100%;
    background-color: var(--accent);
    color: white;
    font-weight: 700;
    border-radius: 6px;
    border: none;
    padding: 12px;
    text-transform: uppercase;
    letter-spacing: .055em;
}
.stDownloadButton>button:hover, .stButton>button:hover { background-color: var(--accent-soft); color: white; }
div[data-testid="stExpander"] { border: 1px solid var(--border); border-radius: 6px; background-color: var(--bg-panel); }
hr { border-color: var(--border); }
</style>
""", unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def format_money(value) -> str:
    try:
        return f"${float(value):,.0f}"
    except Exception:
        return "$0"


def format_pct(value) -> str:
    try:
        return f"{float(value) * 100:.1f}%"
    except Exception:
        return "0.0%"


def format_x(value) -> str:
    try:
        return f"{float(value):.2f}x"
    except Exception:
        return "0.00x"


def safe_divide(numerator, denominator) -> float:
    try:
        if denominator in (0, None):
            return 0.0
        return float(numerator) / float(denominator)
    except Exception:
        return 0.0


def clean_filename(value: str) -> str:
    value = str(value).strip().replace(" ", "_").replace("/", "_").replace("\\", "_")
    value = re.sub(r"[^A-Za-z0-9_\\-]", "", value)
    return value or "Client"


def money_to_float(value) -> Optional[float]:
    if value is None:
        return None
    cleaned = str(value).replace("$", "").replace(",", "").replace("(", "-").replace(")", "").strip()
    try:
        return float(cleaned)
    except Exception:
        return None



# ============================================================
# DEAL CONCIERGE: DILIGENCE & PROPERTY INTELLIGENCE
# ============================================================

DILIGENCE_REQUIREMENTS = {
    "Financials (T12)": ["t12", "operating", "p&l", "pnl", "financial", "income statement", "trailing"],
    "Rent Roll": ["rent roll", "rentroll", "tenant", "rr"],
    "Appraisal": ["appraisal", "valuation", "value report"],
    "Environmental": ["phase", "enviro", "environmental", "esa"],
    "Sponsor Info": ["bio", "sreo", "schedule real estate owned", "experience", "sponsor"],
    "Purchase Agreement": ["purchase agreement", "psa", "contract"],
    "Insurance": ["insurance", "certificate", "coi"],
    "Title / Survey": ["title", "survey", "alta"]
}


def audit_deal_room(files):
    """Analyzes uploaded filenames to determine what diligence items are present or missing."""
    if not files:
        return [], list(DILIGENCE_REQUIREMENTS.keys()), pd.DataFrame(columns=["Requirement", "Status", "Matched File"])

    found = []
    missing = []
    rows = []

    for category, keywords in DILIGENCE_REQUIREMENTS.items():
        matched_file = None

        for file in files:
            file_name = getattr(file, "name", "").lower()
            if any(keyword in file_name for keyword in keywords):
                matched_file = getattr(file, "name", "Uploaded file")
                break

        if matched_file:
            found.append(category)
            rows.append({
                "Requirement": category,
                "Status": "Received",
                "Matched File": matched_file
            })
        else:
            missing.append(category)
            rows.append({
                "Requirement": category,
                "Status": "Missing",
                "Matched File": ""
            })

    return found, missing, pd.DataFrame(rows)



def nrcan_geolocate_canada(address: str) -> Dict[str, Any]:
    """Canada-first no-key geolocation helper using NRCan/Geo.ca style endpoints.

    This is intentionally safe: if the endpoint changes, fails, or rate-limits, the app falls back without crashing.
    """
    if not address:
        return {"status": "NO_ADDRESS", "message": "No address provided."}

    # NRCan geolocator endpoint family. Keep fallback-safe because public endpoint formats can change.
    query = urlencode({"q": address})
    url = f"https://geolocator.api.geo.ca/geolocations?{query}"

    data = http_get_json(url, timeout=8)

    if data.get("_error"):
        return {
            "status": "FALLBACK",
            "message": data.get("_error"),
            "source": "NRCan / Geo.ca public geolocation"
        }

    return {
        "status": "OK",
        "source": "NRCan / Geo.ca public geolocation",
        "raw": data
    }


def canadian_property_fallback(address: str) -> Dict[str, Any]:
    """Manual/API-ready Canada fallback facts.

    Values are not asserted as real parcel facts. They are placeholders for underwriting workflow until
    municipal, ATTOM, parcel, zoning, or flood APIs are connected.
    """
    return {
        "Address": address or "Not provided",
        "Country": "Canada",
        "Province": "Manual entry required",
        "Municipality": "Manual entry required",
        "Lot Size": "Manual entry required",
        "Year Built": "Manual entry required",
        "Last Sale": "Manual entry required",
        "Assessment / Roll Number": "Manual entry required",
        "Zoning": "Manual entry required",
        "Flood / Hazard": "Check NRCan / municipal flood data",
        "Environmental": "Check Phase I ESA / provincial records",
        "Data Mode": "Canada fallback / manual"
    }


def get_property_intelligence(address):
    """Canada-first property intelligence.

    No private API keys are bundled. If keys are configured, production calls can be added here.
    Otherwise, the app uses safe Canadian fallbacks and public/no-key lookup attempts.
    """
    address_text = str(address or "").strip()
    fallback = canadian_property_fallback(address_text)

    google_key = get_secret_value("GOOGLE_MAPS_API_KEY")
    attom_key = get_secret_value("ATTOM_API_KEY")
    parcel_key = get_secret_value("CANADA_PARCEL_API_KEY")
    zoning_key = get_secret_value("CANADA_ZONING_API_KEY")

    # No-key Canada geolocation attempt. This is best-effort and never required.
    geo_result = nrcan_geolocate_canada(address_text)

    if geo_result.get("status") == "OK":
        fallback["Geolocation Source"] = geo_result.get("source", "NRCan / Geo.ca")
        fallback["Geolocation Status"] = "Attempted"
    else:
        fallback["Geolocation Source"] = "Manual / Maps link"
        fallback["Geolocation Status"] = geo_result.get("status", "Fallback")

    fallback["Google Maps Key"] = "Configured" if google_key else "Not configured"
    fallback["ATTOM Key"] = "Configured" if attom_key else "Not configured"
    fallback["Parcel API Key"] = "Configured" if parcel_key else "Not configured"
    fallback["Zoning API Key"] = "Configured" if zoning_key else "Not configured"

    return fallback


def build_google_maps_url(address):
    """Returns a safely encoded Google Maps search URL."""
    encoded = quote_plus(str(address or ""))
    return f"https://www.google.com/maps/search/{encoded}"


def build_google_maps_embed_url(address):
    """Returns a safely encoded Google Maps embed URL."""
    encoded = quote_plus(str(address or ""))
    return f"https://www.google.com/maps?q={encoded}&output=embed"


def add_audit(action: str, detail: str) -> None:
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    st.session_state.audit_log.append({"Timestamp": now_str(), "Action": action, "Detail": detail})



# ============================================================
# PLATFORM ENGINE: AUTH, ROLES, DATABASE, DOCUMENT ROOM
# ============================================================

DATA_DIR = Path("alenza_data")
DOC_DIR = DATA_DIR / "documents"
DB_PATH = DATA_DIR / "alenza_platform.db"

ROLE_PERMISSIONS = {
    "Admin": [
        "view_all_deals",
        "manage_users",
        "manage_documents",
        "view_admin",
        "view_borrower_portal",
        "view_lender_portal",
        "export_data"
    ],
    "Broker": [
        "view_all_deals",
        "manage_documents",
        "view_borrower_portal",
        "view_lender_portal",
        "export_data"
    ],
    "Borrower": [
        "view_borrower_portal",
        "upload_documents"
    ],
    "Lender": [
        "view_lender_portal",
        "review_deals"
    ]
}

SOC2_SECURITY_CONTROLS = [
    "User access reviews documented",
    "Role-based permissions enforced",
    "Sensitive documents stored in controlled repository",
    "Audit trail available for key actions",
    "Backups / exports available",
    "Incident response contact identified",
    "Vendor/API key inventory maintained",
    "PII/client document retention policy documented",
    "Encryption in transit confirmed for hosted app",
    "Admin-only system settings protected"
]


def ensure_platform_storage():
    """Creates local SQLite database and document storage folder."""
    DATA_DIR.mkdir(exist_ok=True)
    DOC_DIR.mkdir(exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            role TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS deals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deal_name TEXT NOT NULL,
            sponsor TEXT,
            property_address TEXT,
            property_type TEXT,
            transaction_type TEXT,
            loan_amount REAL,
            noi REAL,
            appraised_value REAL,
            score INTEGER,
            classification TEXT,
            status TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deal_id INTEGER,
            filename TEXT NOT NULL,
            category TEXT,
            file_path TEXT,
            uploaded_by TEXT,
            uploaded_at TEXT NOT NULL,
            status TEXT DEFAULT 'Received',
            notes TEXT,
            FOREIGN KEY(deal_id) REFERENCES deals(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS audit_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            action TEXT NOT NULL,
            detail TEXT,
            timestamp TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS api_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            provider TEXT UNIQUE NOT NULL,
            status TEXT NOT NULL,
            notes TEXT,
            updated_at TEXT NOT NULL
        )
    """)

    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        create_user_db(cur, "admin", "alenza-admin", "Admin")
        create_user_db(cur, "broker", "alenza-broker", "Broker")
        create_user_db(cur, "borrower", "alenza-borrower", "Borrower")
        create_user_db(cur, "lender", "alenza-lender", "Lender")

    conn.commit()
    conn.close()


def hash_password(password: str, salt: str) -> str:
    """PBKDF2 password hash. Suitable for prototype/local use."""
    return hashlib.pbkdf2_hmac(
        "sha256",
        str(password).encode("utf-8"),
        salt.encode("utf-8"),
        120000
    ).hex()


def create_user_db(cur, username: str, password: str, role: str):
    salt = secrets.token_hex(16)
    password_hash = hash_password(password, salt)
    cur.execute(
        "INSERT OR IGNORE INTO users (username, password_hash, salt, role, created_at) VALUES (?, ?, ?, ?, ?)",
        (username, password_hash, salt, role, now_str())
    )


def verify_login(username: str, password: str) -> Optional[Dict[str, Any]]:
    ensure_platform_storage()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE username = ?", (username,))
    row = cur.fetchone()
    conn.close()

    if not row:
        return None

    test_hash = hash_password(password, row["salt"])

    if secrets.compare_digest(test_hash, row["password_hash"]):
        return {
            "username": row["username"],
            "role": row["role"],
            "permissions": ROLE_PERMISSIONS.get(row["role"], [])
        }

    return None


def platform_audit(action: str, detail: str = ""):
    """Writes audit event to local DB and session log."""
    ensure_platform_storage()
    username = st.session_state.get("current_user", {}).get("username", "anonymous")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO audit_events (username, action, detail, timestamp) VALUES (?, ?, ?, ?)",
        (username, action, detail, now_str())
    )
    conn.commit()
    conn.close()

    add_audit(action, detail)


def has_permission(permission: str) -> bool:
    user = st.session_state.get("current_user")
    if not user:
        return False
    return permission in user.get("permissions", [])


def upsert_current_deal_record(inputs: Dict[str, Any], loan_amt: float, score: int, classification: str):
    """Saves the current underwriting snapshot into SQLite."""
    ensure_platform_storage()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    deal_name = f"{inputs.get('sponsor', 'Client')} - {inputs.get('property_address', 'Property')}"
    timestamp = now_str()

    cur.execute("""
        INSERT INTO deals (
            deal_name, sponsor, property_address, property_type, transaction_type,
            loan_amount, noi, appraised_value, score, classification, status,
            created_at, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        deal_name,
        inputs.get("sponsor"),
        inputs.get("property_address"),
        inputs.get("property_type"),
        inputs.get("transaction_type"),
        float(loan_amt or 0),
        float(inputs.get("noi") or 0),
        float(inputs.get("appraisal") or 0),
        int(score or 0),
        classification,
        "Active",
        timestamp,
        timestamp
    ))

    deal_id = cur.lastrowid
    conn.commit()
    conn.close()
    platform_audit("Deal Snapshot Saved", f"Deal ID {deal_id} | {deal_name}")
    return deal_id


def get_deals_df() -> pd.DataFrame:
    ensure_platform_storage()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM deals ORDER BY updated_at DESC", conn)
    conn.close()
    return df


def get_documents_df() -> pd.DataFrame:
    ensure_platform_storage()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM documents ORDER BY uploaded_at DESC", conn)
    conn.close()
    return df


def get_audit_df() -> pd.DataFrame:
    ensure_platform_storage()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM audit_events ORDER BY id DESC LIMIT 500", conn)
    conn.close()
    return df


def save_uploaded_documents(files, deal_id: Optional[int], category: str, uploaded_by: str):
    """Saves uploaded documents locally and records metadata in SQLite."""
    ensure_platform_storage()

    if not files:
        return []

    saved = []
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for file in files:
        safe_name = clean_filename(file.name)
        timestamp_prefix = datetime.now().strftime("%Y%m%d%H%M%S%f")
        stored_name = f"{timestamp_prefix}_{safe_name}"
        dest = DOC_DIR / stored_name
        dest.write_bytes(file.getbuffer())

        cur.execute("""
            INSERT INTO documents (
                deal_id, filename, category, file_path, uploaded_by, uploaded_at, status, notes
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            deal_id,
            file.name,
            category,
            str(dest),
            uploaded_by,
            now_str(),
            "Received",
            ""
        ))

        saved.append(file.name)

    conn.commit()
    conn.close()

    platform_audit("Documents Uploaded", f"{len(saved)} file(s): {', '.join(saved[:5])}")
    return saved


def parse_uploaded_table(file) -> pd.DataFrame:
    """Reads CSV/XLS/XLSX uploaded financial or rent-roll table."""
    name = file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(file)

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file)

    raise ValueError("Only CSV, XLS, or XLSX table files are supported for structured parsing.")


def normalize_rent_roll_table(df: pd.DataFrame) -> pd.DataFrame:
    """Maps common rent-roll column names into Alenza's schema."""
    if df is None or df.empty:
        return default_rent_roll_df()

    normalized = df.copy()
    lower_map = {str(c).lower().strip(): c for c in normalized.columns}

    def find_col(candidates):
        for candidate in candidates:
            for col_lower, original in lower_map.items():
                if candidate in col_lower:
                    return original
        return None

    tenant_col = find_col(["tenant", "lessee", "occupant", "name"])
    sf_col = find_col(["sf", "sq ft", "square feet", "area", "size"])
    term_col = find_col(["remaining term", "term remaining", "years left", "lease term", "walt"])
    rent_col = find_col(["monthly rent", "rent/month", "monthly", "base rent", "rent"])

    out = pd.DataFrame()
    out["Tenant"] = normalized[tenant_col] if tenant_col else [f"Tenant {i+1}" for i in range(len(normalized))]
    out["SF"] = pd.to_numeric(normalized[sf_col], errors="coerce").fillna(0) if sf_col else 0
    out["Remaining Term"] = pd.to_numeric(normalized[term_col], errors="coerce").fillna(0) if term_col else 0
    out["Monthly Rent"] = pd.to_numeric(normalized[rent_col], errors="coerce").fillna(0) if rent_col else 0

    return out


def normalize_t12_table(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, float]]:
    """Very lightweight T12 normalization for uploaded CSV/XLSX.

    Expected useful columns: line item/category/account and amount/actual/value.
    """
    if df is None or df.empty:
        return pd.DataFrame(), {"Income": 0, "Expenses": 0, "NOI": 0}

    data = df.copy()
    lower_map = {str(c).lower().strip(): c for c in data.columns}

    line_col = None
    amount_col = None

    for key in ["line", "account", "category", "description", "item"]:
        for col_lower, original in lower_map.items():
            if key in col_lower:
                line_col = original
                break
        if line_col:
            break

    for key in ["amount", "actual", "value", "total", "t12"]:
        for col_lower, original in lower_map.items():
            if key in col_lower:
                amount_col = original
                break
        if amount_col:
            break

    if not line_col or not amount_col:
        return pd.DataFrame(), {"Income": 0, "Expenses": 0, "NOI": 0}

    out = pd.DataFrame({
        "Line Item": data[line_col].astype(str),
        "Amount": pd.to_numeric(data[amount_col], errors="coerce").fillna(0)
    })

    def classify_line(text):
        t = text.lower()
        if any(k in t for k in ["income", "revenue", "rent", "egi", "gross"]):
            return "Income"
        if any(k in t for k in ["expense", "tax", "insurance", "repair", "utility", "payroll", "maintenance", "management"]):
            return "Expense"
        if "noi" in t or "net operating" in t:
            return "NOI"
        return "Other"

    out["Category"] = out["Line Item"].apply(classify_line)

    income = out.loc[out["Category"] == "Income", "Amount"].sum()
    expenses = abs(out.loc[out["Category"] == "Expense", "Amount"].sum())
    explicit_noi = out.loc[out["Category"] == "NOI", "Amount"].sum()
    noi = explicit_noi if explicit_noi else income - expenses

    summary = {
        "Income": float(income),
        "Expenses": float(expenses),
        "NOI": float(noi)
    }

    return out, summary


def property_api_status_df() -> pd.DataFrame:
    """Shows API readiness for 950-platform roadmap."""
    ensure_platform_storage()
    providers = [
        ("Google Places / Maps", "Ready for key", "Address, maps, nearby context"),
        ("ATTOM / County Records", "Ready for key", "Parcel, ownership, sale history"),
        ("Zoning API", "Ready for key", "Zoning / allowed use"),
        ("FEMA Flood", "Ready for key", "Flood zone and risk"),
        ("EPA / Enviro", "Ready for key", "Environmental proximity signals")
    ]

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for provider, status, notes in providers:
        cur.execute(
            "INSERT OR IGNORE INTO api_settings (provider, status, notes, updated_at) VALUES (?, ?, ?, ?)",
            (provider, status, notes, now_str())
        )

    conn.commit()
    df = pd.read_sql_query("SELECT provider, status, notes, updated_at FROM api_settings ORDER BY provider", conn)
    conn.close()
    return df




def get_secret_value(key: str) -> Optional[str]:
    """Loads a secret from Streamlit Secrets first, then environment variables.

    Never print or expose returned values in the UI.
    """
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass

    return os.environ.get(key)


def masked_secret_status(key: str) -> str:
    """Returns a safe yes/no status without exposing secret values."""
    return "Yes" if bool(get_secret_value(key)) else "No"


def api_key_status_df() -> pd.DataFrame:
    """Checks whether optional Canada-focused production API keys are configured.

    The app does not need these keys to run. Public/no-key fallbacks and manual underwriting remain available.
    """
    keys = [
        ("Google Maps / Places", "GOOGLE_MAPS_API_KEY"),
        ("ATTOM Canada / Property Data", "ATTOM_API_KEY"),
        ("Canadian County / Municipal Parcel Data", "CANADA_PARCEL_API_KEY"),
        ("Canadian Zoning Data", "CANADA_ZONING_API_KEY"),
        ("NRCan / Flood or Geospatial Services", "NRCAN_API_KEY"),
        ("Environment / Climate / Contamination Data", "CANADA_ENVIRONMENT_API_KEY")
    ]

    rows = []
    for provider, env_key in keys:
        configured = bool(get_secret_value(env_key))
        rows.append({
            "Provider": provider,
            "Environment Variable": env_key,
            "Configured": "Yes" if configured else "No",
            "Status": "Live-ready" if configured else "Canada fallback / manual mode"
        })

    return pd.DataFrame(rows)



def find_evidence_line(text: str, field_name: str, value) -> str:
    """Finds the most relevant OCR line behind an extracted field."""
    if not text or value is None:
        return ""

    keyword_map = {
        "Purchase Price / Cost Basis": ["purchase price", "acquisition price", "cost basis", "contract price"],
        "Appraised Value": ["appraised value", "market value", "as-is value", "as stabilized value"],
        "Gross Income": ["gross income", "rental income", "total income", "revenue", "egi"],
        "Vacancy / Credit Loss": ["vacancy", "credit loss"],
        "Operating Expenses": ["operating expenses", "total expenses", "opex", "property expenses"],
        "Stabilized NOI": ["net operating income", "noi"],
        "Debt Service": ["debt service", "mortgage payment"],
        "Cap Rate": ["cap rate", "capitalization rate"]
    }

    keywords = keyword_map.get(field_name, [field_name.lower()])

    for idx, line in enumerate(text.splitlines(), start=1):
        line_lower = line.lower()
        if any(keyword in line_lower for keyword in keywords):
            return f"Line {idx}: {line.strip()}"

    return ""


def build_extraction_evidence_df(text: str, extracted_fields: Dict[str, Any]) -> pd.DataFrame:
    """Creates a source-evidence table for extracted OCR fields."""
    if not extracted_fields:
        return pd.DataFrame(columns=["Field", "Extracted Value", "Confidence", "Evidence"])

    rows = []
    for field, value in extracted_fields.items():
        rows.append({
            "Field": field,
            "Extracted Value": "" if value is None else (format_pct(value) if field == "Cap Rate" else format_money(value)),
            "Confidence": field_confidence(field, value, text or ""),
            "Evidence": find_evidence_line(text or "", field, value)
        })

    return pd.DataFrame(rows)


def platform_readiness_score(
    deals_df: pd.DataFrame,
    documents_df: pd.DataFrame,
    diligence_pct_value: float,
    api_keys_df: pd.DataFrame,
    extraction_evidence_df: Optional[pd.DataFrame] = None
) -> Tuple[int, pd.DataFrame]:
    """Scores platform readiness toward a 950-grade commercial system."""
    components = []

    deal_db_points = 15 if deals_df is not None and not deals_df.empty else 5
    doc_room_points = 15 if documents_df is not None and not documents_df.empty else 8
    diligence_points = int(min(15, diligence_pct_value * 15))

    api_configured_count = 0
    if api_keys_df is not None and not api_keys_df.empty:
        api_configured_count = (api_keys_df["Configured"] == "Yes").sum()
    api_points = min(15, int(api_configured_count * 2.5))

    extraction_points = 8
    if extraction_evidence_df is not None and not extraction_evidence_df.empty:
        evidence_count = extraction_evidence_df["Evidence"].astype(str).str.len().gt(0).sum()
        extraction_points = min(15, int(8 + evidence_count))

    auth_points = 15 if st.session_state.get("current_user") else 7
    audit_points = 15 if get_audit_df() is not None and not get_audit_df().empty else 8
    export_points = 10

    components.append(("Database-backed deals", deal_db_points, 15))
    components.append(("Document room", doc_room_points, 15))
    components.append(("Diligence workflow", diligence_points, 15))
    components.append(("Live API keys configured", api_points, 15))
    components.append(("Extraction evidence", extraction_points, 15))
    components.append(("Role-based access", auth_points, 15))
    components.append(("Audit trail", audit_points, 15))
    components.append(("Export / backup controls", export_points, 10))

    total = sum(x[1] for x in components)
    max_total = sum(x[2] for x in components)
    score = int(100 * safe_divide(total, max_total))

    return score, pd.DataFrame(components, columns=["Component", "Points", "Max"])


def create_platform_backup_zip(current_deal_json: bytes, excel_file: io.BytesIO, report_text: str) -> io.BytesIO:
    """Creates a platform backup ZIP with DB, current deal JSON, workbook, and manifest."""
    ensure_platform_storage()
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("current_deal.json", current_deal_json)
        z.writestr("current_report.txt", report_text)
        z.writestr("current_workbook.xlsx", excel_file.getvalue())

        if DB_PATH.exists():
            z.write(DB_PATH, "alenza_platform.db")

        docs_df = get_documents_df()
        z.writestr("document_manifest.csv", docs_df.to_csv(index=False))

        deals_df = get_deals_df()
        z.writestr("deals_export.csv", deals_df.to_csv(index=False))

        audit_df = get_audit_df()
        z.writestr("audit_export.csv", audit_df.to_csv(index=False))

        z.writestr("README.txt", (
            "Alenza Platform Backup\n"
            "Includes current deal JSON, report text, workbook, SQLite database, "
            "document manifest, deals export, and audit export.\n"
            "For production, store backups in encrypted cloud storage with retention controls.\n"
        ))

    buffer.seek(0)
    return buffer


def portfolio_metrics(deals_df: pd.DataFrame) -> Dict[str, Any]:
    if deals_df is None or deals_df.empty:
        return {"Deals": 0, "Total Loan Amount": 0, "Average Score": 0, "Average NOI": 0}

    return {
        "Deals": len(deals_df),
        "Total Loan Amount": deals_df["loan_amount"].fillna(0).sum(),
        "Average Score": deals_df["score"].fillna(0).mean(),
        "Average NOI": deals_df["noi"].fillna(0).mean()
    }



# ============================================================
# SESSION INIT
# ============================================================

def default_rent_roll_df() -> pd.DataFrame:
    return pd.DataFrame([
        {"Tenant": "Anchor Tenant A", "SF": 25000, "Remaining Term": 8.5, "Monthly Rent": 45000},
        {"Tenant": "In-Line Shop B", "SF": 4500, "Remaining Term": 3.0, "Monthly Rent": 12000},
        {"Tenant": "Vacant Suite 101", "SF": 2000, "Remaining Term": 0.0, "Monthly Rent": 0},
    ])


def default_lender_quotes_df() -> pd.DataFrame:
    return pd.DataFrame([
        {"Lender": "Bank A", "Loan Amount": 8500000, "Rate": 0.0575, "Term": 5, "Amortization": 25, "Fees": 0.010, "Recourse": "Partial", "Status": "Indication"},
        {"Lender": "Credit Union B", "Loan Amount": 9000000, "Rate": 0.0600, "Term": 5, "Amortization": 30, "Fees": 0.0075, "Recourse": "Limited", "Status": "Quoted"},
        {"Lender": "Debt Fund C", "Loan Amount": 9750000, "Rate": 0.0850, "Term": 3, "Amortization": 0, "Fees": 0.020, "Recourse": "Non-Recourse", "Status": "Indicative"},
    ])


def init_diligence_tracker():
    """Enterprise checklist for nCino-style workflow."""
    items = [
        "Appraisal",
        "Phase I Enviro",
        "Rent Roll",
        "T12 Financials",
        "Sponsor Bio",
        "KYC/AML",
        "Insurance",
        "Title / Survey",
        "Purchase Agreement",
        "Lender Term Sheet"
    ]
    if "diligence" not in st.session_state:
        st.session_state.diligence = {item: "Pending" for item in items}


def init_state() -> None:
    for key, val in DEFAULTS.items():
        st.session_state.setdefault(key, val)
    st.session_state.setdefault("raw_ocr_text", None)
    st.session_state.setdefault("extracted_fields", None)
    st.session_state.setdefault("offline_ai_result", None)
    st.session_state.setdefault("ollama_review_text", None)
    st.session_state.setdefault("audit_log", [])
    st.session_state.setdefault("rent_roll_df", default_rent_roll_df())
    st.session_state.setdefault("lender_quotes_df", default_lender_quotes_df())
    st.session_state.setdefault("property_address", "100 King Street West, Toronto, ON")
    st.session_state.setdefault("concierge_requests", [])
    init_diligence_tracker()


init_state()


# ============================================================
# OCR + DOCUMENT INTAKE
# ============================================================

def extract_text_from_image(uploaded_file) -> str:
    if Image is None or pytesseract is None:
        raise RuntimeError("Image OCR requires pillow and pytesseract.")
    image = Image.open(uploaded_file).convert("RGB")
    return pytesseract.image_to_string(image)


def extract_text_from_pdf(uploaded_file) -> str:
    if fitz is None:
        raise RuntimeError("PDF intake requires pymupdf.")
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    extracted_text = []
    for page in doc:
        native_text = page.get_text("text")
        if native_text and len(native_text.strip()) > 50:
            extracted_text.append(native_text)
        elif Image is not None and pytesseract is not None:
            pix = page.get_pixmap(dpi=225)
            img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            extracted_text.append(pytesseract.image_to_string(img))
        else:
            extracted_text.append("")
    return "\n".join(extracted_text)


def find_money_near_keywords(text: str, keywords: List[str]) -> Optional[float]:
    money_pattern = r"\(?\$?\s*-?\d[\d,]*(?:\.\d+)?\)?"
    for line in text.splitlines():
        normalized = line.lower()
        if any(keyword in normalized for keyword in keywords):
            matches = re.findall(money_pattern, line)
            if matches:
                value = money_to_float(matches[-1])
                if value is not None:
                    return value
    return None


def find_percent_near_keywords(text: str, keywords: List[str]) -> Optional[float]:
    for line in text.splitlines():
        normalized = line.lower()
        if any(keyword in normalized for keyword in keywords):
            matches = re.findall(r"(\d{1,2}(?:\.\d+)?)\s*%", normalized)
            if matches:
                return float(matches[-1]) / 100
    return None


def parse_financials_from_text(text: str) -> Dict[str, Optional[float]]:
    gross_income = find_money_near_keywords(text, ["gross potential income", "gross rental income", "rental income", "total income", "effective gross income", "egi", "revenue"])
    vacancy = find_money_near_keywords(text, ["vacancy", "credit loss", "vacancy loss", "vacancy and credit"])
    operating_expenses = find_money_near_keywords(text, ["operating expenses", "total expenses", "property expenses", "opex", "repairs and maintenance", "taxes and insurance"])
    noi = find_money_near_keywords(text, ["net operating income", "noi", "net income before debt service"])
    debt_service = find_money_near_keywords(text, ["debt service", "annual debt service", "mortgage payment"])
    purchase_price = find_money_near_keywords(text, ["purchase price", "acquisition price", "cost basis", "contract price"])
    appraised_value = find_money_near_keywords(text, ["appraised value", "market value", "as-is value", "as stabilized value"])
    cap_rate = find_percent_near_keywords(text, ["cap rate", "capitalization rate"])

    if noi is None and gross_income is not None and operating_expenses is not None:
        noi = gross_income - abs(vacancy or 0) - operating_expenses

    return {
        "Purchase Price / Cost Basis": purchase_price,
        "Appraised Value": appraised_value,
        "Gross Income": gross_income,
        "Vacancy / Credit Loss": vacancy,
        "Operating Expenses": operating_expenses,
        "Stabilized NOI": noi,
        "Debt Service": debt_service,
        "Cap Rate": cap_rate
    }


def process_uploaded_financial(uploaded_file) -> Tuple[str, Dict[str, Optional[float]]]:
    file_name = uploaded_file.name.lower()
    if file_name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        text = extract_text_from_image(uploaded_file)
    elif file_name.endswith(".pdf"):
        text = extract_text_from_pdf(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Upload PDF, PNG, JPG, JPEG, or WEBP.")
    return text, parse_financials_from_text(text)


# ============================================================
# OFFLINE DOCUMENT INTELLIGENCE + OLLAMA
# ============================================================

def field_confidence(field_name: str, value, text: str) -> int:
    if value is None:
        return 0
    text_lower = text.lower()
    score = 45
    keyword_map = {
        "Purchase Price / Cost Basis": ["purchase price", "acquisition price", "cost basis", "contract price"],
        "Appraised Value": ["appraised value", "market value", "as-is value", "as stabilized value"],
        "Gross Income": ["gross income", "rental income", "total income", "revenue", "egi"],
        "Vacancy / Credit Loss": ["vacancy", "credit loss"],
        "Operating Expenses": ["operating expenses", "total expenses", "opex", "property expenses"],
        "Stabilized NOI": ["net operating income", "noi"],
        "Debt Service": ["debt service", "mortgage payment"],
        "Cap Rate": ["cap rate", "capitalization rate"]
    }
    if any(k in text_lower for k in keyword_map.get(field_name, [])):
        score += 35
    if value:
        score += 10
    if field_name == "Stabilized NOI" and value and value > 0:
        score += 10
    return min(score, 100)


def offline_ai_verify_fields(text: str, extracted_fields: Dict[str, Any]) -> Dict[str, Any]:
    warnings, suggestions, blockers = [], [], []
    pp = extracted_fields.get("Purchase Price / Cost Basis")
    av = extracted_fields.get("Appraised Value")
    gi = extracted_fields.get("Gross Income")
    vac = extracted_fields.get("Vacancy / Credit Loss")
    oe = extracted_fields.get("Operating Expenses")
    noi = extracted_fields.get("Stabilized NOI")
    ds = extracted_fields.get("Debt Service")
    cap = extracted_fields.get("Cap Rate")

    confidence_rows = []
    for field, value in extracted_fields.items():
        display = "" if value is None else (format_pct(value) if field == "Cap Rate" else format_money(value))
        confidence_rows.append({"Field": field, "Extracted Value": display, "Confidence": field_confidence(field, value, text)})

    if noi is None:
        blockers.append("NOI was not confidently extracted. Underwriting should not rely on OCR until reviewed.")

    if gi and oe:
        implied_noi = gi - abs(vac or 0) - oe
        if noi:
            var_pct = safe_divide(abs(noi - implied_noi), max(abs(noi), 1))
            if var_pct > 0.10:
                warnings.append(f"NOI reconciliation variance exceeds 10%. Extracted NOI {format_money(noi)} vs implied NOI {format_money(implied_noi)}.")
            else:
                suggestions.append("NOI reconciles directionally with income and expense fields.")
        else:
            suggestions.append(f"NOI missing; implied NOI from extracted lines is approximately {format_money(implied_noi)}.")

    if gi and oe:
        expense_ratio = safe_divide(oe, gi)
        if expense_ratio < 0.15:
            warnings.append(f"Expense ratio appears unusually low at {format_pct(expense_ratio)}.")
        elif expense_ratio > 0.60:
            warnings.append(f"Expense ratio appears unusually high at {format_pct(expense_ratio)}.")
        else:
            suggestions.append(f"Expense ratio appears reasonable at {format_pct(expense_ratio)}.")

    if noi and av:
        implied_cap = safe_divide(noi, av)
        if implied_cap < 0.03:
            warnings.append(f"Implied cap rate is low at {format_pct(implied_cap)}.")
        elif implied_cap > 0.12:
            warnings.append(f"Implied cap rate is high at {format_pct(implied_cap)}.")
        else:
            suggestions.append(f"Implied cap rate appears reasonable at {format_pct(implied_cap)}.")
        if cap and abs(implied_cap - cap) > 0.015:
            warnings.append(f"Extracted cap rate {format_pct(cap)} differs from NOI/value implied cap rate {format_pct(implied_cap)}.")

    if pp and av:
        variance = safe_divide(av - pp, pp)
        if abs(variance) > 0.20:
            warnings.append(f"Appraised value differs from purchase/cost basis by {format_pct(variance)}.")
        else:
            suggestions.append("Purchase price and appraised value appear directionally aligned.")

    if noi and ds:
        implied_dscr = safe_divide(noi, ds)
        if implied_dscr < 1.00:
            warnings.append(f"Implied DSCR from OCR is below 1.00x at {format_x(implied_dscr)}.")
        else:
            suggestions.append(f"Implied DSCR from OCR is {format_x(implied_dscr)}.")

    return {
        "confidence_df": pd.DataFrame(confidence_rows),
        "warnings": warnings,
        "suggestions": suggestions,
        "blockers": blockers
    }


def test_ollama_connection(ollama_url: str, model: str) -> Tuple[bool, str]:
    try:
        payload = {"model": model, "prompt": "Reply with only OK.", "stream": False}
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            f"{ollama_url.rstrip('/')}/api/generate",
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=20) as response:
            result = json.loads(response.read().decode("utf-8"))
        return True, result.get("response", "").strip()
    except Exception as exc:
        return False, str(exc)


def ollama_review_ocr(ollama_url: str, model: str, ocr_text: str, extracted_fields: Dict[str, Any]) -> str:
    compact_fields = {k: v for k, v in extracted_fields.items() if v is not None}
    prompt = f"""
You are a local offline CRE underwriting analyst. Review OCR output and extracted fields.

Rules:
- Be concise.
- Do not invent facts.
- Flag uncertain values.
- Suggest corrections only if supported by OCR text.
- Focus on NOI, income, expenses, value, cap rate, debt service, lease risk, and underwriting risk.

Extracted fields:
{json.dumps(compact_fields, indent=2)}

OCR text:
{ocr_text[:9000]}

Return:
SUMMARY:
POSSIBLE ISSUES:
SUGGESTED CORRECTIONS:
UNDERWRITING NOTES:
"""
    payload = {"model": model, "prompt": prompt, "stream": False, "options": {"temperature": 0.1}}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        f"{ollama_url.rstrip('/')}/api/generate",
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))
        return result.get("response", "No response returned from Ollama.")
    except Exception as exc:
        return f"Ollama review failed: {exc}"


# ============================================================
# UNDERWRITING ENGINE
# ============================================================

def monthly_payment_amortizing(loan_amount: float, rate: float, amort_years: int) -> float:
    monthly_rate = rate / 12
    periods = amort_years * 12
    if loan_amount <= 0 or periods <= 0:
        return 0
    if monthly_rate <= 0:
        return loan_amount / periods
    return (loan_amount * monthly_rate) / (1 - (1 + monthly_rate) ** -periods)


def monthly_payment_interest_only(loan_amount: float, rate: float) -> float:
    if loan_amount <= 0 or rate <= 0:
        return 0
    return loan_amount * rate / 12


def loan_balance_at_maturity(loan_amount: float, rate: float, amort_years: int, term_years: int, debt_structure: str) -> float:
    if debt_structure == "Interest-Only":
        return loan_amount
    monthly_rate = rate / 12
    paid = min(term_years * 12, amort_years * 12)
    if loan_amount <= 0:
        return 0
    if monthly_rate <= 0:
        return max(loan_amount - (loan_amount / (amort_years * 12)) * paid, 0)
    pmt = monthly_payment_amortizing(loan_amount, rate, amort_years)
    balance = loan_amount * (1 + monthly_rate) ** paid - pmt * (((1 + monthly_rate) ** paid - 1) / monthly_rate)
    return max(balance, 0)


def calculate_monthly_payment(loan_amount: float, rate: float, amort_years: int, debt_structure: str) -> float:
    if debt_structure == "Interest-Only":
        return monthly_payment_interest_only(loan_amount, rate)
    return monthly_payment_amortizing(loan_amount, rate, amort_years)


def size_loan(noi: float, appraisal: float, total_cost: float, rate: float, amort_years: int,
              target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float,
              debt_structure: str) -> Tuple[float, str, Dict[str, float]]:
    monthly_rate = rate / 12
    periods = amort_years * 12

    ltv_limit = appraisal * target_ltv
    ltc_limit = total_cost * target_ltc
    monthly_dscr_capacity = (noi / target_dscr) / 12 if target_dscr > 0 else 0

    if debt_structure == "Interest-Only":
        dscr_limit = monthly_dscr_capacity / monthly_rate if monthly_rate > 0 else 0
    else:
        dscr_limit = monthly_dscr_capacity * ((1 - (1 + monthly_rate) ** -periods) / monthly_rate) if monthly_rate > 0 and periods > 0 else 0

    debt_yield_limit = noi / target_dy if target_dy > 0 else 0

    gates = {"LTV": ltv_limit, "LTC": ltc_limit, "DSCR": dscr_limit, "Debt Yield": debt_yield_limit}
    supportable_loan = min(gates.values())
    binding_gate = min(gates, key=gates.get)
    return supportable_loan, binding_gate, gates


# ============================================================
# WORKFLOW ENGINES: RENT ROLL & AMORTIZATION
# ============================================================

def calculate_rent_roll_metrics(df: pd.DataFrame) -> Dict[str, Any]:
    """Computes institutional lease exposure metrics."""
    if df is None or df.empty:
        return {"WALT": 0, "Expiring_1yr_pct": 0, "Total_SF": 0, "Occupancy": 0, "Annual_Rent": 0, "Average_Rent_PSF": 0, "Vacancy_SF": 0}

    # Supports both requested schema and older internal schema.
    df = df.copy()
    if "Remaining_Term" in df.columns and "Remaining Term" not in df.columns:
        df["Remaining Term"] = df["Remaining_Term"]
    if "Rent_PSF" in df.columns and "Monthly Rent" not in df.columns:
        df["Monthly Rent"] = pd.to_numeric(df["Rent_PSF"], errors="coerce").fillna(0) * pd.to_numeric(df.get("SF", 0), errors="coerce").fillna(0) / 12

    required = {"SF", "Remaining Term", "Monthly Rent"}
    if not required.issubset(df.columns):
        return {"WALT": 0, "Expiring_1yr_pct": 0, "Total_SF": 0, "Occupancy": 0, "Annual_Rent": 0, "Average_Rent_PSF": 0, "Vacancy_SF": 0}

    df["SF"] = pd.to_numeric(df["SF"], errors="coerce").fillna(0)
    df["Remaining Term"] = pd.to_numeric(df["Remaining Term"], errors="coerce").fillna(0)
    df["Monthly Rent"] = pd.to_numeric(df["Monthly Rent"], errors="coerce").fillna(0)

    total_sf = df["SF"].sum()
    if total_sf <= 0:
        return {"WALT": 0, "Expiring_1yr_pct": 0, "Total_SF": 0, "Occupancy": 0, "Annual_Rent": 0, "Average_Rent_PSF": 0, "Vacancy_SF": 0}

    occupied_df = df[df["Monthly Rent"] > 0].copy()
    occupied_sf = occupied_df["SF"].sum()
    vacancy_sf = total_sf - occupied_sf

    if occupied_sf > 0:
        walt = (occupied_df["Remaining Term"] * (occupied_df["SF"] / occupied_sf)).sum()
    else:
        walt = 0

    expiring_sf = occupied_df[occupied_df["Remaining Term"] <= 1]["SF"].sum()
    annual_rent = occupied_df["Monthly Rent"].sum() * 12

    return {
        "WALT": walt,
        "Expiring_1yr_pct": safe_divide(expiring_sf, occupied_sf),
        "Total_SF": total_sf,
        "Occupancy": safe_divide(occupied_sf, total_sf),
        "Annual_Rent": annual_rent,
        "Average_Rent_PSF": safe_divide(annual_rent, occupied_sf),
        "Vacancy_SF": vacancy_sf
    }


def generate_amortization_schedule(loan_amt: float, rate: float, amort_years: int, term_years: int, debt_structure: str) -> pd.DataFrame:
    """Generates month-by-month debt service breakdown."""
    m_rate = rate / 12
    term_months = int(term_years * 12)
    amort_months = int(amort_years * 12)
    rows = []
    curr_balance = float(loan_amt)

    if loan_amt <= 0 or term_months <= 0:
        return pd.DataFrame(columns=["Period", "Opening Balance", "Total Payment", "Principal", "Interest", "Closing Balance"])

    if debt_structure == "Interest-Only":
        pmt = loan_amt * m_rate
    else:
        if m_rate > 0 and amort_months > 0:
            pmt = (loan_amt * m_rate) / (1 - (1 + m_rate) ** -amort_months)
        elif amort_months > 0:
            pmt = loan_amt / amort_months
        else:
            pmt = 0

    for i in range(1, term_months + 1):
        interest_payment = curr_balance * m_rate
        principal_payment = 0 if debt_structure == "Interest-Only" else pmt - interest_payment
        principal_payment = max(0, min(principal_payment, curr_balance))
        closing_balance = curr_balance - principal_payment

        rows.append({
            "Period": i,
            "Opening Balance": curr_balance,
            "Total Payment": pmt,
            "Principal": principal_payment,
            "Interest": interest_payment,
            "Closing Balance": max(0, closing_balance)
        })

        curr_balance = closing_balance
        if curr_balance <= 0:
            break

    return pd.DataFrame(rows)


def diligence_progress() -> float:
    init_diligence_tracker()
    statuses = list(st.session_state.diligence.values())
    if not statuses:
        return 0
    complete = sum(1 for status in statuses if status in ["Reviewed", "Waived"])
    return complete / len(statuses)


def evaluate_lender_quotes(quotes_df: pd.DataFrame, noi: float, appraisal: float, total_uses: float) -> pd.DataFrame:
    if quotes_df is None or quotes_df.empty:
        return pd.DataFrame()

    df = quotes_df.copy()
    for col in ["Loan Amount", "Rate", "Term", "Amortization", "Fees"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    dscrs, ltvs, ltcs, dys, ads_list, fee_dollars = [], [], [], [], [], []

    for _, row in df.iterrows():
        loan_amount = row["Loan Amount"]
        rate = row["Rate"]
        amort = int(row["Amortization"])
        if amort <= 0:
            monthly_payment = monthly_payment_interest_only(loan_amount, rate)
        else:
            monthly_payment = monthly_payment_amortizing(loan_amount, rate, amort)

        ads = monthly_payment * 12
        ads_list.append(ads)
        dscrs.append(safe_divide(noi, ads))
        ltvs.append(safe_divide(loan_amount, appraisal))
        ltcs.append(safe_divide(loan_amount, total_uses))
        dys.append(safe_divide(noi, loan_amount))
        fee_dollars.append(loan_amount * row["Fees"])

    df["Annual Debt Service"] = ads_list
    df["DSCR"] = dscrs
    df["LTV"] = ltvs
    df["LTC"] = ltcs
    df["Debt Yield"] = dys
    df["Fee Dollars"] = fee_dollars

    df["Quote Score"] = (
        df["Loan Amount"].rank(pct=True) * 35
        + df["DSCR"].rank(pct=True) * 25
        + (1 - df["Rate"].rank(pct=True)) * 25
        + (1 - df["Fees"].rank(pct=True)) * 15
    ).round(0).astype(int)

    return df.sort_values("Quote Score", ascending=False)


def constraint_advice(binding_gate: str) -> str:
    if binding_gate == "LTV":
        return "Leverage-constrained. Increase value, lower requested leverage, add collateral, or find a lender willing to advance higher LTV."
    if binding_gate == "LTC":
        return "Cost-constrained. Reduce total cost, increase equity, or use a lender with higher LTC tolerance."
    if binding_gate == "DSCR":
        return "Cash-flow constrained. Improve NOI, lower rate, use interest-only, extend amortization, or reduce DSCR requirement."
    return "Debt-yield constrained. Improve NOI, lower debt-yield threshold, or evidence stronger compensating credit factors."


def classify_deal(score: int) -> str:
    if score >= 925:
        return "Tier 1A | Institutional Core Credit"
    if score >= 850:
        return "Tier 1 | High Bankability"
    if score >= 725:
        return "Tier 2 | Bankable / Credit Union / Select Alternative"
    if score >= 575:
        return "Tier 3 | Alternative / Structured Credit"
    return "Tier 4 | Private / Bridge / Restructure Required"


def pass_fail(actual: float, threshold: float, mode: str = "gte") -> str:
    if mode == "gte":
        return "PASS" if actual >= threshold else "FAIL"
    return "PASS" if actual <= threshold else "FAIL"


def score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, equity_pct, property_type, risk_flags_count) -> Tuple[int, Dict[str, int]]:
    risk_weight = PROPERTY_RISK_WEIGHTS.get(property_type, 1.10)
    safe_weight = max(risk_weight, 0.1)

    ltv_score = 230 if actual_ltv <= 0.60 else 200 if actual_ltv <= 0.65 else 165 if actual_ltv <= 0.70 else 125 if actual_ltv <= 0.75 else 65 if actual_ltv <= 0.80 else 0
    ltc_score = 130 if actual_ltc <= 0.65 else 110 if actual_ltc <= 0.70 else 85 if actual_ltc <= 0.75 else 55 if actual_ltc <= 0.80 else 25 if actual_ltc <= 0.85 else 0
    dscr_score = 250 if actual_dscr >= 1.50 else 220 if actual_dscr >= 1.40 else 175 if actual_dscr >= 1.30 else 130 if actual_dscr >= 1.25 else 60 if actual_dscr >= 1.15 else 0
    dy_score = 190 if actual_dy >= 0.10 else 165 if actual_dy >= 0.09 else 125 if actual_dy >= 0.08 else 70 if actual_dy >= 0.07 else 0
    equity_score = 120 if equity_pct >= 0.35 else 100 if equity_pct >= 0.30 else 80 if equity_pct >= 0.25 else 50 if equity_pct >= 0.20 else 20 if equity_pct >= 0.15 else 0
    asset_score = min(80, int(80 / safe_weight))
    penalty = min(100, risk_flags_count * 20)

    components = {
        "Loan-to-Value": ltv_score,
        "Loan-to-Cost": ltc_score,
        "Debt Service Coverage": dscr_score,
        "Debt Yield": dy_score,
        "Equity Contribution": equity_score,
        "Asset Type Risk": asset_score,
        "Risk Flag Penalty": -penalty
    }
    return max(0, min(1000, sum(components.values()))), components


def generate_risk_flags(actual_ltv, actual_ltc, actual_dscr, actual_dy, equity_pct, appraisal_premium, property_type, debt_structure, rent_roll_metrics) -> List[str]:
    flags = []
    if actual_ltv > 0.80:
        flags.append("LTV exceeds 80%; proceeds likely require stretch, bridge, or private capital.")
    if actual_ltc > 0.85:
        flags.append("LTC exceeds 85%; sponsor equity contribution appears thin.")
    if actual_dscr < 1.15:
        flags.append("DSCR below 1.15x; coverage risk is elevated.")
    if actual_dy < 0.07:
        flags.append("Debt yield below 7.0%; lender proceeds may be materially constrained.")
    if equity_pct < 0.15:
        flags.append("Equity contribution below 15%; alignment and capitalization risk.")
    if abs(appraisal_premium) > 0.25:
        flags.append("Appraised value differs from purchase/cost basis by more than 25%.")
    if property_type in ["Office", "Hospitality"]:
        flags.append(f"{property_type} carries elevated sector scrutiny in many lender processes.")
    if debt_structure == "Interest-Only" and actual_dscr < 1.25:
        flags.append("Interest-only structure is masking amortizing DSCR weakness.")
    if rent_roll_metrics.get("WALT", 0) < 2 and rent_roll_metrics.get("Total_SF", 0) > 0:
        flags.append("WALT below 2.0 years; rollover risk is elevated.")
    if rent_roll_metrics.get("Expiring_1yr_pct", 0) > 0.25:
        flags.append("More than 25% of occupied SF expires within 12 months.")
    if rent_roll_metrics.get("Occupancy", 1) < 0.85 and rent_roll_metrics.get("Total_SF", 0) > 0:
        flags.append("Physical occupancy below 85%.")
    return flags


# ============================================================
# EXPORT ENGINES
# ============================================================

def create_excel_workbook(sheets: Dict[str, pd.DataFrame], metadata: Dict[str, Any], raw_ocr_text=None, offline_notes=None, ollama_review_text=None) -> io.BytesIO:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        header_format = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1D4ED8", "border": 1})
        title_format = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1D4ED8"})
        normal_format = workbook.add_format({"border": 1})

        cover = workbook.add_worksheet("Cover")
        cover.write("A1", "ALENZA CAPITAL UNDERWRITING WORKBOOK", title_format)
        row = 2
        for k, v in metadata.items():
            cover.write(row, 0, k, header_format)
            cover.write(row, 1, v, normal_format)
            row += 1
        cover.set_column("A:A", 30)
        cover.set_column("B:B", 46)

        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            include_index = sheet_name in ["Sensitivity Proceeds", "Sensitivity Gates", "Sensitivity Score"]
            df.to_excel(writer, sheet_name=safe_name, index=include_index)
            ws = writer.sheets[safe_name]
            ws.freeze_panes(1, 0)
            ws.set_column("A:A", 30)
            ws.set_column("B:Z", 22)
            headers = df.reset_index().columns if include_index else df.columns
            for col_num, value in enumerate(headers):
                ws.write(0, col_num, value, header_format)

        if offline_notes:
            ws = workbook.add_worksheet("Offline AI Notes")
            ws.write("A1", "Offline AI Verification Notes", title_format)
            row = 3
            for section, items in offline_notes.items():
                ws.write(row, 0, section, header_format)
                row += 1
                for item in items:
                    ws.write(row, 0, item)
                    row += 1
                row += 2
            ws.set_column("A:A", 130)

        if ollama_review_text:
            ws = workbook.add_worksheet("Ollama Review")
            ws.write("A1", "Local Ollama Review", title_format)
            ws.write("A3", ollama_review_text)
            ws.set_column("A:A", 130)

        if raw_ocr_text:
            ws = workbook.add_worksheet("Raw OCR Text")
            ws.write("A1", "Raw OCR Output", title_format)
            ws.write("A3", raw_ocr_text)
            ws.set_column("A:A", 130)

    output.seek(0)
    return output


def create_pdf_summary(metadata: Dict[str, Any], preview_df: pd.DataFrame, covenant_df: pd.DataFrame, score_df: pd.DataFrame, risk_flags: List[str], offline_warnings: List[str], ollama_review_text: Optional[str]) -> Optional[io.BytesIO]:
    if SimpleDocTemplate is None:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [Paragraph("ALENZA CAPITAL UNDERWRITING SUMMARY", styles["Title"]), Spacer(1, 10)]
    meta_html = "<br/>".join([f"<b>{k}:</b> {v}" for k, v in metadata.items()])
    story.append(Paragraph(meta_html, styles["Normal"]))
    story.append(Spacer(1, 14))

    def df_to_table(title: str, df: pd.DataFrame):
        story.append(Paragraph(title, styles["Heading2"]))
        table_data = [list(df.columns)] + df.astype(str).values.tolist()
        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

    df_to_table("Executive Summary", preview_df)
    df_to_table("Covenant Testing", covenant_df)
    df_to_table("Scorecard", score_df)

    if risk_flags:
        story.append(Paragraph("Risk Flags", styles["Heading2"]))
        story.append(Paragraph("<br/>".join([f"- {f}" for f in risk_flags]), styles["Normal"]))
        story.append(Spacer(1, 12))

    if offline_warnings:
        story.append(Paragraph("Offline AI Review Items", styles["Heading2"]))
        story.append(Paragraph("<br/>".join([f"- {w}" for w in offline_warnings]), styles["Normal"]))
        story.append(Spacer(1, 12))

    if ollama_review_text:
        story.append(Paragraph("Local Ollama Review", styles["Heading2"]))
        story.append(Paragraph(ollama_review_text[:2400].replace("\n", "<br/>"), styles["Normal"]))
        story.append(Spacer(1, 12))

    disclaimer = "This summary is indicative only and is not a loan commitment, credit approval, appraisal, legal opinion, or final underwriting decision. All terms are subject to lender diligence, borrower review, third-party reports, credit approval, committee review, and final documentation."
    story.append(Paragraph("Disclaimer", styles["Heading2"]))
    story.append(Paragraph(disclaimer, styles["Normal"]))
    doc.build(story)
    buffer.seek(0)
    return buffer


def build_deal_json(data: Dict[str, Any]) -> bytes:
    return json.dumps(data, indent=2).encode("utf-8")



# ============================================================
# PLATFORM LOGIN
# ============================================================

if "current_user" not in st.session_state:
    st.session_state.current_user = None

with st.sidebar:
    st.markdown("### Platform Access")

    if st.session_state.current_user:
        st.success(f"{st.session_state.current_user['role']}: {st.session_state.current_user['username']}")
        if st.button("Log Out"):
            platform_audit("Logout", st.session_state.current_user["username"])
            st.session_state.current_user = None
            st.rerun()
    else:
        with st.form("login_form"):
            login_username = st.text_input("Username", "admin")
            login_password = st.text_input("Password", "alenza-admin", type="password")
            login_submit = st.form_submit_button("Log In")

            if login_submit:
                user = verify_login(login_username, login_password)
                if user:
                    st.session_state.current_user = user
                    platform_audit("Login", login_username)
                    st.success("Logged in.")
                    st.rerun()
                else:
                    st.error("Invalid login.")

        st.caption("Demo users: admin / broker / borrower / lender. Password format: alenza-role.")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.title("ALENZA CAPITAL")
    st.caption(f"Institutional CRE Debt Sizing | {APP_VERSION}")
    st.markdown("---")

    with st.expander("Deal File", expanded=False):
        uploaded_deal = st.file_uploader("Load saved Alenza deal JSON", type=["json"])
        if uploaded_deal is not None:
            try:
                loaded = json.loads(uploaded_deal.read().decode("utf-8"))
                for k, v in loaded.get("inputs", {}).items():
                    if k in DEFAULTS:
                        st.session_state[k] = v
                if loaded.get("rent_roll"):
                    st.session_state.rent_roll_df = pd.DataFrame(loaded["rent_roll"])
                if loaded.get("diligence"):
                    st.session_state.diligence = loaded["diligence"]
                if loaded.get("lender_quotes"):
                    st.session_state.lender_quotes_df = pd.DataFrame(loaded["lender_quotes"])
                add_audit("Deal Loaded", uploaded_deal.name)
                st.success("Deal loaded. Review inputs below.")
            except Exception as exc:
                st.error(f"Could not load deal JSON: {exc}")

    with st.expander("Auto Intake + Offline AI", expanded=False):
        uploaded_financial = st.file_uploader("Upload financial statement", type=["pdf", "png", "jpg", "jpeg", "webp"])
        offline_ai_enabled = st.checkbox("Enable offline AI verifier", value=True)

        with st.expander("Advanced AI Settings (Local Only)", expanded=False):
            use_ollama = st.checkbox("Connect to Local Ollama", value=False)
            ollama_url = st.text_input("Ollama URL", "http://localhost:11434")
            ollama_model = st.text_input("Ollama Model", "llama3.2:3b")
            if use_ollama and st.button("Verify Connection"):
                ok, msg = test_ollama_connection(ollama_url, ollama_model)
                st.success("Connected") if ok else st.error("Local Server Not Found")

        if uploaded_financial is not None:
            try:
                raw_ocr_text, extracted_fields = process_uploaded_financial(uploaded_financial)
                st.session_state.raw_ocr_text = raw_ocr_text
                st.session_state.extracted_fields = extracted_fields
                add_audit("Document OCR", uploaded_financial.name)

                extracted_review_df = pd.DataFrame({
                    "Field": list(extracted_fields.keys()),
                    "Extracted Value": ["" if v is None else (format_pct(v) if k == "Cap Rate" else format_money(v)) for k, v in extracted_fields.items()]
                })
                st.success("Document processed.")
                st.dataframe(extracted_review_df, hide_index=True, use_container_width=True)

                if offline_ai_enabled:
                    ai_result = offline_ai_verify_fields(raw_ocr_text, extracted_fields)
                    st.session_state.offline_ai_result = ai_result
                    st.write("Offline AI Confidence")
                    st.dataframe(ai_result["confidence_df"], hide_index=True, use_container_width=True)
                    for b in ai_result.get("blockers", []):
                        st.error(b)
                    for w in ai_result.get("warnings", []):
                        st.warning(w)
                    for s in ai_result.get("suggestions", []):
                        st.info(s)

                col_apply, col_ollama = st.columns(2)
                with col_apply:
                    if st.button("Apply OCR Values"):
                        if extracted_fields.get("Purchase Price / Cost Basis"):
                            st.session_state.purchase_price = int(extracted_fields["Purchase Price / Cost Basis"])
                        if extracted_fields.get("Appraised Value"):
                            st.session_state.appraisal = int(extracted_fields["Appraised Value"])
                        if extracted_fields.get("Stabilized NOI"):
                            st.session_state.noi = int(extracted_fields["Stabilized NOI"])
                        if extracted_fields.get("Gross Income"):
                            st.session_state.gross_income = int(extracted_fields["Gross Income"])
                        if extracted_fields.get("Operating Expenses"):
                            st.session_state.operating_expenses = int(extracted_fields["Operating Expenses"])
                        if extracted_fields.get("Vacancy / Credit Loss"):
                            st.session_state.vacancy_loss = int(abs(extracted_fields["Vacancy / Credit Loss"]))
                        add_audit("OCR Values Applied", uploaded_financial.name)
                        st.success("OCR values applied.")

                with col_ollama:
                    if use_ollama and st.button("Run Ollama Review"):
                        with st.spinner("Running local Ollama review..."):
                            ollama_text = ollama_review_ocr(ollama_url, ollama_model, raw_ocr_text, extracted_fields)
                        st.session_state.ollama_review_text = ollama_text
                        add_audit("Ollama Review", f"{ollama_model} @ {ollama_url}")
                        st.markdown(ollama_text)

                with st.expander("Raw OCR Text", expanded=False):
                    st.text_area("OCR Output", raw_ocr_text, height=260)

            except Exception as exc:
                st.error(f"OCR processing failed: {exc}")
                st.caption("Install pillow, pytesseract, pymupdf, and add tesseract-ocr to packages.txt for Streamlit Cloud.")

    with st.expander("Asset Information", expanded=True):
        st.session_state.sponsor = st.text_input("Sponsor / Borrower", st.session_state.sponsor)
        st.session_state.property_address = st.text_input("Property Address", st.session_state.property_address)
        property_types = list(PROPERTY_RISK_WEIGHTS.keys())
        st.session_state.property_type = st.selectbox("Property Type", property_types, index=property_types.index(st.session_state.property_type) if st.session_state.property_type in property_types else 0)
        st.session_state.transaction_type = st.selectbox("Transaction Type", ["Acquisition", "Refinance"], index=0 if st.session_state.transaction_type == "Acquisition" else 1)
        st.session_state.purchase_price = st.number_input("Purchase Price / Cost Basis ($)", value=int(st.session_state.purchase_price), min_value=1, step=100000)
        st.session_state.appraisal = st.number_input("Appraised Value ($)", value=int(st.session_state.appraisal), min_value=1, step=100000)
        if st.session_state.transaction_type == "Refinance":
            st.session_state.existing_debt = st.number_input("Existing Debt Payoff ($)", value=int(st.session_state.existing_debt), min_value=0, step=100000)
        st.session_state.gross_income = st.number_input("Gross Income / EGI ($)", value=int(st.session_state.gross_income), min_value=1, step=10000)
        st.session_state.vacancy_loss = st.number_input("Vacancy / Credit Loss ($)", value=int(st.session_state.vacancy_loss), min_value=0, step=5000)
        st.session_state.operating_expenses = st.number_input("Operating Expenses ($)", value=int(st.session_state.operating_expenses), min_value=0, step=10000)
        st.session_state.noi = st.number_input("Stabilized NOI ($)", value=int(st.session_state.noi), min_value=1, step=10000)

    with st.expander("Underwriting Criteria", expanded=True):
        profiles = ["Conservative", "Market", "Aggressive", "Custom"]
        st.session_state.credit_profile = st.selectbox("Credit Profile", profiles, index=profiles.index(st.session_state.credit_profile) if st.session_state.credit_profile in profiles else 1)
        if st.button("Apply Credit Profile Preset"):
            preset = PROFILE_PRESETS.get(st.session_state.credit_profile, {})
            for k, v in preset.items():
                st.session_state[k] = v
            add_audit("Profile Applied", st.session_state.credit_profile)
            st.success(f"{st.session_state.credit_profile} preset applied.")
        st.session_state.target_ltv = st.slider("Maximum LTV (%)", 50, 85, int(st.session_state.target_ltv * 100)) / 100
        st.session_state.target_ltc = st.slider("Maximum LTC (%)", 50, 90, int(st.session_state.target_ltc * 100)) / 100
        st.session_state.target_dscr = st.slider("Minimum DSCR (x)", 1.10, 1.75, float(st.session_state.target_dscr), 0.05)
        st.session_state.target_dy = st.slider("Minimum Debt Yield (%)", 5.0, 15.0, float(st.session_state.target_dy * 100), 0.25) / 100

    with st.expander("Loan Terms", expanded=True):
        st.session_state.debt_structure = st.selectbox("Debt Service Structure", ["Amortizing", "Interest-Only"], index=0 if st.session_state.debt_structure == "Amortizing" else 1)
        st.session_state.rate = st.slider("Interest Rate (%)", 3.0, 12.0, float(st.session_state.rate * 100), 0.125) / 100
        st.session_state.amort = st.number_input("Amortization (Years)", value=int(st.session_state.amort), min_value=1, max_value=40)
        st.session_state.loan_term = st.number_input("Loan Term (Years)", value=int(st.session_state.loan_term), min_value=1, max_value=30)
        st.session_state.fees = st.slider("Origination / Financing Fees (%)", 0.0, 5.0, float(st.session_state.fees * 100), 0.25) / 100
        st.session_state.closing_costs = st.number_input("Other Closing Costs ($)", value=int(st.session_state.closing_costs), min_value=0, step=5000)

    with st.expander("Reserves / Adjustments", expanded=False):
        st.session_state.capex_reserve = st.number_input("CapEx / TI-LC Reserve ($)", value=int(st.session_state.capex_reserve), min_value=0, step=25000)
        st.session_state.interest_reserve = st.number_input("Interest Reserve ($)", value=int(st.session_state.interest_reserve), min_value=0, step=25000)


# ============================================================
# CALCULATIONS
# ============================================================

inputs = {k: st.session_state[k] for k in DEFAULTS.keys()}
base_uses = inputs["purchase_price"] if inputs["transaction_type"] == "Acquisition" else inputs["existing_debt"]
initial_total_cost = base_uses + inputs["closing_costs"] + inputs["capex_reserve"] + inputs["interest_reserve"]

loan_tmp, _, _ = size_loan(
    inputs["noi"], inputs["appraisal"], initial_total_cost, inputs["rate"], inputs["amort"],
    inputs["target_ltv"], inputs["target_ltc"], inputs["target_dscr"], inputs["target_dy"], inputs["debt_structure"]
)
total_uses = initial_total_cost + loan_tmp * inputs["fees"]

loan_amt, gate, gates = size_loan(
    inputs["noi"], inputs["appraisal"], total_uses, inputs["rate"], inputs["amort"],
    inputs["target_ltv"], inputs["target_ltc"], inputs["target_dscr"], inputs["target_dy"], inputs["debt_structure"]
)

financing_fees = loan_amt * inputs["fees"]
total_uses = initial_total_cost + financing_fees
required_equity = total_uses - loan_amt

monthly_payment = calculate_monthly_payment(loan_amt, inputs["rate"], inputs["amort"], inputs["debt_structure"])
annual_debt_service = monthly_payment * 12
balloon_balance = loan_balance_at_maturity(loan_amt, inputs["rate"], inputs["amort"], inputs["loan_term"], inputs["debt_structure"])
amort_schedule_df = generate_amortization_schedule(loan_amt, inputs["rate"], inputs["amort"], inputs["loan_term"], inputs["debt_structure"])

rent_roll_metrics = calculate_rent_roll_metrics(st.session_state.rent_roll_df)
diligence_pct = diligence_progress()
lender_quotes_eval_df = evaluate_lender_quotes(st.session_state.lender_quotes_df, inputs["noi"], inputs["appraisal"], total_uses)

actual_ltv = safe_divide(loan_amt, inputs["appraisal"])
actual_ltc = safe_divide(loan_amt, total_uses)
actual_dscr = safe_divide(inputs["noi"], annual_debt_service)
actual_dy = safe_divide(inputs["noi"], loan_amt)
equity_pct = safe_divide(required_equity, total_uses)
expense_ratio = safe_divide(inputs["operating_expenses"], inputs["gross_income"])
implied_cap_rate = safe_divide(inputs["noi"], inputs["appraisal"])
debt_service_cushion = actual_dscr - inputs["target_dscr"]
appraisal_premium = safe_divide(inputs["appraisal"] - inputs["purchase_price"], inputs["purchase_price"])

ltv_status = pass_fail(actual_ltv, inputs["target_ltv"], "lte")
ltc_status = pass_fail(actual_ltc, inputs["target_ltc"], "lte")
dscr_status = pass_fail(actual_dscr, inputs["target_dscr"], "gte")
dy_status = pass_fail(actual_dy, inputs["target_dy"], "gte")

risk_flags = generate_risk_flags(actual_ltv, actual_ltc, actual_dscr, actual_dy, equity_pct, appraisal_premium, inputs["property_type"], inputs["debt_structure"], rent_roll_metrics)
score, score_components = score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, equity_pct, inputs["property_type"], len(risk_flags))
classification = classify_deal(score)
generated_at = now_str()


# ============================================================
# DATAFRAMES
# ============================================================

sizing_df = pd.DataFrame({
    "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
    "Threshold": [format_pct(inputs["target_ltv"]), format_pct(inputs["target_ltc"]), format_x(inputs["target_dscr"]), format_pct(inputs["target_dy"])],
    "Max Proceeds": [gates["LTV"], gates["LTC"], gates["DSCR"], gates["Debt Yield"]],
    "Proceeds Gap": [gates["LTV"] - loan_amt, gates["LTC"] - loan_amt, gates["DSCR"] - loan_amt, gates["Debt Yield"] - loan_amt],
    "Binding": ["YES" if gate == "LTV" else "", "YES" if gate == "LTC" else "", "YES" if gate == "DSCR" else "", "YES" if gate == "Debt Yield" else ""]
})

uses_df = pd.DataFrame({
    "Project Uses": [
        "Purchase Price / Cost Basis" if inputs["transaction_type"] == "Acquisition" else "Existing Debt Payoff",
        "Origination / Financing Fees",
        "Other Closing Costs",
        "CapEx / TI-LC Reserve",
        "Interest Reserve",
        "Total Uses"
    ],
    "Amount": [base_uses, financing_fees, inputs["closing_costs"], inputs["capex_reserve"], inputs["interest_reserve"], total_uses]
})

sources_df = pd.DataFrame({
    "Project Sources": ["Supportable Senior Debt", "Required Sponsor Equity", "Total Sources"],
    "Amount": [loan_amt, required_equity, loan_amt + required_equity]
})

covenant_df = pd.DataFrame({
    "Covenant": ["Maximum LTV", "Maximum LTC", "Minimum DSCR", "Minimum Debt Yield"],
    "Required": [f"≤ {format_pct(inputs['target_ltv'])}", f"≤ {format_pct(inputs['target_ltc'])}", f"≥ {format_x(inputs['target_dscr'])}", f"≥ {format_pct(inputs['target_dy'])}"],
    "Actual": [format_pct(actual_ltv), format_pct(actual_ltc), format_x(actual_dscr), format_pct(actual_dy)],
    "Status": [ltv_status, ltc_status, dscr_status, dy_status]
})

assumptions_df = pd.DataFrame({
    "Assumption": [
        "Sponsor / Borrower", "Property Type", "Transaction Type", "Credit Profile", "Purchase Price / Cost Basis",
        "Appraised Value", "Existing Debt Payoff", "Gross Income / EGI", "Vacancy / Credit Loss", "Operating Expenses",
        "Stabilized NOI", "Maximum LTV", "Maximum LTC", "Minimum DSCR", "Minimum Debt Yield", "Debt Service Structure",
        "Interest Rate", "Amortization", "Loan Term", "Origination / Financing Fees", "Other Closing Costs",
        "CapEx / TI-LC Reserve", "Interest Reserve", "Generated"
    ],
    "Value": [
        inputs["sponsor"], inputs["property_type"], inputs["transaction_type"], inputs["credit_profile"], format_money(inputs["purchase_price"]),
        format_money(inputs["appraisal"]), format_money(inputs["existing_debt"]), format_money(inputs["gross_income"]),
        format_money(inputs["vacancy_loss"]), format_money(inputs["operating_expenses"]), format_money(inputs["noi"]),
        format_pct(inputs["target_ltv"]), format_pct(inputs["target_ltc"]), format_x(inputs["target_dscr"]),
        format_pct(inputs["target_dy"]), inputs["debt_structure"], format_pct(inputs["rate"]), f"{inputs['amort']} years",
        f"{inputs['loan_term']} years", format_pct(inputs["fees"]), format_money(inputs["closing_costs"]),
        format_money(inputs["capex_reserve"]), format_money(inputs["interest_reserve"]), generated_at
    ]
})

score_df = pd.DataFrame({
    "Component": list(score_components.keys()),
    "Score": list(score_components.values()),
    "Maximum": [230, 130, 250, 190, 120, 80, 0]
})

preview_df = pd.DataFrame({
    "Field": [
        "Sponsor / Borrower", "Property Address", "Property Type", "Transaction Type", "Supportable Proceeds", "Binding Constraint",
        "Actual LTV", "Actual LTC", "Actual DSCR", "Debt Yield", "Required Equity", "Balloon Balance",
        "WALT", "Occupancy", "Diligence Progress", "Deal Score", "Classification"
    ],
    "Value": [
        inputs["sponsor"], inputs["property_address"], inputs["property_type"], inputs["transaction_type"], format_money(loan_amt), gate,
        format_pct(actual_ltv), format_pct(actual_ltc), format_x(actual_dscr), format_pct(actual_dy),
        format_money(required_equity), format_money(balloon_balance), f"{rent_roll_metrics['WALT']:.2f} yrs",
        format_pct(rent_roll_metrics["Occupancy"]), format_pct(diligence_pct), f"{score}/1000", classification
    ]
})

capital_metrics_df = pd.DataFrame({
    "Metric": [
        "Loan-to-Value", "Loan-to-Cost", "Equity Contribution", "Expense Ratio", "Implied Cap Rate",
        "Financing Fee Rate", "Financing Fees", "Appraisal Premium / Discount", "Balloon Balance"
    ],
    "Value": [
        format_pct(actual_ltv), format_pct(actual_ltc), format_pct(equity_pct), format_pct(expense_ratio),
        format_pct(implied_cap_rate), format_pct(inputs["fees"]), format_money(financing_fees),
        format_pct(appraisal_premium), format_money(balloon_balance)
    ]
})

rent_roll_summary_df = pd.DataFrame({
    "Metric": ["Total SF", "Vacant SF", "Annual Rent", "Average Rent PSF", "WALT", "12-Month Expiry", "Physical Occupancy"],
    "Value": [
        f"{rent_roll_metrics['Total_SF']:,.0f}", f"{rent_roll_metrics.get('Vacancy_SF', 0):,.0f}", format_money(rent_roll_metrics.get("Annual_Rent", 0)),
        f"${rent_roll_metrics.get('Average_Rent_PSF', 0):,.2f}", f"{rent_roll_metrics['WALT']:.2f} years",
        format_pct(rent_roll_metrics["Expiring_1yr_pct"]), format_pct(rent_roll_metrics["Occupancy"])
    ]
})

diligence_df = pd.DataFrame({"Item": list(st.session_state.diligence.keys()), "Status": list(st.session_state.diligence.values())})
risk_flags_df = pd.DataFrame({"Risk Flag": risk_flags}) if risk_flags else pd.DataFrame({"Risk Flag": ["No major automated risk flags."]})


# Scenarios
scenario_defs = {
    "Base": {"noi_mult": 1.00, "rate_delta": 0.000, "value_mult": 1.00},
    "Downside": {"noi_mult": 0.90, "rate_delta": 0.010, "value_mult": 0.95},
    "Severe Stress": {"noi_mult": 0.80, "rate_delta": 0.020, "value_mult": 0.90},
    "Upside": {"noi_mult": 1.10, "rate_delta": -0.005, "value_mult": 1.03},
}
scenario_rows = []
for name, sc in scenario_defs.items():
    sc_noi = inputs["noi"] * sc["noi_mult"]
    sc_rate = max(inputs["rate"] + sc["rate_delta"], 0.0025)
    sc_value = inputs["appraisal"] * sc["value_mult"]
    sc_loan, sc_gate, _ = size_loan(sc_noi, sc_value, total_uses, sc_rate, inputs["amort"], inputs["target_ltv"], inputs["target_ltc"], inputs["target_dscr"], inputs["target_dy"], inputs["debt_structure"])
    sc_pmt = calculate_monthly_payment(sc_loan, sc_rate, inputs["amort"], inputs["debt_structure"])
    sc_dscr = safe_divide(sc_noi, sc_pmt * 12)
    sc_ltv = safe_divide(sc_loan, sc_value)
    sc_ltc = safe_divide(sc_loan, total_uses)
    sc_dy = safe_divide(sc_noi, sc_loan)
    scenario_rows.append({
        "Scenario": name, "NOI": sc_noi, "Rate": sc_rate, "Value": sc_value,
        "Supportable Loan": sc_loan, "Binding": sc_gate, "LTV": sc_ltv,
        "LTC": sc_ltc, "DSCR": sc_dscr, "Debt Yield": sc_dy
    })
scenario_df = pd.DataFrame(scenario_rows)

# Sensitivity
noi_scenarios = [inputs["noi"] * x for x in [0.85, 0.90, 0.95, 1.00, 1.05, 1.10, 1.15]]
rate_scenarios = [max(inputs["rate"] + x, 0.0025) for x in [-0.015, -0.010, -0.005, 0.000, 0.005, 0.010, 0.015]]
matrix, gate_matrix = [], []
for sr in rate_scenarios:
    row, grow = [], []
    for sn in noi_scenarios:
        sl, sg, _ = size_loan(sn, inputs["appraisal"], total_uses, sr, inputs["amort"], inputs["target_ltv"], inputs["target_ltc"], inputs["target_dscr"], inputs["target_dy"], inputs["debt_structure"])
        row.append(sl)
        grow.append(sg)
    matrix.append(row)
    gate_matrix.append(grow)

col_names = ["NOI -15%", "NOI -10%", "NOI -5%", "Base NOI", "NOI +5%", "NOI +10%", "NOI +15%"]
row_names = [f"{r * 100:.2f}%" for r in rate_scenarios]
sensitivity_df = pd.DataFrame(matrix, index=row_names, columns=col_names)
sensitivity_gate_df = pd.DataFrame(gate_matrix, index=row_names, columns=col_names)


# ============================================================
# MAIN UI
# ============================================================

st.title("ALENZA CAPITAL")
st.subheader("Commercial Real Estate Underwriting Suite")
st.caption(f"{APP_VERSION} | Generated: {generated_at} | Transaction: {inputs['transaction_type']} | Active Constraint: {gate}")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Max Proceeds", format_money(loan_amt))
m2.metric("Actual LTV", format_pct(actual_ltv))
m3.metric("Actual LTC", format_pct(actual_ltc))
m4.metric("Actual DSCR", format_x(actual_dscr))
m5.metric("Debt Yield", format_pct(actual_dy))
m6.metric("Deal Score", f"{score}/1000")

st.markdown("---")

tabs = st.tabs([
    "Executive", "Platform", "Borrower Portal", "Lender Portal", "Document Room", "Parsing",
    "Intelligence", "Rent Roll", "Sizing", "Pipeline", "Lender Quotes", "Scenarios",
    "Sensitivity", "Amortization", "Portfolio", "Admin", "Capital Stack", "Covenants",
    "Assumptions", "Scorecard", "Audit", "Report"
])


with tabs[0]:
    left, right = st.columns([1.2, 1])
    with left:
        st.subheader("Executive Summary")
        st.dataframe(preview_df, hide_index=True, use_container_width=True)
        st.subheader("Risk Flags")
        st.dataframe(risk_flags_df, hide_index=True, use_container_width=True)
    with right:
        st.subheader("Underwriting Verdict")
        st.info(f"Supportable proceeds are constrained by {gate}.")
        st.write(constraint_advice(gate))
        if score >= 850:
            st.success(classification)
        elif score >= 725:
            st.info(classification)
        elif score >= 575:
            st.warning(classification)
        else:
            st.error(classification)
        e1, e2 = st.columns(2)
        e1.metric("WALT", f"{rent_roll_metrics['WALT']:.2f} yrs")
        e2.metric("Diligence", format_pct(diligence_pct))


with tabs[1]:
    st.subheader("Platform Control Center")
    st.caption("Database-backed workflow controls for the 950-target Alenza platform.")

    p1, p2, p3, p4 = st.columns(4)
    deals_df_platform = get_deals_df()
    docs_df_platform = get_documents_df()
    pm = portfolio_metrics(deals_df_platform)
    p1.metric("Saved Deals", f"{pm['Deals']:,.0f}")
    p2.metric("Documents", f"{len(docs_df_platform):,.0f}")
    p3.metric("Portfolio Loan Amt", format_money(pm["Total Loan Amount"]))
    p4.metric("Avg Score", f"{pm['Average Score']:.0f}")

    if st.button("Save Current Deal Snapshot to Database"):
        saved_id = upsert_current_deal_record(inputs, loan_amt, score, classification)
        st.success(f"Saved deal snapshot #{saved_id}.")

    st.write("Recent Deals")
    st.dataframe(deals_df_platform, hide_index=True, use_container_width=True)


with tabs[2]:
    st.subheader("Borrower Portal")
    st.caption("Borrower-facing upload and package completion workspace.")

    if not has_permission("view_borrower_portal"):
        st.warning("Your current role does not have borrower portal access.")
    else:
        bp1, bp2 = st.columns([1, 1])
        with bp1:
            st.write("Required Package")
            st.dataframe(diligence_df, hide_index=True, use_container_width=True)
            st.progress(diligence_pct, text=f"Package {diligence_pct * 100:.0f}% complete")

        with bp2:
            st.write("Borrower Upload")
            borrower_category = st.selectbox("Document Category", list(DILIGENCE_REQUIREMENTS.keys()), key="borrower_doc_category")
            borrower_files = st.file_uploader(
                "Upload borrower documents",
                accept_multiple_files=True,
                type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "csv", "docx", "txt"],
                key="borrower_uploads"
            )
            if st.button("Submit Borrower Documents"):
                saved = save_uploaded_documents(
                    borrower_files,
                    deal_id=None,
                    category=borrower_category,
                    uploaded_by=st.session_state.current_user["username"] if st.session_state.current_user else "borrower"
                )
                st.success(f"Submitted {len(saved)} document(s).")


with tabs[3]:
    st.subheader("Lender Portal")
    st.caption("Lender-facing deal summary, covenant view, and quote comparison.")

    if not has_permission("view_lender_portal"):
        st.warning("Your current role does not have lender portal access.")
    else:
        st.write("Lender Summary")
        st.dataframe(preview_df, hide_index=True, use_container_width=True)

        st.write("Covenant Package")
        st.dataframe(covenant_df, hide_index=True, use_container_width=True)

        st.write("Available Quotes")
        st.dataframe(
            lender_quotes_eval_df.style.format({
                "Loan Amount": "${:,.0f}",
                "Rate": "{:.2%}",
                "Fees": "{:.2%}",
                "Annual Debt Service": "${:,.0f}",
                "DSCR": "{:.2f}x",
                "LTV": "{:.1%}",
                "LTC": "{:.1%}",
                "Debt Yield": "{:.1%}",
                "Fee Dollars": "${:,.0f}"
            }) if not lender_quotes_eval_df.empty else lender_quotes_eval_df,
            hide_index=True,
            use_container_width=True
        )


with tabs[4]:
    st.subheader("Integrated Document Room")
    st.caption("Controlled document repository with SQLite metadata and local file storage.")

    if not has_permission("manage_documents") and not has_permission("upload_documents"):
        st.warning("Your current role does not have document-room access.")
    else:
        doc_col_1, doc_col_2 = st.columns([1, 1])
        with doc_col_1:
            document_category = st.selectbox("Category", list(DILIGENCE_REQUIREMENTS.keys()) + ["Other"])
            document_files = st.file_uploader(
                "Upload deal documents",
                accept_multiple_files=True,
                type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "csv", "docx", "txt"],
                key="document_room_uploads"
            )
            if st.button("Save to Document Room"):
                saved = save_uploaded_documents(
                    document_files,
                    deal_id=None,
                    category=document_category,
                    uploaded_by=st.session_state.current_user["username"] if st.session_state.current_user else "anonymous"
                )
                st.success(f"Saved {len(saved)} file(s).")

        with doc_col_2:
            st.write("Diligence Gap Audit")
            if document_files:
                found, missing, audit_df_room = audit_deal_room(document_files)
                st.dataframe(audit_df_room, hide_index=True, use_container_width=True)
                if missing:
                    st.error(f"Missing: {', '.join(missing)}")
                else:
                    st.success("No major diligence gaps detected in this upload batch.")

        st.write("Document Repository")
        st.dataframe(get_documents_df(), hide_index=True, use_container_width=True)


with tabs[5]:
    st.subheader("AI Extraction & Structured Parsing")
    st.caption("Source-backed extraction scaffold for T12 and rent-roll tables.")

    if st.session_state.get("extracted_fields"):
        st.write("OCR Field Evidence")
        st.dataframe(
            build_extraction_evidence_df(st.session_state.get("raw_ocr_text") or "", st.session_state.get("extracted_fields") or {}),
            hide_index=True,
            use_container_width=True
        )

    parse_type = st.selectbox("Parser", ["Rent Roll Table", "T12 / Operating Statement"])
    parse_file = st.file_uploader("Upload CSV/XLS/XLSX for structured parsing", type=["csv", "xlsx", "xls"], key="structured_parser_upload")

    if parse_file is not None:
        try:
            parsed_source_df = parse_uploaded_table(parse_file)
            st.write("Raw Uploaded Table")
            st.dataframe(parsed_source_df, hide_index=True, use_container_width=True)

            if parse_type == "Rent Roll Table":
                normalized_rr = normalize_rent_roll_table(parsed_source_df)
                st.write("Normalized Rent Roll")
                st.dataframe(normalized_rr, hide_index=True, use_container_width=True)

                rr_metrics_uploaded = calculate_rent_roll_metrics(normalized_rr)
                c1, c2, c3 = st.columns(3)
                c1.metric("WALT", f"{rr_metrics_uploaded['WALT']:.2f} yrs")
                c2.metric("Occupancy", format_pct(rr_metrics_uploaded["Occupancy"]))
                c3.metric("12-Mo Rollover", format_pct(rr_metrics_uploaded["Expiring_1yr_pct"]))

                if st.button("Apply Parsed Rent Roll"):
                    st.session_state.rent_roll_df = normalized_rr
                    platform_audit("Parsed Rent Roll Applied", parse_file.name)
                    st.success("Parsed rent roll applied to Rent Roll tab.")

            else:
                normalized_t12, t12_summary = normalize_t12_table(parsed_source_df)
                st.write("Normalized T12 Lines")
                st.dataframe(normalized_t12, hide_index=True, use_container_width=True)

                t1, t2, t3 = st.columns(3)
                t1.metric("Income", format_money(t12_summary["Income"]))
                t2.metric("Expenses", format_money(t12_summary["Expenses"]))
                t3.metric("NOI", format_money(t12_summary["NOI"]))

                if st.button("Apply Parsed T12 NOI"):
                    if t12_summary["Income"] > 0:
                        st.session_state.gross_income = int(t12_summary["Income"])
                    if t12_summary["Expenses"] > 0:
                        st.session_state.operating_expenses = int(t12_summary["Expenses"])
                    if t12_summary["NOI"] > 0:
                        st.session_state.noi = int(t12_summary["NOI"])
                    platform_audit("Parsed T12 Applied", parse_file.name)
                    st.success("Parsed T12 values applied to underwriting inputs.")

        except Exception as exc:
            st.error(f"Could not parse file: {exc}")


with tabs[6]:
    st.subheader("Deal Concierge & Intelligence")
    st.caption("Bulk diligence audit, property intelligence, and satellite context for faster borrower/lender readiness.")

    col_int_1, col_int_2 = st.columns([1.5, 1])

    with col_int_1:
        addr = st.text_input("Property Address", st.session_state.property_address, key="intel_property_address")
        st.session_state.property_address = addr

        if addr:
            maps_url = build_google_maps_url(addr)
            embed_url = build_google_maps_embed_url(addr)

            st.markdown(f"### [Satellite View for {addr}]({maps_url})")
            components.iframe(embed_url, height=320)

    with col_int_2:
        if addr:
            intel = get_property_intelligence(addr)
            st.write("**Property Intelligence (Auto-Fetch Ready)**")
            st.dataframe(pd.DataFrame(intel.items(), columns=["Fact", "Value"]), hide_index=True, use_container_width=True)

            st.write("**Live API Readiness**")
            st.dataframe(property_api_status_df(), hide_index=True, use_container_width=True)

    st.divider()

    st.write("**Document Audit Room**")
    all_files = st.file_uploader(
        "Drop all deal documents here",
        accept_multiple_files=True,
        type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "csv", "docx", "txt"],
        key="concierge_deal_room"
    )

    if all_files:
        found, missing, deal_room_audit_df = audit_deal_room(all_files)
        add_audit("Deal Room Audit", f"{len(all_files)} files uploaded")

        c1, c2 = st.columns(2)
        c1.success(f"Received: {', '.join(found) if found else 'No required items detected'}")

        if missing:
            c2.error(f"Missing: {', '.join(missing)}")
            st.warning("Action Required: lenders cannot price this deal accurately without the missing items.")

            with st.expander("Need help collecting these or want a pro review?"):
                with st.form("concierge_contact"):
                    u_name = st.text_input("Name")
                    u_email = st.text_input("Email")
                    u_notes = st.text_area("Notes", "I need help completing this deal package.")
                    submitted = st.form_submit_button("Request Alenza Concierge Support")

                    if submitted:
                        request = {
                            "Timestamp": now_str(),
                            "Name": u_name,
                            "Email": u_email,
                            "Address": addr,
                            "Missing Items": missing,
                            "Notes": u_notes
                        }
                        st.session_state.concierge_requests.append(request)
                        add_audit("Concierge Request", f"{u_name} | {u_email}")
                        st.balloons()
                        st.success("Request recorded. Export the deal JSON or Excel workbook to retain this lead record.")
        else:
            c2.success("No major diligence gaps detected.")

        st.dataframe(deal_room_audit_df, hide_index=True, use_container_width=True)

    if st.session_state.concierge_requests:
        st.write("**Concierge Requests Captured This Session**")
        st.dataframe(pd.DataFrame(st.session_state.concierge_requests), hide_index=True, use_container_width=True)


with tabs[7]:
    st.subheader("Rent Roll Analyzer")
    st.caption("Institutional Lease Exposure & WALT Calculation")
    edited_rr = st.data_editor(st.session_state.rent_roll_df, num_rows="dynamic", use_container_width=True, key="rent_roll_editor")
    st.session_state.rent_roll_df = edited_rr
    rr_m = calculate_rent_roll_metrics(edited_rr)

    rc1, rc2, rr3, rr4 = st.columns(4)
    rc1.metric("WALT", f"{rr_m['WALT']:.2f} Yrs")
    rc2.metric("Occupancy", f"{rr_m['Occupancy']*100:.1f}%")
    rr3.metric("Rollover (12mo)", f"{rr_m['Expiring_1yr_pct']*100:.1f}%", delta="Risk Signal", delta_color="inverse")
    rr4.metric("Avg Rent PSF", f"${rr_m['Average_Rent_PSF']:,.2f}")
    st.dataframe(pd.DataFrame({
        "Metric": ["Total SF", "Vacant SF", "Annual Rent", "Average Rent PSF"],
        "Value": [f"{rr_m['Total_SF']:,.0f}", f"{rr_m.get('Vacancy_SF', 0):,.0f}", format_money(rr_m["Annual_Rent"]), f"${rr_m['Average_Rent_PSF']:,.2f}"]
    }), hide_index=True, use_container_width=True)


with tabs[8]:
    left, right = st.columns([1.5, 1])
    with left:
        st.subheader("Loan Sizing Constraints")
        st.dataframe(sizing_df.style.format({"Max Proceeds": "${:,.0f}", "Proceeds Gap": "${:,.0f}"}), hide_index=True, use_container_width=True)
        chart_df = pd.DataFrame({"Constraint": list(gates.keys()), "Max Proceeds": list(gates.values())})
        st.bar_chart(chart_df, x="Constraint", y="Max Proceeds", color="#1D4ED8")
    with right:
        st.subheader("Cash Flow / Debt Service")
        debt_df = pd.DataFrame({
            "Metric": ["Stabilized NOI", "Monthly Payment", "Annual Debt Service", "DSCR Cushion", "Loan Term", "Balloon Balance"],
            "Value": [format_money(inputs["noi"]), format_money(monthly_payment), format_money(annual_debt_service), format_x(debt_service_cushion), f"{inputs['loan_term']} years", format_money(balloon_balance)]
        })
        st.dataframe(debt_df, hide_index=True, use_container_width=True)


with tabs[9]:
    st.subheader("Deal Diligence & Pipeline")
    st.caption("Workflow integration to compete with nCino/Rockport")
    init_diligence_tracker()
    for item, status in list(st.session_state.diligence.items()):
        col_name, col_status = st.columns([3, 1])
        col_name.write(f"**{item}**")
        new_status = col_status.selectbox(
            "Status",
            ["Pending", "Received", "Reviewed", "Waived"],
            index=["Pending", "Received", "Reviewed", "Waived"].index(status) if status in ["Pending", "Received", "Reviewed", "Waived"] else 0,
            key=f"check_{item}"
        )
        st.session_state.diligence[item] = new_status

    st.divider()
    diligence_pct = diligence_progress()
    st.write("Deal Velocity Signal")
    st.progress(diligence_pct, text=f"Diligence {diligence_pct*100:.0f}% Complete")
    diligence_df = pd.DataFrame({"Item": list(st.session_state.diligence.keys()), "Status": list(st.session_state.diligence.values())})
    st.dataframe(diligence_df, hide_index=True, use_container_width=True)


with tabs[10]:
    st.subheader("Lender Quote Comparison")
    st.caption("Side-by-side term sheet analysis for borrower/lender strategy.")
    edited_quotes_df = st.data_editor(st.session_state.lender_quotes_df, num_rows="dynamic", use_container_width=True, key="lender_quote_editor")
    st.session_state.lender_quotes_df = edited_quotes_df
    lender_quotes_eval_df = evaluate_lender_quotes(edited_quotes_df, inputs["noi"], inputs["appraisal"], total_uses)

    if lender_quotes_eval_df.empty:
        st.info("Add lender quotes to compare proceeds, rate, DSCR, LTV, LTC, fees, and quote score.")
    else:
        st.dataframe(
            lender_quotes_eval_df.style.format({
                "Loan Amount": "${:,.0f}", "Rate": "{:.2%}", "Fees": "{:.2%}",
                "Annual Debt Service": "${:,.0f}", "DSCR": "{:.2f}x", "LTV": "{:.1%}",
                "LTC": "{:.1%}", "Debt Yield": "{:.1%}", "Fee Dollars": "${:,.0f}"
            }),
            hide_index=True,
            use_container_width=True
        )
        best_quote = lender_quotes_eval_df.iloc[0]
        st.success(f"Best current quote by weighted score: {best_quote['Lender']} | Score {best_quote['Quote Score']}/100")


with tabs[11]:
    st.subheader("Scenario Manager")
    st.dataframe(
        scenario_df.style.format({
            "NOI": "${:,.0f}", "Rate": "{:.2%}", "Value": "${:,.0f}", "Supportable Loan": "${:,.0f}",
            "LTV": "{:.1%}", "LTC": "{:.1%}", "DSCR": "{:.2f}x", "Debt Yield": "{:.1%}"
        }),
        hide_index=True,
        use_container_width=True
    )


with tabs[12]:
    st.subheader("Sensitivity Analysis")
    st.write("Supportable Proceeds")
    st.dataframe(sensitivity_df.style.format("${:,.0f}"), use_container_width=True)
    st.write("Binding Constraint")
    st.dataframe(sensitivity_gate_df, use_container_width=True)


with tabs[13]:
    st.subheader("Full Amortization Schedule")
    st.caption(f"Term: {inputs['loan_term']} Years | Structure: {inputs['debt_structure']}")
    st.dataframe(
        amort_schedule_df.style.format({
            "Opening Balance": "${:,.2f}", "Total Payment": "${:,.2f}",
            "Principal": "${:,.2f}", "Interest": "${:,.2f}", "Closing Balance": "${:,.2f}"
        }),
        hide_index=True,
        use_container_width=True
    )
    total_principal = amort_schedule_df["Principal"].sum() if not amort_schedule_df.empty else 0
    total_interest = amort_schedule_df["Interest"].sum() if not amort_schedule_df.empty else 0
    ending_balance = amort_schedule_df["Closing Balance"].iloc[-1] if not amort_schedule_df.empty else loan_amt
    a1, a2, a3 = st.columns(3)
    a1.metric("Principal Paid During Term", format_money(total_principal))
    a2.metric("Interest Paid During Term", format_money(total_interest))
    a3.metric("Ending / Balloon Balance", format_money(ending_balance))
    if inputs["debt_structure"] != "Interest-Only" and not amort_schedule_df.empty:
        st.write("Cumulative Capital Deployment")
        st.area_chart(amort_schedule_df[["Principal", "Interest"]].cumsum(), color=["#1D4ED8", "#94A3B8"])


with tabs[14]:
    st.subheader("Portfolio Monitoring")
    deals_df_portfolio = get_deals_df()
    metrics = portfolio_metrics(deals_df_portfolio)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Deals", f"{metrics['Deals']:,.0f}")
    c2.metric("Total Loan Amount", format_money(metrics["Total Loan Amount"]))
    c3.metric("Average Score", f"{metrics['Average Score']:.0f}")
    c4.metric("Average NOI", format_money(metrics["Average NOI"]))

    if not deals_df_portfolio.empty:
        st.dataframe(deals_df_portfolio, hide_index=True, use_container_width=True)
        st.write("Loan Amount by Deal")
        chart_df = deals_df_portfolio[["deal_name", "loan_amount"]].copy()
        st.bar_chart(chart_df, x="deal_name", y="loan_amount")
    else:
        st.info("Save deal snapshots from the Platform tab to populate portfolio monitoring.")


with tabs[15]:
    st.subheader("Admin Dashboard")
    st.caption("Security posture, users, APIs, platform readiness, and audit evidence.")

    if not has_permission("view_admin"):
        st.warning("Your current role does not have admin access.")
    else:
        api_keys_df_admin = api_key_status_df()
        readiness_score, readiness_df = platform_readiness_score(
            get_deals_df(),
            get_documents_df(),
            diligence_progress(),
            api_keys_df_admin,
            build_extraction_evidence_df(st.session_state.get("raw_ocr_text") or "", st.session_state.get("extracted_fields") or {})
        )

        a0, a1, a2 = st.columns(3)
        a0.metric("Platform Readiness", f"{readiness_score}/100")
        a1.metric("Configured APIs", f"{(api_keys_df_admin['Configured'] == 'Yes').sum()}/{len(api_keys_df_admin)}")
        a2.metric("Audit Events", f"{len(get_audit_df()):,.0f}")

        st.write("950-Target Readiness Breakdown")
        st.dataframe(readiness_df, hide_index=True, use_container_width=True)

        admin_1, admin_2 = st.columns(2)

        with admin_1:
            st.write("SOC2-Style Security Controls")
            controls_df = pd.DataFrame({
                "Control": SOC2_SECURITY_CONTROLS,
                "Status": ["Needs evidence" for _ in SOC2_SECURITY_CONTROLS]
            })
            st.dataframe(controls_df, hide_index=True, use_container_width=True)

        with admin_2:
            st.write("API Readiness")
            st.dataframe(property_api_status_df(), hide_index=True, use_container_width=True)
            st.write("API Key Status")
            st.dataframe(api_keys_df_admin, hide_index=True, use_container_width=True)

        st.write("Recent Platform Audit Events")
        st.dataframe(get_audit_df(), hide_index=True, use_container_width=True)


with tabs[16]:
    st.subheader("Sources and Uses")
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(uses_df.style.format({"Amount": "${:,.0f}"}), hide_index=True, use_container_width=True)
    with c2:
        st.dataframe(sources_df.style.format({"Amount": "${:,.0f}"}), hide_index=True, use_container_width=True)
    st.subheader("Capital Stack Metrics")
    st.dataframe(capital_metrics_df, hide_index=True, use_container_width=True)


with tabs[17]:
    st.subheader("Covenant Compliance")
    st.dataframe(covenant_df, hide_index=True, use_container_width=True)
    if "FAIL" in covenant_df["Status"].values:
        st.error("One or more covenants fail under current assumptions.")
    else:
        st.success("All automated covenant tests pass.")


with tabs[18]:
    st.subheader("Underwriting Assumptions")
    st.dataframe(assumptions_df, hide_index=True, use_container_width=True)


with tabs[19]:
    st.subheader("Alenza Deal Score")
    l, r = st.columns([1, 2])
    with l:
        st.metric("Score", f"{score}/1000")
        st.write(f"**Classification:** {classification}")
    with r:
        st.dataframe(score_df, hide_index=True, use_container_width=True)
    st.caption("Score is indicative only and does not replace lender diligence, sponsor review, environmental diligence, legal review, appraisal review, or final credit approval.")


with tabs[20]:
    st.subheader("Audit Trail")
    audit_df = pd.DataFrame(st.session_state.audit_log) if st.session_state.audit_log else pd.DataFrame({"Timestamp": [], "Action": [], "Detail": []})
    st.dataframe(audit_df, hide_index=True, use_container_width=True)
    st.write("Database Audit Events")
    st.dataframe(get_audit_df(), hide_index=True, use_container_width=True)


with tabs[21]:
    st.subheader("Executive Summary Preview")
    st.dataframe(preview_df, hide_index=True, use_container_width=True)

    raw_ocr_text = st.session_state.get("raw_ocr_text")
    extracted_fields = st.session_state.get("extracted_fields")
    offline_ai_result = st.session_state.get("offline_ai_result")
    ollama_review_text = st.session_state.get("ollama_review_text")

    extracted_review_df = None
    if extracted_fields:
        extracted_review_df = pd.DataFrame({
            "Field": list(extracted_fields.keys()),
            "Extracted Value": ["" if v is None else (format_pct(v) if k == "Cap Rate" else format_money(v)) for k, v in extracted_fields.items()]
        })

    offline_ai_df = offline_ai_result.get("confidence_df") if offline_ai_result else None
    offline_warnings = offline_ai_result.get("warnings", []) if offline_ai_result else []
    offline_suggestions = offline_ai_result.get("suggestions", []) if offline_ai_result else []
    offline_blockers = offline_ai_result.get("blockers", []) if offline_ai_result else []
    extraction_evidence_df = build_extraction_evidence_df(raw_ocr_text or "", extracted_fields or {})

    report_text = f"""ALENZA CAPITAL UNDERWRITING SUMMARY
Generated: {generated_at}
Application Version: {APP_VERSION}

EXECUTIVE SUMMARY
Sponsor / Borrower: {inputs['sponsor']}
Property Address: {inputs['property_address']}
Property Type: {inputs['property_type']}
Transaction Type: {inputs['transaction_type']}
Credit Profile: {inputs['credit_profile']}
Supportable Proceeds: {format_money(loan_amt)}
Binding Constraint: {gate}
Classification: {classification}
Deal Score: {score}/1000

LEASE / RENT ROLL METRICS
Portfolio WALT: {rent_roll_metrics['WALT']:.2f} years
Physical Occupancy: {format_pct(rent_roll_metrics['Occupancy'])}
12-Month Expiry Concentration: {format_pct(rent_roll_metrics['Expiring_1yr_pct'])}
Annual Rent: {format_money(rent_roll_metrics['Annual_Rent'])}
Average Rent PSF: ${rent_roll_metrics['Average_Rent_PSF']:,.2f}

PIPELINE / DILIGENCE
Diligence Progress: {format_pct(diligence_pct)}

ASSET PROFILE
Purchase Price / Cost Basis: {format_money(inputs['purchase_price'])}
Appraised Value: {format_money(inputs['appraisal'])}
Existing Debt Payoff: {format_money(inputs['existing_debt'])}
Gross Income / EGI: {format_money(inputs['gross_income'])}
Vacancy / Credit Loss: {format_money(inputs['vacancy_loss'])}
Operating Expenses: {format_money(inputs['operating_expenses'])}
Stabilized NOI: {format_money(inputs['noi'])}
Expense Ratio: {format_pct(expense_ratio)}
Implied Cap Rate: {format_pct(implied_cap_rate)}

LOAN METRICS
Actual LTV: {format_pct(actual_ltv)}
Actual LTC: {format_pct(actual_ltc)}
Actual DSCR: {format_x(actual_dscr)}
Debt Yield: {format_pct(actual_dy)}
Monthly Payment: {format_money(monthly_payment)}
Annual Debt Service: {format_money(annual_debt_service)}
Balloon Balance: {format_money(balloon_balance)}

CAPITAL STACK
Total Uses: {format_money(total_uses)}
Supportable Senior Debt: {format_money(loan_amt)}
Required Sponsor Equity: {format_money(required_equity)}
Equity Contribution: {format_pct(equity_pct)}

SIZING CONSTRAINTS
LTV Limit: {format_money(gates['LTV'])}
LTC Limit: {format_money(gates['LTC'])}
DSCR Limit: {format_money(gates['DSCR'])}
Debt Yield Limit: {format_money(gates['Debt Yield'])}
Binding Constraint: {gate}

RISK FLAGS
{chr(10).join('- ' + f for f in risk_flags) if risk_flags else '- No major automated risk flags.'}

UNDERWRITING VERDICT
{constraint_advice(gate)}

OFFLINE AI REVIEW
Blockers:
{chr(10).join('- ' + b for b in offline_blockers) if offline_blockers else '- None'}
Warnings:
{chr(10).join('- ' + w for w in offline_warnings) if offline_warnings else '- None'}
Suggestions:
{chr(10).join('- ' + s for s in offline_suggestions) if offline_suggestions else '- None'}

LOCAL OLLAMA REVIEW
{ollama_review_text if ollama_review_text else 'No Ollama review generated.'}

DISCLAIMER
This summary is indicative only and is not a loan commitment, credit approval, investment advice, appraisal, legal opinion, or final underwriting decision. All terms are subject to lender diligence, borrower review, third-party reports, credit approval, committee review, and final documentation.
"""

    metadata = {
        "Sponsor / Borrower": inputs["sponsor"],
        "Property Address": inputs["property_address"],
        "Property Type": inputs["property_type"],
        "Transaction Type": inputs["transaction_type"],
        "Supportable Proceeds": format_money(loan_amt),
        "Binding Constraint": gate,
        "Deal Score": f"{score}/1000",
        "Classification": classification,
        "Generated": generated_at,
        "Version": APP_VERSION
    }

    sheets = {
        "Executive Summary": preview_df,
        "Assumptions": assumptions_df,
        "Property Intelligence": pd.DataFrame(get_property_intelligence(inputs["property_address"]).items(), columns=["Fact", "Value"]),
        "API Readiness": property_api_status_df(),
        "API Key Status": api_key_status_df(),
        "Extraction Evidence": extraction_evidence_df,
        "Concierge Requests": pd.DataFrame(st.session_state.concierge_requests),
        "Deals Database": get_deals_df(),
        "Documents Database": get_documents_df(),
        "Rent Roll": st.session_state.rent_roll_df,
        "Rent Roll Summary": rent_roll_summary_df,
        "Diligence Tracker": diligence_df,
        "Lender Quotes": lender_quotes_eval_df,
        "Sizing": sizing_df,
        "Amortization Schedule": amort_schedule_df,
        "Uses": uses_df,
        "Sources": sources_df,
        "Capital Metrics": capital_metrics_df,
        "Covenants": covenant_df,
        "Scorecard": score_df,
        "Risk Flags": risk_flags_df,
        "Scenarios": scenario_df,
        "Sensitivity Proceeds": sensitivity_df,
        "Sensitivity Gates": sensitivity_gate_df,
        "Audit Trail": pd.DataFrame(st.session_state.audit_log),
        "Platform Audit": get_audit_df()
    }

    if extracted_review_df is not None:
        sheets["OCR Extract"] = extracted_review_df
    if offline_ai_df is not None:
        sheets["Offline AI Confidence"] = offline_ai_df

    offline_notes = {"Blockers": offline_blockers, "Warnings": offline_warnings, "Suggestions": offline_suggestions}

    safe_sponsor = clean_filename(inputs["sponsor"])
    deal_json = build_deal_json({
        "version": APP_VERSION,
        "generated": generated_at,
        "inputs": inputs,
        "score": score,
        "classification": classification,
        "rent_roll": st.session_state.rent_roll_df.to_dict(orient="records"),
        "diligence": st.session_state.diligence,
        "lender_quotes": st.session_state.lender_quotes_df.to_dict(orient="records"),
        "property_intelligence": get_property_intelligence(inputs["property_address"]),
        "concierge_requests": st.session_state.concierge_requests
    })
    excel_file = create_excel_workbook(sheets, metadata, raw_ocr_text, offline_notes, ollama_review_text)
    pdf_file = create_pdf_summary(metadata, preview_df, covenant_df, score_df, risk_flags, offline_warnings + offline_blockers, ollama_review_text)
    backup_zip = create_platform_backup_zip(deal_json, excel_file, report_text)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.download_button("Download Text", report_text, file_name=f"Alenza_Summary_{safe_sponsor}.txt", mime="text/plain")
    with c2:
        st.download_button("Download Excel", excel_file, file_name=f"Alenza_Workbook_{safe_sponsor}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        if pdf_file:
            st.download_button("Download PDF", pdf_file, file_name=f"Alenza_Summary_{safe_sponsor}.pdf", mime="application/pdf")
        else:
            st.warning("PDF export requires reportlab.")
    with c4:
        st.download_button("Save Deal JSON", deal_json, file_name=f"Alenza_Deal_{safe_sponsor}.json", mime="application/json")
    with c5:
        st.download_button("Backup ZIP", backup_zip, file_name=f"Alenza_Backup_{safe_sponsor}.zip", mime="application/zip")

    with st.expander("Deployment Notes", expanded=False):
        st.code(
            """requirements.txt:
streamlit
pandas
numpy
openpyxl
xlsxwriter
reportlab
pillow
pytesseract
pymupdf

packages.txt:
tesseract-ocr

Demo logins:
admin / alenza-admin
broker / alenza-broker
borrower / alenza-borrower
lender / alenza-lender

Optional local Ollama:
ollama serve
ollama pull llama3.2:3b
streamlit run app.py

For Streamlit Cloud, localhost will not reach the user's machine.
Use a reachable Ollama endpoint URL if deployed remotely.
""",
            language="text"
        )
