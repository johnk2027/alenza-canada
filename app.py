import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import io
import re
import json
import sqlite3
import zipfile
import urllib.request
from urllib.parse import quote_plus, urlencode
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple

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
# ALENZA CAPITAL OS
# Sovereign Canada Build
# Local-first Canadian CRE underwriting workstation
# ============================================================

APP_VERSION = "Sovereign Canada Build 5.0 | Zero-Account"

BASE_DIR = Path("alenza_data")
DOC_DIR = BASE_DIR / "documents"
BACKUP_DIR = BASE_DIR / "backups"
EXPORT_DIR = BASE_DIR / "exports"
DB_PATH = BASE_DIR / "alenza_platform.db"

DEFAULT_DEAL_INBOX_EMAIL = "resourcefulcapital@gmail.com"

BOC_VALET_BASE = "https://www.bankofcanada.ca/valet"
BOC_OVERNIGHT_TARGET_SERIES = "STATIC_ATABLE_V39079"
BOC_OVERNIGHT_RATE_SERIES = "V122514"
BOC_TBILL_3M_SERIES = "V122512"
BOC_PRIME_SERIES = "V122530"

CORPORATIONS_CANADA_BASE = "https://www.ic.gc.ca/app/scr/cc/CorporationsCanada/api/corporations"


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Alenza Capital OS",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
:root {
    --bg-main: #05080F;
    --bg-panel: #0B1220;
    --border: #1E293B;
    --text-muted: #94A3B8;
    --accent: #1D4ED8;
    --accent-soft: #2563EB;
}

.main {
    background-color: var(--bg-main);
    font-family: "Helvetica Neue", Arial, sans-serif;
}

section[data-testid="stSidebar"] {
    background-color: var(--bg-panel) !important;
    border-right: 1px solid var(--border);
}

h1, h2, h3 {
    letter-spacing: -0.025em;
}

[data-testid="stMetricValue"] {
    font-size: 26px !important;
    font-weight: 800 !important;
    color: var(--accent-soft) !important;
}

[data-testid="stMetricLabel"] {
    font-size: 11px !important;
    text-transform: uppercase;
    letter-spacing: 1.3px;
    color: var(--text-muted) !important;
}

div[data-testid="stMetric"] {
    background-color: var(--bg-panel);
    padding: 16px;
    border-radius: 8px;
    border: 1px solid var(--border);
}

.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    border-bottom: 1px solid var(--border);
}

.stTabs [data-baseweb="tab"] {
    background-color: var(--bg-panel);
    border: 1px solid var(--border);
    border-bottom: none;
    border-radius: 5px 5px 0 0;
    padding: 9px 14px;
    color: var(--text-muted);
    font-weight: 700;
    letter-spacing: .04em;
    text-transform: uppercase;
    font-size: 10px;
}

.stTabs [aria-selected="true"] {
    background-color: var(--accent) !important;
    color: #fff !important;
    border-color: var(--accent) !important;
}

.stDownloadButton>button, .stButton>button {
    width: 100%;
    background-color: var(--accent);
    color: white;
    font-weight: 700;
    border-radius: 6px;
    border: none;
    padding: 12px;
    text-transform: uppercase;
    letter-spacing: .04em;
}

.stDownloadButton>button:hover, .stButton>button:hover {
    background-color: var(--accent-soft);
    color: white;
}

div[data-testid="stExpander"] {
    border: 1px solid var(--border);
    border-radius: 6px;
    background-color: var(--bg-panel);
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# BASIC HELPERS
# ============================================================

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def ensure_dirs() -> None:
    BASE_DIR.mkdir(exist_ok=True)
    DOC_DIR.mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)
    EXPORT_DIR.mkdir(exist_ok=True)


def get_secret_value(key: str) -> Optional[str]:
    try:
        if key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return None


def get_deal_inbox_email() -> str:
    configured = get_secret_value("ALENZA_DEAL_INBOX_EMAIL")
    return configured or DEFAULT_DEAL_INBOX_EMAIL


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
    value = str(value or "Client").strip()
    value = value.replace(" ", "_").replace("/", "_").replace("\\", "_")
    value = re.sub(r"[^A-Za-z0-9_\\-]", "", value)
    return value or "Client"


def money_to_float(value) -> Optional[float]:
    if value is None:
        return None
    cleaned = (
        str(value)
        .replace("$", "")
        .replace(",", "")
        .replace("(", "-")
        .replace(")", "")
        .strip()
    )
    try:
        return float(cleaned)
    except Exception:
        return None


def http_get_json(url: str, timeout: int = 10) -> Any:
    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "AlenzaCanadaPlatform/1.0",
                "Accept": "application/json",
            },
            method="GET",
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
        return json.loads(raw)
    except Exception as exc:
        return {"_error": str(exc)}


def build_deal_email_link(subject: str, body: str) -> str:
    to_email = get_deal_inbox_email()
    return f"mailto:{to_email}?subject={quote_plus(subject)}&body={quote_plus(body)}"


def build_concierge_email_body(name: str, email: str, address: str, missing: List[str]) -> str:
    return f"""ALENZA CAPITAL CONCIERGE REQUEST

Sponsor / Contact: {name or "Not provided"}
Email: {email or "Not provided"}
Property: {address or "Not provided"}
Missing Diligence: {", ".join(missing) if missing else "None"}

Generated: {now_str()}

Generated via Alenza Canada Platform.
"""


# ============================================================
# LOCAL DATABASE
# ============================================================

def init_db() -> None:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS deals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sponsor TEXT,
            property_address TEXT,
            property_type TEXT,
            transaction_type TEXT,
            loan_amount REAL,
            noi REAL,
            appraised_value REAL,
            score INTEGER,
            classification TEXT,
            created_at TEXT
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            category TEXT,
            file_path TEXT,
            uploaded_at TEXT
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT,
            detail TEXT,
            timestamp TEXT
        )
        """
    )

    conn.commit()
    conn.close()


def add_audit(action: str, detail: str = "") -> None:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO audit_events (action, detail, timestamp) VALUES (?, ?, ?)",
        (action, detail, now_str()),
    )
    conn.commit()
    conn.close()


def save_deal_snapshot(inputs: Dict[str, Any], loan_amt: float, score: int, classification: str) -> int:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO deals (
            sponsor, property_address, property_type, transaction_type,
            loan_amount, noi, appraised_value, score, classification, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            inputs["sponsor"],
            inputs["property_address"],
            inputs["property_type"],
            inputs["transaction_type"],
            float(loan_amt),
            float(inputs["noi"]),
            float(inputs["appraisal"]),
            int(score),
            classification,
            now_str(),
        ),
    )
    deal_id = cur.lastrowid
    conn.commit()
    conn.close()
    add_audit("Deal Snapshot Saved", f"Deal ID {deal_id}")
    return deal_id


def get_deals_df() -> pd.DataFrame:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM deals ORDER BY id DESC", conn)
    conn.close()
    return df


def get_audit_df() -> pd.DataFrame:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM audit_events ORDER BY id DESC LIMIT 500", conn)
    conn.close()
    return df


def save_uploaded_documents(files, category: str = "Deal Document") -> List[str]:
    if not files:
        return []

    init_db()
    saved = []
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for file in files:
        safe_name = clean_filename(file.name)
        stored_name = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{safe_name}"
        dest = DOC_DIR / stored_name
        dest.write_bytes(file.getbuffer())

        cur.execute(
            """
            INSERT INTO documents (filename, category, file_path, uploaded_at)
            VALUES (?, ?, ?, ?)
            """,
            (file.name, category, str(dest), now_str()),
        )
        saved.append(file.name)

    conn.commit()
    conn.close()
    add_audit("Documents Saved", f"{len(saved)} file(s)")
    return saved


# ============================================================
# CANADIAN OPEN DATA CONNECTORS
# ============================================================

def nrcan_standardize_address(query: str) -> List[Dict[str, Any]]:
    query = str(query or "").strip()
    if len(query) < 5:
        return []

    params = {
        "q": query,
        "lang": "en",
        "keys": "geonames,nominatim,locate,fsa",
    }
    url = f"https://geolocator.api.geo.ca/?{urlencode(params)}"
    data = http_get_json(url, timeout=8)

    if isinstance(data, list):
        return data

    return []


def canadian_property_fallback(address: str) -> Dict[str, Any]:
    return {
        "Address": address or "Not provided",
        "Country": "Canada",
        "Province": "Manual entry / geocoder result",
        "Municipality": "Manual entry required",
        "Lot Size": "Manual entry required",
        "Year Built": "Manual entry required",
        "Last Sale": "Manual entry required",
        "Assessment / Roll Number": "Manual entry required",
        "Zoning": "Manual entry required",
        "Flood / Hazard": "Check NRCan / municipal flood data",
        "Environmental": "Check Phase I ESA / provincial records",
        "Data Mode": "Sovereign fallback / manual",
    }


def get_property_intelligence(address: str) -> Dict[str, Any]:
    fallback = canadian_property_fallback(address)
    results = nrcan_standardize_address(address)

    if results:
        best = results[0]
        fallback["Geolocation Source"] = "NRCan / Geo.ca"
        fallback["Matched Location"] = best.get("name", "")
        fallback["Province"] = best.get("province", "")
        fallback["Location Type"] = best.get("category", "")
        fallback["Latitude"] = best.get("lat", "")
        fallback["Longitude"] = best.get("lng", "")
        fallback["Result Count"] = len(results)
    else:
        fallback["Geolocation Source"] = "Manual / fallback"
        fallback["Result Count"] = 0

    return fallback


def boc_valet_latest(series: str) -> Dict[str, Any]:
    url = f"{BOC_VALET_BASE}/observations/{quote_plus(series)}/json?recent=10"
    data = http_get_json(url, timeout=10)

    if isinstance(data, dict) and data.get("_error"):
        return {
            "Status": "Error",
            "Series": series,
            "Value": None,
            "Date": "",
            "Endpoint": url,
            "Error": data.get("_error"),
        }

    observations = data.get("observations", []) if isinstance(data, dict) else []

    for obs in reversed(observations):
        if series in obs and isinstance(obs[series], dict):
            value_text = obs[series].get("v")
            try:
                return {
                    "Status": "OK",
                    "Series": series,
                    "Value": float(value_text),
                    "Date": obs.get("d", ""),
                    "Endpoint": url,
                }
            except Exception:
                continue

    return {
        "Status": "No Value",
        "Series": series,
        "Value": None,
        "Date": "",
        "Endpoint": url,
    }


def fetch_bank_of_canada_sovereign_rates() -> pd.DataFrame:
    series_map = {
        "Overnight Target": BOC_OVERNIGHT_TARGET_SERIES,
        "Overnight Rate": BOC_OVERNIGHT_RATE_SERIES,
        "3-Month T-Bill": BOC_TBILL_3M_SERIES,
        "Prime Business Rate": BOC_PRIME_SERIES,
    }

    rows = []

    for label, series in series_map.items():
        latest = boc_valet_latest(series)
        value = latest.get("Value")
        rows.append(
            {
                "Benchmark": label,
                "Series": series,
                "Date": latest.get("Date", ""),
                "Rate": "" if value is None else f"{value:.2f}%",
                "Status": latest.get("Status", "Unavailable"),
                "Endpoint": latest.get("Endpoint", ""),
            }
        )

    return pd.DataFrame(rows)


def sovereign_market_pulse_df(deal_rate: float) -> pd.DataFrame:
    rates = fetch_bank_of_canada_sovereign_rates()
    rows = []

    for _, row in rates.iterrows():
        rate_text = str(row.get("Rate", ""))
        spread_bps = ""

        try:
            benchmark = float(rate_text.replace("%", "")) / 100
            spread_bps = f"{(float(deal_rate) - benchmark) * 10000:,.0f}"
        except Exception:
            pass

        rows.append(
            {
                "Market Signal": row.get("Benchmark"),
                "Current Status": row.get("Rate") or row.get("Status"),
                "Source": "Bank of Canada Valet",
                "Deal Spread bps": spread_bps,
            }
        )

    rows.extend(
        [
            {
                "Market Signal": "Geolocation",
                "Current Status": "Live / no key",
                "Source": "NRCan / Geo.ca",
                "Deal Spread bps": "",
            },
            {
                "Market Signal": "Corp Verification",
                "Current Status": "Live / no key",
                "Source": "Corporations Canada",
                "Deal Spread bps": "",
            },
            {
                "Market Signal": "Market Metadata",
                "Current Status": "Open / no key",
                "Source": "Statistics Canada / Open Canada",
                "Deal Spread bps": "",
            },
        ]
    )

    return pd.DataFrame(rows)


def corporations_canada_lookup(identifier: str, lang: str = "eng") -> Dict[str, Any]:
    identifier = str(identifier or "").strip()
    lang = "fra" if lang == "fra" else "eng"

    if not identifier:
        return {"Status": "Missing Identifier"}

    url = f"{CORPORATIONS_CANADA_BASE}/{quote_plus(identifier)}.json?{urlencode({'lang': lang})}"
    data = http_get_json(url, timeout=10)

    if isinstance(data, dict) and data.get("_error"):
        return {"Status": "Error", "Error": data.get("_error"), "Endpoint": url}

    if isinstance(data, list):
        if data and all(isinstance(x, str) for x in data):
            return {"Status": "Not Found", "Message": data[0], "Endpoint": url}

        obj = None
        if lang == "fra" and len(data) > 1 and isinstance(data[1], dict):
            obj = data[1]
        elif len(data) > 0 and isinstance(data[0], dict):
            obj = data[0]
        elif len(data) > 1 and isinstance(data[1], dict):
            obj = data[1]

        if not obj:
            return {"Status": "No Data", "Endpoint": url}

        names = []
        for item in obj.get("corporationNames", []) or []:
            name_obj = item.get("CorporationName", {}) if isinstance(item, dict) else {}
            names.append(name_obj)

        current_name = ""
        for item in names:
            if item.get("current"):
                current_name = item.get("name", "")
                break

        if not current_name and names:
            current_name = names[0].get("name", "")

        addresses = []
        for item in obj.get("adresses", []) or obj.get("addresses", []) or []:
            addr = item.get("address", {}) if isinstance(item, dict) else {}
            addresses.append(addr)

        returns = []
        for item in obj.get("annualReturns", []) or []:
            ar = item.get("annualReturn", {}) if isinstance(item, dict) else {}
            returns.append(ar)

        activities = []
        for item in obj.get("activities", []) or []:
            activity = item.get("activity", {}) if isinstance(item, dict) else {}
            activities.append(activity)

        return {
            "Status": "Found",
            "Corporation ID": obj.get("corporationId", ""),
            "Business Number": (obj.get("businessNumbers") or {}).get("businessNumber", ""),
            "Current Name": current_name,
            "Act": obj.get("act", ""),
            "Corporate Status": obj.get("status", ""),
            "Director Minimum": (obj.get("directorLimits") or {}).get("minimum", ""),
            "Director Maximum": (obj.get("directorLimits") or {}).get("maximum", ""),
            "Addresses": addresses,
            "Annual Returns": returns,
            "Activities": activities,
            "Endpoint": url,
        }

    return {"Status": "Unexpected Response", "Endpoint": url}


def corporation_due_diligence_flags(result: Dict[str, Any]) -> List[str]:
    flags = []

    if not result or result.get("Status") != "Found":
        return ["Federal corporation record not verified. Confirm corporation ID or BN9."]

    corp_status = str(result.get("Corporate Status", "")).lower()

    if "active" not in corp_status and "actif" not in corp_status:
        flags.append(f"Corporation status is not clearly active: {result.get('Corporate Status', '')}")

    annual_returns = result.get("Annual Returns", []) or []

    if not annual_returns:
        flags.append("No annual return history returned. Confirm filing status manually.")
    else:
        years = []
        for item in annual_returns:
            try:
                years.append(int(item.get("yearOfFiling")))
            except Exception:
                pass
        if years and max(years) < datetime.now().year - 2:
            flags.append(f"Most recent annual return appears older than two years: {max(years)}")

    if not result.get("Addresses"):
        flags.append("No registered address returned. Confirm registered office manually.")

    if not flags:
        flags.append("No automated federal registry flags detected.")

    return flags


def sovereign_stack_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Layer": "Geolocation",
                "Source": "NRCan / Geo.ca",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Canadian address matching",
            },
            {
                "Layer": "Rates",
                "Source": "Bank of Canada Valet",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Benchmark-rate context",
            },
            {
                "Layer": "Market Metadata",
                "Source": "Statistics Canada / Open Canada",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Market-rate and economic source references",
            },
            {
                "Layer": "Corporate Registry",
                "Source": "Corporations Canada",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Federal corporation verification",
            },
            {
                "Layer": "Storage",
                "Source": "SQLite / local folders",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Deals, documents, audit trail, exports",
            },
            {
                "Layer": "Lead Routing",
                "Source": "mailto email draft",
                "Account Required": "No",
                "Key Required": "No",
                "Use": "Send deal request to inbox",
            },
        ]
    )


def sovereign_privacy_note() -> str:
    return (
        "Local mode stores the database and uploaded documents in the local alenza_data folder. "
        "If deployed to a hosted environment, uploaded files and calculations are processed by that host."
    )


# ============================================================
# OCR / DOCUMENT INTAKE
# ============================================================

def find_money_near_keywords(text: str, keywords: List[str]) -> Optional[float]:
    for line in text.splitlines():
        normalized = line.lower()
        if any(keyword in normalized for keyword in keywords):
            matches = re.findall(r"\(?\$?\s*-?\d[\d,]*\.?\d*\)?", line)
            if matches:
                value = money_to_float(matches[-1])
                if value is not None:
                    return value
    return None


def extract_text_from_image(uploaded_file) -> str:
    if Image is None or pytesseract is None:
        raise RuntimeError("OCR dependencies are missing. Install pillow and pytesseract.")
    image = Image.open(uploaded_file).convert("RGB")
    return pytesseract.image_to_string(image)


def extract_text_from_pdf(uploaded_file) -> str:
    if fitz is None:
        raise RuntimeError("PDF intake requires PyMuPDF. Install pymupdf.")

    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    extracted = []

    for page in doc:
        native_text = page.get_text("text")
        if native_text and len(native_text.strip()) > 50:
            extracted.append(native_text)
        else:
            if Image is None or pytesseract is None:
                extracted.append("")
            else:
                pix = page.get_pixmap(dpi=200)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                extracted.append(pytesseract.image_to_string(img))

    return "\n".join(extracted)


def parse_financials_from_text(text: str) -> Dict[str, Optional[float]]:
    gross_income = find_money_near_keywords(
        text,
        ["gross income", "gross rental income", "rental income", "total income", "egi", "revenue"],
    )
    vacancy = find_money_near_keywords(text, ["vacancy", "credit loss"])
    expenses = find_money_near_keywords(
        text,
        ["operating expenses", "total expenses", "property expenses", "opex"],
    )
    noi = find_money_near_keywords(text, ["net operating income", "noi"])
    purchase_price = find_money_near_keywords(text, ["purchase price", "acquisition price", "cost basis"])
    appraised_value = find_money_near_keywords(text, ["appraised value", "market value", "as-is value"])

    if noi is None and gross_income is not None and expenses is not None:
        noi = gross_income - abs(vacancy or 0) - expenses

    return {
        "Purchase Price / Cost Basis": purchase_price,
        "Appraised Value": appraised_value,
        "Gross Income": gross_income,
        "Vacancy / Credit Loss": vacancy,
        "Operating Expenses": expenses,
        "Stabilized NOI": noi,
    }


def process_uploaded_financial(uploaded_file) -> Tuple[str, Dict[str, Optional[float]]]:
    name = uploaded_file.name.lower()

    if name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        text = extract_text_from_image(uploaded_file)
    elif name.endswith(".pdf"):
        text = extract_text_from_pdf(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Upload PDF, PNG, JPG, JPEG, or WEBP.")

    return text, parse_financials_from_text(text)


# ============================================================
# UNDERWRITING ENGINE
# ============================================================

def monthly_payment_amortizing(loan_amount: float, rate: float, amort_years: int) -> float:
    monthly_rate = rate / 12
    periods = amort_years * 12

    if loan_amount <= 0 or monthly_rate <= 0 or periods <= 0:
        return 0

    return (loan_amount * monthly_rate) / (1 - (1 + monthly_rate) ** -periods)


def monthly_payment_interest_only(loan_amount: float, rate: float) -> float:
    if loan_amount <= 0 or rate <= 0:
        return 0
    return loan_amount * rate / 12


def calculate_monthly_payment(loan_amount: float, rate: float, amort_years: int, debt_structure: str) -> float:
    if debt_structure == "Interest-Only":
        return monthly_payment_interest_only(loan_amount, rate)
    return monthly_payment_amortizing(loan_amount, rate, amort_years)


def size_loan(
    noi: float,
    appraisal: float,
    total_uses_before_fees: float,
    rate: float,
    amort_years: int,
    target_ltv: float,
    target_ltc: float,
    target_dscr: float,
    target_dy: float,
    debt_structure: str,
    fees: float,
) -> Tuple[float, str, Dict[str, float]]:
    """Multi-pass loan sizing with rough circular fee treatment."""
    loan_guess = min(appraisal * target_ltv, noi / target_dy if target_dy else 0)

    for _ in range(8):
        total_uses = total_uses_before_fees + loan_guess * fees

        ltv_limit = appraisal * target_ltv
        ltc_limit = total_uses * target_ltc
        monthly_rate = rate / 12
        periods = amort_years * 12
        monthly_dscr_capacity = (noi / target_dscr) / 12 if target_dscr > 0 else 0

        if debt_structure == "Interest-Only":
            dscr_limit = monthly_dscr_capacity / monthly_rate if monthly_rate > 0 else 0
        else:
            if monthly_rate > 0 and periods > 0:
                dscr_limit = monthly_dscr_capacity * ((1 - (1 + monthly_rate) ** -periods) / monthly_rate)
            else:
                dscr_limit = 0

        debt_yield_limit = noi / target_dy if target_dy > 0 else 0

        gates = {
            "LTV": ltv_limit,
            "LTC": ltc_limit,
            "DSCR": dscr_limit,
            "Debt Yield": debt_yield_limit,
        }

        new_loan = max(0, min(gates.values()))

        if abs(new_loan - loan_guess) < 1:
            loan_guess = new_loan
            break

        loan_guess = new_loan

    binding_gate = min(gates, key=gates.get)
    return loan_guess, binding_gate, gates


def generate_amortization_schedule(
    loan_amt: float,
    rate: float,
    amort_years: int,
    term_years: int,
    debt_structure: str,
) -> pd.DataFrame:
    m_rate = rate / 12
    term_months = int(term_years * 12)
    amort_months = int(amort_years * 12)

    rows = []
    balance = float(loan_amt)

    if debt_structure == "Interest-Only":
        payment = loan_amt * m_rate
    else:
        if m_rate > 0:
            payment = (loan_amt * m_rate) / (1 - (1 + m_rate) ** -amort_months)
        else:
            payment = loan_amt / amort_months

    for period in range(1, term_months + 1):
        interest = balance * m_rate
        principal = 0 if debt_structure == "Interest-Only" else payment - interest
        principal = min(principal, balance)
        closing = max(0, balance - principal)

        rows.append(
            {
                "Period": period,
                "Opening Balance": balance,
                "Total Payment": payment,
                "Principal": principal,
                "Interest": interest,
                "Closing Balance": closing,
            }
        )

        balance = closing
        if balance <= 0:
            break

    return pd.DataFrame(rows)


def classify_deal(score: int) -> str:
    if score >= 900:
        return "Tier 1A | Institutional Core Credit"
    if score >= 800:
        return "Tier 1 | High Bankability"
    if score >= 675:
        return "Tier 2 | Bankable / Select Alternative"
    if score >= 525:
        return "Tier 3 | Alternative / Structured Credit"
    return "Tier 4 | Private / Bridge / Restructure Required"


def constraint_advice(binding_gate: str) -> str:
    if binding_gate == "LTV":
        return "The transaction is leverage-constrained. Higher proceeds require a higher valuation, additional collateral, or a lender willing to advance at a higher LTV."
    if binding_gate == "LTC":
        return "The transaction is cost-constrained. Higher proceeds require lower total uses, lower reserves/fees, or a lender willing to advance at a higher LTC."
    if binding_gate == "DSCR":
        return "The transaction is cash-flow constrained. Higher proceeds require higher NOI, lower rate, interest-only debt, longer amortization, or a lower DSCR requirement."
    return "The transaction is debt-yield constrained. Higher proceeds require stronger NOI or a lower lender debt-yield threshold."


def pass_fail(actual: float, threshold: float, mode: str = "gte") -> str:
    if mode == "gte":
        return "PASS" if actual >= threshold else "FAIL"
    return "PASS" if actual <= threshold else "FAIL"


def calculate_deal_score(actual_ltv: float, actual_ltc: float, actual_dscr: float, actual_dy: float, equity_pct: float) -> int:
    ltv_score = 260 if actual_ltv <= 0.65 else 210 if actual_ltv <= 0.70 else 160 if actual_ltv <= 0.75 else 80 if actual_ltv <= 0.80 else 0
    ltc_score = 140 if actual_ltc <= 0.70 else 110 if actual_ltc <= 0.75 else 75 if actual_ltc <= 0.80 else 30 if actual_ltc <= 0.85 else 0
    dscr_score = 260 if actual_dscr >= 1.45 else 210 if actual_dscr >= 1.35 else 160 if actual_dscr >= 1.25 else 75 if actual_dscr >= 1.15 else 0
    dy_score = 200 if actual_dy >= 0.095 else 160 if actual_dy >= 0.085 else 110 if actual_dy >= 0.075 else 50 if actual_dy >= 0.065 else 0
    equity_score = 140 if equity_pct >= 0.30 else 110 if equity_pct >= 0.25 else 75 if equity_pct >= 0.20 else 35 if equity_pct >= 0.15 else 0
    return int(ltv_score + ltc_score + dscr_score + dy_score + equity_score)


# ============================================================
# RENT ROLL + DILIGENCE
# ============================================================

def default_rent_roll_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"Tenant": "Anchor A", "SF": 25000, "Remaining Term": 8.5, "Monthly Rent": 45833},
            {"Tenant": "In-Line B", "SF": 5000, "Remaining Term": 2.0, "Monthly Rent": 14583},
            {"Tenant": "Vacant Unit", "SF": 2000, "Remaining Term": 0.0, "Monthly Rent": 0},
        ]
    )


def calculate_rent_roll_metrics(df: pd.DataFrame) -> Dict[str, float]:
    if df is None or df.empty:
        return {
            "WALT": 0,
            "Total_SF": 0,
            "Vacancy_SF": 0,
            "Occupancy": 0,
            "Expiring_1yr_pct": 0,
            "Annual_Rent": 0,
            "Average_Rent_PSF": 0,
        }

    work = df.copy()
    work["SF"] = pd.to_numeric(work.get("SF", 0), errors="coerce").fillna(0)
    work["Remaining Term"] = pd.to_numeric(work.get("Remaining Term", 0), errors="coerce").fillna(0)
    work["Monthly Rent"] = pd.to_numeric(work.get("Monthly Rent", 0), errors="coerce").fillna(0)

    total_sf = work["SF"].sum()
    occupied_sf = work.loc[work["Monthly Rent"] > 0, "SF"].sum()
    vacant_sf = max(0, total_sf - occupied_sf)

    if total_sf > 0:
        walt = (work["Remaining Term"] * (work["SF"] / total_sf)).sum()
        expiring_pct = safe_divide(work.loc[work["Remaining Term"] <= 1, "SF"].sum(), total_sf)
        occupancy = safe_divide(occupied_sf, total_sf)
    else:
        walt = 0
        expiring_pct = 0
        occupancy = 0

    annual_rent = work["Monthly Rent"].sum() * 12
    avg_rent_psf = safe_divide(annual_rent, occupied_sf)

    return {
        "WALT": float(walt),
        "Total_SF": float(total_sf),
        "Vacancy_SF": float(vacant_sf),
        "Occupancy": float(occupancy),
        "Expiring_1yr_pct": float(expiring_pct),
        "Annual_Rent": float(annual_rent),
        "Average_Rent_PSF": float(avg_rent_psf),
    }


DILIGENCE_REQUIREMENTS = {
    "Financials (T12)": ["t12", "operating", "p&l", "pnl", "financial", "income statement", "trailing"],
    "Rent Roll": ["rent roll", "rentroll", "tenant", "rr"],
    "Appraisal": ["appraisal", "valuation", "value report"],
    "Environmental": ["phase", "enviro", "environmental", "esa"],
    "Sponsor Info": ["bio", "sreo", "schedule real estate owned", "experience", "sponsor"],
    "Purchase Agreement": ["purchase agreement", "psa", "contract"],
    "Insurance": ["insurance", "certificate", "coi"],
    "Title / Survey": ["title", "survey", "alta"],
}


def audit_deal_room(files) -> Tuple[List[str], List[str], pd.DataFrame]:
    if not files:
        return [], list(DILIGENCE_REQUIREMENTS.keys()), pd.DataFrame(
            columns=["Requirement", "Status", "Matched File"]
        )

    found = []
    missing = []
    rows = []

    for category, keywords in DILIGENCE_REQUIREMENTS.items():
        matched_file = None

        for file in files:
            name = getattr(file, "name", "").lower()
            if any(keyword in name for keyword in keywords):
                matched_file = getattr(file, "name", "Uploaded file")
                break

        if matched_file:
            found.append(category)
            rows.append({"Requirement": category, "Status": "Received", "Matched File": matched_file})
        else:
            missing.append(category)
            rows.append({"Requirement": category, "Status": "Missing", "Matched File": ""})

    return found, missing, pd.DataFrame(rows)


# ============================================================
# EXPORTS
# ============================================================

def create_excel_workbook(sheets: Dict[str, pd.DataFrame], metadata: Dict[str, str]) -> io.BytesIO:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": "#1D4ED8", "border": 1}
        )
        title_format = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1D4ED8"})

        cover = workbook.add_worksheet("Cover")
        cover.write("A1", "ALENZA CAPITAL OS", title_format)
        cover.write("A2", APP_VERSION)

        row = 4
        for key, value in metadata.items():
            cover.write(row, 0, key, header_format)
            cover.write(row, 1, value)
            row += 1

        cover.set_column("A:A", 32)
        cover.set_column("B:B", 48)

        for sheet_name, df in sheets.items():
            safe_sheet = sheet_name[:31]
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            ws = writer.sheets[safe_sheet]
            ws.set_column("A:Z", 22)
            for col_num, value in enumerate(df.columns):
                ws.write(0, col_num, value, header_format)

    output.seek(0)
    return output


def create_pdf_summary(metadata: Dict[str, str], preview_df: pd.DataFrame, covenant_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if SimpleDocTemplate is None:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("ALENZA CAPITAL OS", styles["Title"]))
    story.append(Paragraph("Canadian CRE Underwriting Summary", styles["Heading2"]))
    story.append(Spacer(1, 12))

    for key, value in metadata.items():
        story.append(Paragraph(f"<b>{key}:</b> {value}", styles["Normal"]))

    story.append(Spacer(1, 14))

    def add_table(title: str, df: pd.DataFrame):
        story.append(Paragraph(title, styles["Heading2"]))
        data = [list(df.columns)] + df.astype(str).values.tolist()
        table = Table(data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        story.append(table)
        story.append(Spacer(1, 12))

    add_table("Executive Summary", preview_df)
    add_table("Covenant Testing", covenant_df)

    disclaimer = (
        "This summary is indicative only and is not a loan commitment, appraisal, legal opinion, "
        "tax advice, investment recommendation, or final credit decision."
    )
    story.append(Paragraph("Important Notice", styles["Heading2"]))
    story.append(Paragraph(disclaimer, styles["Normal"]))

    doc.build(story)
    buffer.seek(0)
    return buffer


def build_deal_json(data: Dict[str, Any]) -> bytes:
    return json.dumps(data, indent=2, default=str).encode("utf-8")


def create_backup_zip(deal_json: bytes, excel_file: io.BytesIO, report_text: str) -> io.BytesIO:
    init_db()
    buffer = io.BytesIO()

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("current_deal.json", deal_json)
        z.writestr("current_report.txt", report_text)
        z.writestr("current_workbook.xlsx", excel_file.getvalue())

        if DB_PATH.exists():
            z.write(DB_PATH, "alenza_platform.db")

        z.writestr("deals_export.csv", get_deals_df().to_csv(index=False))
        z.writestr("audit_export.csv", get_audit_df().to_csv(index=False))
        z.writestr("README.txt", "Alenza local backup package.\n")

    buffer.seek(0)
    return buffer


# ============================================================
# SESSION STATE
# ============================================================

def init_state() -> None:
    defaults = {
        "sponsor": "Client Name",
        "property_address": "100 King Street West, Toronto, ON",
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
    }

    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

    if "rent_roll_df" not in st.session_state:
        st.session_state.rent_roll_df = default_rent_roll_df()

    st.session_state.setdefault("raw_ocr_text", None)
    st.session_state.setdefault("extracted_fields", None)


init_db()
init_state()


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.title("ALENZA CAPITAL OS")
    st.caption("Sovereign Canada Build")
    st.caption(f"Deal inbox: {get_deal_inbox_email()}")
    st.markdown("---")

    with st.expander("Auto Intake / OCR", expanded=False):
        uploaded_financial = st.file_uploader(
            "Upload PDF or image financials",
            type=["pdf", "png", "jpg", "jpeg", "webp"],
            key="ocr_upload",
        )

        if uploaded_financial is not None:
            try:
                raw_text, fields = process_uploaded_financial(uploaded_financial)
                st.session_state.raw_ocr_text = raw_text
                st.session_state.extracted_fields = fields

                review_df = pd.DataFrame(
                    {
                        "Field": list(fields.keys()),
                        "Extracted Value": ["" if v is None else format_money(v) for v in fields.values()],
                    }
                )

                st.success("Document processed.")
                st.dataframe(review_df, hide_index=True, use_container_width=True)

                if fields.get("Purchase Price / Cost Basis"):
                    st.session_state.purchase_price = int(fields["Purchase Price / Cost Basis"])
                if fields.get("Appraised Value"):
                    st.session_state.appraisal = int(fields["Appraised Value"])
                if fields.get("Stabilized NOI"):
                    st.session_state.noi = int(fields["Stabilized NOI"])

                with st.expander("Raw OCR Text", expanded=False):
                    st.text_area("OCR Output", raw_text, height=220)

            except Exception as exc:
                st.error(f"OCR processing failed: {exc}")

    with st.expander("Asset Information", expanded=True):
        st.session_state.sponsor = st.text_input("Sponsor / Borrower", st.session_state.sponsor)
        st.session_state.property_address = st.text_input("Property Address", st.session_state.property_address)

        st.session_state.property_type = st.selectbox(
            "Property Type",
            [
                "Multifamily",
                "Industrial",
                "Retail",
                "Office",
                "Mixed-Use",
                "Hospitality",
                "Self-Storage",
                "Medical Office",
                "Other",
            ],
            index=[
                "Multifamily",
                "Industrial",
                "Retail",
                "Office",
                "Mixed-Use",
                "Hospitality",
                "Self-Storage",
                "Medical Office",
                "Other",
            ].index(st.session_state.property_type)
            if st.session_state.property_type
            in [
                "Multifamily",
                "Industrial",
                "Retail",
                "Office",
                "Mixed-Use",
                "Hospitality",
                "Self-Storage",
                "Medical Office",
                "Other",
            ]
            else 0,
        )

        st.session_state.transaction_type = st.selectbox(
            "Transaction Type",
            ["Acquisition", "Refinance"],
            index=0 if st.session_state.transaction_type == "Acquisition" else 1,
        )

        st.session_state.purchase_price = st.number_input(
            "Purchase Price / Cost Basis ($)",
            value=int(st.session_state.purchase_price),
            min_value=1,
            step=100000,
        )

        st.session_state.appraisal = st.number_input(
            "Appraised Value ($)",
            value=int(st.session_state.appraisal),
            min_value=1,
            step=100000,
        )

        if st.session_state.transaction_type == "Refinance":
            st.session_state.existing_debt = st.number_input(
                "Existing Debt Payoff ($)",
                value=int(st.session_state.existing_debt),
                min_value=0,
                step=100000,
            )
        else:
            st.session_state.existing_debt = 0

        st.session_state.gross_income = st.number_input(
            "Gross Income / EGI ($)",
            value=int(st.session_state.gross_income),
            min_value=0,
            step=10000,
        )
        st.session_state.vacancy_loss = st.number_input(
            "Vacancy / Credit Loss ($)",
            value=int(st.session_state.vacancy_loss),
            min_value=0,
            step=5000,
        )
        st.session_state.operating_expenses = st.number_input(
            "Operating Expenses ($)",
            value=int(st.session_state.operating_expenses),
            min_value=0,
            step=10000,
        )
        calculated_noi = st.session_state.gross_income - st.session_state.vacancy_loss - st.session_state.operating_expenses
        st.session_state.noi = st.number_input(
            "Stabilized NOI ($)",
            value=int(st.session_state.noi or calculated_noi),
            min_value=1,
            step=10000,
        )

    with st.expander("Underwriting Criteria", expanded=True):
        st.session_state.target_ltv = st.slider("Maximum LTV (%)", 50, 85, int(st.session_state.target_ltv * 100)) / 100
        st.session_state.target_ltc = st.slider("Maximum LTC (%)", 50, 90, int(st.session_state.target_ltc * 100)) / 100
        st.session_state.target_dscr = st.slider("Minimum DSCR (x)", 1.10, 1.75, float(st.session_state.target_dscr), 0.05)
        st.session_state.target_dy = st.slider("Minimum Debt Yield (%)", 5.0, 15.0, float(st.session_state.target_dy * 100), 0.25) / 100

    with st.expander("Loan Terms", expanded=True):
        st.session_state.debt_structure = st.selectbox(
            "Debt Service Structure",
            ["Amortizing", "Interest-Only"],
            index=0 if st.session_state.debt_structure == "Amortizing" else 1,
        )
        st.session_state.rate = st.slider("Interest Rate (%)", 3.0, 12.0, float(st.session_state.rate * 100), 0.125) / 100
        st.session_state.amort = st.number_input("Amortization (Years)", value=int(st.session_state.amort), min_value=1, max_value=40)
        st.session_state.loan_term = st.number_input("Loan Term (Years)", value=int(st.session_state.loan_term), min_value=1, max_value=30)
        st.session_state.fees = st.slider("Origination / Financing Fees (%)", 0.0, 5.0, float(st.session_state.fees * 100), 0.25) / 100
        st.session_state.closing_costs = st.number_input("Other Closing Costs ($)", value=int(st.session_state.closing_costs), min_value=0, step=5000)
        st.session_state.capex_reserve = st.number_input("CapEx / TI-LC Reserve ($)", value=int(st.session_state.capex_reserve), min_value=0, step=25000)
        st.session_state.interest_reserve = st.number_input("Interest Reserve ($)", value=int(st.session_state.interest_reserve), min_value=0, step=25000)


# ============================================================
# CALCULATIONS
# ============================================================

inputs = {
    "sponsor": st.session_state.sponsor,
    "property_address": st.session_state.property_address,
    "property_type": st.session_state.property_type,
    "transaction_type": st.session_state.transaction_type,
    "purchase_price": st.session_state.purchase_price,
    "appraisal": st.session_state.appraisal,
    "existing_debt": st.session_state.existing_debt,
    "gross_income": st.session_state.gross_income,
    "vacancy_loss": st.session_state.vacancy_loss,
    "operating_expenses": st.session_state.operating_expenses,
    "noi": st.session_state.noi,
    "target_ltv": st.session_state.target_ltv,
    "target_ltc": st.session_state.target_ltc,
    "target_dscr": st.session_state.target_dscr,
    "target_dy": st.session_state.target_dy,
    "rate": st.session_state.rate,
    "amort": st.session_state.amort,
    "loan_term": st.session_state.loan_term,
    "fees": st.session_state.fees,
    "closing_costs": st.session_state.closing_costs,
    "capex_reserve": st.session_state.capex_reserve,
    "interest_reserve": st.session_state.interest_reserve,
    "debt_structure": st.session_state.debt_structure,
}

base_uses = inputs["purchase_price"] if inputs["transaction_type"] == "Acquisition" else inputs["existing_debt"]
uses_before_fees = base_uses + inputs["closing_costs"] + inputs["capex_reserve"] + inputs["interest_reserve"]

loan_amt, gate, gates = size_loan(
    noi=inputs["noi"],
    appraisal=inputs["appraisal"],
    total_uses_before_fees=uses_before_fees,
    rate=inputs["rate"],
    amort_years=inputs["amort"],
    target_ltv=inputs["target_ltv"],
    target_ltc=inputs["target_ltc"],
    target_dscr=inputs["target_dscr"],
    target_dy=inputs["target_dy"],
    debt_structure=inputs["debt_structure"],
    fees=inputs["fees"],
)

financing_fees = loan_amt * inputs["fees"]
total_uses = uses_before_fees + financing_fees
required_equity = total_uses - loan_amt

monthly_payment = calculate_monthly_payment(loan_amt, inputs["rate"], inputs["amort"], inputs["debt_structure"])
annual_debt_service = monthly_payment * 12

actual_ltv = safe_divide(loan_amt, inputs["appraisal"])
actual_ltc = safe_divide(loan_amt, total_uses)
actual_dscr = safe_divide(inputs["noi"], annual_debt_service)
actual_dy = safe_divide(inputs["noi"], loan_amt)
equity_pct = safe_divide(required_equity, total_uses)
expense_ratio = safe_divide(inputs["operating_expenses"], inputs["gross_income"])
implied_cap_rate = safe_divide(inputs["noi"], inputs["appraisal"])

score = calculate_deal_score(actual_ltv, actual_ltc, actual_dscr, actual_dy, equity_pct)
classification = classify_deal(score)
generated_at = now_str()

amort_schedule_df = generate_amortization_schedule(
    loan_amt=loan_amt,
    rate=inputs["rate"],
    amort_years=int(inputs["amort"]),
    term_years=int(inputs["loan_term"]),
    debt_structure=inputs["debt_structure"],
)

balloon_balance = amort_schedule_df["Closing Balance"].iloc[-1] if not amort_schedule_df.empty else loan_amt
rent_roll_metrics = calculate_rent_roll_metrics(st.session_state.rent_roll_df)

preview_df = pd.DataFrame(
    {
        "Field": [
            "Sponsor / Borrower",
            "Property Address",
            "Property Type",
            "Transaction Type",
            "Supportable Proceeds",
            "Binding Constraint",
            "Actual LTV",
            "Actual LTC",
            "Actual DSCR",
            "Debt Yield",
            "Required Equity",
            "Deal Score",
            "Classification",
        ],
        "Value": [
            inputs["sponsor"],
            inputs["property_address"],
            inputs["property_type"],
            inputs["transaction_type"],
            format_money(loan_amt),
            gate,
            format_pct(actual_ltv),
            format_pct(actual_ltc),
            format_x(actual_dscr),
            format_pct(actual_dy),
            format_money(required_equity),
            f"{score}/1000",
            classification,
        ],
    }
)

sizing_df = pd.DataFrame(
    {
        "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
        "Threshold": [
            format_pct(inputs["target_ltv"]),
            format_pct(inputs["target_ltc"]),
            format_x(inputs["target_dscr"]),
            format_pct(inputs["target_dy"]),
        ],
        "Max Proceeds": [gates["LTV"], gates["LTC"], gates["DSCR"], gates["Debt Yield"]],
        "Binding": ["YES" if gate == x else "" for x in ["LTV", "LTC", "DSCR", "Debt Yield"]],
    }
)

uses_df = pd.DataFrame(
    {
        "Project Uses": [
            "Purchase Price / Existing Debt",
            "Origination / Financing Fees",
            "Other Closing Costs",
            "CapEx / TI-LC Reserve",
            "Interest Reserve",
            "Total Uses",
        ],
        "Amount": [
            base_uses,
            financing_fees,
            inputs["closing_costs"],
            inputs["capex_reserve"],
            inputs["interest_reserve"],
            total_uses,
        ],
    }
)

sources_df = pd.DataFrame(
    {
        "Project Sources": ["Supportable Senior Debt", "Required Sponsor Equity", "Total Sources"],
        "Amount": [loan_amt, required_equity, loan_amt + required_equity],
    }
)

covenant_df = pd.DataFrame(
    {
        "Covenant": ["Maximum LTV", "Maximum LTC", "Minimum DSCR", "Minimum Debt Yield"],
        "Required": [
            f"<= {format_pct(inputs['target_ltv'])}",
            f"<= {format_pct(inputs['target_ltc'])}",
            f">= {format_x(inputs['target_dscr'])}",
            f">= {format_pct(inputs['target_dy'])}",
        ],
        "Actual": [
            format_pct(actual_ltv),
            format_pct(actual_ltc),
            format_x(actual_dscr),
            format_pct(actual_dy),
        ],
        "Status": [
            pass_fail(actual_ltv, inputs["target_ltv"], "lte"),
            pass_fail(actual_ltc, inputs["target_ltc"], "lte"),
            pass_fail(actual_dscr, inputs["target_dscr"], "gte"),
            pass_fail(actual_dy, inputs["target_dy"], "gte"),
        ],
    }
)

assumptions_df = pd.DataFrame(
    {
        "Assumption": list(inputs.keys()),
        "Value": [str(v) for v in inputs.values()],
    }
)

score_df = pd.DataFrame(
    {
        "Component": ["Loan-to-Value", "Loan-to-Cost", "Debt Service Coverage", "Debt Yield", "Equity Contribution"],
        "Actual": [format_pct(actual_ltv), format_pct(actual_ltc), format_x(actual_dscr), format_pct(actual_dy), format_pct(equity_pct)],
        "Maximum": [260, 140, 260, 200, 140],
    }
)

risk_flags = []
if actual_ltv > inputs["target_ltv"]:
    risk_flags.append("LTV exceeds target.")
if actual_ltc > inputs["target_ltc"]:
    risk_flags.append("LTC exceeds target.")
if actual_dscr < inputs["target_dscr"]:
    risk_flags.append("DSCR is below target.")
if actual_dy < inputs["target_dy"]:
    risk_flags.append("Debt yield is below target.")
if expense_ratio > 0.55:
    risk_flags.append("Expense ratio appears elevated.")
if not risk_flags:
    risk_flags.append("No major automated risk flags.")


# ============================================================
# MAIN UI
# ============================================================

st.title("Alenza Capital OS")
st.subheader("Canadian CRE Debt Underwriting Workstation")
st.caption(f"{APP_VERSION} | Generated: {generated_at} | Active Constraint: {gate}")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Max Proceeds", format_money(loan_amt))
m2.metric("Actual LTV", format_pct(actual_ltv))
m3.metric("Actual LTC", format_pct(actual_ltc))
m4.metric("Actual DSCR", format_x(actual_dscr))
m5.metric("Debt Yield", format_pct(actual_dy))
m6.metric("Deal Score", f"{score}/1000")

st.markdown("---")

tabs = st.tabs(
    [
        "Executive",
        "Sovereign Intelligence",
        "Market Rates",
        "Corporate Registry",
        "Rent Roll",
        "Sizing",
        "Amortization",
        "Diligence",
        "Portfolio",
        "Report",
    ]
)


with tabs[0]:
    left, right = st.columns([1.4, 1])

    with left:
        st.subheader("Executive Summary")
        st.dataframe(preview_df, hide_index=True, use_container_width=True)

        st.subheader("Risk Flags")
        st.dataframe(pd.DataFrame({"Flag": risk_flags}), hide_index=True, use_container_width=True)

    with right:
        st.subheader("Underwriting Verdict")
        st.info(f"Supportable proceeds are constrained by {gate}.")
        st.write(constraint_advice(gate))

        if score >= 850:
            st.success(classification)
        elif score >= 675:
            st.warning(classification)
        else:
            st.error(classification)

        if st.button("Save Deal Snapshot Locally"):
            deal_id = save_deal_snapshot(inputs, loan_amt, score, classification)
            st.success(f"Saved deal snapshot #{deal_id}.")

        body = f"""Alenza Deal Snapshot

Sponsor: {inputs['sponsor']}
Property: {inputs['property_address']}
Supportable Proceeds: {format_money(loan_amt)}
Binding Constraint: {gate}
Score: {score}/1000
Classification: {classification}
"""
        st.link_button("Email Deal Snapshot", build_deal_email_link("Alenza Deal Snapshot", body))


with tabs[1]:
    st.subheader("Sovereign Property Intelligence")
    st.caption("Open-access Canadian data sources. No mandatory accounts. No mandatory API keys.")

    st.info(sovereign_privacy_note())

    st.write("Sovereign Canada Stack")
    st.dataframe(sovereign_stack_df(), hide_index=True, use_container_width=True)

    address = st.text_input("Enter Canadian Address", value=st.session_state.property_address, key="intel_address")
    st.session_state.property_address = address

    if address:
        results = nrcan_standardize_address(address)

        if results:
            labels = [
                f"{r.get('name', '')} | {r.get('province', '')} | {r.get('category', '')}"
                for r in results[:10]
            ]
            selected = st.selectbox("Official NRCan Standard", labels)
            selected_index = labels.index(selected)
            record = results[selected_index]
            st.session_state.property_address = f"{record.get('name', address)}, {record.get('province', '')}".strip(", ")
            with st.expander("NRCan Results"):
                st.dataframe(pd.DataFrame(results[:10]), hide_index=True, use_container_width=True)

    c1, c2 = st.columns([1.5, 1])

    with c1:
        st.write(f"Satellite Analysis: {st.session_state.property_address}")
        iframe_src = f"https://www.google.com/maps?q={quote_plus(st.session_state.property_address)}&output=embed&t=k"
        components.iframe(iframe_src, height=350)

    with c2:
        st.write("Market Pulse")
        st.dataframe(sovereign_market_pulse_df(inputs["rate"]), hide_index=True, use_container_width=True)

        st.write("Property Intelligence")
        intel = get_property_intelligence(st.session_state.property_address)
        st.dataframe(pd.DataFrame(intel.items(), columns=["Fact", "Value"]), hide_index=True, use_container_width=True)

    st.divider()

    st.write("Diligence Audit & Partnership Request")
    bulk_files = st.file_uploader(
        "Upload current deal files for a gap audit",
        accept_multiple_files=True,
        type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "csv", "docx", "txt"],
        key="gap_audit",
    )

    if bulk_files:
        found, missing, audit_df = audit_deal_room(bulk_files)
        st.dataframe(audit_df, hide_index=True, use_container_width=True)

        if missing:
            st.error(f"Diligence gaps detected: {', '.join(missing)}")
            lead_body = build_concierge_email_body("User", "Email", st.session_state.property_address, missing)
            st.link_button("Request Alenza Concierge Help", build_deal_email_link("Alenza Concierge Lead", lead_body))
        else:
            st.success("No major diligence gaps detected.")


with tabs[2]:
    st.subheader("Canadian Market Rates")
    st.caption("Live no-key federal rate context from Bank of Canada Valet.")

    st.write("Live Bank of Canada Rates")
    boc_df = fetch_bank_of_canada_sovereign_rates()
    st.dataframe(boc_df, hide_index=True, use_container_width=True)

    st.write("Deal Rate Comparison")
    st.dataframe(sovereign_market_pulse_df(inputs["rate"]), hide_index=True, use_container_width=True)


with tabs[3]:
    st.subheader("Federal Corporation Registry")
    st.caption("No-key Corporations Canada lookup by corporation ID or 9-digit business number.")

    corp_id = st.text_input("Corporation ID or BN9")
    corp_lang = st.selectbox("Language", ["eng", "fra"], format_func=lambda x: "English" if x == "eng" else "French")

    if corp_id:
        result = corporations_canada_lookup(corp_id, corp_lang)
        st.dataframe(pd.DataFrame(result.items(), columns=["Field", "Value"]), hide_index=True, use_container_width=True)

        flags = corporation_due_diligence_flags(result)
        st.write("Due Diligence Flags")
        st.dataframe(pd.DataFrame({"Flag": flags}), hide_index=True, use_container_width=True)

        if result.get("Addresses"):
            st.write("Registered Addresses")
            st.dataframe(pd.DataFrame(result["Addresses"]), hide_index=True, use_container_width=True)

        if result.get("Annual Returns"):
            st.write("Annual Returns")
            st.dataframe(pd.DataFrame(result["Annual Returns"]), hide_index=True, use_container_width=True)


with tabs[4]:
    st.subheader("Rent Roll Analyzer")
    st.session_state.rent_roll_df = st.data_editor(
        st.session_state.rent_roll_df,
        num_rows="dynamic",
        use_container_width=True,
    )

    rr = calculate_rent_roll_metrics(st.session_state.rent_roll_df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("WALT", f"{rr['WALT']:.2f} yrs")
    c2.metric("Occupancy", format_pct(rr["Occupancy"]))
    c3.metric("12-Mo Rollover", format_pct(rr["Expiring_1yr_pct"]))
    c4.metric("Avg Rent PSF", f"${rr['Average_Rent_PSF']:,.2f}")

    st.dataframe(pd.DataFrame(rr.items(), columns=["Metric", "Value"]), hide_index=True, use_container_width=True)


with tabs[5]:
    st.subheader("Loan Sizing Constraints")
    st.dataframe(
        sizing_df.style.format({"Max Proceeds": "${:,.0f}"}),
        hide_index=True,
        use_container_width=True,
    )

    st.subheader("Sources and Uses")
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(uses_df.style.format({"Amount": "${:,.0f}"}), hide_index=True, use_container_width=True)
    with c2:
        st.dataframe(sources_df.style.format({"Amount": "${:,.0f}"}), hide_index=True, use_container_width=True)

    st.subheader("Covenants")
    st.dataframe(covenant_df, hide_index=True, use_container_width=True)


with tabs[6]:
    st.subheader("Amortization Schedule")
    st.caption(f"{inputs['loan_term']}-year term | {inputs['debt_structure']}")

    st.dataframe(
        amort_schedule_df.style.format(
            {
                "Opening Balance": "${:,.2f}",
                "Total Payment": "${:,.2f}",
                "Principal": "${:,.2f}",
                "Interest": "${:,.2f}",
                "Closing Balance": "${:,.2f}",
            }
        ),
        hide_index=True,
        use_container_width=True,
    )

    a1, a2, a3 = st.columns(3)
    a1.metric("Principal Paid", format_money(amort_schedule_df["Principal"].sum() if not amort_schedule_df.empty else 0))
    a2.metric("Interest Paid", format_money(amort_schedule_df["Interest"].sum() if not amort_schedule_df.empty else 0))
    a3.metric("Balloon Balance", format_money(balloon_balance))


with tabs[7]:
    st.subheader("Document Room / Diligence")
    files = st.file_uploader(
        "Upload deal documents",
        accept_multiple_files=True,
        type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "xls", "csv", "docx", "txt"],
        key="document_room",
    )

    if st.button("Save Uploaded Documents Locally"):
        saved = save_uploaded_documents(files, "Deal Document")
        st.success(f"Saved {len(saved)} file(s).")

    if files:
        found, missing, audit_df = audit_deal_room(files)
        st.write("Gap Audit")
        st.dataframe(audit_df, hide_index=True, use_container_width=True)

    st.write("Audit Trail")
    st.dataframe(get_audit_df(), hide_index=True, use_container_width=True)


with tabs[8]:
    st.subheader("Portfolio")
    deals_df = get_deals_df()

    if deals_df.empty:
        st.info("No saved deal snapshots yet.")
    else:
        st.dataframe(deals_df, hide_index=True, use_container_width=True)
        p1, p2, p3 = st.columns(3)
        p1.metric("Saved Deals", f"{len(deals_df):,.0f}")
        p2.metric("Total Loan Amount", format_money(deals_df["loan_amount"].fillna(0).sum()))
        p3.metric("Average Score", f"{deals_df['score'].fillna(0).mean():.0f}")


with tabs[9]:
    st.subheader("Report & Exports")
    st.dataframe(preview_df, hide_index=True, use_container_width=True)

    report_text = f"""ALENZA CAPITAL OS
Generated: {generated_at}
Version: {APP_VERSION}

EXECUTIVE SUMMARY
Sponsor / Borrower: {inputs['sponsor']}
Property Address: {inputs['property_address']}
Property Type: {inputs['property_type']}
Transaction Type: {inputs['transaction_type']}

Supportable Proceeds: {format_money(loan_amt)}
Binding Constraint: {gate}
Classification: {classification}
Deal Score: {score}/1000

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
{chr(10).join("- " + flag for flag in risk_flags)}

UNDERWRITING VERDICT
{constraint_advice(gate)}

IMPORTANT NOTICE
This summary is indicative only and is not a loan commitment, appraisal, legal opinion, tax advice, investment recommendation, or final credit decision.
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
        "Version": APP_VERSION,
    }

    sheets = {
        "Executive Summary": preview_df,
        "Assumptions": assumptions_df,
        "Sizing": sizing_df,
        "Uses": uses_df,
        "Sources": sources_df,
        "Covenants": covenant_df,
        "Rent Roll": st.session_state.rent_roll_df,
        "Rent Roll Metrics": pd.DataFrame(rent_roll_metrics.items(), columns=["Metric", "Value"]),
        "Amortization": amort_schedule_df,
        "Market Pulse": sovereign_market_pulse_df(inputs["rate"]),
        "BoC Rates": fetch_bank_of_canada_sovereign_rates(),
        "Property Intelligence": pd.DataFrame(get_property_intelligence(inputs["property_address"]).items(), columns=["Fact", "Value"]),
        "Sovereign Stack": sovereign_stack_df(),
        "Audit Trail": get_audit_df(),
    }

    excel_file = create_excel_workbook(sheets, metadata)
    pdf_file = create_pdf_summary(metadata, preview_df, covenant_df)
    deal_json = build_deal_json(
        {
            "version": APP_VERSION,
            "generated": generated_at,
            "inputs": inputs,
            "loan_amount": loan_amt,
            "binding_constraint": gate,
            "score": score,
            "classification": classification,
            "rent_roll": st.session_state.rent_roll_df.to_dict(orient="records"),
            "property_intelligence": get_property_intelligence(inputs["property_address"]),
        }
    )
    backup_zip = create_backup_zip(deal_json, excel_file, report_text)

    safe_sponsor = clean_filename(inputs["sponsor"])

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.download_button(
            "Download Text",
            report_text,
            file_name=f"Alenza_Summary_{safe_sponsor}.txt",
            mime="text/plain",
        )

    with c2:
        st.download_button(
            "Download Excel",
            excel_file,
            file_name=f"Alenza_Workbook_{safe_sponsor}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with c3:
        if pdf_file:
            st.download_button(
                "Download PDF",
                pdf_file,
                file_name=f"Alenza_Summary_{safe_sponsor}.pdf",
                mime="application/pdf",
            )
        else:
            st.warning("PDF export requires reportlab.")

    with c4:
        st.download_button(
            "Save JSON",
            deal_json,
            file_name=f"Alenza_Deal_{safe_sponsor}.json",
            mime="application/json",
        )

    with c5:
        st.download_button(
            "Backup ZIP",
            backup_zip,
            file_name=f"Alenza_Backup_{safe_sponsor}.zip",
            mime="application/zip",
        )

    email_body = f"""Alenza Deal Submission

Sponsor: {inputs['sponsor']}
Property: {inputs['property_address']}
Supportable Proceeds: {format_money(loan_amt)}
Binding Constraint: {gate}
Score: {score}/1000
Classification: {classification}

Generated: {generated_at}
"""
    st.link_button("Email Deal Snapshot", build_deal_email_link("Alenza Deal Submission", email_body))
    st.caption(f"Deal submissions route to: {get_deal_inbox_email()}")

    with st.expander("Deployment Notes"):
        st.code(
            """Folder name:
alenza-canada

Run:
streamlit run app.py

requirements.txt:
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

Optional Streamlit Secrets:
ALENZA_DEAL_INBOX_EMAIL = "resourcefulcapital@gmail.com"
""",
            language="text",
        )
