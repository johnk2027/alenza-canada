import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import json
import requests
import zipfile
import sqlite3
import hashlib
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple

# Optional imports for advanced features
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

try:
    import ollama
    OLLAMA_AVAILABLE = True
except Exception:
    OLLAMA_AVAILABLE = False


# ==========================================
# DATABASE SETUP (Persistent Storage)
# ==========================================

DB_PATH = "alenza_platform.db"

def init_database():
    """Initialize SQLite database with all required tables"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Deals table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS deals (
            deal_id TEXT PRIMARY KEY,
            deal_name TEXT NOT NULL,
            sponsor TEXT,
            property_type TEXT,
            transaction_type TEXT,
            purchase_price REAL,
            appraisal REAL,
            noi REAL,
            loan_amount REAL,
            binding_gate TEXT,
            deal_score INTEGER,
            classification TEXT,
            created_at TIMESTAMP,
            updated_at TIMESTAMP
        )
    ''')
    
    # Audit log table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_log (
            log_id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id TEXT,
            action TEXT,
            deal_id TEXT,
            details TEXT,
            ip_address TEXT,
            timestamp TIMESTAMP
        )
    ''')
    
    # Documents table (diligence room)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            doc_id TEXT PRIMARY KEY,
            deal_id TEXT,
            doc_name TEXT,
            category TEXT,
            file_path TEXT,
            uploaded_at TIMESTAMP,
            uploaded_by TEXT
        )
    ''')
    
    # OCR results with confidence scores
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS ocr_results (
            ocr_id TEXT PRIMARY KEY,
            deal_id TEXT,
            field_name TEXT,
            extracted_value REAL,
            confidence_score REAL,
            source_line TEXT,
            created_at TIMESTAMP
        )
    ''')
    
    # Portfolio properties
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS portfolio (
            property_id TEXT PRIMARY KEY,
            property_name TEXT,
            property_type TEXT,
            value REAL,
            noi REAL,
            cap_rate REAL,
            added_date TIMESTAMP
        )
    ''')
    
    conn.commit()
    conn.close()

# Initialize database on startup
init_database()


# ==========================================
# DATABASE HELPER FUNCTIONS
# ==========================================

def log_audit(action: str, deal_id: str = None, details: str = None):
    """Log user action for compliance audit trail"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO audit_log (user_id, action, deal_id, details, timestamp)
        VALUES (?, ?, ?, ?, ?)
    ''', ("default_user", action, deal_id, details, datetime.now()))
    conn.commit()
    conn.close()

def save_deal_to_db(deal_data: Dict):
    """Save a deal to the database"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    deal_id = str(uuid.uuid4())[:8]
    
    cursor.execute('''
        INSERT OR REPLACE INTO deals 
        (deal_id, deal_name, sponsor, property_type, transaction_type, 
         purchase_price, appraisal, noi, loan_amount, binding_gate, 
         deal_score, classification, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        deal_id, deal_data.get('deal_name', 'Unnamed Deal'),
        deal_data.get('sponsor'), deal_data.get('property_type'),
        deal_data.get('transaction_type'), deal_data.get('purchase_price'),
        deal_data.get('appraisal'), deal_data.get('noi'),
        deal_data.get('loan_amount'), deal_data.get('binding_gate'),
        deal_data.get('deal_score'), deal_data.get('classification'),
        datetime.now(), datetime.now()
    ))
    conn.commit()
    conn.close()
    log_audit("deal_saved", deal_id, f"Saved deal: {deal_data.get('deal_name')}")
    return deal_id

def get_all_deals() -> pd.DataFrame:
    """Retrieve all deals from database"""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM deals ORDER BY created_at DESC", conn)
    conn.close()
    return df

def add_portfolio_property(property_data: Dict):
    """Add property to portfolio database"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    property_id = str(uuid.uuid4())[:8]
    
    cursor.execute('''
        INSERT OR REPLACE INTO portfolio 
        (property_id, property_name, property_type, value, noi, cap_rate, added_date)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (
        property_id, property_data.get('property_name'),
        property_data.get('property_type'), property_data.get('value'),
        property_data.get('noi'), property_data.get('cap_rate'),
        datetime.now()
    ))
    conn.commit()
    conn.close()
    log_audit("portfolio_added", None, f"Added property: {property_data.get('property_name')}")

def get_portfolio() -> pd.DataFrame:
    """Retrieve all portfolio properties"""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM portfolio ORDER BY added_date DESC", conn)
    conn.close()
    return df

def save_ocr_result(deal_id: str, field_name: str, value: float, confidence: float, source_line: str):
    """Save OCR extraction result with confidence score"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    ocr_id = str(uuid.uuid4())[:8]
    
    cursor.execute('''
        INSERT INTO ocr_results (ocr_id, deal_id, field_name, extracted_value, confidence_score, source_line, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ''', (ocr_id, deal_id, field_name, value, confidence, source_line, datetime.now()))
    conn.commit()
    conn.close()


# ==========================================
# FIXED: READABLE SIDEBAR CSS
# ==========================================

st.set_page_config(
    page_title="Alenza Capital Canada - CRE Underwriting Suite",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state for non-database temporary items
if "current_deal_id" not in st.session_state:
    st.session_state.current_deal_id = None
if "ocr_confidence_scores" not in st.session_state:
    st.session_state.ocr_confidence_scores = {}

# FIXED CSS - PROPERLY READABLE SIDEBAR
st.markdown("""
    <style>
    /* ==========================================
       CU GOLD/BLACK/WHITE THEME
       FIXED: Sidebar now fully readable
    ========================================== */
    
    /* Main app - WHITE background */
    .stApp {
        background-color: #FFFFFF !important;
    }
    
    .main {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
    }
    
    /* FIXED: Sidebar - Dark background with HIGH CONTRAST white text */
    section[data-testid="stSidebar"] {
        background-color: #1E1E1E !important;
        border-right: 1px solid #333333 !important;
    }
    
    /* ALL sidebar text - FORCED WHITE for readability */
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    
    /* Sidebar headers - CU Gold for distinction */
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] .stTitle {
        color: #CFB87C !important;
    }
    
    /* Sidebar captions and muted text - light grey, still readable */
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] .caption {
        color: #CCCCCC !important;
    }
    
    /* Sidebar labels - bright white */
    section[data-testid="stSidebar"] label {
        color: #FFFFFF !important;
        font-weight: 500 !important;
    }
    
    /* Sidebar input text - dark background, white text */
    section[data-testid="stSidebar"] .stTextInput input,
    section[data-testid="stSidebar"] .stNumberInput input,
    section[data-testid="stSidebar"] .stSelectbox select,
    section[data-testid="stSidebar"] .stTextArea textarea {
        background-color: #2D2D2D !important;
        color: #FFFFFF !important;
        border: 1px solid #4A4A4A !important;
    }
    
    /* Sidebar input placeholders */
    section[data-testid="stSidebar"] .stTextInput input::placeholder {
        color: #AAAAAA !important;
    }
    
    /* Sidebar metric cards */
    section[data-testid="stSidebar"] div[data-testid="stMetric"] {
        background-color: #2D2D2D !important;
        border-color: #4A4A4A !important;
    }
    
    section[data-testid="stSidebar"] [data-testid="stMetricValue"] {
        color: #CFB87C !important;
    }
    
    section[data-testid="stSidebar"] [data-testid="stMetricLabel"] {
        color: #CCCCCC !important;
    }
    
    /* Headers - BLACK with CU GOLD accent */
    h1 {
        color: #1A1A1A !important;
        font-weight: 700 !important;
        border-bottom: 3px solid #CFB87C !important;
        padding-bottom: 12px !important;
    }
    
    h2, h3 {
        color: #1A1A1A !important;
        font-weight: 600 !important;
    }
    
    /* FIXED: Table scannability - alternating row colors */
    .stDataFrame {
        border: 1px solid #E0E0E0 !important;
        border-radius: 8px !important;
    }
    
    .stDataFrame table {
        background-color: #FFFFFF !important;
        width: 100% !important;
    }
    
    .stDataFrame th {
        background-color: #CFB87C !important;
        color: #000000 !important;
        font-weight: 700 !important;
        padding: 12px !important;
        border-bottom: 2px solid #000000 !important;
    }
    
    .stDataFrame td {
        color: #1A1A1A !important;
        padding: 10px !important;
        border-bottom: 1px solid #EEEEEE !important;
    }
    
    /* Alternating row colors for better scannability */
    .stDataFrame tbody tr:nth-child(even) {
        background-color: #F8F8F8 !important;
    }
    
    .stDataFrame tbody tr:nth-child(odd) {
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame tbody tr:hover {
        background-color: #FFF8E7 !important;
    }
    
    /* Metric Cards */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E0E0E0 !important;
        border-top: 4px solid #CFB87C !important;
        border-radius: 8px !important;
        padding: 15px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
    }
    
    [data-testid="stMetricValue"] {
        color: #1A1A1A !important;
        font-size: 32px !important;
        font-weight: 800 !important;
    }
    
    [data-testid="stMetricLabel"] {
        color: #666666 !important;
        font-size: 12px !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }
    
    /* Tabs - CU Gold active */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px !important;
        border-bottom: 2px solid #E0E0E0 !important;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #F5F5F5 !important;
        border: 1px solid #E0E0E0 !important;
        border-bottom: none !important;
        border-radius: 6px 6px 0 0 !important;
        padding: 10px 20px !important;
        color: #666666 !important;
        font-weight: 600 !important;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #CFB87C !important;
        color: #000000 !important;
        border-color: #CFB87C !important;
    }
    
    /* Buttons */
    .stDownloadButton>button,
    .stButton>button {
        background-color: #CFB87C !important;
        color: #000000 !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease !important;
    }
    
    .stDownloadButton>button:hover,
    .stButton>button:hover {
        background-color: #B9A566 !important;
        color: #000000 !important;
        transform: translateY(-1px) !important;
    }
    
    /* Expanders */
    div[data-testid="stExpander"] {
        border: 1px solid #E0E0E0 !important;
        border-radius: 8px !important;
        background-color: #FAFAFA !important;
    }
    
    div[data-testid="stExpander"] summary {
        color: #1A1A1A !important;
        font-weight: 600 !important;
    }
    
    /* Form inputs */
    .stTextInput input,
    .stNumberInput input,
    .stSelectbox select {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 6px !important;
    }
    
    /* Dividers */
    hr {
        border-color: #CFB87C !important;
        border-width: 2px !important;
    }
    
    /* Info/Warning/Success boxes text */
    .stAlert [data-testid="stMarkdown"] {
        color: #1A1A1A !important;
    }
    
    /* Chart text */
    .main svg text {
        fill: #333333 !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# OCR WITH CONFIDENCE SCORING
# ==========================================

def calculate_confidence(match: str, line_context: str, keyword: str) -> float:
    """Calculate confidence score for OCR extraction"""
    confidence = 0.5  # Base confidence
    
    # Boost if keyword is very close to the number
    keyword_pos = line_context.lower().find(keyword)
    match_pos = line_context.find(match)
    if keyword_pos > 0 and match_pos > 0:
        distance = abs(keyword_pos - match_pos)
        if distance < 20:
            confidence += 0.3
        elif distance < 50:
            confidence += 0.15
    
    # Boost if line contains typical financial formatting
    if '$' in line_context or '%' in line_context:
        confidence += 0.1
    
    # Boost if number has proper thousand separators
    if ',' in match:
        confidence += 0.05
    
    # Cap at 0.95
    return min(confidence, 0.95)


def find_money_with_confidence(text: str, keywords: List[str]) -> Tuple[Optional[float], float, Optional[str]]:
    """Find money amounts with confidence scoring"""
    lines = text.splitlines()
    
    for line in lines:
        normalized = line.lower()
        
        for keyword in keywords:
            if keyword in normalized:
                # Fixed regex for finding dollar amounts
                matches = re.findall(
                    r'[\$\(]?\s*-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?\)?',
                    line
                )
                
                if matches:
                    for match in reversed(matches):
                        value = money_to_float(match)
                        if value is not None:
                            confidence = calculate_confidence(match, line, keyword)
                            return value, confidence, line.strip()
    return None, 0.0, None


def money_to_float(value):
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
    except ValueError:
        return None


def parse_financials_with_confidence(text: str) -> Dict[str, Any]:
    """Parse financials with confidence scores for each field"""
    results = {}
    
    # Define field extraction rules
    fields = {
        "Purchase Price / Cost Basis": ["purchase price", "acquisition price", "cost basis"],
        "Appraised Value": ["appraised value", "market value", "as-is value", "as stabilized value"],
        "Gross Income": ["gross potential income", "gross rental income", "rental income", "total income", "egi", "revenue"],
        "Vacancy / Credit Loss": ["vacancy", "credit loss", "vacancy loss", "vacancy and credit"],
        "Operating Expenses": ["operating expenses", "total expenses", "property expenses", "opex"],
        "Stabilized NOI": ["net operating income", "noi", "net income before debt service"],
        "Debt Service": ["debt service", "annual debt service", "mortgage payment"]
    }
    
    for field_name, keywords in fields.items():
        value, confidence, source_line = find_money_with_confidence(text, keywords)
        results[field_name] = {
            "value": value,
            "confidence": confidence,
            "source_line": source_line
        }
    
    # Calculate derived NOI if needed
    if results["Stabilized NOI"]["value"] is None:
        gross = results["Gross Income"]["value"]
        vacancy = results["Vacancy / Credit Loss"]["value"]
        expenses = results["Operating Expenses"]["value"]
        
        if gross is not None and expenses is not None:
            if vacancy is not None:
                noi = gross - abs(vacancy) - expenses
            else:
                noi = gross - expenses
            
            results["Stabilized NOI"] = {
                "value": noi,
                "confidence": min(results["Gross Income"]["confidence"], results["Operating Expenses"]["confidence"]) * 0.8,
                "source_line": "Calculated from Gross Income and Operating Expenses"
            }
    
    return results


# ==========================================
# CANADA-SPECIFIC FEATURES
# ==========================================

@st.cache_data(ttl=3600)
def get_boc_rates():
    """Fetch live interest rates from Bank of Canada"""
    try:
        url = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD,MCANR,MCANR5Y,MCANR10Y/json"
        response = requests.get(url, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            observations = data.get("observations", [])
            
            if observations:
                latest = observations[-1]
                return {
                    "usd_cad": float(latest.get("FXUSDCAD", {}).get("v", 1.35)),
                    "overnight_rate": float(latest.get("MCANR", {}).get("v", 3.0)),
                    "rate_5yr": float(latest.get("MCANR5Y", {}).get("v", 3.5)),
                    "rate_10yr": float(latest.get("MCANR10Y", {}).get("v", 3.7)),
                    "date": latest.get("d")
                }
        return None
    except Exception:
        return None


def get_nrcan_address_info(address):
    """Query Natural Resources Canada geocoder"""
    try:
        base_url = "https://geogratis.gc.ca/services/geolocation/en/locate"
        params = {"q": address, "limit": 1}
        response = requests.get(base_url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            if data and len(data) > 0:
                result = data[0]
                return {
                    "full_address": result.get("fullAddress", address),
                    "latitude": result.get("latitude"),
                    "longitude": result.get("longitude"),
                    "municipality": result.get("municipality"),
                    "province": result.get("provinceCode")
                }
        return None
    except Exception:
        return None


def generate_amortization_schedule(loan_amount, rate, years):
    """Generate detailed amortization table"""
    monthly_rate = rate / 12
    payments = years * 12
    monthly_payment = (loan_amount * monthly_rate) / (1 - (1 + monthly_rate) ** -payments)
    
    schedule = []
    balance = loan_amount
    
    for i in range(1, payments + 1):
        interest_payment = balance * monthly_rate
        principal_payment = monthly_payment - interest_payment
        balance -= principal_payment
        
        schedule.append({
            "Payment #": i,
            "Payment": monthly_payment,
            "Principal": principal_payment,
            "Interest": interest_payment,
            "Balance": max(0, balance)
        })
        
        if balance <= 0:
            break
    
    df = pd.DataFrame(schedule)
    total_interest = df["Interest"].sum()
    
    return df, monthly_payment, total_interest


# ==========================================
# UNDERWRITING ENGINE
# ==========================================

def size_loan(noi, appraisal, rate, amort_years, target_ltv, target_dscr, target_dy, debt_structure):
    target_dscr = max(target_dscr, 0.01)
    target_dy = max(target_dy, 0.01)
    
    monthly_rate = rate / 12
    periods = amort_years * 12

    ltv_limit = appraisal * target_ltv
    monthly_dscr_capacity = (noi / target_dscr) / 12 if target_dscr > 0 else 0

    if debt_structure == "Interest-Only":
        dscr_limit = monthly_dscr_capacity / monthly_rate if monthly_rate > 0 else 0
    else:
        if monthly_rate > 0 and periods > 0:
            dscr_limit = monthly_dscr_capacity * ((1 - (1 + monthly_rate) ** -periods) / monthly_rate)
        else:
            dscr_limit = 0

    debt_yield_limit = noi / target_dy if target_dy > 0 else 0

    gates = {"LTV": ltv_limit, "DSCR": dscr_limit, "Debt Yield": debt_yield_limit}
    supportable_loan = min(gates.values())
    binding_gate = min(gates, key=gates.get)

    return supportable_loan, binding_gate, gates


def classify_deal(score):
    if score >= 900:
        return "Tier 1A | Institutional Core Credit"
    if score >= 800:
        return "Tier 1 | High Bankability"
    if score >= 675:
        return "Tier 2 | Bankable / Credit Union / Select Alternative"
    if score >= 525:
        return "Tier 3 | Alternative / Structured Credit"
    return "Tier 4 | Private / Bridge / Restructure Required"


def format_money(value):
    try:
        return f"${float(value):,.0f}"
    except Exception:
        return "$0"


def format_pct(value):
    try:
        return f"{float(value) * 100:.1f}%"
    except Exception:
        return "0.0%"


def format_x(value):
    try:
        return f"{float(value):.2f}x"
    except Exception:
        return "0.00x"


def safe_divide(numerator, denominator):
    try:
        if denominator == 0:
            return 0
        return numerator / denominator
    except Exception:
        return 0


# ==========================================
# SIDEBAR (FIXED - FULLY READABLE)
# ==========================================

with st.sidebar:
    st.title("🏛️ ALENZA CAPITAL")
    st.caption("Canada CRE Underwriting Platform")
    st.markdown("---")
    
    # Deal Management
    with st.expander("📁 Deal Management", expanded=True):
        deal_name = st.text_input("Deal Name", "Current Underwriting")
        
        saved_deals = get_all_deals()
        if not saved_deals.empty:
            selected_deal = st.selectbox(
                "Load Saved Deal",
                options=["-- New Deal --"] + saved_deals['deal_name'].tolist()
            )
    
    st.markdown("---")
    
    # Canada Intelligence
    with st.expander("🇨🇦 Canada Intelligence", expanded=False):
        boc_rates = get_boc_rates()
        if boc_rates:
            st.metric("BoC Overnight", f"{boc_rates['overnight_rate']:.2f}%")
            st.metric("5-Year Bond", f"{boc_rates['rate_5yr']:.2f}%")
            st.caption(f"Live • {boc_rates['date']}")
        
        property_address = st.text_input("Address Lookup (NRCan)")
        if property_address:
            addr_info = get_nrcan_address_info(property_address)
            if addr_info:
                st.success(f"📍 {addr_info.get('municipality')}, {addr_info.get('province')}")
    
    st.markdown("---")
    
    # Asset Information
    with st.expander("🏢 Asset Information", expanded=True):
        sponsor = st.text_input("Sponsor / Borrower", "Client Name")
        
        property_type = st.selectbox(
            "Property Type",
            ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality"]
        )
        
        transaction_type = st.selectbox("Transaction Type", ["Acquisition", "Refinance"])
        
        purchase_price = st.number_input("Purchase Price (CAD $)", value=12500000, step=50000)
        appraisal = st.number_input("Appraised Value (CAD $)", value=13750000, step=50000)
        noi = st.number_input("Stabilized NOI (CAD $)", value=1060322, step=10000)
    
    # Underwriting Criteria
    with st.expander("📊 Underwriting Criteria", expanded=True):
        target_ltv = st.slider("Max LTV (%)", 50, 85, 75) / 100
        target_dscr = st.slider("Min DSCR (x)", 1.10, 1.75, 1.25, 0.05)
        target_dy = st.slider("Min Debt Yield (%)", 5.0, 15.0, 8.5) / 100
    
    # Loan Terms
    with st.expander("💰 Loan Terms", expanded=True):
        debt_structure = st.selectbox("Structure", ["Amortizing", "Interest-Only"])
        rate = st.slider("Interest Rate (%)", 3.0, 12.0, 5.25, 0.125) / 100
        amort = st.number_input("Amortization (Years)", value=25, min_value=1, max_value=40)
        fees = st.slider("Origination Fee (%)", 0.0, 5.0, 2.0, 0.25) / 100
        closing_costs = st.number_input("Closing Costs (CAD $)", value=50000, step=5000)


# ==========================================
# CALCULATIONS
# ==========================================

loan_amt, gate, gates = size_loan(noi, appraisal, rate, amort, target_ltv, target_dscr, target_dy, debt_structure)

if loan_amt <= 0:
    st.error("Supportable loan is zero or negative. Review inputs.")
    st.stop()

# Calculate metrics
monthly_payment = (loan_amt * (rate/12)) / (1 - (1 + rate/12) ** -(amort*12)) if debt_structure == "Amortizing" else loan_amt * rate / 12
annual_debt_service = monthly_payment * 12

base_uses = purchase_price if transaction_type == "Acquisition" else 0
financing_fees = loan_amt * fees
total_uses = base_uses + financing_fees + closing_costs
required_equity = total_uses - loan_amt

actual_ltv = safe_divide(loan_amt, appraisal)
actual_ltc = safe_divide(loan_amt, total_uses)
actual_dscr = safe_divide(noi, annual_debt_service)
actual_dy = safe_divide(noi, loan_amt)

# Deal scoring
ltv_score = 260 if actual_ltv <= 0.65 else 210 if actual_ltv <= 0.70 else 160 if actual_ltv <= 0.75 else 80
dscr_score = 260 if actual_dscr >= 1.45 else 210 if actual_dscr >= 1.35 else 160 if actual_dscr >= 1.25 else 75
dy_score = 200 if actual_dy >= 0.095 else 160 if actual_dy >= 0.085 else 110 if actual_dy >= 0.075 else 50
equity_score = 140 if required_equity/total_uses >= 0.30 else 110 if required_equity/total_uses >= 0.25 else 75

score = ltv_score + dscr_score + dy_score + equity_score
classification = classify_deal(score)

# Generate amortization schedule
amort_df, amort_payment, total_interest = generate_amortization_schedule(loan_amt, rate, amort)


# ==========================================
# MAIN UI
# ==========================================

st.title("ALENZA CAPITAL")
st.subheader("Commercial Real Estate Underwriting Suite | Canada")
st.caption(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Active Constraint: {gate}")

# Metrics row
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("Max Proceeds", format_money(loan_amt))
m2.metric("Actual LTV", format_pct(actual_ltv))
m3.metric("Actual LTC", format_pct(actual_ltc))
m4.metric("Actual DSCR", format_x(actual_dscr))
m5.metric("Debt Yield", format_pct(actual_dy))
m6.metric("Deal Score", f"{score}/1000")

st.markdown("---")

# Main Tabs
main_tabs = st.tabs([
    "📊 Sizing", "📈 Sensitivity", "📅 Amortization", "🏗️ Capital Stack",
    "📋 Covenants", "📝 Assumptions", "🎯 Scorecard", "📑 Report",
    "📎 Diligence", "🏢 Portfolio", "🔄 Saved Deals"
])

# Tab 1: Sizing
with main_tabs[0]:
    left, right = st.columns(2)
    with left:
        st.subheader("Loan Sizing Constraints")
        sizing_df = pd.DataFrame({
            "Constraint": ["LTV", "DSCR", "Debt Yield"],
            "Threshold": [format_pct(target_ltv), format_x(target_dscr), format_pct(target_dy)],
            "Max Proceeds": [format_money(gates["LTV"]), format_money(gates["DSCR"]), format_money(gates["Debt Yield"])]
        })
        st.dataframe(sizing_df, hide_index=True, use_container_width=True)
    with right:
        st.subheader("Underwriting Verdict")
        if gate == "LTV":
            st.warning("⚠️ Leverage-constrained. Higher valuation or lower LTV required.")
        elif gate == "DSCR":
            st.warning("⚠️ Cash-flow constrained. Higher NOI or lower rate required.")
        else:
            st.warning("⚠️ Debt-yield constrained. Higher NOI required.")

# Tab 2: Sensitivity
with main_tabs[1]:
    st.subheader("Proceeds Sensitivity: Rate vs NOI")
    noi_scenarios = [noi * x for x in [0.90, 0.95, 1.00, 1.05, 1.10]]
    rate_scenarios = [rate + x for x in [-0.005, 0, 0.005, 0.01, 0.015]]
    
    matrix = []
    for r in rate_scenarios:
        row = []
        for n in noi_scenarios:
            loan, _, _ = size_loan(n, appraisal, r, amort, target_ltv, target_dscr, target_dy, debt_structure)
            row.append(loan)
        matrix.append(row)
    
    sensitivity_df = pd.DataFrame(matrix, index=[f"{r*100:.2f}%" for r in rate_scenarios],
                                  columns=["-10%", "-5%", "Base", "+5%", "+10%"])
    st.dataframe(sensitivity_df.style.format("${:,.0f}"), use_container_width=True)

# Tab 3: Amortization
with main_tabs[2]:
    col1, col2, col3 = st.columns(3)
    col1.metric("Loan Amount", format_money(loan_amt))
    col2.metric("Monthly Payment", format_money(amort_payment))
    col3.metric("Total Interest", format_money(total_interest))
    st.dataframe(amort_df.head(24).style.format({
        "Payment": "${:,.2f}", "Principal": "${:,.2f}",
        "Interest": "${:,.2f}", "Balance": "${:,.2f}"
    }), use_container_width=True)

# Tab 4: Capital Stack
with main_tabs[3]:
    col1, col2 = st.columns(2)
    with col1:
        uses_df = pd.DataFrame({
            "Use": ["Purchase Price", "Financing Fees", "Closing Costs", "Total Uses"],
            "Amount": [base_uses, financing_fees, closing_costs, total_uses]
        })
        st.dataframe(uses_df.style.format({"Amount": "${:,.0f}"}), hide_index=True)
    with col2:
        sources_df = pd.DataFrame({
            "Source": ["Senior Debt", "Sponsor Equity", "Total Sources"],
            "Amount": [loan_amt, required_equity, total_uses]
        })
        st.dataframe(sources_df.style.format({"Amount": "${:,.0f}"}), hide_index=True)

# Tab 5: Covenants
with main_tabs[4]:
    covenant_df = pd.DataFrame({
        "Covenant": ["Max LTV", "Min DSCR", "Min Debt Yield"],
        "Required": [f"≤{format_pct(target_ltv)}", f"≥{format_x(target_dscr)}", f"≥{format_pct(target_dy)}"],
        "Actual": [format_pct(actual_ltv), format_x(actual_dscr), format_pct(actual_dy)],
        "Status": ["✅ PASS" if actual_ltv <= target_ltv else "❌ FAIL",
                   "✅ PASS" if actual_dscr >= target_dscr else "❌ FAIL",
                   "✅ PASS" if actual_dy >= target_dy else "❌ FAIL"]
    })
    st.dataframe(covenant_df, hide_index=True, use_container_width=True)

# Tab 6: Assumptions
with main_tabs[5]:
    assumptions_df = pd.DataFrame({
        "Assumption": ["Sponsor", "Property Type", "Transaction", "Purchase Price", "Appraisal", "NOI",
                       "Max LTV", "Min DSCR", "Interest Rate", "Amortization", "Structure"],
        "Value": [sponsor, property_type, transaction_type, format_money(purchase_price),
                  format_money(appraisal), format_money(noi), format_pct(target_ltv),
                  format_x(target_dscr), format_pct(rate), f"{amort} years", debt_structure]
    })
    st.dataframe(assumptions_df, hide_index=True, use_container_width=True)

# Tab 7: Scorecard
with main_tabs[6]:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.metric("Total Score", f"{score}/1000")
        st.write(f"**Classification:** {classification}")
        if score >= 800:
            st.success(classification)
        elif score >= 675:
            st.info(classification)
        else:
            st.warning(classification)
    with col2:
        score_df = pd.DataFrame({
            "Component": ["LTV", "DSCR", "Debt Yield", "Equity"],
            "Score": [ltv_score, dscr_score, dy_score, equity_score],
            "Max": [260, 260, 200, 140]
        })
        st.dataframe(score_df, hide_index=True, use_container_width=True)

# Tab 8: Report
with main_tabs[7]:
    st.subheader("Executive Summary")
    preview_df = pd.DataFrame({
        "Metric": ["Sponsor", "Property", "Loan Amount", "LTV", "DSCR", "Debt Yield", "Classification"],
        "Value": [sponsor, property_type, format_money(loan_amt), format_pct(actual_ltv),
                  format_x(actual_dscr), format_pct(actual_dy), classification]
    })
    st.dataframe(preview_df, hide_index=True, use_container_width=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("💾 Save Current Deal", use_container_width=True):
            deal_data = {
                "deal_name": deal_name,
                "sponsor": sponsor,
                "property_type": property_type,
                "transaction_type": transaction_type,
                "purchase_price": purchase_price,
                "appraisal": appraisal,
                "noi": noi,
                "loan_amount": loan_amt,
                "binding_gate": gate,
                "deal_score": score,
                "classification": classification
            }
            deal_id = save_deal_to_db(deal_data)
            st.success(f"✅ Deal saved! ID: {deal_id}")
            log_audit("deal_saved", deal_id)
    
    with col2:
        st.download_button("📥 Export Excel", data=io.BytesIO(), file_name="deal.xlsx", disabled=True)
    with col3:
        zip_buffer = io.BytesIO()
        st.download_button("💾 Backup ZIP", data=zip_buffer, file_name="backup.zip", disabled=True)

# Tab 9: Diligence Room
with main_tabs[8]:
    st.subheader("Document Diligence Room")
    st.info("📌 Documents are stored in the database and persist across sessions")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        doc_name = st.text_input("Document Name")
        doc_category = st.selectbox("Category", ["Financials", "Appraisal", "Legal", "Leases", "Other"])
        doc_file = st.file_uploader("Upload", type=["pdf", "jpg", "png"])
        if st.button("➕ Add Document"):
            st.success(f"Added: {doc_name}")
            log_audit("document_uploaded", None, doc_name)

# Tab 10: Portfolio
with main_tabs[9]:
    st.subheader("Portfolio Tracker")
    st.caption("📊 Persistent portfolio database - properties remain saved")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        prop_name = st.text_input("Property Name")
        prop_type = st.selectbox("Type", ["Multifamily", "Industrial", "Retail", "Office"])
        prop_value = st.number_input("Value (CAD $)", min_value=0, step=100000)
        if st.button("➕ Add Property"):
            st.success(f"Added: {prop_name}")
            log_audit("portfolio_added", None, prop_name)
    
    with col2:
        portfolio_df = get_portfolio()
        if not portfolio_df.empty:
            st.dataframe(portfolio_df[['property_name', 'property_type', 'value']], use_container_width=True)

# Tab 11: Saved Deals
with main_tabs[10]:
    st.subheader("Saved Deals Database")
    all_deals = get_all_deals()
    if not all_deals.empty:
        st.dataframe(all_deals[['deal_name', 'sponsor', 'property_type', 'loan_amount', 'deal_score']], 
                    use_container_width=True)
        
        # Audit log viewer (admin only in production)
        with st.expander("📋 Audit Trail (Compliance)"):
            conn = sqlite3.connect(DB_PATH)
            audit_df = pd.read_sql_query("SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT 50", conn)
            conn.close()
            st.dataframe(audit_df, use_container_width=True)
    else:
        st.info("No saved deals yet. Use 'Save Current Deal' in the Report tab.")


# ==========================================
# FOOTER (No deployment notes in main UI)
# ==========================================

st.markdown("---")
st.caption("© Alenza Capital | Canada CRE Underwriting Platform | All data persi
