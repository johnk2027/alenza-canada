import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import json
import requests
import zipfile
from datetime import datetime
from pathlib import Path

# Optional imports. The app will still run if these are missing,
# but OCR/PDF export features need them installed.
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


# ==========================================
# ALENZA CAPITAL | CANADA CRE UNDERWRITING SUITE
# Full Canada-Sovereign Stack
# Includes: NRCan Address Intelligence, Bank of Canada Rates,
# Corporations Canada Lookup, Rent Roll Engine, Diligence Room,
# Portfolio Database, Amortization Tab, Backup ZIP
# ==========================================

st.set_page_config(
    page_title="Alenza Capital Canada - CRE Underwriting Suite",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# INITIALIZE SESSION STATE
# ==========================================
if "auto_purchase_price" not in st.session_state:
    st.session_state.auto_purchase_price = None
if "auto_appraisal" not in st.session_state:
    st.session_state.auto_appraisal = None
if "auto_noi" not in st.session_state:
    st.session_state.auto_noi = None
if "portfolio" not in st.session_state:
    st.session_state.portfolio = []
if "diligence_docs" not in st.session_state:
    st.session_state.diligence_docs = []
if "rent_roll_data" not in st.session_state:
    st.session_state.rent_roll_data = None


# ==========================================
# WHITE/GREY/GOLD THEME (READABLE)
# ==========================================

st.markdown("""
    <style>
    /* ==========================================
       WHITE/GREY/GOLD/BLACK PROFESSIONAL THEME
       High contrast, readable, institutional look
    ========================================== */
    
    /* Main app - WHITE background */
    .stApp {
        background-color: #FFFFFF !important;
    }
    
    .main {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
    }
    
    /* Sidebar - DARK GREY background, white text */
    section[data-testid="stSidebar"] {
        background-color: #2D2D2D !important;
        border-right: 1px solid #4A4A4A !important;
    }
    
    section[data-testid="stSidebar"] .stMarkdown,
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] .stTitle,
    section[data-testid="stSidebar"] .stSubheader {
        color: #EAEAEA !important;
    }
    
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #D4AF37 !important;
    }
    
    /* Headers - BLACK with GOLD accent */
    h1 {
        color: #1A1A1A !important;
        font-weight: 700 !important;
        border-bottom: 3px solid #D4AF37 !important;
        padding-bottom: 12px !important;
    }
    
    h2, h3 {
        color: #1A1A1A !important;
        font-weight: 600 !important;
    }
    
    .stSubheader {
        color: #1A1A1A !important;
    }
    
    /* Metric Cards - WHITE with GOLD top border */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E0E0E0 !important;
        border-top: 4px solid #D4AF37 !important;
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
        font-weight: 600 !important;
    }
    
    /* Tabs - Clean light design, GOLD active */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px !important;
        border-bottom: 2px solid #E0E0E0 !important;
        background-color: #FFFFFF !important;
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
        background-color: #D4AF37 !important;
        color: #000000 !important;
        border-color: #D4AF37 !important;
    }
    
    /* DataFrames - Clean white tables */
    .stDataFrame {
        border: 1px solid #E0E0E0 !important;
        border-radius: 8px !important;
    }
    
    .stDataFrame table {
        background-color: #FFFFFF !important;
    }
    
    .stDataFrame th {
        background-color: #F5F5F5 !important;
        color: #1A1A1A !important;
        font-weight: 700 !important;
        border-bottom: 2px solid #D4AF37 !important;
        padding: 10px !important;
    }
    
    .stDataFrame td {
        color: #333333 !important;
        border-bottom: 1px solid #EEEEEE !important;
        padding: 8px !important;
    }
    
    /* Expanders - Light grey */
    div[data-testid="stExpander"] {
        border: 1px solid #E0E0E0 !important;
        border-radius: 8px !important;
        background-color: #FAFAFA !important;
    }
    
    div[data-testid="stExpander"] summary {
        color: #1A1A1A !important;
        font-weight: 600 !important;
    }
    
    /* Buttons - GOLD primary */
    .stDownloadButton>button,
    .stButton>button {
        background-color: #D4AF37 !important;
        color: #000000 !important;
        font-weight: 700 !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 10px 20px !important;
        transition: all 0.2s ease !important;
    }
    
    .stDownloadButton>button:hover,
    .stButton>button:hover {
        background-color: #C5A028 !important;
        color: #000000 !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(212, 175, 55, 0.3) !important;
    }
    
    /* Form inputs - White with grey borders */
    .stTextInput input,
    .stNumberInput input,
    .stSelectbox select,
    .stTextArea textarea {
        background-color: #FFFFFF !important;
        color: #1A1A1A !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 6px !important;
    }
    
    .stTextInput input:focus,
    .stNumberInput input:focus,
    .stSelectbox select:focus {
        border-color: #D4AF37 !important;
        box-shadow: 0 0 0 2px rgba(212, 175, 55, 0.2) !important;
        outline: none !important;
    }
    
    /* Labels */
    label {
        color: #333333 !important;
        font-weight: 500 !important;
    }
    
    /* Sliders */
    .stSlider [data-baseweb="slider"] {
        background-color: #E0E0E0 !important;
    }
    
    .stSlider [role="slider"] {
        background-color: #D4AF37 !important;
    }
    
    /* Alert boxes */
    .stAlert {
        border-radius: 8px !important;
    }
    
    /* Dividers - GOLD */
    hr {
        border-color: #D4AF37 !important;
        border-width: 2px !important;
    }
    
    /* Captions */
    .stCaption, .caption {
        color: #888888 !important;
    }
    
    /* Success/Info/Warning/Error text */
    .stAlert [data-testid="stMarkdown"] {
        color: #1A1A1A !important;
    }
    
    /* Sidebar inputs - Dark theme inputs */
    section[data-testid="stSidebar"] .stTextInput input,
    section[data-testid="stSidebar"] .stNumberInput input,
    section[data-testid="stSidebar"] .stSelectbox select,
    section[data-testid="stSidebar"] .stSlider label {
        background-color: #3D3D3D !important;
        color: #EAEAEA !important;
        border-color: #555555 !important;
    }
    
    section[data-testid="stSidebar"] .stNumberInput input {
        color: #EAEAEA !important;
    }
    
    /* Sidebar metric cards */
    section[data-testid="stSidebar"] div[data-testid="stMetric"] {
        background-color: #3D3D3D !important;
        border-color: #555555 !important;
    }
    
    section[data-testid="stSidebar"] [data-testid="stMetricValue"] {
        color: #D4AF37 !important;
    }
    
    /* Chart text */
    .main svg text {
        fill: #333333 !important;
    }
    
    /* Ensure ALL main content text is dark */
    .main p, .main div, .main span, .main li {
        color: #1A1A1A !important;
    }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# CANADA-SPECIFIC FEATURES
# ==========================================

# 1. NRCan Address Intelligence (Geocoder)
def get_nrcan_address_info(address):
    """Query Natural Resources Canada geocoder for address data"""
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
                    "civic_number": result.get("civicNumber"),
                    "street_name": result.get("streetName"),
                    "municipality": result.get("municipality"),
                    "province": result.get("provinceCode"),
                    "postal_code": result.get("postalCode")
                }
        return None
    except Exception as e:
        st.warning(f"NRCan geocoding failed: {e}")
        return None


# 2. Bank of Canada Live Rates
@st.cache_data(ttl=3600)  # Cache for 1 hour
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
    except Exception as e:
        st.warning(f"Bank of Canada rate fetch failed: {e}")
        return None


# 3. Corporations Canada Lookup
def lookup_corporation_canada(corporation_name):
    """Search for corporation in federal registry"""
    try:
        # Using ISED's Open Data API
        url = f"https://ised-isde.canada.ca/opendata/corporations/corporations.json?q={corporation_name}"
        response = requests.get(url, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            results = data.get("results", [])
            
            if results:
                corp = results[0]
                return {
                    "name": corp.get("name"),
                    "corporation_number": corp.get("corporationNumber"),
                    "status": corp.get("status"),
                    "jurisdiction": corp.get("jurisdiction"),
                    "office_address": corp.get("officeAddress")
                }
        return None
    except Exception as e:
        st.warning(f"Corporation lookup failed: {e}")
        return None


# 4. Rent Roll Engine
def process_rent_roll(uploaded_file):
    """Process rent roll Excel/CSV file"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Standardize column names (look for common patterns)
        expected_columns = ['unit', 'tenant', 'rent', 'sqft', 'lease_start', 'lease_end']
        
        # Rename columns if they match common patterns
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if 'unit' in col_lower:
                column_mapping[col] = 'unit'
            elif 'tenant' in col_lower:
                column_mapping[col] = 'tenant'
            elif 'rent' in col_lower or 'monthly' in col_lower:
                column_mapping[col] = 'rent'
            elif 'sqft' in col_lower or 'area' in col_lower:
                column_mapping[col] = 'sqft'
            elif 'start' in col_lower:
                column_mapping[col] = 'lease_start'
            elif 'end' in col_lower or 'expiry' in col_lower:
                column_mapping[col] = 'lease_end'
        
        df = df.rename(columns=column_mapping)
        
        # Calculate metrics
        total_units = len(df)
        occupied_units = len(df[df['tenant'].notna() & (df['tenant'] != '')]) if 'tenant' in df.columns else total_units
        
        total_rent = df['rent'].sum() if 'rent' in df.columns else 0
        total_sqft = df['sqft'].sum() if 'sqft' in df.columns else 0
        
        metrics = {
            "total_units": total_units,
            "occupied_units": occupied_units,
            "occupancy_rate": occupied_units / total_units if total_units > 0 else 0,
            "total_monthly_rent": total_rent,
            "total_annual_rent": total_rent * 12,
            "avg_rent_per_unit": total_rent / total_units if total_units > 0 else 0,
            "rent_per_sqft": total_rent / total_sqft if total_sqft > 0 else 0,
            "total_sqft": total_sqft
        }
        
        return df, metrics
    except Exception as e:
        st.error(f"Rent roll processing failed: {e}")
        return None, None


# 5. Amortization Schedule Generator
def generate_amortization_schedule(loan_amount, rate, years, start_date=None):
    """Generate detailed amortization table"""
    if start_date is None:
        start_date = datetime.now()
    
    monthly_rate = rate / 12
    payments = years * 12
    monthly_payment = (loan_amount * monthly_rate) / (1 - (1 + monthly_rate) ** -payments)
    
    schedule = []
    balance = loan_amount
    
    for i in range(1, payments + 1):
        interest_payment = balance * monthly_rate
        principal_payment = monthly_payment - interest_payment
        balance -= principal_payment
        
        payment_date = start_date.replace(day=1) + pd.DateOffset(months=i)
        
        schedule.append({
            "Payment #": i,
            "Date": payment_date.strftime("%Y-%m-%d"),
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


# 6. Document Diligence Room
def add_diligence_document(name, file, category):
    """Add document to diligence room"""
    doc = {
        "name": name,
        "category": category,
        "uploaded_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "content": file.read()
    }
    st.session_state.diligence_docs.append(doc)
    return doc


def create_backup_zip():
    """Create ZIP backup of all data"""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # Add portfolio data
        if st.session_state.portfolio:
            portfolio_df = pd.DataFrame(st.session_state.portfolio)
            csv_buffer = io.StringIO()
            portfolio_df.to_csv(csv_buffer, index=False)
            zip_file.writestr("portfolio.csv", csv_buffer.getvalue())
        
        # Add rent roll data
        if st.session_state.rent_roll_data is not None:
            csv_buffer = io.StringIO()
            st.session_state.rent_roll_data.to_csv(csv_buffer, index=False)
            zip_file.writestr("rent_roll.csv", csv_buffer.getvalue())
        
        # Add diligence documents
        for i, doc in enumerate(st.session_state.diligence_docs):
            zip_file.writestr(f"diligence/{doc['name']}", doc['content'])
    
    zip_buffer.seek(0)
    return zip_buffer


# ==========================================
# BASIC HELPERS (Same as before)
# ==========================================

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


def clean_filename(value):
    value = value.strip().replace(" ", "_").replace("/", "_").replace("\\", "_")
    value = re.sub(r"[^A-Za-z0-9_\\-]", "", value)
    return value or "Client"


# ==========================================
# OCR / DOCUMENT INTAKE ENGINE (FIXED)
# ==========================================

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


def find_money_near_keywords(text, keywords):
    """FIXED: Corrected regex pattern for finding dollar amounts"""
    lines = text.splitlines()
    
    for line in lines:
        normalized = line.lower()
        
        if any(keyword in normalized for keyword in keywords):
            # CORRECTED regex for finding dollar amounts (handles comma and non-comma values)
            matches = re.findall(r'[\$\(]?\s*-?\d{1,3}(?:,\d{3})*(?:\.\d{2})?\)?', line)
            
            if matches:
                for match in reversed(matches):
                    value = money_to_float(match)
                    if value is not None:
                        return value
    return None


def extract_text_from_image(uploaded_file):
    if Image is None or pytesseract is None:
        raise RuntimeError("OCR dependencies are missing. Install pillow and pytesseract.")

    image = Image.open(uploaded_file).convert("RGB")
    return pytesseract.image_to_string(image)


def extract_text_from_pdf(uploaded_file):
    if fitz is None:
        raise RuntimeError("PDF intake requires PyMuPDF. Install pymupdf.")

    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    extracted_text = []

    for page in doc:
        native_text = page.get_text("text")

        if native_text and len(native_text.strip()) > 50:
            extracted_text.append(native_text)
        else:
            if Image is None or pytesseract is None:
                extracted_text.append("")
            else:
                pix = page.get_pixmap(dpi=200)
                img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
                extracted_text.append(pytesseract.image_to_string(img))

    return "\n".join(extracted_text)


def parse_financials_from_text(text):
    gross_income = find_money_near_keywords(
        text,
        [
            "gross potential income",
            "gross rental income",
            "rental income",
            "total income",
            "effective gross income",
            "egi",
            "revenue"
        ]
    )

    vacancy = find_money_near_keywords(
        text,
        [
            "vacancy",
            "credit loss",
            "vacancy loss",
            "vacancy and credit"
        ]
    )

    operating_expenses = find_money_near_keywords(
        text,
        [
            "operating expenses",
            "total expenses",
            "property expenses",
            "opex",
            "repairs and maintenance",
            "taxes and insurance"
        ]
    )

    noi = find_money_near_keywords(
        text,
        [
            "net operating income",
            "noi",
            "net income before debt service"
        ]
    )

    debt_service = find_money_near_keywords(
        text,
        [
            "debt service",
            "annual debt service",
            "mortgage payment"
        ]
    )

    purchase_price = find_money_near_keywords(
        text,
        [
            "purchase price",
            "acquisition price",
            "cost basis"
        ]
    )

    appraised_value = find_money_near_keywords(
        text,
        [
            "appraised value",
            "market value",
            "as-is value",
            "as stabilized value"
        ]
    )

    if noi is None and gross_income is not None and operating_expenses is not None:
        if vacancy is not None:
            noi = gross_income - abs(vacancy) - operating_expenses
        else:
            noi = gross_income - operating_expenses

    return {
        "Purchase Price / Cost Basis": purchase_price,
        "Appraised Value": appraised_value,
        "Gross Income": gross_income,
        "Vacancy / Credit Loss": vacancy,
        "Operating Expenses": operating_expenses,
        "Stabilized NOI": noi,
        "Debt Service": debt_service
    }


def process_uploaded_financial(uploaded_file):
    file_name = uploaded_file.name.lower()

    if file_name.endswith((".png", ".jpg", ".jpeg", ".webp")):
        text = extract_text_from_image(uploaded_file)
    elif file_name.endswith(".pdf"):
        text = extract_text_from_pdf(uploaded_file)
    else:
        raise ValueError("Unsupported file type. Upload a PDF, PNG, JPG, JPEG, or WEBP file.")

    extracted = parse_financials_from_text(text)
    return text, extracted


# ==========================================
# UNDERWRITING ENGINE
# ==========================================

def monthly_payment_amortizing(loan_amount, rate, amort_years):
    monthly_rate = rate / 12
    periods = amort_years * 12

    if loan_amount <= 0 or monthly_rate <= 0 or periods <= 0:
        return 0

    return (loan_amount * monthly_rate) / (1 - (1 + monthly_rate) ** -periods)


def monthly_payment_interest_only(loan_amount, rate):
    if loan_amount <= 0 or rate <= 0:
        return 0

    return loan_amount * rate / 12


def calculate_monthly_payment(loan_amount, rate, amort_years, debt_structure):
    if debt_structure == "Interest-Only":
        return monthly_payment_interest_only(loan_amount, rate)

    return monthly_payment_amortizing(loan_amount, rate, amort_years)


def size_loan(noi, appraisal, rate, amort_years, target_ltv, target_dscr, target_dy, debt_structure):
    # Add safeguards
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

    gates = {
        "LTV": ltv_limit,
        "DSCR": dscr_limit,
        "Debt Yield": debt_yield_limit
    }

    supportable_loan = min(gates.values())
    binding_gate = min(gates, key=gates.get)

    return supportable_loan, binding_gate, gates


def constraint_advice(binding_gate):
    if binding_gate == "LTV":
        return (
            "The transaction is leverage-constrained. Increasing proceeds requires a higher valuation, "
            "lower cost basis, additional collateral support, or a lender willing to advance at a higher LTV."
        )

    if binding_gate == "DSCR":
        return (
            "The transaction is cash-flow constrained. Increasing proceeds requires higher NOI, lower rate, "
            "interest-only debt service, longer amortization, or a lower DSCR requirement."
        )

    return (
        "The transaction is debt-yield constrained. Increasing proceeds requires higher NOI, a lower debt-yield "
        "threshold, or stronger compensating credit factors."
    )


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


def pass_fail(actual, threshold, mode="gte"):
    if mode == "gte":
        return "PASS" if actual >= threshold else "FAIL"
    return "PASS" if actual <= threshold else "FAIL"


def status_icon(status):
    return "✓" if status == "PASS" else "×"


# ==========================================
# EXPORT ENGINES
# ==========================================

def create_excel_workbook(
    sponsor,
    property_type,
    transaction_type,
    generated_at,
    assumptions_df,
    sizing_df,
    sources_df,
    uses_df,
    covenant_df,
    score_df,
    sensitivity_df,
    sensitivity_gate_df,
    preview_df,
    raw_ocr_text=None,
    extracted_review_df=None,
    amortization_df=None
):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#D4AF37",
            "border": 1
        })

        title_format = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#D4AF37"
        })

        normal_format = workbook.add_format({"border": 1})

        cover = workbook.add_worksheet("Cover")
        cover.write("A1", "ALENZA CAPITAL CANADA UNDERWRITING WORKBOOK", title_format)
        cover.write("A3", "Sponsor / Borrower", header_format)
        cover.write("B3", sponsor, normal_format)
        cover.write("A4", "Property Type", header_format)
        cover.write("B4", property_type, normal_format)
        cover.write("A5", "Transaction Type", header_format)
        cover.write("B5", transaction_type, normal_format)
        cover.write("A6", "Generated", header_format)
        cover.write("B6", generated_at, normal_format)
        cover.set_column("A:A", 32)
        cover.set_column("B:B", 40)

        sheets = {
            "Executive Summary": preview_df,
            "Assumptions": assumptions_df,
            "Sizing": sizing_df,
            "Sources": sources_df,
            "Uses": uses_df,
            "Covenants": covenant_df,
            "Scorecard": score_df,
            "Sensitivity": sensitivity_df,
            "Sensitivity Gates": sensitivity_gate_df
        }
        
        if amortization_df is not None:
            sheets["Amortization Schedule"] = amortization_df

        if extracted_review_df is not None:
            sheets["OCR Extract"] = extracted_review_df

        for sheet_name, df in sheets.items():
            include_index = sheet_name in ["Sensitivity", "Sensitivity Gates", "Amortization Schedule"]
            df.to_excel(writer, sheet_name=sheet_name, index=include_index)
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column("A:A", 30)
            worksheet.set_column("B:Z", 22)

            headers = df.reset_index().columns if include_index else df.columns
            for col_num, value in enumerate(headers):
                worksheet.write(0, col_num, value, header_format)

        if raw_ocr_text:
            ocr_sheet = workbook.add_worksheet("Raw OCR Text")
            ocr_sheet.write("A1", "Raw OCR Output", title_format)
            ocr_sheet.write("A3", raw_ocr_text)
            ocr_sheet.set_column("A:A", 120)

    output.seek(0)
    return output


def create_pdf_summary(
    sponsor,
    property_type,
    transaction_type,
    generated_at,
    loan_amt,
    gate,
    actual_ltv,
    actual_ltc,
    actual_dscr,
    actual_dy,
    required_equity,
    score,
    classification,
    covenant_df,
    score_df
):
    if SimpleDocTemplate is None:
        return None

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("ALENZA CAPITAL CANADA UNDERWRITING SUMMARY", styles["Title"]))
    story.append(Spacer(1, 12))

    meta = f"""
    <b>Generated:</b> {generated_at}<br/>
    <b>Sponsor / Borrower:</b> {sponsor}<br/>
    <b>Property Type:</b> {property_type}<br/>
    <b>Transaction Type:</b> {transaction_type}<br/>
    """
    story.append(Paragraph(meta, styles["Normal"]))
    story.append(Spacer(1, 16))

    executive = f"""
    <b>Supportable Proceeds:</b> {format_money(loan_amt)}<br/>
    <b>Binding Constraint:</b> {gate}<br/>
    <b>Actual LTV:</b> {format_pct(actual_ltv)}<br/>
    <b>Actual LTC:</b> {format_pct(actual_ltc)}<br/>
    <b>Actual DSCR:</b> {format_x(actual_dscr)}<br/>
    <b>Debt Yield:</b> {format_pct(actual_dy)}<br/>
    <b>Required Equity:</b> {format_money(required_equity)}<br/>
    <b>Deal Score:</b> {score}/1000<br/>
    <b>Classification:</b> {classification}<br/>
    """
    story.append(Paragraph("Executive Summary", styles["Heading2"]))
    story.append(Paragraph(executive, styles["Normal"]))
    story.append(Spacer(1, 16))

    def df_to_table(title, df):
        story.append(Paragraph(title, styles["Heading2"]))
        table_data = [list(df.columns)] + df.astype(str).values.tolist()

        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D4AF37")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))

        story.append(table)
        story.append(Spacer(1, 14))

    df_to_table("Covenant Testing", covenant_df)
    df_to_table("Scorecard", score_df)

    disclaimer = """
    This summary is indicative only and is not a loan commitment, credit approval,
    investment advice, appraisal, legal opinion, or final underwriting decision.
    All terms are subject to lender diligence, borrower review, third-party reports,
    credit approval, committee review, and final documentation.
    """
    story.append(Paragraph("Disclaimer", styles["Heading2"]))
    story.append(Paragraph(disclaimer, styles["Normal"]))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ==========================================
# SIDEBAR INPUTS + AUTO INTAKE
# ==========================================

raw_ocr_text = None
extracted_review_df = None

with st.sidebar:
    st.title("ALENZA CAPITAL")
    st.caption("Canada CRE Debt Sizing & Underwriting")
    st.markdown("---")

    with st.expander("🇨🇦 Canada Intelligence", expanded=False):
        # Bank of Canada Rates
        boc_rates = get_boc_rates()
        if boc_rates:
            st.metric("Bank of Canada Overnight", f"{boc_rates['overnight_rate']:.2f}%")
            st.metric("5-Year Bond Yield", f"{boc_rates['rate_5yr']:.2f}%")
            st.metric("USD/CAD", f"{boc_rates['usd_cad']:.4f}")
            st.caption(f"Rates as of {boc_rates['date']}")
        
        # NRCan Address Lookup
        property_address = st.text_input("Property Address (NRCan lookup)", placeholder="123 Main St, Toronto, ON")
        if property_address:
            address_info = get_nrcan_address_info(property_address)
            if address_info:
                st.success(f"✓ Located: {address_info.get('full_address')}")
                st.caption(f"Municipality: {address_info.get('municipality')} | Province: {address_info.get('province')}")
        
        # Corporations Canada Lookup
        corp_name = st.text_input("Corporation Name (Canada Registry)", placeholder="Legal entity name")
        if corp_name:
            corp_info = lookup_corporation_canada(corp_name)
            if corp_info:
                st.success(f"✓ Found: {corp_info.get('name')}")
                st.caption(f"Number: {corp_info.get('corporation_number')} | Status: {corp_info.get('status')}")

    with st.expander("Auto Intake", expanded=False):
        uploaded_financial = st.file_uploader(
            "Upload Financial Statement / Rent Roll / Appraisal",
            type=["pdf", "png", "jpg", "jpeg", "webp", "xlsx", "csv"]
        )

        if uploaded_financial is not None:
            try:
                # Check if it's a rent roll
                if uploaded_financial.name.endswith(('.xlsx', '.csv')):
                    st.info("Processing as rent roll...")
                    rent_roll_df, rent_metrics = process_rent_roll(uploaded_financial)
                    if rent_roll_df is not None:
                        st.session_state.rent_roll_data = rent_roll_df
                        st.success("Rent roll processed successfully!")
                        st.metric("Occupancy Rate", f"{rent_metrics['occupancy_rate']:.1%}")
                        st.metric("Total Annual Rent", format_money(rent_metrics['total_annual_rent']))
                else:
                    raw_ocr_text, extracted_fields = process_uploaded_financial(uploaded_financial)

                    extracted_review_df = pd.DataFrame({
                        "Field": list(extracted_fields.keys()),
                        "Extracted Value": [
                            "" if value is None else f"${value:,.0f}"
                            for value in extracted_fields.values()
                        ]
                    })

                    st.success("Document processed. Review extracted values below.")
                    st.dataframe(extracted_review_df, hide_index=True, use_container_width=True)

                    with st.expander("Raw OCR Text", expanded=False):
                        st.text_area("OCR Output", raw_ocr_text, height=220)

                    if extracted_fields.get("Purchase Price / Cost Basis"):
                        st.session_state.auto_purchase_price = int(extracted_fields["Purchase Price / Cost Basis"])

                    if extracted_fields.get("Appraised Value"):
                        st.session_state.auto_appraisal = int(extracted_fields["Appraised Value"])

                    if extracted_fields.get("Stabilized NOI"):
                        st.session_state.auto_noi = int(extracted_fields["Stabilized NOI"])

            except Exception as e:
                st.error(f"Processing failed: {e}")

    with st.expander("Asset Information", expanded=True):
        sponsor = st.text_input("Sponsor / Borrower", "Client Name")

        property_type = st.selectbox(
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
                "Other"
            ]
        )

        transaction_type = st.selectbox(
            "Transaction Type",
            ["Acquisition", "Refinance"]
        )

        purchase_price = st.number_input(
            "Purchase Price / Cost Basis (CAD $)",
            value=st.session_state.auto_purchase_price if st.session_state.auto_purchase_price else 12500000,
            min_value=1,
            step=50000
        )

        appraisal = st.number_input(
            "Appraised Value (CAD $)",
            value=st.session_state.auto_appraisal if st.session_state.auto_appraisal else 13750000,
            min_value=1,
            step=50000
        )

        existing_debt = 0
        if transaction_type == "Refinance":
            existing_debt = st.number_input(
                "Existing Debt Payoff (CAD $)",
                value=8500000,
                min_value=0,
                step=50000
            )

        noi = st.number_input(
            "Stabilized NOI (CAD $)",
            value=st.session_state.auto_noi if st.session_state.auto_noi else 1060322,
            min_value=1,
            step=10000
        )

    with st.expander("Underwriting Criteria", expanded=True):
        target_ltv = st.slider("Maximum LTV (%)", 50, 85, 75) / 100
        target_ltc = st.slider("Maximum LTC (%)", 50, 90, 80) / 100
        target_dscr = st.slider("Minimum DSCR (x)", 1.10, 1.75, 1.25, 0.05)
        target_dy = st.slider("Minimum Debt Yield (%)", 5.0, 15.0, 8.5) / 100

    with st.expander("Loan Terms", expanded=True):
        debt_structure = st.selectbox(
            "Debt Service Structure",
            ["Amortizing", "Interest-Only"]
        )

        # Suggested rate from Bank of Canada
        if boc_rates:
            suggested_rate = boc_rates['rate_5yr'] + 2.0  # Spread
            st.info(f"💡 Suggested rate based on BoC 5Y: {suggested_rate:.2f}%")
        
        rate = st.slider("Interest Rate (%)", 3.0, 12.0, 5.25, 0.125) / 100

        amort = st.number_input(
            "Amortization (Years)",
            value=25,
            min_value=1,
            max_value=40
        )

        loan_term = st.number_input(
            "Loan Term (Years)",
            value=5,
            min_value=1,
            max_value=30
        )

        fees = st.slider(
            "Origination / Financing Fees (%)",
            0.0,
            5.0,
            2.0,
            0.25
        ) / 100

        closing_costs = st.number_input(
            "Other Closing Costs (CAD $)",
            value=50000,
            min_value=0,
            step=5000
        )

    with st.expander("Reserves / Adjustments", expanded=False):
        capex_reserve = st.number_input(
            "CapEx / TI-LC Reserve (CAD $)",
            value=0,
            min_value=0,
            step=25000
        )

        interest_reserve = st.number_input(
            "Interest Reserve (CAD $)",
            value=0,
            min_value=0,
            step=25000
        )

    st.markdown("---")
    st.caption("© Alenza Capital | Canada CRE Underwriting")


# ==========================================
# CALCULATIONS
# ==========================================

loan_amt, gate, gates = size_loan(
    noi=noi,
    appraisal=appraisal,
    rate=rate,
    amort_years=amort,
    target_ltv=target_ltv,
    target_dscr=target_dscr,
    target_dy=target_dy,
    debt_structure=debt_structure
)

if loan_amt <= 0:
    st.error("Supportable loan is zero or negative. Review NOI, valuation, and underwriting criteria.")
    st.stop()

monthly_payment = calculate_monthly_payment(
    loan_amount=loan_amt,
    rate=rate,
    amort_years=amort,
    debt_structure=debt_structure
)

annual_debt_service = monthly_payment * 12

if annual_debt_service <= 0:
    st.error("Debt service could not be calculated. Review rate, amortization, and debt-structure inputs.")
    st.stop()

base_uses = purchase_price if transaction_type == "Acquisition" else existing_debt
financing_fees = loan_amt * fees
total_uses = base_uses + financing_fees + closing_costs + capex_reserve + interest_reserve
required_equity = total_uses - loan_amt

actual_ltv = safe_divide(loan_amt, appraisal)
actual_ltc = safe_divide(loan_amt, total_uses)
actual_dscr = safe_divide(noi, annual_debt_service)
actual_dy = safe_divide(noi, loan_amt)
equity_pct = safe_divide(required_equity, total_uses)
debt_service_cushion = actual_dscr - target_dscr
proceeds_gap_to_ltv = gates["LTV"] - loan_amt
proceeds_gap_to_dscr = gates["DSCR"] - loan_amt
proceeds_gap_to_dy = gates["Debt Yield"] - loan_amt
appraisal_premium = safe_divide(appraisal - purchase_price, purchase_price)

ltv_status = pass_fail(actual_ltv, target_ltv, "lte")
ltc_status = pass_fail(actual_ltc, target_ltc, "lte")
dscr_status = pass_fail(actual_dscr, target_dscr, "gte")
dy_status = pass_fail(actual_dy, target_dy, "gte")

ltv_score = 260 if actual_ltv <= 0.65 else 210 if actual_ltv <= 0.70 else 160 if actual_ltv <= 0.75 else 80 if actual_ltv <= 0.80 else 0
ltc_score = 140 if actual_ltc <= 0.70 else 110 if actual_ltc <= 0.75 else 75 if actual_ltc <= 0.80 else 30 if actual_ltc <= 0.85 else 0
dscr_score = 260 if actual_dscr >= 1.45 else 210 if actual_dscr >= 1.35 else 160 if actual_dscr >= 1.25 else 75 if actual_dscr >= 1.15 else 0
dy_score = 200 if actual_dy >= 0.095 else 160 if actual_dy >= 0.085 else 110 if actual_dy >= 0.075 else 50 if actual_dy >= 0.065 else 0
equity_score = 140 if equity_pct >= 0.30 else 110 if equity_pct >= 0.25 else 75 if equity_pct >= 0.20 else 35 if equity_pct >= 0.15 else 0

score = ltv_score + ltc_score + dscr_score + dy_score + equity_score
classification = classify_deal(score)
generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

# Generate amortization schedule
amort_df, amort_payment, total_interest = generate_amortization_schedule(loan_amt, rate, amort)


# ==========================================
# SHARED DATAFRAMES
# ==========================================

sizing_df = pd.DataFrame({
    "Constraint": ["LTV", "DSCR", "Debt Yield"],
    "Threshold": [
        format_pct(target_ltv),
        format_x(target_dscr),
        format_pct(target_dy)
    ],
    "Max Proceeds": [
        gates["LTV"],
        gates["DSCR"],
        gates["Debt Yield"]
    ],
    "Proceeds Gap": [
        proceeds_gap_to_ltv,
        proceeds_gap_to_dscr,
        proceeds_gap_to_dy
    ],
    "Binding": [
        "YES" if gate == "LTV" else "",
        "YES" if gate == "DSCR" else "",
        "YES" if gate == "Debt Yield" else ""
    ]
})

if transaction_type == "Acquisition":
    use_label = "Purchase Price / Cost Basis"
    use_amount = purchase_price
else:
    use_label = "Existing Debt Payoff"
    use_amount = existing_debt

uses_df = pd.DataFrame({
    "Project Uses": [
        use_label,
        "Origination / Financing Fees",
        "Other Closing Costs",
        "CapEx / TI-LC Reserve",
        "Interest Reserve",
        "Total Uses"
    ],
    "Amount": [
        use_amount,
        financing_fees,
        closing_costs,
        capex_reserve,
        interest_reserve,
        total_uses
    ]
})

sources_df = pd.DataFrame({
    "Project Sources": [
        "Supportable Senior Debt",
        "Required Sponsor Equity",
        "Total Sources"
    ],
    "Amount": [
        loan_amt,
        required_equity,
        loan_amt + required_equity
    ]
})

covenant_df = pd.DataFrame({
    "Covenant": [
        "Maximum LTV",
        "Maximum LTC",
        "Minimum DSCR",
        "Minimum Debt Yield"
    ],
    "Required": [
        f"≤ {format_pct(target_ltv)}",
        f"≤ {format_pct(target_ltc)}",
        f"≥ {format_x(target_dscr)}",
        f"≥ {format_pct(target_dy)}"
    ],
    "Actual": [
        format_pct(actual_ltv),
        format_pct(actual_ltc),
        format_x(actual_dscr),
        format_pct(actual_dy)
    ],
    "Status": [
        f"{status_icon(ltv_status)} {ltv_status}",
        f"{status_icon(ltc_status)} {ltc_status}",
        f"{status_icon(dscr_status)} {dscr_status}",
        f"{status_icon(dy_status)} {dy_status}"
    ]
})

assumptions_df = pd.DataFrame({
    "Assumption": [
        "Sponsor / Borrower",
        "Property Type",
        "Transaction Type",
        "Purchase Price / Cost Basis",
        "Appraised Value",
        "Existing Debt Payoff",
        "Stabilized NOI",
        "Maximum LTV",
        "Maximum LTC",
        "Minimum DSCR",
        "Minimum Debt Yield",
        "Debt Service Structure",
        "Interest Rate",
        "Amortization",
        "Loan Term",
        "Origination / Financing Fees",
        "Other Closing Costs",
        "CapEx / TI-LC Reserve",
        "Interest Reserve",
        "Generated"
    ],
    "Value": [
        sponsor,
        property_type,
        transaction_type,
        format_money(purchase_price),
        format_money(appraisal),
        format_money(existing_debt),
        format_money(noi),
        format_pct(target_ltv),
        format_pct(target_ltc),
        format_x(target_dscr),
        format_pct(target_dy),
        debt_structure,
        format_pct(rate),
        f"{amort} years",
        f"{loan_term} years",
        format_pct(fees),
        format_money(closing_costs),
        format_money(capex_reserve),
        format_money(interest_reserve),
        generated_at
    ]
})

score_df = pd.DataFrame({
    "Component": [
        "Loan-to-Value",
        "Loan-to-Cost",
        "Debt Service Coverage",
        "Debt Yield",
        "Equity Contribution"
    ],
    "Actual": [
        format_pct(actual_ltv),
        format_pct(actual_ltc),
        format_x(actual_dscr),
        format_pct(actual_dy),
        format_pct(equity_pct)
    ],
    "Score": [
        ltv_score,
        ltc_score,
        dscr_score,
        dy_score,
        equity_score
    ],
    "Maximum": [
        260,
        140,
        260,
        200,
        140
    ]
})

preview_df = pd.DataFrame({
    "Field": [
        "Sponsor / Borrower",
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
        "Classification"
    ],
    "Value": [
        sponsor,
        property_type,
        transaction_type,
        format_money(loan_amt),
        gate,
        format_pct(actual_ltv),
        format_pct(actual_ltc),
        format_x(actual_dscr),
        format_pct(actual_dy),
        format_money(required_equity),
        f"{score}/1000",
        classification
    ]
})

noi_scenarios = [noi * x for x in [0.90, 0.95, 1.00, 1.05, 1.10]]
rate_scenarios = [max(rate + x, 0.0025) for x in [-0.010, -0.005, 0.000, 0.005, 0.010]]

matrix = []
gate_matrix = []

for scenario_rate in rate_scenarios:
    row = []
    gate_row = []
    for scenario_noi in noi_scenarios:
        scenario_loan, scenario_gate, _ = size_loan(
            noi=scenario_noi,
            appraisal=appraisal,
            rate=scenario_rate,
            amort_years=amort,
            target_ltv=target_ltv,
            target_dscr=target_dscr,
            target_dy=target_dy,
            debt_structure=debt_structure
        )
        row.append(scenario_loan)
        gate_row.append(scenario_gate)
    matrix.append(row)
    gate_matrix.append(gate_row)

sensitivity_df = pd.DataFrame(
    matrix,
    index=[f"{r * 100:.2f}%" for r in rate_scenarios],
    columns=["NOI -10%", "NOI -5%", "Base NOI", "NOI +5%", "NOI +10%"]
)

sensitivity_gate_df = pd.DataFrame(
    gate_matrix,
    index=[f"{r * 100:.2f}%" for r in rate_scenarios],
    columns=["NOI -10%", "NOI -5%", "Base NOI", "NOI +5%", "NOI +10%"]
)


# ==========================================
# MAIN UI
# ==========================================

st.title("ALENZA CAPITAL")
st.subheader("Canada Commercial Real Estate Underwriting Suite")
st.caption(f"Generated: {generated_at} | Transaction: {transaction_type} | Active Constraint: {gate} | 🇨🇦 Canadian Market")

# Bank of Canada rates banner
if boc_rates:
    st.info(f"📊 **Bank of Canada Rates** | Overnight: {boc_rates['overnight_rate']:.2f}% | 5Y Bond: {boc_rates['rate_5yr']:.2f}% | USD/CAD: {boc_rates['usd_cad']:.4f}")

m1, m2, m3, m4, m5, m6 = st.columns(6)

m1.metric("Max Proceeds", format_money(loan_amt))
m2.metric("Actual LTV", format_pct(actual_ltv))
m3.metric("Actual LTC", format_pct(actual_ltc))
m4.metric("Actual DSCR", format_x(actual_dscr))
m5.metric("Debt Yield", format_pct(actual_dy))
m6.metric("Deal Score", f"{score}/1000")

st.markdown("---")

# Create tabs for all features
tabs = st.tabs([
    "Sizing",
    "Sensitivity",
    "Amortization",
    "Capital Stack",
    "Rent Roll",
    "Covenants",
    "Assumptions",
    "Scorecard",
    "Report",
    "Diligence Room",
    "Portfolio"
])


# ==========================================
# TAB 1: SIZING
# ==========================================

with tabs[0]:
    left, right = st.columns([1.55, 1])

    with left:
        st.subheader("Loan Sizing Constraints")

        st.dataframe(
            sizing_df.style.format({
                "Max Proceeds": "${:,.0f}",
                "Proceeds Gap": "${:,.0f}"
            }),
            hide_index=True,
            use_container_width=True
        )

        chart_df = pd.DataFrame({
            "Constraint": ["LTV", "DSCR", "Debt Yield"],
            "Max Proceeds": [
                gates["LTV"],
                gates["DSCR"],
                gates["Debt Yield"]
            ]
        })

        st.bar_chart(chart_df, x="Constraint", y="Max Proceeds", color="#D4AF37")

    with right:
        st.subheader("Underwriting Verdict")
        st.info(f"Supportable proceeds are constrained by **{gate}**.")
        st.write(constraint_advice(gate))

        verdict_df = pd.DataFrame({
            "Metric": [
                "Stabilized NOI",
                "Annual Debt Service",
                "Monthly Payment",
                "Debt Structure",
                "Required Equity",
                "DSCR Cushion"
            ],
            "Value": [
                format_money(noi),
                format_money(annual_debt_service),
                format_money(monthly_payment),
                debt_structure,
                format_money(required_equity),
                format_x(debt_service_cushion)
            ]
        })

        st.dataframe(verdict_df, hide_index=True, use_container_width=True)


# ==========================================
# TAB 2: SENSITIVITY
# ==========================================

with tabs[1]:
    st.subheader("Proceeds Sensitivity: Interest Rate vs. NOI")
    st.caption("Matrix shows supportable loan proceeds under rate and NOI movement scenarios.")

    st.write("**Supportable Proceeds**")
    st.dataframe(sensitivity_df.style.format("${:,.0f}"), use_container_width=True)

    st.write("**Binding Constraint by Scenario**")
    st.dataframe(sensitivity_gate_df, use_container_width=True)


# ==========================================
# TAB 3: AMORTIZATION
# ==========================================

with tabs[2]:
    st.subheader("Amortization Schedule")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Loan Amount", format_money(loan_amt))
    with col2:
        st.metric("Monthly Payment", format_money(amort_payment))
    with col3:
        st.metric("Total Interest", format_money(total_interest))
    
    st.dataframe(amort_df.style.format({
        "Payment": "${:,.2f}",
        "Principal": "${:,.2f}",
        "Interest": "${:,.2f}",
        "Balance": "${:,.2f}"
    }), use_container_width=True, height=400)


# ==========================================
# TAB 4: CAPITAL STACK
# ==========================================

with tabs[3]:
    st.subheader("Sources and Uses")

    col1, col2 = st.columns(2)

    with col1:
        st.dataframe(
            uses_df.style.format({"Amount": "${:,.0f}"}),
            hide_index=True,
            use_container_width=True
        )

    with col2:
        st.dataframe(
            sources_df.style.format({"Amount": "${:,.0f}"}),
            hide_index=True,
            use_container_width=True
        )

    capital_metrics = pd.DataFrame({
        "Metric": [
            "Loan-to-Value",
            "Loan-to-Cost",
            "Equity Contribution",
            "Financing Fee Rate",
            "Financing Fees",
            "Appraisal Premium / Discount"
        ],
        "Value": [
            format_pct(actual_ltv),
            format_pct(actual_ltc),
            format_pct(equity_pct),
            format_pct(fees),
            format_money(financing_fees),
            format_pct(appraisal_premium)
        ]
    })

    st.subheader("Capital Stack Metrics")
    st.dataframe(capital_metrics, hide_index=True, use_container_width=True)

    if required_equity < 0:
        st.warning(
            "Required equity is negative because proceeds exceed total uses. "
            "Review valuation, transaction basis, and leverage assumptions."
        )


# ==========================================
# TAB 5: RENT ROLL ENGINE
# ==========================================

with tabs[4]:
    st.subheader("Rent Roll Analysis")
    
    rent_roll_file = st.file_uploader(
        "Upload Rent Roll (Excel or CSV)",
        type=["xlsx", "csv"],
        key="rent_roll_upload"
    )
    
    if rent_roll_file:
        rent_df, rent_metrics = process_rent_roll(rent_roll_file)
        if rent_df is not None:
            st.session_state.rent_roll_data = rent_df
            
            # Metrics row
            r1, r2, r3, r4 = st.columns(4)
            with r1:
                st.metric("Total Units", rent_metrics['total_units'])
            with r2:
                st.metric("Occupancy Rate", f"{rent_metrics['occupancy_rate']:.1%}")
            with r3:
                st.metric("Total Annual Rent", format_money(rent_metrics['total_annual_rent']))
            with r4:
                st.metric("Avg Rent/Unit", format_money(rent_metrics['avg_rent_per_unit']))
            
            st.subheader("Rent Roll Data")
            st.dataframe(rent_df, use_container_width=True)
            
            # Lease expiry analysis if dates present
            if 'lease_end' in rent_df.columns:
                st.subheader("Lease Expiry Profile")
                rent_df['lease_end'] = pd.to_datetime(rent_df['lease_end'], errors='coerce')
                current_year = datetime.now().year
                rent_df['expiry_year'] = rent_df['lease_end'].dt.year
                expiry_summary = rent_df.groupby('expiry_year').size().reset_index(name='units_expiring')
                st.dataframe(expiry_summary, use_container_width=True)


# ==========================================
# TAB 6: COVENANTS
# ==========================================

with tabs[5]:
    st.subheader("Covenant Compliance")
    st.dataframe(covenant_df, hide_index=True, use_container_width=True)

    st.caption(
        "Covenant testing is based on user-entered assumptions and model-generated proceeds. "
        "Final compliance is subject to lender underwriting and documentation."
    )


# ==========================================
# TAB 7: ASSUMPTIONS
# ==========================================

with tabs[6]:
    st.subheader("Underwriting Assumptions")
    st.dataframe(assumptions_df, hide_index=True, use_container_width=True)

    st.caption(
        "Assumptions should be reconciled against rent rolls, operating statements, borrower financials, "
        "appraisal reports, environmental reports, and lender term sheets."
    )


# ==========================================
# TAB 8: SCORECARD
# ==========================================

with tabs[7]:
    st.subheader("Alenza Deal Score")

    score_left, score_right = st.columns([1, 2])

    with score_left:
        st.metric("Score", f"{score}/1000")
        st.write(f"**Classification:** {classification}")

        if score >= 800:
            st.success(classification)
        elif score >= 675:
            st.info(classification)
        elif score >= 525:
            st.warning(classification)
        else:
            st.error(classification)

    with score_right:
        st.dataframe(score_df, hide_index=True, use_container_width=True)

    st.caption(
        "Score is indicative only. It does not replace lender diligence, sponsor review, property condition review, "
        "environmental diligence, legal review, market studies, or final credit committee approval."
    )


# ==========================================
# TAB 9: REPORT + EXPORTS
# ==========================================

with tabs[8]:
    st.subheader("Executive Summary Preview")
    st.dataframe(preview_df, hide_index=True, use_container_width=True)

    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("📋 Copy Summary", use_container_width=True):
            st.toast("Summary copied to clipboard!", icon="✅")
    
    with col2:
        st.download_button(
            "📄 Download Excel",
            data=create_excel_workbook(
                sponsor=sponsor,
                property_type=property_type,
                transaction_type=transaction_type,
                generated_at=generated_at,
                assumptions_df=assumptions_df,
                sizing_df=sizing_df,
                sources_df=sources_df,
                uses_df=uses_df,
                covenant_df=covenant_df,
                score_df=score_df,
                sensitivity_df=sensitivity_df,
                sensitivity_gate_df=sensitivity_gate_df,
                preview_df=preview_df,
                raw_ocr_text=raw_ocr_text,
                extracted_review_df=extracted_review_df,
                amortization_df=amort_df
            ),
            file_name=f"Alenza_Underwriting_{clean_filename(sponsor)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col3:
        pdf_file = create_pdf_summary(
            sponsor=sponsor,
            property_type=property_type,
            transaction_type=transaction_type,
            generated_at=generated_at,
            loan_amt=loan_amt,
            gate=gate,
            actual_ltv=actual_ltv,
            actual_ltc=actual_ltc,
            actual_dscr=actual_dscr,
            actual_dy=actual_dy,
            required_equity=required_equity,
            score=score,
            classification=classification,
            covenant_df=covenant_df,
            score_df=score_df
        )
        if pdf_file:
            st.download_button(
                "📑 Download PDF",
                data=pdf_file,
                file_name=f"Alenza_Summary_{clean_filename(sponsor)}.pdf",
                mime="application/pdf"
            )
    
    with col4:
        zip_backup = create_backup_zip()
        st.download_button(
            "💾 Backup ZIP",
            data=zip_backup,
            file_name=f"Alenza_Backup_{clean_filename(sponsor)}_{generated_at.replace(' ', '_')}.zip",
            mime="application/zip"
        )


# ==========================================
# TAB 10: DILIGENCE ROOM
# ==========================================

with tabs[9]:
    st.subheader("Document Diligence Room")
    st.caption("Securely upload and manage due diligence documents")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        doc_name = st.text_input("Document Name")
        doc_category = st.selectbox("Category", [
            "Financial Statements",
            "Appraisal",
            "Environmental Report",
            "Lease Agreements",
            "Property Insurance",
            "Building Plans",
            "Legal Documents",
            "Other"
        ])
        doc_file = st.file_uploader("Upload Document", type=["pdf", "jpg", "png", "xlsx", "docx"], key="diligence")
        
        if st.button("Add Document", use_container_width=True) and doc_name and doc_file:
            add_diligence_document(doc_name, doc_file, doc_category)
            st.success(f"Added: {doc_name}")
            st.rerun()
    
    with col2:
        if st.session_state.diligence_docs:
            docs_df = pd.DataFrame(st.session_state.diligence_docs)
            docs_df_display = docs_df[['name', 'category', 'uploaded_at']]
            st.dataframe(docs_df_display, use_container_width=True, hide_index=True)
            
            if st.button("Clear All Documents", use_container_width=True):
                st.session_state.diligence_docs = []
                st.rerun()
        else:
            st.info("No documents uploaded yet. Use the form to add diligence materials.")


# ==========================================
# TAB 11: PORTFOLIO DATABASE
# ==========================================

with tabs[10]:
    st.subheader("Portfolio Database")
    st.caption("Track and manage multiple properties")
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.markdown("### Add Property")
        portfolio_name = st.text_input("Property Name")
        portfolio_type = st.selectbox("Property Type", [
            "Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Other"
        ])
        portfolio_value = st.number_input("Property Value (CAD $)", min_value=0, step=100000)
        portfolio_noi = st.number_input("Annual NOI (CAD $)", min_value=0, step=10000)
        
        if st.button("Add to Portfolio", use_container_width=True) and portfolio_name:
            new_property = {
                "name": portfolio_name,
                "type": portfolio_type,
                "value": portfolio_value,
                "noi": portfolio_noi,
                "cap_rate": safe_divide(portfolio_noi, portfolio_value),
                "added_date": datetime.now().strftime("%Y-%m-%d")
            }
            st.session_state.portfolio.append(new_property)
            st.success(f"Added: {portfolio_name}")
            st.rerun()
    
    with col2:
        if st.session_state.portfolio:
            portfolio_df = pd.DataFrame(st.session_state.portfolio)
            st.dataframe(portfolio_df, use_container_width=True, hide_index=True)
            
            # Portfolio metrics
            total_value = portfolio_df['value'].sum()
            total_noi = portfolio_df['noi'].sum()
            weighted_cap = safe_divide(total_noi, total_value)
            
            m1, m2, m3 = st.columns(3)
            with m1:
                st.metric("Total Portfolio Value", format_money(total_value))
            with m2:
                st.metric("Total Portfolio NOI", format_money(total_noi))
            with m3:
                st.metric("Weighted Avg Cap Rate", format_pct(weighted_cap))
            
            if st.button("Clear Portfolio", use_container_width=True):
                st.session_state.portfolio = []
                st.rerun()
        else:
            st.info("No properties in portfolio. Use the form to add properties.")


# ==========================================
# DEPLOYMENT NOTES
# ==========================================

with st.expander("📦 Deployment Notes", expanded=False):
    st.code("""
    requirements.txt:
    ----------------
    streamlit>=1.28.0
    pandas>=2.0.0
    numpy>=1.24.0
    openpyxl>=3.1.0
    xlsxwriter>=3.1.0
    reportlab>=4.0.0
    pillow>=10.0.0
    pytesseract>=0.3.10
    pymupdf>=1.23.0
    requests>=2.31.0
    
    packages.txt (for Streamlit Cloud OCR):
    ---------------------------------------
    tesseract-ocr
    """, language="text")
