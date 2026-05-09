import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

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
# ALENZA CAPITAL | CRE UNDERWRITING SUITE
# Single-File Institutional Build
# Includes OCR Intake, Excel Export, PDF Export
# ==========================================

st.set_page_config(
    page_title="Alenza Capital Underwriting Suite",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ==========================================
# ARGUS-STYLE INSTITUTIONAL UI
# ==========================================

st.markdown("""
    <style>
    :root {
        --bg-main: #05080F;
        --bg-panel: #0F172A;
        --bg-panel-alt: #111827;
        --border: #1E293B;
        --text-main: #F3F4F6;
        --text-muted: #9CA3AF;
        --accent: #1D4ED8;
        --accent-soft: #2563EB;
        --success: #16A34A;
        --warning: #D97706;
        --danger: #DC2626;
    }

    .main {
        background-color: var(--bg-main);
        color: var(--text-main);
        font-family: "Helvetica Neue", Arial, sans-serif;
    }

    section[data-testid="stSidebar"] {
        background-color: var(--bg-panel) !important;
        border-right: 1px solid var(--border);
    }

    h1, h2, h3 {
        letter-spacing: -0.02em;
    }

    [data-testid="stMetricValue"] {
        font-size: 27px !important;
        font-weight: 800 !important;
        color: var(--accent-soft) !important;
    }

    [data-testid="stMetricLabel"] {
        font-size: 12px !important;
        text-transform: uppercase;
        letter-spacing: 1.4px;
        color: var(--text-muted) !important;
    }

    div[data-testid="stMetric"] {
        background-color: var(--bg-panel);
        padding: 19px;
        border-radius: 6px;
        border: 1px solid var(--border);
        box-shadow: 0 3px 14px rgba(0, 0, 0, 0.35);
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 5px;
        border-bottom: 1px solid var(--border);
    }

    .stTabs [data-baseweb="tab"] {
        background-color: var(--bg-panel);
        border: 1px solid var(--border);
        border-bottom: none;
        border-radius: 4px 4px 0 0;
        padding: 10px 18px;
        color: var(--text-muted);
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        font-size: 12px;
    }

    .stTabs [aria-selected="true"] {
        background-color: var(--accent) !important;
        color: #FFFFFF !important;
        border-color: var(--accent) !important;
    }

    .stDataFrame {
        border: 1px solid var(--border);
        border-radius: 6px;
    }

    .stDownloadButton>button {
        width: 100%;
        background-color: var(--accent);
        color: white;
        font-weight: 700;
        border-radius: 6px;
        border: none;
        padding: 13px;
        text-transform: uppercase;
        letter-spacing: 0.06em;
    }

    .stDownloadButton>button:hover {
        background-color: var(--accent-soft);
        color: white;
    }

    div[data-testid="stExpander"] {
        border: 1px solid var(--border);
        border-radius: 6px;
        background-color: var(--bg-panel);
    }

    hr {
        border-color: var(--border);
    }
    </style>
    """, unsafe_allow_html=True)


# ==========================================
# BASIC HELPERS
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
# OCR / DOCUMENT INTAKE ENGINE
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
    lines = text.splitlines()

    for line in lines:
        normalized = line.lower()

        if any(keyword in normalized for keyword in keywords):
            matches = re.findall(r"\\(?\\$?\\s*-?\\d[\\d,]*\\.?\\d*\\)?", line)

            if matches:
                value = money_to_float(matches[-1])
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
    extracted_review_df=None
):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1D4ED8",
            "border": 1
        })

        title_format = workbook.add_format({
            "bold": True,
            "font_size": 16,
            "font_color": "#1D4ED8"
        })

        normal_format = workbook.add_format({"border": 1})

        cover = workbook.add_worksheet("Cover")
        cover.write("A1", "ALENZA CAPITAL UNDERWRITING WORKBOOK", title_format)
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

        if extracted_review_df is not None:
            sheets["OCR Extract"] = extracted_review_df

        for sheet_name, df in sheets.items():
            include_index = sheet_name in ["Sensitivity", "Sensitivity Gates"]
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

    story.append(Paragraph("ALENZA CAPITAL UNDERWRITING SUMMARY", styles["Title"]))
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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1D4ED8")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
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
    st.caption("Institutional CRE Debt Sizing")
    st.markdown("---")

    with st.expander("Auto Intake", expanded=False):
        uploaded_financial = st.file_uploader(
            "Upload Financial Statement / Rent Roll / Appraisal Snapshot",
            type=["pdf", "png", "jpg", "jpeg", "webp"]
        )

        if uploaded_financial is not None:
            try:
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
                    st.session_state["auto_purchase_price"] = int(extracted_fields["Purchase Price / Cost Basis"])

                if extracted_fields.get("Appraised Value"):
                    st.session_state["auto_appraisal"] = int(extracted_fields["Appraised Value"])

                if extracted_fields.get("Stabilized NOI"):
                    st.session_state["auto_noi"] = int(extracted_fields["Stabilized NOI"])

            except Exception as e:
                st.error(f"OCR processing failed: {e}")
                st.caption("For Streamlit Cloud, add pytesseract, pillow, and pymupdf to requirements. Add tesseract-ocr to packages.txt.")

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
            "Purchase Price / Cost Basis ($)",
            value=st.session_state.get("auto_purchase_price", 12500000),
            min_value=1,
            step=100000
        )

        appraisal = st.number_input(
            "Appraised Value ($)",
            value=st.session_state.get("auto_appraisal", 13750000),
            min_value=1,
            step=100000
        )

        existing_debt = 0
        if transaction_type == "Refinance":
            existing_debt = st.number_input(
                "Existing Debt Payoff ($)",
                value=8500000,
                min_value=0,
                step=100000
            )

        noi = st.number_input(
            "Stabilized NOI ($)",
            value=st.session_state.get("auto_noi", 1060322),
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
            "Other Closing Costs ($)",
            value=50000,
            min_value=0,
            step=5000
        )

    with st.expander("Reserves / Adjustments", expanded=False):
        capex_reserve = st.number_input(
            "CapEx / TI-LC Reserve ($)",
            value=0,
            min_value=0,
            step=25000
        )

        interest_reserve = st.number_input(
            "Interest Reserve ($)",
            value=0,
            min_value=0,
            step=25000
        )

    st.markdown("---")
    st.caption("Outputs are indicative and subject to lender diligence.")


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
st.subheader("Commercial Real Estate Underwriting Suite")
st.caption(f"Generated: {generated_at} | Transaction: {transaction_type} | Active Constraint: {gate}")

m1, m2, m3, m4, m5, m6 = st.columns(6)

m1.metric("Max Proceeds", format_money(loan_amt))
m2.metric("Actual LTV", format_pct(actual_ltv))
m3.metric("Actual LTC", format_pct(actual_ltc))
m4.metric("Actual DSCR", format_x(actual_dscr))
m5.metric("Debt Yield", format_pct(actual_dy))
m6.metric("Deal Score", f"{score}/1000")

st.markdown("---")

tabs = st.tabs([
    "Sizing",
    "Sensitivity",
    "Capital Stack",
    "Covenants",
    "Assumptions",
    "Scorecard",
    "Report"
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

        st.bar_chart(chart_df, x="Constraint", y="Max Proceeds", color="#1D4ED8")

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
# TAB 3: CAPITAL STACK
# ==========================================

with tabs[2]:
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
# TAB 4: COVENANTS
# ==========================================

with tabs[3]:
    st.subheader("Covenant Compliance")
    st.dataframe(covenant_df, hide_index=True, use_container_width=True)

    st.caption(
        "Covenant testing is based on user-entered assumptions and model-generated proceeds. "
        "Final compliance is subject to lender underwriting and documentation."
    )


# ==========================================
# TAB 5: ASSUMPTIONS
# ==========================================

with tabs[4]:
    st.subheader("Underwriting Assumptions")
    st.dataframe(assumptions_df, hide_index=True, use_container_width=True)

    st.caption(
        "Assumptions should be reconciled against rent rolls, operating statements, borrower financials, "
        "appraisal reports, environmental reports, and lender term sheets."
    )


# ==========================================
# TAB 6: SCORECARD
# ==========================================

with tabs[5]:
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
# TAB 7: REPORT + EXPORTS
# ==========================================

with tabs[6]:
    st.subheader("Executive Summary Preview")
    st.dataframe(preview_df, hide_index=True, use_container_width=True)

    report_text = f"""ALENZA CAPITAL UNDERWRITING SUMMARY
Generated: {generated_at}

============================================================
EXECUTIVE SUMMARY
============================================================

Sponsor / Borrower: {sponsor}
Property Type: {property_type}
Transaction Type: {transaction_type}

Supportable Proceeds: {format_money(loan_amt)}
Binding Constraint: {gate}
Classification: {classification}
Deal Score: {score}/1000

============================================================
ASSET PROFILE
============================================================

Purchase Price / Cost Basis: {format_money(purchase_price)}
Appraised Value: {format_money(appraisal)}
Existing Debt Payoff: {format_money(existing_debt)}
Stabilized NOI: {format_money(noi)}

============================================================
UNDERWRITING CRITERIA
============================================================

Maximum LTV: {format_pct(target_ltv)}
Maximum LTC: {format_pct(target_ltc)}
Minimum DSCR: {format_x(target_dscr)}
Minimum Debt Yield: {format_pct(target_dy)}
Debt Service Structure: {debt_structure}
Interest Rate: {format_pct(rate)}
Amortization: {amort} years
Loan Term: {loan_term} years

============================================================
LOAN METRICS
============================================================

Actual LTV: {format_pct(actual_ltv)}
Actual LTC: {format_pct(actual_ltc)}
Actual DSCR: {format_x(actual_dscr)}
Debt Yield: {format_pct(actual_dy)}
Monthly Payment: {format_money(monthly_payment)}
Annual Debt Service: {format_money(annual_debt_service)}
DSCR Cushion: {format_x(debt_service_cushion)}

============================================================
CAPITAL STACK
============================================================

Base Uses: {format_money(base_uses)}
Origination / Financing Fees: {format_money(financing_fees)}
Other Closing Costs: {format_money(closing_costs)}
CapEx / TI-LC Reserve: {format_money(capex_reserve)}
Interest Reserve: {format_money(interest_reserve)}
Total Uses: {format_money(total_uses)}

Supportable Senior Debt: {format_money(loan_amt)}
Required Sponsor Equity: {format_money(required_equity)}
Loan-to-Cost: {format_pct(actual_ltc)}
Equity Contribution: {format_pct(equity_pct)}

============================================================
SIZING CONSTRAINTS
============================================================

LTV Limit: {format_money(gates["LTV"])}
DSCR Limit: {format_money(gates["DSCR"])}
Debt Yield Limit: {format_money(gates["Debt Yield"])}
Binding Constraint: {gate}

============================================================
COVENANT TESTING
============================================================

Maximum LTV: Required <= {format_pct(target_ltv)} | Actual {format_pct(actual_ltv)} | {ltv_status}
Maximum LTC: Required <= {format_pct(target_ltc)} | Actual {format_pct(actual_ltc)} | {ltc_status}
Minimum DSCR: Required >= {format_x(target_dscr)} | Actual {format_x(actual_dscr)} | {dscr_status}
Minimum Debt Yield: Required >= {format_pct(target_dy)} | Actual {format_pct(actual_dy)} | {dy_status}

============================================================
SCORECARD
============================================================

Loan-to-Value Score: {ltv_score}/260
Loan-to-Cost Score: {ltc_score}/140
DSCR Score: {dscr_score}/260
Debt Yield Score: {dy_score}/200
Equity Contribution Score: {equity_score}/140

Total Score: {score}/1000
Classification: {classification}

============================================================
UNDERWRITING VERDICT
============================================================

{constraint_advice(gate)}

============================================================
DISCLAIMER
============================================================

This summary is indicative only and is not a loan commitment, credit approval,
investment advice, appraisal, legal opinion, or final underwriting decision.
All terms are subject to lender diligence, borrower review, third-party reports,
credit approval, committee review, and final documentation.
"""

    safe_sponsor = clean_filename(sponsor)

    text_col, excel_col, pdf_col = st.columns(3)

    with text_col:
        st.download_button(
            "Download Text Summary",
            report_text,
            file_name=f"Alenza_Underwriting_Summary_{safe_sponsor}.txt",
            mime="text/plain"
        )

    excel_file = create_excel_workbook(
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
        extracted_review_df=extracted_review_df
    )

    with excel_col:
        st.download_button(
            "Download Excel Workbook",
            data=excel_file,
            file_name=f"Alenza_Underwriting_Workbook_{safe_sponsor}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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

    with pdf_col:
        if pdf_file is not None:
            st.download_button(
                "Download PDF Summary",
                data=pdf_file,
                file_name=f"Alenza_Underwriting_Summary_{safe_sponsor}.pdf",
                mime="application/pdf"
            )
        else:
            st.warning("PDF export requires reportlab. Install reportlab to enable PDF downloads.")

    with st.expander("Deployment Notes", expanded=False):
        st.code(
            """Recommended requirements.txt:

streamlit
pandas
numpy
openpyxl
xlsxwriter
reportlab
pillow
pytesseract
pymupdf

Recommended packages.txt for Streamlit Cloud OCR:

tesseract-ocr
""",
            language="text"
        )

