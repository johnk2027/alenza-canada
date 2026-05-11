Your README is **95% complete and excellent!** Just a few small fixes needed:

## **ISSUES TO FIX:**

### 1. **Badges Section is Broken**
Lines 7-10 show raw text:
```
Version Python License Streamlit
```

### 2. **Formatting Glitch at Bottom**
Last line is missing the "B" in "Built" and has stray text.

### 3. **Quick Start Section Formatting**
The run command is separated from the code block.

---

## **HERE'S THE COMPLETE FIXED README:**

```markdown
# 🏛️ Alenza Capital OS v3.0.1

**Canadian Commercial Real Estate Debt Underwriting Workstation**

![Version](https://img.shields.io/badge/version-3.0.1-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-green)
![Streamlit](https://img.shields.io/badge/streamlit-1.28%2B-red)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

Alenza Capital OS is a local-first underwriting tool for Canadian CRE mortgage analysis, lender package preparation, and deal triage. It brings together loan sizing, rent-roll review, document-gap checks, federal registry verification, live market-rate context with interactive charts, and exportable reporting in one Streamlit application.

The platform is built around a simple principle: **core underwriting should work without paid APIs, cloud databases, or third-party accounts.**

---

## 📊 What It Does

Alenza helps a broker, analyst, or lender quickly answer the core questions on a commercial mortgage file:

- 💰 **How much debt can the property support?**
- 🔒 **Which constraint is binding?** (LTV, LTC, DSCR, or Debt Yield)
- 💎 **How much equity is required?**
- ✅ **Does the file pass basic lender covenants?**
- 📋 **What diligence is missing?**
- 🏢 **Is the sponsor's federal corporation record active?**
- 📈 **How does the proposed rate compare with live Bank of Canada benchmark data?**
- 📦 **Can the analysis be exported into a lender-ready package?**

---

## ✨ Key Features

### 📐 Loan Sizing Engine
The underwriting engine sizes debt against the main credit constraints used in Canadian commercial mortgage placement:

- **Loan-to-Value (LTV)**
- **Loan-to-Cost (LTC)**
- **Debt Service Coverage Ratio (DSCR)**
- **Debt Yield (DY)**
- Amortizing and interest-only structures
- Financing fees, reserves, closing costs, and other uses of funds

The app identifies the **binding constraint** and produces:
- Sources-and-uses summary
- Required equity amount
- Covenant comparison table
- Full amortization schedule
- Deal score (0-1000) with tier classification

### 🧪 Sensitivity & Stress Testing
- **Scenario Matrix**: Rate +1%, Rate -1%, NOI -10%, NOI +10%, Combined Stress
- **Custom Stress Test**: Adjust rate shocks (bps), NOI shocks (%), and LTV adjustments
- Real-time proceeds impact analysis
- Breakeven occupancy calculation

### 📝 Rent Roll Review
The rent-roll module provides a quick lease-risk snapshot:

- **Weighted Average Lease Term (WALT)**
- Physical occupancy percentage
- Vacancy exposure
- 12-month rollover concentration
- Annual rent and rent per square foot
- Interactive data editor with import/export
- Auto-column mapping for CSV/Excel imports

### 🇨🇦 Canadian Sovereign Intelligence (LIVE DATA)

| Source | Data Provided |
|--------|---------------|
| **Bank of Canada Valet** | Live FX rates, 2Y/5Y/10Y bond yields |
| **Statistics Canada** | National unemployment rate (Table 14-10-0287-01) |
| **Corporations Canada** | Federal corporation verification by BN9 |

#### Interactive Charts:
- 📈 **Current Yield Curve** - 2Y, 5Y, 10Y visualization
- 📊 **90-Day Yield Trend** - Historical bond yield movements
- 💱 **USD/CAD Exchange Rate** - 90-day trend with statistics
- 🎯 **Deal Rate vs Benchmark** - Spread comparison with commentary

### 📋 Diligence Gap Check
The deal-room audit tracks required documents:

- T12 financials
- Rent roll
- Appraisal
- Phase I Environmental report
- Sponsor biography
- Purchase agreement
- Insurance documentation
- Title/survey materials

Uploaded documents are tracked with status indicators (✅ Uploaded / ❌ Missing).

### 🔍 Federal Corporation Verification
Verify Canadian federal corporations by corporation number or 9-digit business number:

- Legal name
- Corporate status
- Governing act
- Registered address
- Annual return history

*This is a diligence aid. It does not replace legal, KYC, AML, or beneficial ownership review.*

### 🤖 OCR Document Extraction (Optional)
Process uploaded PDFs or image files for financial data extraction:

1. Upload financials or scanned documents
2. Extract text through OCR (Tesseract)
3. Review detected values with confidence scores
4. Apply parameters to underwriting model
5. Export results with the underwriting package

### 💾 Export Capabilities

| Format | Contents |
|--------|----------|
| **Excel (.xlsx)** | Summary, Rent Roll, Amortization Schedule |
| **JSON (.json)** | Complete deal state (optionally encrypted) |
| **ZIP Package** | Excel + JSON + backup database |

---

## 🏗️ Architecture

**Local-First Design** - Runs without cloud infrastructure:

- **Database**: Local SQLite with WAL mode and versioning
- **Documents**: Local filesystem storage
- **Backups**: ZIP exports with deal packages
- **Data Sources**: Public Canadian government APIs
- **Email**: Pre-filled mailto links for deal routing

### Technical Stack

| Area | Technology |
|------|------------|
| UI Framework | Streamlit |
| Database | SQLite (WAL mode) |
| Data Processing | Pandas, NumPy |
| Charts | Plotly (interactive), Streamlit (native) |
| Excel | XlsxWriter, OpenPyXL, XLRD |
| OCR | Tesseract OCR, PyMuPDF, Pillow |
| PDF Export | ReportLab |
| Security | Cryptography (Fernet AES-128) |
| Public Data | Bank of Canada, Statistics Canada, Corporations Canada |

---

## 📁 Project Structure

```
alenza-canada/
├── app.py                    # Main application (self-contained)
├── requirements.txt          # Python dependencies
├── packages.txt              # System packages (for Streamlit Cloud)
├── secrets_template.toml     # Template for secrets configuration
├── README.md                 # This documentation
├── .gitignore               # Git ignore rules
└── .streamlit/
    └── config.toml           # Theme and server configuration
```

**Auto-generated at runtime:**
```
alenza_data/
├── alenza_platform.db        # SQLite database
└── documents/                # Uploaded deal documents
```

---

## 🚀 Quick Start

### Prerequisites
- **Python 3.10+**
- **Tesseract OCR** (optional, for document extraction)

### Installation

```bash
# Clone the repository
git clone https://github.com/johnk2027/alenza-canada.git
cd alenza-canada

# Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Run Locally

```bash
streamlit run app.py
```

The app will open at `http://localhost:8501` and automatically create all necessary directories and the database.

---

## ☁️ Streamlit Cloud Deployment

1. Push this repository to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Deploy with these settings:
   - **Repository**: `johnk2027/alenza-canada`
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. The app runs without secrets by default

### Optional Configuration

Add these to Streamlit Secrets (or `.streamlit/secrets.toml` locally):

```toml
APP_USER = "Your Name"
ALENZA_DEAL_INBOX_EMAIL = "resourcefulcapital@gmail.com"
ALENZA_SECRET_KEY = "your-random-secret-key"
```

---

## 🎨 Theme

Alenza Capital OS uses a custom dark theme:

| Element | Color |
|---------|-------|
| Primary (Accent) | `#CFB87C` CU Gold |
| Background | `#0B0F19` Midnight Slate |
| Secondary | `#0F172A` Dark Navy |
| Text | `#F3F4F6` Light Gray |

---

## 👥 Intended Users

Alenza is built for:

- 🇨🇦 Canadian commercial mortgage brokers
- 💼 Debt advisors and placement agents
- 🏦 Private lenders and credit unions
- 📊 CRE analysts and underwriters
- 📋 Borrower intake and screening teams
- 🔍 Acquisition and refinance screening
- 📁 Lender package preparation

---

## ⚠️ Important Notice

**Alenza Capital OS is an underwriting and workflow tool.** It is not:

- ❌ A loan commitment or credit approval
- ❌ An appraisal or valuation
- ❌ Legal, tax, or investment advice
- ❌ A replacement for professional due diligence

All outputs should be reviewed by qualified professionals before being relied upon for financing, legal, tax, or investment purposes.

---

## 📄 License

MIT License - See LICENSE file for details.

---

## 📧 Contact

For broker workflow, partnership, or institutional inquiries:

- **Email**: [resourcefulcapital@gmail.com](mailto:resourcefulcapital@gmail.com)
- **GitHub**: [github.com/johnk2027/alenza-canada](https://github.com/johnk2027/alenza-canada)

---

**Built for Canadian CRE professionals. No paid APIs required.** 🍁
```
