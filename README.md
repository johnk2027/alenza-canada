# 🏛️ Alenza Capital OS v3.0.3

**Canadian Commercial Real Estate Debt Underwriting Workstation**

![Version](https://img.shields.io/badge/version-3.0.3-blue)
![Python](https://img.shields.io/badge/python-3.11-green)
![Streamlit](https://img.shields.io/badge/streamlit-1.33%2B-red)
![License](https://img.shields.io/badge/license-MIT-lightgrey)

Alenza Capital OS is a local-first underwriting tool for Canadian commercial real estate mortgage analysis, lender package preparation, and deal triage.

It brings together loan sizing, rent-roll review, sensitivity testing, pro forma return analysis, capital stack modeling, document-gap checks, federal registry verification, live Canadian market-rate context, PDF memo generation, and exportable lender packages in one Streamlit application.

The platform is built around a simple principle: **core underwriting should work without paid APIs, cloud databases, or third-party accounts.**

---

## 📊 What It Does

Alenza helps a broker, analyst, lender, or debt advisor quickly answer the core questions on a commercial mortgage file:

- 💰 **How much debt can the property support?**
- 🔒 **Which constraint is binding?** LTV, LTC, DSCR, or Debt Yield
- 💎 **How much equity is required, or is there surplus cash-out proceeds?**
- ✅ **Does the file pass basic lender covenants?**
- 🧪 **How sensitive is the deal to interest-rate and NOI movement?**
- 📈 **What do levered returns look like under realistic revenue and expense growth?**
- 🏗️ **How does senior debt, mezzanine debt, preferred equity, and sponsor equity fit together?**
- 📋 **What diligence is missing?**
- 🏢 **Can the sponsor’s federal corporation record be checked?**
- 🇨🇦 **How does the proposed rate compare with live Bank of Canada benchmark data?**
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
- Iterative sizing that accounts for financing fees inside total uses

The app identifies the **binding constraint** and produces:

- Sources-and-uses summary
- Required equity or surplus cash-out amount
- Covenant comparison table
- Full amortization schedule
- Balloon exposure by term
- Deal score from 0 to 1000
- Tier classification for credit positioning

---

### 🧪 Sensitivity & Stress Testing

The sensitivity module helps users assess how fragile or resilient a transaction is.

Included tools:

- **Scenario Matrix**
  - Base Case
  - Rate +1%
  - Rate -1%
  - NOI -10%
  - NOI +10%
  - Combined Stress

- **Custom Stress Test**
  - Rate shock in basis points
  - NOI shock percentage
  - LTV adjustment percentage

- **Proceeds Heatmap**
  - Max proceeds under rate and NOI shocks
  - Base-case cell highlighting
  - Visual risk mapping using Plotly when available

- **Breakeven Occupancy**
  - Calculates the occupancy level required to maintain a 1.00x DSCR

---

### 📝 Rent Roll Review

The rent-roll module provides a quick lease-risk snapshot.

Metrics include:

- **Weighted Average Lease Term (WALT)**
- Physical occupancy percentage
- Vacancy exposure
- 12-month rollover concentration
- Annual rent
- Rent per square foot
- Total square footage

Functionality includes:

- Interactive data editor
- CSV import
- XLSX import through OpenPyXL
- XLS import through XLRD
- Auto-column mapping for common rent-roll headers
- In-tab save button
- Recalculate button to refresh full model metrics immediately

---

### 📅 Amortization & Balloon Analysis

The amortization tab produces both monthly and annual debt-service schedules.

Included outputs:

- Monthly payment
- Annual debt service
- Balloon balance
- Balloon as a percentage of original loan
- Principal and interest breakdown
- Outstanding balance curve
- Annual summary table
- Full monthly schedule

The app also includes a **Term vs. Balloon Analysis** showing how balloon exposure changes across multiple loan terms.

---

### 📈 Pro Forma & Returns

The pro forma module models investment performance with more realistic operating assumptions.

Instead of using a single flat NOI growth assumption, Alenza separates:

- Revenue growth
- Expense growth
- Initial expense ratio
- Terminal NOI growth
- Exit cap rate
- Selling costs
- Projection period

This allows the model to show margin compression when expenses grow faster than revenue.

Outputs include:

- Projected revenue
- Projected expenses
- Projected NOI
- NOI margin
- Cash flow after debt service
- Levered IRR
- Equity multiple
- Gross exit value
- Net exit proceeds
- Exit NOI
- Exit cap / terminal growth sensitivity

---

### 🏗️ Capital Stack Modeling

Alenza supports a more complete capital stack beyond simple senior debt.

Included layers:

- Senior debt
- Mezzanine debt
- Preferred equity
- Sponsor equity
- Cash-out / surplus proceeds

Outputs include:

- Total capital
- Senior debt cost
- Mezzanine cost
- Preferred equity cost
- Total fixed charges
- Fixed-charge coverage ratio

---

### 🇨🇦 Canadian Sovereign Intelligence

The Canada Intel tab provides Canadian macro and market context for credit decisions.

| Source | Data Provided |
|--------|---------------|
| **Bank of Canada Valet API** | USD/CAD, EUR/CAD, overnight rate, overnight target, bank rate, 2Y/5Y/10Y Government of Canada yields |
| **Statistics Canada** | National unemployment rate |
| **Corporations Canada** | Federal corporation verification by corporation number or BN9 |

Charts and analysis include:

- Current yield curve
- Government of Canada bond-yield history
- Policy-rate history
- USD/CAD and EUR/CAD exchange-rate history
- Canadian unemployment history
- 3-month unemployment moving average
- Deal rate vs. 5Y GoC benchmark
- Thin-risk-premium warning
- Yield-curve commentary
- Labour-market commentary
- Simulated commercial vacancy context by property type

The app uses public Canadian data sources and does not require paid data feeds.

---

### ⚖️ Live Yield Benchmarking

Users can optionally benchmark the deal rate to the live 5Y Government of Canada yield.

The app can display:

- Current 5Y GoC benchmark
- User-defined credit spread in basis points
- Indexed deal rate based on benchmark plus spread
- Spread to GoC
- Risk-premium commentary

---

### 📈 Market Comparables

The Market Comps tab creates a simulated comparison stack based on the selected property type.

Supported property types include:

- Multifamily
- Industrial
- Retail
- Office
- Mixed-Use
- Hospitality
- Self-Storage

Outputs include:

- Simulated comparable sale names
- Sale dates
- Cap rates
- Estimated NOI
- Estimated value
- Price per unit or square foot
- Subject implied cap rate
- Average market cap rate
- Median market cap rate
- Value at average market cap rate
- Subject vs. comp map

**Important:** Market comps are simulated placeholders and should be replaced with verified broker, appraisal, CoStar, Altus, CMHC, CBRE, or internal data before final underwriting.

---

### 📋 Diligence Room & Document Vault

The Diligence Room tracks uploaded deal documents and missing items.

Tracked categories include:

- Appraisal
- Phase I ESA
- T12 Financials
- Rent Roll
- Sponsor Bio
- Purchase Agreement
- Environmental Report
- Structural Report
- Other

Functionality includes:

- File upload
- File size validation
- Local vault storage
- Document inventory
- Required document checklist
- Missing document warnings
- Document delete confirmation
- Diligence notes

Uploaded documents are tracked with status indicators:

- ✅ Uploaded
- ❌ Missing

---

### 🔎 PDF Context Scan

Uploaded PDFs can be scanned for underwriting keywords.

The app returns:

- Keyword counts
- Page references
- Context snippets
- Text preview

Default scan keywords include:

- NOI
- Total
- Lease
- Tenant
- Rent
- Occupancy
- Appraisal
- Value
- Environmental
- Phase

This is a diligence aid. It does not replace full document review or legal review.

---

### 🔍 Federal Corporation Verification

Users can verify Canadian federal corporations by corporation number or 9-digit business number.

Depending on registry availability, the app may return:

- Corporation number
- Raw registry response
- Basic corporation record

This is a diligence aid. It does not replace legal, KYC, AML, beneficial ownership, sanctions, fraud, or corporate counsel review.

---

### 💾 Save & Export

Alenza supports database saving and multiple export formats.

| Format | Contents |
|--------|----------|
| **Excel (.xlsx)** | Summary, Constraints, Capital Stack, Rent Roll, Amortization, Sensitivity, Heatmap Data, Documents, Versions |
| **JSON (.json)** | Complete deal state |
| **Encrypted JSON** | Password-encrypted deal export when cryptography is available |
| **PDF Memo** | One-page indicative underwriting memo with key credit metrics and Canada Intel commentary |
| **ZIP Package** | Deal state, summary CSV, constraints CSV, capital stack JSON, rent roll CSV, amortization CSV, sensitivity CSV, heatmap CSV, document inventory, version history, audit log, and PDF memo when available |

Export files are timestamped at download.

---

### ✅ QA & Health

The QA & Health tab provides internal model checks and system diagnostics.

Included:

- Financial regression self-tests
- Pass/fail test table
- Database connectivity check
- Local storage check
- Dependency availability check
- Encryption-key status
- Current deal validation
- Current model snapshot
- Deal version history
- Audit log
- Manual QA checklist

This helps confirm that core sizing, amortization, rent-roll math, pro forma logic, and capital stack calculations are working as expected.

---

### 🔐 Audit Log & Version History

The app records key workflow events, including:

- Deal saves
- Deal loads
- Deal deletes
- Document uploads
- Document deletes
- Manual saves
- Autosaves

Saved deals also create version records with change summaries.

---

## 🏗️ Architecture

**Local-first design.** Alenza runs without cloud infrastructure by default.

- **Database:** SQLite with WAL mode
- **Documents:** Local filesystem storage
- **Versioning:** Deal-version table
- **Audit Trail:** Audit-log table
- **Data Sources:** Public Canadian government APIs
- **Exports:** Excel, JSON, encrypted JSON, ZIP, and PDF memo
- **Security:** Optional encryption for stored deal state and exported JSON

---

## 🧰 Technical Stack

| Area | Technology |
|------|------------|
| UI Framework | Streamlit |
| Database | SQLite with WAL mode |
| Data Processing | Pandas, NumPy |
| Financial Math | NumPy Financial |
| Charts | Plotly, Streamlit native charts |
| Excel Read | OpenPyXL, XLRD |
| Excel Write | XlsxWriter |
| OCR / PDF Text | Tesseract OCR, PyMuPDF, Pillow |
| PDF Export | ReportLab |
| Security | Cryptography |
| Public Data | Bank of Canada, Statistics Canada, Corporations Canada |
| Autosave Timer | streamlit-autorefresh |

---

## 📁 Project Structure

```text
alenza-canada/
├── app.py                    # Main application
├── requirements.txt          # Python dependencies
├── packages.txt              # System packages for Streamlit Cloud
├── runtime.txt               # Python runtime pin
├── secrets_template.toml     # Template for optional secrets
├── README.md                 # Documentation
├── .gitignore                # Git ignore rules
└── .streamlit/
    └── config.toml           # Theme and server configuration
