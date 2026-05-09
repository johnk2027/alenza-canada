# Alenza Capital OS

**Canadian commercial real estate debt underwriting workstation.**

Alenza Capital OS is a local-first underwriting tool for Canadian CRE mortgage analysis, lender package preparation, and deal triage. It brings together loan sizing, rent-roll review, document-gap checks, federal registry verification, market-rate context, and exportable reporting in one Streamlit application.

The platform is built around a simple principle: core underwriting should work without paid APIs, cloud databases, or third-party accounts.

---

## What It Does

Alenza helps a broker, analyst, or lender quickly answer the core questions on a commercial mortgage file:

- How much debt can the property support?
- Which constraint is binding?
- How much equity is required?
- Does the file pass basic lender covenants?
- What diligence is missing?
- Is the sponsor’s federal corporation record active?
- How does the proposed rate compare with current Canadian benchmark-rate context?
- Can the analysis be exported into a lender-ready package?

---

## Key Features

### Loan Sizing

The underwriting engine sizes debt against the main credit constraints used in Canadian commercial mortgage placement:

- Loan-to-Value
- Loan-to-Cost
- Debt Service Coverage Ratio
- Debt Yield
- amortizing and interest-only structures
- financing fees, reserves, closing costs, and other uses of funds

The app identifies the binding constraint and produces a sources-and-uses summary, required equity amount, covenant table, amortization schedule, and deal score.

### Rent Roll Review

The rent-roll module provides a quick lease-risk snapshot:

- Weighted Average Lease Term
- physical occupancy
- vacancy exposure
- 12-month rollover concentration
- annual rent
- average rent per square foot

This is intended for early-stage screening and lender conversations, not as a replacement for a full valuation platform.

### Diligence Gap Check

The deal-room audit reviews uploaded file names and flags missing items from a typical commercial mortgage package, including:

- T12 financials
- rent roll
- appraisal
- Phase I environmental report
- sponsor biography
- Schedule of Real Estate Owned
- purchase agreement
- insurance
- title or survey materials

If documents are missing, the app can create a pre-filled email draft for follow-up.

### Canadian Public Data

Alenza uses Canadian public data sources where possible.

| Source | Use |
|---|---|
| NRCan / Geo.ca | Canadian address matching and geolocation |
| Bank of Canada Valet | benchmark-rate context |
| Statistics Canada / Open Canada | market-rate and economic dataset references |
| Corporations Canada | federal corporation verification |

Core features do not require paid data subscriptions.

### Federal Corporation Verification

The app can check a Canadian federal corporation by corporation number or 9-digit business number.

It can display:

- legal name
- corporate status
- governing act
- registered address
- annual return history
- director limits
- corporate activity history

This is a diligence aid. It does not replace legal, KYC, AML, or beneficial ownership review.

### Market Rate Context

The Market Rates section pulls live/open Canadian rate context and compares deal pricing against available benchmark-rate data. This helps support lender discussions and internal credit notes.

### Local OCR and Optional Local AI

Alenza can process uploaded PDFs or image files for OCR-based financial review.

It also supports optional local Ollama review for users who want to validate extracted text without sending sensitive property documents to a third-party AI service.

Typical workflow:

1. Upload financials or scanned documents.
2. Extract text through OCR.
3. Review detected values.
4. Run optional local model review.
5. Export the results with the underwriting package.

---

## Local-First Architecture

Alenza is designed to run without requiring cloud infrastructure.

By default, it uses:

- local SQLite database
- local document folder
- local backup ZIP exports
- public Canadian data sources
- email draft links for deal routing

Recommended folder structure:

```text
alenza-canada/
├── app.py
├── requirements.txt
├── packages.txt
├── secrets_template.toml
├── README.md
├── .gitignore
├── .streamlit/
│   └── config.toml
└── alenza_data/
    ├── documents/
    │   └── .gitkeep
    ├── backups/
    │   └── .gitkeep
    └── exports/
        └── .gitkeep

In local mode, deal files and the database stay on the user’s machine.

When deployed to a hosted environment, uploaded files and calculations are handled by that hosting environment. Confidential borrower documents should only be used in hosted mode if appropriate privacy and security controls are in place.

Exports

Alenza can generate:

Excel underwriting workbook
PDF executive summary
text summary
JSON deal snapshot
backup ZIP
audit trail
amortization schedule
rent-roll summary
diligence tracker
lender quote comparison
Technical Stack
Area	Tooling
UI	Streamlit
Database	SQLite
PDF reporting	ReportLab
Excel reporting	XlsxWriter / OpenPyXL
OCR	Tesseract OCR
PDF parsing	PyMuPDF
Image handling	Pillow
Optional local model	Ollama
Public data	NRCan, Bank of Canada, Statistics Canada, Corporations Canada
Quick Start
Requirements
Python 3.10+
Tesseract OCR

Install
git clone https://github.com/johnk2027/alenza-canada.git
cd alenza-canada
pip install -r requirements.txt
Run
streamlit run app.py
Streamlit Cloud Deployment

Use these settings:

Repository: johnk2027/alenza-canada
Branch: main
Main file path: app.py

The app can run without secrets.

Configuration

To set the default deal inbox, add this to Streamlit Secrets or your environment:

ALENZA_DEAL_INBOX_EMAIL = "resourcefulcapital@gmail.com"

Optional paid integrations can be added later through secrets_template.toml.

Intended Users

Alenza is built for:

Canadian commercial mortgage brokers
debt advisors
private lenders
CRE analysts
borrower intake teams
acquisition and refinance screening
lender package preparation
Important Notice

Alenza Capital OS is an underwriting and workflow tool. It is not a loan commitment, appraisal, legal opinion, tax advice, investment recommendation, or final credit decision. All outputs should be reviewed by qualified professionals before being relied on for financing, legal, tax, or investment purposes.

Contact

For broker workflow, partnership, or institutional inquiries:

resourcefulcapital@gmail.com
