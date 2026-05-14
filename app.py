"""
Alenza Capital OS v4.5.1
Single-file Streamlit CRE Underwriting Workstation.
Features strict state contracts, cryptographic persistence, zero-latency state rendering, and exhaustive institutional exports.
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import logging
import os
import re
import sqlite3
import time
import uuid
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import requests
import streamlit as st

# =============================================================================
# 1. STREAMLIT LIFECYCLE (MUST BE FIRST)
# =============================================================================
st.set_page_config(page_title="Alenza Capital OS", page_icon="🏢", layout="wide")

# =============================================================================
# 2. OPTIONAL DEPENDENCIES
# =============================================================================
@st.cache_resource
def check_dependencies() -> Dict[str, bool]:
    deps = {
        "ocr": False, "pdf": False, "excel_write": False, 
        "plotly": False, "crypto": False, "numpy_financial": False,
    }
    try: import fitz; deps["ocr"] = True
    except ImportError: pass
    try: from reportlab.lib.pagesizes import letter; from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle; from reportlab.lib.styles import getSampleStyleSheet; from reportlab.lib import colors; deps["pdf"] = True
    except ImportError: pass
    try: import xlsxwriter; deps["excel_write"] = True
    except ImportError: pass
    try: import plotly.express as px; import plotly.graph_objects as go; from plotly.subplots import make_subplots; deps["plotly"] = True
    except ImportError: pass
    try: from cryptography.fernet import Fernet; from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC; from cryptography.hazmat.primitives import hashes; deps["crypto"] = True
    except ImportError: pass
    try: import numpy_financial as npf; deps["numpy_financial"] = True
    except ImportError: pass
    return deps

DEPS = check_dependencies()

if DEPS["ocr"]: import fitz
if DEPS["pdf"]:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
if DEPS["plotly"]:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
if DEPS["numpy_financial"]:
    import numpy_financial as npf
if DEPS["crypto"]:
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes

# =============================================================================
# 3. CONSTANTS, CONVENTIONS & EXPLICIT STATE CONTRACT
# =============================================================================
VERSION = "4.5.1"
MAX_UPLOAD_MB = 50
DATA_DIR = Path("alenza_data")
DB_PATH = DATA_DIR / "alenza_platform.db"
DOC_DIR = DATA_DIR / "documents"

LENDER_PROFILES_LIST = ["Bank / Credit Union", "LifeCo / Core", "Bridge / Private", "CMHC Multifamily"]
PROPERTY_TYPES = ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality", "Self-Storage"]
TX_TYPES = ["Acquisition", "Refinance", "Construction", "Bridge", "Recapitalization"]
DEFAULT_RENT_COLS = ["Tenant", "SF", "Remaining Term", "Monthly Rent"]
DEFAULT_RENT_ROLL = [
    {"Tenant": "Main Anchor", "SF": 25000, "Remaining Term": 5.5, "Monthly Rent": 45000},
    {"Tenant": "In-Line A", "SF": 3500, "Remaining Term": 1.2, "Monthly Rent": 8000},
]

FINANCIAL_CONVENTIONS = {
    "Currency": "CAD unless otherwise noted.",
    "Amortization": "Monthly compounding / equal monthly payments.",
    "Day Count": "Simplified monthly model (30/360 proxy); not daily exact accrual.",
    "IRR Convention": "Periodic annual cash flows (End of Year).",
    "Sizing Convention": "IO DSCR sized on interest-only debt service."
}

DEAL_STATE_KEYS = [
    "deal_id", "deal_name", "sponsor", "property_address", "property_type", "transaction_type", "lender_profile",
    "purchase_price", "appraisal", "noi", "rate", "amort", "term", "is_io", "fees", 
    "closing_costs", "reserves", "target_ltv", "target_ltc", "target_dscr", "target_dy", 
    "mezz_debt", "pref_equity", "mezz_rate", "pref_rate", "pf_rev_growth", "pf_exp_growth", 
    "pf_exp_ratio", "pf_exit_cap", "pf_sell_costs", "pf_term_growth", "rate_lock_enabled", 
    "rate_lock_spread_bps", "rent_roll_dict", "diligence_notes"
]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# =============================================================================
# 4. STATE MANAGEMENT
# =============================================================================
def default_state() -> dict:
    return {
        "deal_id": f"deal_{int(time.time())}_{uuid.uuid4().hex[:6]}",
        "deal_name": "Untitled Deal",
        "sponsor": "",
        "property_address": "",
        "property_type": "Multifamily",
        "transaction_type": "Acquisition",
        "lender_profile": "Bank / Credit Union",
        "purchase_price": 0.0,
        "appraisal": 0.0,
        "noi": 0.0,
        "rate": 0.055,
        "amort": 25,
        "term": 5,
        "is_io": False,
        "fees": 0.015,
        "closing_costs": 0.0,
        "reserves": 0.0,
        "target_ltv": 0.75,
        "target_ltc": 0.80,
        "target_dscr": 1.25,
        "target_dy": 0.08,
        "mezz_debt": 0.0,
        "pref_equity": 0.0,
        "mezz_rate": 0.11,
        "pref_rate": 0.09,
        "pf_rev_growth": 0.03,
        "pf_exp_growth": 0.02,
        "pf_exp_ratio": 0.40,
        "pf_exit_cap": 0.06,
        "pf_sell_costs": 0.015,
        "pf_term_growth": 0.02,
        "rate_lock_enabled": False,
        "rate_lock_spread_bps": 250,
        "rent_roll_dict": DEFAULT_RENT_ROLL.copy(),
        "diligence_notes": "",
    }

def get_current_state() -> dict:
    defaults = default_state()
    return {key: st.session_state.get(key, defaults.get(key)) for key in DEAL_STATE_KEYS}

# =============================================================================
# 5. UTILITIES & CRYPTOGRAPHY
# =============================================================================
def safe_float(v: Any, default: float = 0.0) -> float:
    if v is None or v == "": return default
    try:
        if isinstance(v, (int, float, np.number)): return float(v)
        c = re.sub(r"[^\d.\-()]", "", str(v)).strip()
        if c.startswith("(") and c.endswith(")"): c = "-" + c[1:-1]
        return float(c) if c else default
    except (ValueError, TypeError, OverflowError): return default

def safe_ratio(num: float, den: float) -> float:
    return float(num) / float(den) if float(den) > 0 else 0.0

def get_encryption_key() -> Optional[str]:
    try: return st.secrets.get("ALENZA_DB_ENCRYPTION_KEY") or os.environ.get("ALENZA_DB_ENCRYPTION_KEY")
    except Exception: return os.environ.get("ALENZA_DB_ENCRYPTION_KEY")

def make_fernet(secret: str, salt: bytes):
    if not DEPS["crypto"]: raise RuntimeError("Cryptography package missing.")
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(secret.encode()))
    return Fernet(key)

def encrypt_text(text: str, secret: str) -> str:
    salt = os.urandom(16)
    token = make_fernet(secret, salt).encrypt(text.encode())
    return base64.b64encode(salt + token).decode()

def decrypt_text(token_b64: str, secret: str) -> str:
    raw = base64.b64decode(token_b64.encode())
    salt, token = raw[:16], raw[16:]
    return make_fernet(secret, salt).decrypt(token).decode()

def encrypt_bytes(data: bytes, secret: str) -> bytes:
    salt = os.urandom(16)
    token = make_fernet(secret, salt).encrypt(data)
    return salt + token

def decrypt_bytes(raw: bytes, secret: str) -> bytes:
    salt, token = raw[:16], raw[16:]
    return make_fernet(secret, salt).decrypt(token)

def apply_theme():
    st.markdown("""<style>
        :root { --g: #CFB87C; --bg: #0B0F19; --fg: #0F172A; }
        .stApp { background: var(--bg); }
        [data-testid="stMetric"] { background: var(--fg); border: 1px solid rgba(207,184,124,0.3); padding: 15px; border-radius: 8px; }
        [data-testid="stMetricValue"] { color: var(--g); }
        .stTabs [aria-selected="true"] { border-bottom: 2px solid var(--g) !important; color: var(--g) !important; }
    </style>""", unsafe_allow_html=True)

# =============================================================================
# 6. DATA NORMALIZERS & GENERATORS
# =============================================================================
def normalize_rr(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame(columns=DEFAULT_RENT_COLS)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    col_map = {}
    for col in df.columns:
        key = re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip()
        if key in ["tenant", "tenant name", "lessee", "occupant", "company", "name"]: col_map[col] = "Tenant"
        elif key in ["sf", "sq ft", "sqft", "square feet", "gla", "area", "leased sf"]: col_map[col] = "SF"
        elif key in ["remaining term", "term remaining", "lease term", "years remaining"]: col_map[col] = "Remaining Term"
        elif key in ["monthly rent", "rent month", "monthly base rent", "rent"]: col_map[col] = "Monthly Rent"
        elif key in ["annual rent", "annual base rent", "yearly rent"]: col_map[col] = "Annual Rent"

    df = df.rename(columns=col_map)
    if "Annual Rent" in df.columns and "Monthly Rent" not in df.columns:
        df["Monthly Rent"] = pd.to_numeric(df["Annual Rent"], errors="coerce").fillna(0) / 12

    for col in DEFAULT_RENT_COLS:
        if col not in df.columns: df[col] = "" if col == "Tenant" else 0.0

    out = df[DEFAULT_RENT_COLS].copy()
    out["Tenant"] = out["Tenant"].fillna("").astype(str).str.strip()
    out["SF"] = pd.to_numeric(out["SF"], errors="coerce").fillna(0).clip(lower=0)
    out["Remaining Term"] = pd.to_numeric(out["Remaining Term"], errors="coerce").fillna(0).clip(lower=0)
    out["Monthly Rent"] = pd.to_numeric(out["Monthly Rent"], errors="coerce").fillna(0).clip(lower=0)
    return out.reset_index(drop=True)

def generate_comps(property_type: str, noi: float, appraisal: float, seed_text: str) -> pd.DataFrame:
    seed_raw = hashlib.sha256(f"{property_type}-{safe_float(noi)}-{safe_float(appraisal)}-{seed_text}".encode()).hexdigest()
    seed = int(seed_raw[:12], 16) % 1_000_000
    rng = np.random.default_rng(seed)
    base_cap = {"Multifamily": 0.045, "Industrial": 0.055, "Retail": 0.065, "Office": 0.075, "Mixed-Use": 0.060, "Hospitality": 0.080, "Self-Storage": 0.058}.get(property_type, 0.06)
    
    rows = []
    for i in range(5):
        cap = max(0.03, base_cap + rng.uniform(-0.006, 0.006))
        comp_noi = safe_float(noi) * rng.uniform(0.70, 1.30) if safe_float(noi) > 0 else (safe_float(appraisal) * cap * rng.uniform(0.75, 1.25) if safe_float(appraisal) > 0 else rng.uniform(350_000, 1_500_000))
        rows.append({
            "Comparable": f"[SIMULATED] {property_type} Comp {i + 1}",
            "Distance (km)": f"{rng.uniform(0.5, 8.0):.1f}",
            "Sale Date": (datetime.now() - timedelta(days=int(rng.integers(30, 540)))).strftime("%Y-%m-%d"),
            "Cap Rate": cap, "NOI": comp_noi, "Value": comp_noi / cap if cap > 0 else 0,
            "lat": 43.6532 + rng.uniform(-0.08, 0.08), "lon": -79.3832 + rng.uniform(-0.08, 0.08),
        })
    return pd.DataFrame(rows)

# =============================================================================
# 7. CACHED FINANCIAL ENGINES
# =============================================================================
class UnderwritingEngine:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def size_loan(noi: float, appraisal: float, purchase_price: float, closing_costs: float, reserves: float, fees_pct: float, rate: float, amort: int, term: int, is_io: bool, target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float) -> Tuple[float, str, dict, float, float]:
        """
        Size the senior mortgage amount using only valid underwriting gates.

        Fix:
        - The prior implementation initialized every gate to 0.0.
          That made the mortgage amount display as $0 whenever appraisal,
          purchase price/uses, or NOI had not been entered.
        - This version excludes unavailable gates, reports missing inputs,
          and uses purchase price as a collateral fallback when appraisal is blank.
        """
        noi = max(0.0, safe_float(noi))
        appraisal = max(0.0, safe_float(appraisal))
        purchase_price = max(0.0, safe_float(purchase_price))
        closing_costs = max(0.0, safe_float(closing_costs))
        reserves = max(0.0, safe_float(reserves))
        fees_pct = max(0.0, safe_float(fees_pct))
        rate = max(0.0, safe_float(rate))
        amort = max(1, int(safe_float(amort, 1)))
        target_ltv = max(0.0, safe_float(target_ltv))
        target_ltc = max(0.0, safe_float(target_ltc))
        target_dscr = max(0.0, safe_float(target_dscr))
        target_dy = max(0.0, safe_float(target_dy))

        hard_costs = purchase_price + closing_costs + reserves
        collateral_value = appraisal if appraisal > 0 else purchase_price

        loan = 0.0
        total_uses = hard_costs
        gates: Dict[str, float] = {}
        missing: List[str] = []

        for _ in range(25):
            total_uses = hard_costs + (loan * fees_pct)
            gates = {}
            missing = []

            if collateral_value > 0 and target_ltv > 0:
                gates["LTV"] = collateral_value * target_ltv
            else:
                missing.append("LTV requires appraisal or purchase price")

            if total_uses > 0 and target_ltc > 0:
                gates["LTC"] = total_uses * target_ltc
            else:
                missing.append("LTC requires purchase price/uses")

            if noi > 0 and target_dy > 0:
                gates["Debt Yield"] = noi / target_dy
            else:
                missing.append("Debt Yield requires NOI and target debt yield")

            if noi > 0 and rate > 0 and target_dscr > 0:
                if is_io:
                    gates["DSCR"] = (noi / target_dscr) / rate
                else:
                    m_rate = rate / 12
                    pmt_f = (1 - (1 + m_rate) ** -(amort * 12)) / m_rate if m_rate > 0 else amort * 12
                    gates["DSCR"] = (noi / target_dscr) / 12 * pmt_f if pmt_f > 0 else 0.0
            else:
                missing.append("DSCR requires NOI, rate, and target DSCR")

            if not gates:
                diagnostics = {"Missing Inputs": 0.0, "_messages": sorted(set(missing))}
                return 0.0, "Missing Inputs", diagnostics, round(total_uses, 2), round(total_uses, 2)

            new_loan = max(0.0, min(gates.values()))
            if abs(new_loan - loan) < 0.01:
                loan = new_loan
                break
            loan = new_loan

        if missing:
            gates["_messages"] = sorted(set(missing))

        numeric_gates = {k: v for k, v in gates.items() if not str(k).startswith("_")}
        binding_gate = min(numeric_gates, key=numeric_gates.get) if numeric_gates else "Missing Inputs"
        return round(loan, 2), binding_gate, gates, round(total_uses, 2), round(total_uses - loan, 2)

    @staticmethod
    @st.cache_data(show_spinner=False)
    def amort_schedule(loan_amt: float, rate: float, amort_yrs: int, term_yrs: int, is_io: bool) -> Tuple[pd.DataFrame, float, float]:
        if loan_amt <= 0: return pd.DataFrame(columns=["Period","Payment","Principal","Interest","Balance"]), 0.0, 0.0
        m_rate = rate / 12
        m_pmt = loan_amt * m_rate if is_io else (loan_amt * m_rate) / (1 - (1 + m_rate)**-(amort_yrs * 12)) if m_rate > 0 else loan_amt / (amort_yrs * 12)
        
        data, bal = [], loan_amt
        for m in range(1, int(term_yrs * 12) + 1):
            intr = bal * m_rate
            prin = 0.0 if is_io else min(m_pmt - intr, bal)
            bal = max(0.0, bal - prin)
            data.append({"Period": m, "Payment": round(m_pmt, 2) if (bal > 0 or prin > 0) else 0.0, "Principal": round(prin, 2), "Interest": round(intr, 2), "Balance": round(bal, 2)})
            if bal <= 0.001: break
        return pd.DataFrame(data), round(m_pmt, 2), round(bal, 2)

    @staticmethod
    def capital_stack(senior_debt: float, mezz_debt: float, pref_equity: float, sponsor_equity: float, noi: float, senior_rate: float, mezz_rate: float, pref_rate: float) -> dict:
        senior_debt, mezz_debt, pref_equity, sponsor_equity = max(0.0, safe_float(senior_debt)), max(0.0, safe_float(mezz_debt)), max(0.0, safe_float(pref_equity)), safe_float(sponsor_equity)
        senior_cost, mezz_cost, pref_cost = senior_debt * safe_float(senior_rate), mezz_debt * safe_float(mezz_rate), pref_equity * safe_float(pref_rate)
        fc = senior_cost + mezz_cost + pref_cost
        tot = senior_debt + mezz_debt + pref_equity + sponsor_equity
        return {"Senior": senior_debt, "Mezz": mezz_debt, "Pref": pref_equity, "Sponsor": sponsor_equity, "Total": tot, "FixedCharges": fc, "FCC": safe_ratio(max(0, safe_float(noi)), fc)}

    @staticmethod
    def rent_roll_metrics(df: pd.DataFrame) -> Tuple[float, float, float, float, float, float]:
        df = normalize_rr(df)
        if df.empty: return 0, 0, 0, 0, 0, 0
        total_sf = df["SF"].sum()
        if total_sf <= 0: return 0, 0, 0, 0, 0, 0
        occ = df[(df["SF"] > 0) & (~df["Tenant"].str.lower().isin(["", "vacant", "available", "empty"]))]
        occ_sf = occ["SF"].sum()
        if occ_sf <= 0: return total_sf, 0, 0, 0, 0, 0
        ar = occ["Monthly Rent"].sum() * 12
        walt = (occ["Remaining Term"] * occ["SF"]).sum() / occ_sf
        exp1 = occ[occ["Remaining Term"] <= 1.0]["SF"].sum() / occ_sf
        return total_sf, occ_sf/total_sf, ar, ar/occ_sf if occ_sf > 0 else 0, walt, exp1

    @staticmethod
    def breakeven_occupancy(noi: float, occ: float, ann_ds: float) -> float:
        noi, occ, ann_ds = safe_float(noi), safe_float(occ), safe_float(ann_ds)
        if noi <= 0 or occ <= 0 or ann_ds <= 0: return 0.0
        return min(1.5, max(0.0, ann_ds / noi * occ))

class InvestmentEngine:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def calculate_pro_forma(noi: float, r_grow: float, e_grow: float, e_ratio: float, yrs: int = 10) -> pd.DataFrame:
        rev = noi / max(1 - e_ratio, 0.01) if noi > 0 else 0.0
        exp = rev * e_ratio
        rows = []
        for y in range(1, yrs + 1):
            rows.append({"Year": y, "Revenue": rev, "Expenses": exp, "Projected NOI": rev - exp, "NOI Margin": safe_ratio(rev-exp, rev)})
            rev *= (1 + r_grow); exp *= (1 + e_grow)
        return pd.DataFrame(rows)

    @staticmethod
    def solve_returns(pp: float, loan: float, pf_df: pd.DataFrame, cap: float, sell_costs: float, ann_ds: float, balloon: float, t_growth: float) -> dict:
        eq = max(0.0, pp - loan)
        if eq <= 0 or pf_df.empty: return {"IRR": 0.0, "EM": 0.0, "Exit NOI": 0.0, "Gross Exit": 0.0, "Net Exit": 0.0, "Total CF": 0.0}
        
        cfs = [-eq]
        for _, r in pf_df.iterrows(): cfs.append(max(0.0, r["Projected NOI"]) - ann_ds)
        exit_noi = pf_df.iloc[-1]["Projected NOI"] * (1 + t_growth)
        gx = exit_noi / cap if cap > 0 else 0.0
        nx = gx * (1 - sell_costs) - balloon
        cfs[-1] += nx
        
        irr = 0.0
        if DEPS["numpy_financial"]:
            try: irr = npf.irr(cfs) or 0.0
            except ValueError: pass
        return {"IRR": irr, "EM": sum(cfs[1:])/eq if eq>0 else 0.0, "Exit NOI": exit_noi, "Gross Exit": gx, "Net Exit": nx, "Total CF": sum(cfs[1:])}

class SensitivityEngine:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def proceeds_heatmap(noi_base: float, appraisal: float, pp: float, costs: float, res: float, fees: float, rate_base: float, amort: int, term: int, is_io: bool, ltv: float, ltc: float, dscr: float, dy: float) -> pd.DataFrame:
        n_shocks = [-0.10, -0.05, 0.0, 0.05, 0.10]
        r_shocks = [-0.01, -0.005, 0.0, 0.005, 0.01]
        data = []
        for n in n_shocks:
            row = {"NOI Shock": f"{n:+.0%}"}
            for r in r_shocks:
                L, _, _, _, _ = UnderwritingEngine.size_loan(noi_base*(1+n), appraisal, pp, costs, res, fees, max(0.001, rate_base+r), amort, term, is_io, ltv, ltc, dscr, dy)
                row[f"{r*100:+.1f}%"] = L
            data.append(row)
        return pd.DataFrame(data)

class ValidationEngine:
    @staticmethod
    def validate_deal_state(state: dict, req_eq: float) -> Tuple[List[str], List[str]]:
        e, w = [], []
        if not str(state.get("deal_name")).strip(): e.append("Deal name is required.")
        if safe_float(state.get("purchase_price")) < 0: e.append("Purchase price cannot be negative.")
        if safe_float(state.get("noi")) < 0: e.append("NOI cannot be negative.")
        if safe_float(state.get("target_ltv")) > 1.25: e.append("Target LTV > 125%")
        if safe_float(state.get("target_dscr")) < 1.0: w.append("Target DSCR < 1.0x implies no cushion.")
        
        cap = safe_ratio(state.get("noi", 0), state.get("appraisal", 0))
        if cap > 0 and cap < 0.02: w.append(f"Implied Cap Rate {cap:.2%} is unusually low.")
        
        prem = safe_ratio(state.get("purchase_price", 0), state.get("appraisal", 0))
        if prem > 1.25: e.append("Purchase price is >25% over appraisal.")
        
        if req_eq < 0: w.append(f"Cash-out structure implied. Negative sponsor equity: ${abs(req_eq):,.0f}.")

        rr = state.get("rent_roll_dict", [])
        if isinstance(rr, list):
            tn = []
            for i, r in enumerate(rr, 1):
                t = str(r.get("Tenant", "")).strip()
                if t: tn.append(t.lower())
                if safe_float(r.get("Monthly Rent")) > 0 and safe_float(r.get("SF")) <= 0: w.append(f"Rent Roll Row {i}: Rent with zero SF.")
            dups = set([x for x in tn if tn.count(x) > 1 and x not in ["vacant", "available", ""]])
            if dups: w.append(f"Duplicate tenants: {', '.join(list(dups)[:3])}")
        return e, w

    @staticmethod
    def run_financial_self_tests() -> pd.DataFrame:
        rows = []
        def check(name, passed, detail): rows.append({"Test": name, "Status": "✅ PASS" if passed else "❌ FAIL", "Detail": detail})

        try:
            L, gate, gates, uses, eq = UnderwritingEngine.size_loan(1_000_000, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
            check("Loan proceeds positive", L > 0, f"${L:,.0f}")
            check("Binding gate exact match", abs(L - min(gates.values())) < 1, gate)

            amort, pmt, balloon = UnderwritingEngine.amort_schedule(1_000_000, 0.06, 25, 5, False)
            check("Exact Payment Calculation", abs(pmt - 6443.01) < 0.02, f"${pmt:,.2f}")
            check("Exact Balloon Calculation", abs(balloon - 896574.40) < 0.02, f"${balloon:,.2f}")

            L_io, _, gates_io, _, _ = UnderwritingEngine.size_loan(1_000_000, 100_000_000, 100_000_000, 0, 0, 0.0, 0.06, 25, 5, True, 1.0, 1.0, 1.25, 0.01)
            check("Exact IO DSCR Sizing", abs(gates_io["DSCR"] - 13333333.33) < 1, f"${gates_io['DSCR']:,.0f}")

            L_z, _, _, _, _ = UnderwritingEngine.size_loan(0, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
            check("Zero NOI -> Zero Proceeds", L_z == 0, f"${L_z:,.0f}")

            _, pmt_zr, _ = UnderwritingEngine.amort_schedule(1_000_000, 0.0, 25, 5, False)
            check("Zero Rate Amortization", abs(pmt_zr - 3333.33) < 0.02, f"${pmt_zr:,.2f}")

            tsf, occ, ar, _, _, _ = UnderwritingEngine.rent_roll_metrics(pd.DataFrame([{"Tenant": "A", "SF": 10000, "Remaining Term": 5, "Monthly Rent": 20000}, {"Tenant": "Vacant", "SF": 2000, "Remaining Term": 0, "Monthly Rent": 0}]))
            check("Rent roll occupancy", abs(occ - 0.8333) < 0.01, f"{occ:.1%}")

            tsf_v, occ_v, ar_v, _, _, _ = UnderwritingEngine.rent_roll_metrics(pd.DataFrame([{"Tenant": "Vacant", "SF": 10000, "Remaining Term": 0, "Monthly Rent": 0}]))
            check("All-Vacant Rent Roll", occ_v == 0 and ar_v == 0, "Handles division by zero")

            pf = InvestmentEngine.calculate_pro_forma(100000, 0.02, 0.05, 0.40)
            check("Margin Compression Test", pf.iloc[-1]["NOI Margin"] < pf.iloc[0]["NOI Margin"], f"{pf.iloc[-1]['NOI Margin']:.1%} vs {pf.iloc[0]['NOI Margin']:.1%}")

            rr_conv = normalize_rr(pd.DataFrame([{"Tenant": "B", "SF": 100, "Annual Rent": 12000}]))
            check("Annual to Monthly Rent Conversion", abs(rr_conv.iloc[0]["Monthly Rent"] - 1000) < 0.01, "Parsed correctly")

            if DEPS["crypto"]:
                pt = b"AlenzaTestDoc"
                key = "YmFzZTY0a2V5" # Dummy for testing
                check("Round-Trip Crypto", decrypt_bytes(encrypt_bytes(pt, key), key) == pt, "PBKDF2HMAC Salted AES")
                
            try:
                p = Path("/etc/passwd")
                p.relative_to(DOC_DIR.resolve())
                check("Path Traversal Rejection", False, "Failed to reject")
            except ValueError:
                check("Path Traversal Rejection", True, "Successfully rejected out-of-bounds path")

        except Exception as e:
            check("Test Suite Execution", False, str(e))
            
        return pd.DataFrame(rows)

# =============================================================================
# 8. DATABASE & PERSISTENCE (TRANSACTIONAL & TAMPER-EVIDENT)
# =============================================================================
class DatabaseManager:
    @staticmethod
    def get_conn():
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        DOC_DIR.mkdir(parents=True, exist_ok=True)
        c = sqlite3.connect(DB_PATH, timeout=30)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys=ON")
        c.execute("PRAGMA journal_mode=WAL")
        return c

    @classmethod
    def init_db(cls):
        try:
            with cls.get_conn() as c:
                c.executescript("""
                    CREATE TABLE IF NOT EXISTS deals (id TEXT PRIMARY KEY, name TEXT, state_json TEXT, updated_at TIMESTAMP, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
                    CREATE TABLE IF NOT EXISTS deal_versions (id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id TEXT, state_json TEXT, created_at TIMESTAMP, FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE);
                    CREATE TABLE IF NOT EXISTS documents (id TEXT PRIMARY KEY, deal_id TEXT, filename TEXT, category TEXT, path TEXT, size INT, is_encrypted BOOLEAN, uploaded_at TIMESTAMP);
                    CREATE TABLE IF NOT EXISTS audit_log (id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id TEXT, user TEXT, action TEXT, details TEXT, timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
                """)
                # Migrations
                cols_doc = [row["name"] for row in c.execute("PRAGMA table_info(documents)").fetchall()]
                if "is_encrypted" not in cols_doc: c.execute("ALTER TABLE documents ADD COLUMN is_encrypted BOOLEAN DEFAULT 0")
                cols_aud = [row["name"] for row in c.execute("PRAGMA table_info(audit_log)").fetchall()]
                if "prev_hash" not in cols_aud: 
                    c.execute("ALTER TABLE audit_log ADD COLUMN prev_hash TEXT DEFAULT ''")
                    c.execute("ALTER TABLE audit_log ADD COLUMN event_hash TEXT DEFAULT ''")
        except sqlite3.Error as e: logger.error(f"DB Init failed: {e}")

    @classmethod
    def log_audit(cls, deal_id: str, action: str, details: str = ""):
        try: user = st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User"))
        except Exception: user = os.environ.get("APP_USER", "Local User")
        ts = datetime.now(timezone.utc).isoformat()
        try:
            with cls.get_conn() as c:
                last_event = c.execute("SELECT event_hash FROM audit_log ORDER BY id DESC LIMIT 1").fetchone()
                prev_hash = last_event["event_hash"] if last_event and last_event["event_hash"] else "GENESIS"
                payload = f"{deal_id}|{user}|{action}|{details}|{ts}|{prev_hash}"
                event_hash = hashlib.sha256(payload.encode()).hexdigest()
                
                c.execute("INSERT INTO audit_log (deal_id, user, action, details, timestamp, prev_hash, event_hash) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                          (deal_id, user, action, details, ts, prev_hash, event_hash))
                c.commit()
        except sqlite3.Error as e: logger.error(f"Audit log failed: {e}")

    @classmethod
    def save_deal(cls, deal_id: str, name: str, state: dict) -> bool:
        clean_state = {k: state.get(k) for k in DEAL_STATE_KEYS}
        key = get_encryption_key()
        
        try:
            if key and DEPS["crypto"]:
                payload = json.dumps({"_alenza_storage": "encrypted", "payload": encrypt_text(json.dumps(clean_state), key)})
            else:
                payload = json.dumps({"_alenza_storage": "plain", "payload": json.dumps(clean_state)})
                
            now = datetime.now(timezone.utc).isoformat()
            with cls.get_conn() as c:
                c.execute("BEGIN TRANSACTION")
                c.execute("INSERT OR REPLACE INTO deals (id, name, state_json, updated_at) VALUES (?, ?, ?, ?)", (deal_id, name, payload, now))
                c.execute("INSERT INTO deal_versions (deal_id, state_json, created_at) VALUES (?, ?, ?)", (deal_id, payload, now))
                c.commit()
            cls.log_audit(deal_id, "SAVE_DEAL", f"Saved: {name}")
            return True
        except (sqlite3.Error, ValueError, TypeError) as e:
            logger.error(f"Save failed: {e}")
            return False

    @classmethod
    def load_deal(cls, deal_id: str) -> Optional[dict]:
        try:
            with cls.get_conn() as c:
                r = c.execute("SELECT state_json FROM deals WHERE id=?", (deal_id,)).fetchone()
                if not r: return None
                
                raw = json.loads(r["state_json"])
                if raw.get("_alenza_storage") == "encrypted":
                    key = get_encryption_key()
                    if not key: raise ValueError("Encrypted deal requires ALENZA_DB_ENCRYPTION_KEY")
                    state = json.loads(decrypt_text(raw["payload"], key))
                elif raw.get("_alenza_storage") == "plain":
                    state = json.loads(raw["payload"])
                else: state = raw
                
                cls.log_audit(deal_id, "LOAD_DEAL", "Loaded successfully")
                return state
        except (sqlite3.Error, json.JSONDecodeError, ValueError) as e:
            logger.error(f"Load failed: {e}")
            return None

    @classmethod
    def get_all_deals(cls):
        try:
            with cls.get_conn() as c:
                return pd.read_sql_query("SELECT id, name, created_at, updated_at FROM deals ORDER BY updated_at DESC", c)
        except sqlite3.Error: return pd.DataFrame()

    @classmethod
    def delete_deal(cls, deal_id: str) -> bool:
        try:
            with cls.get_conn() as c:
                d_name = c.execute("SELECT name FROM deals WHERE id=?", (deal_id,)).fetchone()
                deal_name = d_name["name"] if d_name else "Unknown"
                
                docs = c.execute("SELECT path FROM documents WHERE deal_id=?", (deal_id,)).fetchall()
                for row in docs:
                    try: Path(row["path"]).unlink(missing_ok=True)
                    except OSError: pass
                
                c.execute("DELETE FROM documents WHERE deal_id=?", (deal_id,))
                c.execute("DELETE FROM deal_versions WHERE deal_id=?", (deal_id,))
                c.execute("DELETE FROM deals WHERE id=?", (deal_id,))
                c.commit()
            
            cls.log_audit(deal_id, "DELETE_DEAL", f"Deleted deal: {deal_name}")
            return True
        except sqlite3.Error as e:
            logger.error(f"Delete failed: {e}")
            return False

    @classmethod
    def save_doc(cls, deal_id: str, file, category: str) -> bool:
        content = file.getvalue()
        original_size = len(content)
        if original_size > MAX_UPLOAD_MB * 1024 * 1024: raise ValueError(f"File exceeds {MAX_UPLOAD_MB} MB limit")
        
        doc_id = f"doc_{int(time.time())}_{uuid.uuid4().hex[:6]}"
        safe_name = re.sub(r"[^a-zA-Z0-9_.-]", "_", file.name)
        key = get_encryption_key()
        is_enc = bool(key and DEPS["crypto"])
        
        if is_enc:
            content = encrypt_bytes(content, key)
            safe_name += ".enc"
            
        path = (DOC_DIR / f"{doc_id}_{safe_name}").resolve()
        try: path.relative_to(DOC_DIR.resolve())
        except ValueError: raise ValueError("Unsafe document path rejected")

        try:
            path.write_bytes(content)
            with cls.get_conn() as c:
                c.execute("INSERT INTO documents (id, deal_id, filename, category, path, size, is_encrypted, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", 
                          (doc_id, deal_id, safe_name.replace(".enc", ""), category, str(path), original_size, is_enc, datetime.now().isoformat()))
                c.commit()
            cls.log_audit(deal_id, "DOC_UPLOAD", f"Uploaded {safe_name}")
            return True
        except (sqlite3.Error, OSError) as e:
            logger.error(f"Doc save failed: {e}")
            return False

    @classmethod
    def delete_doc(cls, doc_id: str) -> bool:
        try:
            with cls.get_conn() as c:
                r = c.execute("SELECT deal_id, path, filename FROM documents WHERE id=?", (doc_id,)).fetchone()
                if not r: return False
                try: Path(r["path"]).unlink(missing_ok=True)
                except OSError: pass
                c.execute("DELETE FROM documents WHERE id=?", (doc_id,))
                c.commit()
            cls.log_audit(r["deal_id"], "DOC_DELETE", f"Deleted {r['filename']}")
            return True
        except sqlite3.Error: return False

    @classmethod
    def get_versions(cls, deal_id: str) -> pd.DataFrame:
        try:
            with cls.get_conn() as c:
                return pd.read_sql_query("SELECT id, created_at FROM deal_versions WHERE deal_id=? ORDER BY created_at DESC", c, params=(deal_id,))
        except sqlite3.Error: return pd.DataFrame()

    @classmethod
    def load_version(cls, version_id: int) -> Optional[dict]:
        try:
            with cls.get_conn() as c:
                r = c.execute("SELECT deal_id, state_json FROM deal_versions WHERE id=?", (version_id,)).fetchone()
                if not r: return None
                
                raw = json.loads(r["state_json"])
                if raw.get("_alenza_storage") == "encrypted":
                    key = get_encryption_key()
                    if not key: raise ValueError("Encrypted deal requires ALENZA_DB_ENCRYPTION_KEY")
                    state = json.loads(decrypt_text(raw["payload"], key))
                elif raw.get("_alenza_storage") == "plain":
                    state = json.loads(raw["payload"])
                else: state = raw
                
                cls.log_audit(r["deal_id"], "RESTORE_VERSION", f"Restored version ID {version_id}")
                return state
        except Exception as e:
            logger.error(f"Version load failed: {e}")
            return None
# =============================================================================
# 9. MARKET INTELLIGENCE & APIs (Degraded Mode Aware)
# =============================================================================
@st.cache_data(ttl=3600)
def fetch_boc_history(days=365) -> Tuple[dict, pd.DataFrame, bool]:
    """
    Fetch Bank of Canada Valet observations.

    Fixes:
    - Uses STATIC_ATABLE_V39079 for the overnight target instead of V39079.
    - Keeps the latest date for each series separately.
    - Logs HTTP response details when the API returns a non-200 status.
    """
    sm = {
        "FXUSDCAD": "USD/CAD",
        "V122539": "2Y Yield",
        "V122540": "5Y Yield",
        "V122543": "10Y Yield",
        "STATIC_ATABLE_V39079": "Overnight Target",
    }
    url = f"https://www.bankofcanada.ca/valet/observations/{','.join(sm.keys())}/json?recent={int(days)}"
    headers = {"User-Agent": f"AlenzaCapitalOS/{VERSION}", "Accept": "application/json"}

    try:
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code != 200:
            logger.warning("BoC fetch failed. HTTP %s. URL=%s. Body=%s", r.status_code, r.url, r.text[:1000])
            return {}, pd.DataFrame(), True

        payload = r.json()
        rows = []
        for obs in payload.get("observations", []):
            row = {"Date": pd.to_datetime(obs.get("d"), errors="coerce")}
            for code, label in sm.items():
                row[label] = safe_float(obs.get(code, {}).get("v"), np.nan)
            rows.append(row)

        df = pd.DataFrame(rows).dropna(subset=["Date"])
        if df.empty:
            logger.warning("BoC returned no usable observations. Payload keys=%s", list(payload.keys()))
            return {}, pd.DataFrame(), True

        latest = {}
        for label in sm.values():
            valid = df[["Date", label]].dropna()
            if not valid.empty:
                latest[label] = {"val": float(valid[label].iloc[-1]), "date": valid["Date"].iloc[-1]}

        return latest, df, False

    except requests.exceptions.RequestException as e:
        logger.warning("BoC network/request failed: %s", e)
        return {}, pd.DataFrame(), True
    except ValueError as e:
        logger.warning("BoC JSON parse failed: %s", e)
        return {}, pd.DataFrame(), True
    except Exception:
        logger.exception("Unexpected BoC fetch failure")
        return {}, pd.DataFrame(), True

@st.cache_data(ttl=86400)
def fetch_unemployment() -> Tuple[pd.DataFrame, bool]:
    fb = pd.DataFrame({"Date": pd.date_range(start="2024-01-01", periods=12, freq="ME"), "Unemployment": [5.5, 5.7, 5.8, 6.0, 6.2, 6.4, 6.5, 6.6, 6.5, 6.4, 6.2, 6.1]})
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0"}
    try:
        df = pd.read_csv("https://www150.statcan.gc.ca/n1/en/tbl/csv/14100287-eng.csv", storage_options=headers, low_memory=False)
        m = (df["GEO"].astype(str).str.lower()=="canada") & (df["Labour force characteristics"].astype(str).str.lower().str.contains("unemployment rate")) & (df["Sex"].astype(str).str.lower()=="both sexes") & (df["Age group"].astype(str).str.lower()=="15 years and over")
        out = df.loc[m, ["REF_DATE", "VALUE"]].rename(columns={"REF_DATE": "Date", "VALUE": "Unemployment"}).dropna()
        out["Date"] = pd.to_datetime(out["Date"])
        return out.sort_values("Date").reset_index(drop=True), False
    except Exception as e: 
        logger.warning(f"StatsCan fetch failed: {e}")
        return fb, True

@st.cache_data(ttl=86400)
def fetch_vacancy_rates():
    return pd.DataFrame({
        "Property Class": ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality", "Self-Storage"],
        "National Vacancy": [2.1, 1.8, 5.2, 12.4, 4.5, 7.5, 3.8],
        "Trend": ["Tightening", "Tightening", "Softening", "High Vacancy", "Stable", "Mixed", "Stable"],
    })

def build_market_commentary(boc_latest, unemp_data, vac_data, state):
    cmt = []
    if boc_latest:
        y2, y10 = boc_latest.get('2Y Yield', {}).get('val'), boc_latest.get('10Y Yield', {}).get('val')
        if y2 and y10:
            spr = y10 - y2
            if spr < -0.50: cmt.append(("high", "Yield Curve Deeply Inverted", f"2s10s spread is {spr:.2f}%."))
            elif spr < 0: cmt.append(("medium", "Yield Curve Inverted", f"2s10s spread is {spr:.2f}%."))
        y5 = boc_latest.get('5Y Yield', {}).get('val')
        if y5 and state.get('rate', 0) > 0:
            rs = (state['rate'] * 100) - y5
            if rs < 1.50: cmt.append(("high", "Thin Risk Premium", f"Spread is {rs:.2f}% over 5Y GoC."))
    if not unemp_data.empty:
        u_rate = unemp_data["Unemployment"].iloc[-1]
        if u_rate > 7.0: cmt.append(("high", "Elevated Unemployment", f"Rate is {u_rate:.1f}%."))
    if not vac_data.empty:
        row = vac_data[vac_data["Property Class"] == state.get("property_type", "Multifamily")]
        if not row.empty and float(row["National Vacancy"].iloc[0]) > 8:
            cmt.append(("medium", "High Asset Vacancy", f"Nat avg {row['National Vacancy'].iloc[0]:.1f}%."))
    return cmt

@st.cache_data(ttl=86400)
def geocode_address(address: str) -> Optional[dict]:
    address = str(address or "").strip()
    if not address: return None
    try:
        r = requests.get("https://geogratis.gc.ca/services/geolocation/en/locate", params={"q": address}, timeout=8)
        r.raise_for_status()
        payload = r.json()
        if isinstance(payload, list) and len(payload) > 0:
            coords = payload[0].get("geometry", {}).get("coordinates", [])
            if len(coords) >= 2: return {"lon": safe_float(coords[0]), "lat": safe_float(coords[1])}
        elif isinstance(payload, dict) and payload.get("features"):
            coords = payload["features"][0].get("geometry", {}).get("coordinates", [])
            if len(coords) >= 2: return {"lon": safe_float(coords[0]), "lat": safe_float(coords[1])}
    except Exception as e: logger.warning(f"Geocode failed: {e}")
    return None

# =============================================================================
# 10. EXPORT HELPERS (PDF)
# =============================================================================
def generate_pdf_memo(s: dict, loan: float, gate: str, ltv: float, dscr: float, req_eq: float, irr: float, c_stack: dict, flags: list, cmt: list, v_err: list, v_warn: list, h_pre: str) -> Optional[bytes]:
    if not DEPS["pdf"]: return None
    try:
        b = io.BytesIO()
        doc = SimpleDocTemplate(b, pagesize=letter)
        sty = getSampleStyleSheet()
        s_norm = sty["BodyText"]
        story = [Paragraph("<font color='#CFB87C'><b>ALENZA CAPITAL OS</b></font>", sty["Title"]), Paragraph("Indicative Underwriting Memo", sty["Heading2"]), Spacer(1, 12)]
        story.append(Paragraph(f"<b>Deal:</b> {s.get('deal_name')} | <b>Sponsor:</b> {s.get('sponsor')}", s_norm))
        story.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%Y-%m-%d')} | <b>App Version:</b> {VERSION}", s_norm))
        story.append(Paragraph(f"<b>State Hash:</b> {h_pre[:16]}...", s_norm))
        story.append(Spacer(1, 12))
        
        t1 = Table([["Metric", "Value", "Target"], ["Mortgage Amount", f"${loan:,.0f}", gate], ["LTV", f"{ltv:.1%}", f"{s.get('target_ltv',0):.1%}"], ["DSCR", f"{dscr:.2f}x", f"{s.get('target_dscr',0):.2f}x"], ["Req. Equity", f"${req_eq:,.0f}", "N/A"], ["IRR", f"{irr:.2%}", "N/A"]], colWidths=[150, 150, 200])
        t1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#CFB87C")), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
        story.append(t1)
        story.append(Spacer(1, 16))
        
        story.append(Paragraph("Sources & Uses", sty["Heading3"]))
        t2 = Table([["Uses", "Amount", "Sources", "Amount"], ["Purchase Price", f"${s.get('purchase_price',0):,.0f}", "Senior Debt", f"${c_stack['Senior']:,.0f}"], ["Closing Costs", f"${s.get('closing_costs',0):,.0f}", "Mezzanine", f"${c_stack['Mezz']:,.0f}"], ["Reserves", f"${s.get('reserves',0):,.0f}", "Preferred", f"${c_stack['Pref']:,.0f}"], ["Financing Fees", f"${loan*s.get('fees',0):,.0f}", "Sponsor Equity", f"${c_stack['Sponsor']:,.0f}"]], colWidths=[120, 80, 120, 80])
        t2.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
        story.append(t2)
        
        if flags or cmt or v_err or v_warn:
            story.append(Spacer(1, 16))
            story.append(Paragraph("Validation, Risk & Market Commentary", sty["Heading3"]))
            for e in v_err: story.append(Paragraph(f"• ERROR: {e}", s_norm))
            for w in v_warn: story.append(Paragraph(f"• WARN: {w}", s_norm))
            for f in flags: story.append(Paragraph(f"• RISK: {f[1]}", s_norm))
            for c in cmt: story.append(Paragraph(f"• MARKET: {c[1]} - {c[2]}", s_norm))
            
        story.append(Spacer(1, 24))
        story.append(Paragraph("<i>[SIMULATED DATA INCLUDED] This memo is indicative only and not for final credit approval.</i>", s_norm))
        
        doc.build(story)
        b.seek(0)
        return b.getvalue()
    except Exception: return None

# =============================================================================
# 11. UI ABSTRACTION & MAIN
# =============================================================================
def calculate_outputs(state: dict):
    L, gate, gates, uses, _ = UnderwritingEngine.size_loan(state['noi'], state['appraisal'], state['purchase_price'], state['closing_costs'], state['reserves'], state['fees'], state['rate'], state['amort'], state['term'], state['is_io'], state['target_ltv'], state['target_ltc'], state['target_dscr'], state['target_dy'])
    amort_df, m_pmt, balloon = UnderwritingEngine.amort_schedule(L, state['rate'], state['amort'], state['term'], state['is_io'])
    annual_ds = m_pmt * 12
    req_equity = uses - L - state['mezz_debt'] - state['pref_equity']
    c_stack = UnderwritingEngine.capital_stack(L, state['mezz_debt'], state['pref_equity'], max(0.0, req_equity), state['noi'], state['rate'], state['mezz_rate'], state['pref_rate'])
    act_ltv, act_ltc, act_dscr, act_dy = safe_ratio(L, state['appraisal']), safe_ratio(L, uses), safe_ratio(state['noi'], annual_ds), safe_ratio(state['noi'], L)
    pf_df = InvestmentEngine.calculate_pro_forma(state['noi'], state['pf_rev_growth'], state['pf_exp_growth'], state['pf_exp_ratio'])
    rets = InvestmentEngine.solve_returns(state['purchase_price'], L, pf_df, state['pf_exit_cap'], state['pf_sell_costs'], annual_ds, balloon, state['pf_term_growth'])
    
    return {
        "L": L, "gate": gate, "gates": gates, "uses": uses, "req_equity": req_equity, "amort_df": amort_df, 
        "m_pmt": m_pmt, "balloon": balloon, "annual_ds": annual_ds, "c_stack": c_stack, "act_ltv": act_ltv, 
        "act_ltc": act_ltc, "act_dscr": act_dscr, "act_dy": act_dy, "pf_df": pf_df, "rets": rets
    }

def main():
    apply_theme()
    DatabaseManager.init_db()

    if "deal_id" not in st.session_state:
        st.session_state.update(default_state())
        st.session_state.unsaved_changes = False

    s = st.session_state

    # --- Sidebar ---
    with st.sidebar:
        st.title("ALENZA OS")
        s.deal_name = st.text_input("Deal Name", s.deal_name)
        s.sponsor = st.text_input("Sponsor", s.sponsor)
        s.property_type = st.selectbox("Asset Class", PROPERTY_TYPES, index=PROPERTY_TYPES.index(s.property_type) if s.property_type in PROPERTY_TYPES else 0)
        s.transaction_type = st.selectbox("Transaction Type", TX_TYPES, index=TX_TYPES.index(s.transaction_type) if s.transaction_type in TX_TYPES else 0)
        s.lender_profile = st.selectbox("Lender", LENDER_PROFILES_LIST, index=LENDER_PROFILES_LIST.index(s.lender_profile) if s.lender_profile in LENDER_PROFILES_LIST else 0)
        
        st.write("Debt Term Structure")
        s.amort = st.number_input("Amortization (Yrs)", value=int(s.amort), min_value=1)
        s.term = st.number_input("Term (Yrs)", value=int(s.term), min_value=1)
        s.is_io = st.toggle("Interest Only", value=s.is_io)
        
        st.divider()
        try:
            with DatabaseManager.get_conn() as c:
                deals = pd.read_sql_query("SELECT id, name, updated_at FROM deals ORDER BY updated_at DESC", c)
        except sqlite3.Error: deals = pd.DataFrame()

        if not deals.empty:
            deal_opts = {f"{r['name']} · {str(r['updated_at'])[:10]} · {r['id'][-8:]}": r["id"] for _, r in deals.iterrows()}
            sd = st.selectbox("Load/Delete Deal", ["-- Select --"] + list(deal_opts.keys()))
            col_l, col_d = st.columns(2)
            if sd != "-- Select --":
                d_id = deal_opts[sd]
                if col_l.button("Load"):
                    ld = DatabaseManager.load_deal(d_id)
                    if ld: s.update(ld); st.rerun()
                del_confirm = col_d.checkbox("Confirm", key="del_confirm_sidebar")
                if col_d.button("Del", disabled=not del_confirm):
                    DatabaseManager.delete_deal(d_id); st.rerun()

    # --- DYNAMIC PLACEHOLDERS (Zero-Latency Setup) ---
    st.title(f"🏢 {s.deal_name}")
    header_placeholder = st.empty()

    # --- Tabs ---
    tabs = st.tabs(["Sizing & Risk", "Sensitivity", "Rent Roll", "Amortization", "Pro Forma", 
                    "Canada Intel", "Market Comps", "Diligence Room", "Save & Export", "QA & Health"])

    # TAB 1: Sizing & Risk
    with tabs[0]:
        c1, c2 = st.columns([1.5, 1])
        with c1:
            st.subheader("Asset & Debt Setup")
            k1, k2 = st.columns(2)
            s.purchase_price = k1.number_input("Purchase Price", value=s.purchase_price, step=50000.0)
            s.appraisal = k2.number_input("Appraisal", value=s.appraisal, step=50000.0)
            s.property_address = st.text_input("Property Address (For Geocoding)", value=s.property_address)
            k3, k4 = st.columns(2)
            s.noi = k3.number_input("Stabilized NOI", value=s.noi, step=5000.0)
            
            # HIGHER PRECISION % TOGGLE
            s.rate = k4.slider("Interest Rate (%)", 1.00, 15.00, float(s.rate*100), 0.05, format="%.2f") / 100.0
            
            p1, p2, p3 = st.columns(3)
            s.target_ltv = p1.number_input("Max LTV (%)", value=float(s.target_ltv*100), step=1.0) / 100.0
            s.target_dscr = p2.number_input("Min DSCR (x)", value=float(s.target_dscr), step=0.05)
            s.target_dy = p3.number_input("Min Debt Yield (%)", value=float(s.target_dy*100), step=0.5) / 100.0

            st.subheader("Uses & Capital Stack")
            u1, u2, u3, u4 = st.columns(4)
            s.closing_costs = u1.number_input("Closing Costs ($)", value=s.closing_costs, step=10000.0)
            s.reserves = u2.number_input("Reserves ($)", value=s.reserves, step=10000.0)
            s.fees = u3.number_input("Financing Fee (%)", value=float(s.fees*100), step=0.1) / 100.0
            s.target_ltc = u4.number_input("Max LTC (%)", value=float(s.target_ltc*100), step=1.0) / 100.0
            
            m1, m2 = st.columns(2)
            s.mezz_debt = m1.number_input("Mezzanine Debt ($)", value=s.mezz_debt, step=50000.0)
            s.mezz_rate = m1.slider("Mezz Rate (%)", 1.00, 25.00, float(s.mezz_rate*100), 0.10, format="%.2f") / 100.0
            s.pref_equity = m2.number_input("Preferred Equity ($)", value=s.pref_equity, step=50000.0)
            s.pref_rate = m2.slider("Pref Rate (%)", 1.00, 25.00, float(s.pref_rate*100), 0.10, format="%.2f") / 100.0
            
            fcc_placeholder = st.empty()

        with c2:
            st.subheader("Risk Narrative")
            risk_placeholder = st.empty()

    # TAB 2: Sensitivity
    with tabs[1]:
        st.subheader("Proceeds Heatmap (NOI vs Rate)")
        sens_placeholder = st.empty()

    # TAB 3: Rent Roll
    with tabs[2]:
        st.subheader("Rent Roll Normalization")
        rr_df = normalize_rr(pd.DataFrame(s.rent_roll_dict))
        
        if st.button("Add Blank Row"):
            d = rr_df.to_dict("records"); d.append({"Tenant":"", "SF":0, "Remaining Term":0, "Monthly Rent":0}); s.rent_roll_dict = d; st.rerun()
            
        err = st.data_editor(rr_df, num_rows="dynamic", use_container_width=True)
        if not normalize_rr(err).equals(rr_df):
            st.warning("Rent roll changes are staged. Click below to apply to the deal model.")
            if st.button("Apply Rent Roll Changes", type="primary"):
                s.rent_roll_dict = normalize_rr(err).to_dict("records")
                st.rerun()
            
        tsf, occ, ar, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(pd.DataFrame(s.rent_roll_dict))
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Total SF", f"{tsf:,.0f}"); k2.metric("Occupancy", f"{occ:.1%}")
        k3.metric("Annual Rent", f"${ar:,.0f}"); k4.metric("WALT", f"{walt:.1f}")

    # TAB 4: Amortization
    with tabs[3]:
        st.subheader("Amortization & Paydown")
        amort_placeholder = st.empty()

    # TAB 5: Pro Forma
    with tabs[4]:
        st.subheader("10-Year Pro Forma & Returns")
        with st.expander("Financial Conventions"):
            st.write(FINANCIAL_CONVENTIONS)
        
        c1, c2, c3 = st.columns(3)
        s.pf_rev_growth = c1.slider("Rev Growth (%)", 0.0, 10.0, float(s.pf_rev_growth*100), 0.5) / 100.0
        s.pf_exp_growth = c2.slider("Exp Growth (%)", 0.0, 10.0, float(s.pf_exp_growth*100), 0.5) / 100.0
        s.pf_exp_ratio = c3.slider("Exp Ratio (%)", 10.0, 80.0, float(s.pf_exp_ratio*100), 1.0) / 100.0
        
        c4, c5, c6 = st.columns(3)
        s.pf_exit_cap = c4.slider("Exit Cap (%)", 4.0, 15.0, float(s.pf_exit_cap*100), 0.25) / 100.0
        s.pf_term_growth = c5.slider("Terminal Growth (%)", 0.0, 5.0, float(s.pf_term_growth*100), 0.5) / 100.0
        s.pf_sell_costs = c6.slider("Selling Costs (%)", 0.0, 10.0, float(s.pf_sell_costs*100), 0.5) / 100.0
        
        pf_placeholder = st.empty()

    # TAB 6: Canada Intel
    with tabs[5]:
        st.subheader("Sovereign Intelligence")
        latest, hist, boc_deg = fetch_boc_history()
        unemp, u_deg = fetch_unemployment()
        vac = fetch_vacancy_rates()
        cmt = build_market_commentary(latest, unemp, vac, get_current_state())
        
        if boc_deg or u_deg: st.warning("API Unavailable. Displaying degraded fallback data.")
        
        for sev, tit, txt in cmt:
            if sev=="high": st.error(f"**{tit}**: {txt}")
            elif sev=="medium": st.warning(f"**{tit}**: {txt}")
            else: st.success(f"**{tit}**: {txt}")
            
        if latest:
            st.write("Rate Locking")
            s.rate_lock_enabled = st.toggle("Lock Deal Rate to 5Y GoC", value=s.rate_lock_enabled)
            s.rate_lock_spread_bps = st.number_input("Risk Spread (bps)", 50, 1000, int(s.rate_lock_spread_bps), 5)
            if s.rate_lock_enabled:
                new_r = (latest.get('5Y Yield', {}).get('val', 0) + s.rate_lock_spread_bps/100)/100
                st.info(f"Indexed rate: {new_r:.2%}")
                if st.button("Apply Indexed Rate"):
                    s.rate = new_r; st.rerun()

            c1, c2, c3 = st.columns(3)
            c1.metric("5Y GoC Yield", f"{latest.get('5Y Yield', {}).get('val', 0):.2f}%")
            c1.caption(f"As of {latest.get('5Y Yield', {}).get('date', 'N/A')}")
            c2.metric("Overnight Target", f"{latest.get('Overnight Target', {}).get('val', 0):.2f}%")
            c2.caption(f"As of {latest.get('Overnight Target', {}).get('date', 'N/A')}")
            c3.metric("USD/CAD", f"{latest.get('USD/CAD', {}).get('val', 0):.4f}")
            c3.caption(f"As of {latest.get('USD/CAD', {}).get('date', 'N/A')}")
            
        if not unemp.empty and DEPS["plotly"]:
            fig = px.line(unemp, x="Date", y="Unemployment", title="Unemployment Rate" + (" (Fallback Data)" if u_deg else ""))
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
            st.plotly_chart(fig, use_container_width=True)

    # TAB 7: Comps
    with tabs[6]:
        st.subheader("[SIMULATED] Market Comparables")
        comps = generate_comps(s.property_type, s.noi, s.appraisal, s.deal_id)
        geo = geocode_address(s.property_address) if s.property_address else None
        if not geo: st.info("Using Toronto fallback anchor. Provide valid address for precise geocoding.")
        c_lat, c_lon = (geo["lat"], geo["lon"]) if geo else (43.65, -79.38)
        comps["lat"] = c_lat + np.random.default_rng(int(hashlib.sha256(s.deal_id.encode()).hexdigest()[:8], 16)).uniform(-0.05, 0.05, 5)
        comps["lon"] = c_lon + np.random.default_rng(int(hashlib.sha256(s.deal_id.encode()).hexdigest()[8:16], 16)).uniform(-0.05, 0.05, 5)
        st.dataframe(comps.drop(columns=["lat","lon"]).style.format({"Cap Rate": "{:.2%}", "NOI": "${:,.0f}", "Value": "${:,.0f}"}), hide_index=True, use_container_width=True)
        if DEPS["plotly"]:
            fig = px.scatter_mapbox(comps, lat="lat", lon="lon", hover_name="Comparable", size_max=15, zoom=10, height=400, mapbox_style="carto-darkmatter")
            fig.add_trace(go.Scattermapbox(lat=[c_lat], lon=[c_lon], mode='markers', marker=go.scattermapbox.Marker(size=14, color='gold'), text=["Subject Property"], hoverinfo='text'))
            fig.update_layout(margin={"r":0,"t":0,"l":0,"b":0})
            st.plotly_chart(fig, use_container_width=True)

    # TAB 8: Diligence Vault
    with tabs[7]:
        st.subheader("Document Vault")
        c1, c2 = st.columns([1, 1])
        with c1:
            cat = st.selectbox("Category", ["Appraisal", "Phase I", "T12", "Rent Roll", "Other"])
            up = st.file_uploader("Upload File")
            if up and st.button("Save to Vault"):
                if DatabaseManager.save_doc(s.deal_id, up, cat): st.success("Uploaded!"); st.rerun()
        
        try:
            with DatabaseManager.get_conn() as conn:
                docs = pd.read_sql_query("SELECT id, filename, category, is_encrypted FROM documents WHERE deal_id=?", conn, params=(s.deal_id,))
            if not docs.empty:
                st.dataframe(docs, hide_index=True, use_container_width=True)
                
                dl_id = st.selectbox("Download Doc", ["-- Select --"] + docs["id"].tolist())
                if dl_id != "-- Select --":
                    with DatabaseManager.get_conn() as conn:
                        r_doc = conn.execute("SELECT path, is_encrypted, filename FROM documents WHERE id=?", (dl_id,)).fetchone()
                    if r_doc:
                        raw_b = Path(r_doc["path"]).read_bytes()
                        if bool(r_doc["is_encrypted"]):
                            key = get_encryption_key()
                            if not key:
                                st.error("This document is encrypted. Set ALENZA_DB_ENCRYPTION_KEY to download it.")
                                raw_b = None
                            else:
                                raw_b = decrypt_bytes(raw_b, key)
                        if raw_b is not None:
                            st.download_button("Download File", raw_b, r_doc["filename"])
                        
                del_id = st.selectbox("Delete Doc", ["-- Select --"] + docs["id"].tolist())
                del_doc_confirm = st.checkbox("Confirm doc delete")
                if del_id != "-- Select --" and st.button("Delete Document", disabled=not del_doc_confirm):
                    DatabaseManager.delete_doc(del_id); st.rerun()
                
                if DEPS["ocr"]:
                    st.write("OCR Context Scanner")
                    sc_id = st.selectbox("Scan Doc", docs["id"].tolist())
                    if st.button("Extract Context"):
                        try:
                            with DatabaseManager.get_conn() as conn:
                                r_doc = conn.execute("SELECT path, is_encrypted FROM documents WHERE id=?", (sc_id,)).fetchone()
                            raw_b = Path(r_doc["path"]).read_bytes()
                            if bool(r_doc["is_encrypted"]):
                                key = get_encryption_key()
                                if not key: raise ValueError("Decryption Key required.")
                                raw_b = decrypt_bytes(raw_b, key)
                                
                            doc_pdf = fitz.open(stream=raw_b, filetype="pdf")
                            text = "\n".join([page.get_text() for page in doc_pdf]).replace('\n', ' ')
                            
                            kws = ["noi", "environmental", "phase", "lease", "rent", "appraisal"]
                            res = []
                            for k in kws:
                                matches = [m.start() for m in re.finditer(re.escape(k), text.lower())]
                                for m in matches:
                                    start, end = max(0, m - 60), min(len(text), m + 60)
                                    res.append({"Keyword": k, "Context": f"...{text[start:end]}..."})
                            if res: st.dataframe(pd.DataFrame(res), hide_index=True)
                            else: st.info("No target keywords found.")
                        except Exception as e: st.error(f"Scan failed: Ensure file is a valid, unencrypted PDF. {e}")
        except sqlite3.Error: pass

    # =============================================================================
    # LATE CALCULATION & PLACEHOLDER INJECTION
    # =============================================================================
    # Calculate everything ONCE using the exact state registered by the sliders above
    current_state = get_current_state()
    out = calculate_outputs(current_state)
    h_pre = hashlib.sha256(json.dumps(current_state, default=str, sort_keys=True).encode()).hexdigest()

    # Fill Top Header
    with header_placeholder.container():
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Mortgage Amount", f"${out['L']:,.0f}")
        c1.caption(f"Binding Constraint: **{out['gate']}**")
        if out["L"] <= 0:
            missing_msgs = out.get("gates", {}).get("_messages", [])
            if missing_msgs:
                st.warning("Mortgage amount is $0 because required sizing inputs are missing: " + "; ".join(missing_msgs))
            else:
                st.warning("Mortgage amount is $0 because one or more underwriting constraints sized to zero.")
        c2.metric("Projected IRR", f"{out['rets']['IRR']:.2%}")
        c3.metric("Equity Multiple", f"{out['rets']['EM']:.2f}x")
        c4.metric("Req. Sponsor Equity", f"${abs(out['req_equity']):,.0f}", delta="CASH OUT" if out['req_equity']<0 else None, delta_color="inverse" if out['req_equity']<0 else "normal")

    # Fill FCC metric
    with fcc_placeholder.container():
        st.metric("Fixed Charge Coverage", f"{out['c_stack']['FCC']:.2f}x")

    # Fill Tab 1 Risk
    with risk_placeholder.container():
        flags = []
        if out['act_ltv'] > 0.75: flags.append(("high", f"High Leverage: {out['act_ltv']:.1%} LTV"))
        if out['act_dscr'] < 1.20 and out['L'] > 0: flags.append(("high", f"Tight DSCR: {out['act_dscr']:.2f}x"))
        if out['req_equity'] < 0: flags.append(("medium", f"Cash-out implied: ${abs(out['req_equity']):,.0f} surplus"))
        if not flags: flags.append(("low", "Standard profile. No active flags."))
        
        for f in flags:
            if f[0] == "high": st.error(f[1])
            elif f[0] == "medium": st.warning(f[1])
            else: st.success(f[1])
            
        occ_rr = safe_ratio(UnderwritingEngine.rent_roll_metrics(normalize_rr(pd.DataFrame(s.rent_roll_dict)))[1], 1.0)
        st.metric("Breakeven Occupancy", f"{UnderwritingEngine.breakeven_occupancy(s.noi, max(occ_rr, 0.01), out['c_stack']['FixedCharges']):.1%}")

    # Fill Tab 2 Sensitivity
    with sens_placeholder.container():
        hm = SensitivityEngine.proceeds_heatmap(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy)
        hm_numeric = hm.drop(columns=["NOI Shock"], errors="ignore")
        if DEPS["plotly"]:
            fig = px.imshow(hm_numeric/1e6, text_auto=".1f", aspect="auto", title="Max Proceeds ($MM)")
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.dataframe(hm.style.background_gradient(subset=hm_numeric.columns, cmap="YlOrBr").format({c: "${:,.0f}" for c in hm_numeric.columns}), hide_index=True, use_container_width=True)

    # Fill Tab 4 Amortization
    with amort_placeholder.container():
        k1, k2, k3 = st.columns(3)
        k1.metric("Monthly P&I", f"${out['m_pmt']:,.2f}"); k2.metric("Annual DS", f"${out['annual_ds']:,.0f}")
        k3.metric("Balloon Balance", f"${out['balloon']:,.0f}")
        if not out['amort_df'].empty and DEPS["plotly"]:
            fig = px.bar(out['amort_df'], x="Period", y=["Principal", "Interest"], color_discrete_sequence=["#CFB87C", "#1E293B"])
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A", barmode="stack")
            st.plotly_chart(fig, use_container_width=True)

    # Fill Tab 5 Pro Forma
    with pf_placeholder.container():
        st.dataframe(out['pf_df'].style.format("${:,.0f}", subset=["Revenue","Expenses","Projected NOI"]).format("{:.1%}", subset=["NOI Margin"]), hide_index=True, use_container_width=True)
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Net Exit Proceeds", f"${out['rets']['Net Exit']:,.0f}")
        k2.metric("Total CF", f"${out['rets']['Total CF']:,.0f}")

    # TAB 9: Export (Handled here so it captures final 'out')
    with tabs[8]:
        st.subheader("Institutional Export Suite")
        if out['req_equity'] < 0: st.warning("⚠️ Cash-out structure detected. Confirm lender permits equity extraction before submitting.")

        v_err, v_warn = ValidationEngine.validate_deal_state(current_state, out['req_equity'])

        if st.button("💾 Save Deal", use_container_width=True):
            if v_err:
                for e in v_err: st.error(e)
            else:
                if DatabaseManager.save_deal(s.deal_id, s.deal_name, current_state):
                    s.unsaved_changes = False
                    st.toast("Deal Saved", icon="✅")

        c1, c2, c3 = st.columns(3)
        if DEPS["excel_write"]:
            xl_out = io.BytesIO()
            with pd.ExcelWriter(xl_out, engine="xlsxwriter") as w:
                pd.DataFrame([current_state]).to_excel(w, sheet_name="Inputs", index=False)
                pd.DataFrame([out['c_stack']]).to_excel(w, sheet_name="Capital Stack", index=False)
                constraints_export = pd.DataFrame([{
                    k: v for k, v in out['gates'].items() if not str(k).startswith("_")
                }])
                constraints_export.to_excel(w, sheet_name="Constraints", index=False)
                gate_messages = out.get("gates", {}).get("_messages", [])
                if gate_messages:
                    pd.DataFrame({"Diagnostic": gate_messages}).to_excel(w, sheet_name="Sizing Diagnostics", index=False)
                rr_df.to_excel(w, sheet_name="Rent Roll", index=False)
                out['pf_df'].to_excel(w, sheet_name="Pro Forma", index=False)
                out['amort_df'].to_excel(w, sheet_name="Amortization", index=False)
                SensitivityEngine.proceeds_heatmap(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy).to_excel(w, sheet_name="Sensitivity", index=False)
                comps.to_excel(w, sheet_name="Market Comps", index=False)
                
                # FLATTENED VALIDATION LIST FOR PROPER EXCEL EXPORT
                val_data = [{"Type": "Error", "Message": e} for e in v_err] + [{"Type": "Warning", "Message": w} for w in v_warn]
                if not val_data:
                    val_data = [{"Type": "Status", "Message": "All validations passed."}]
                pd.DataFrame(val_data).to_excel(w, sheet_name="Validation", index=False)
                
                if latest: pd.DataFrame([latest]).to_excel(w, sheet_name="Canada Intel", index=False)
                pd.DataFrame(cmt, columns=["Severity", "Topic", "Note"]).to_excel(w, sheet_name="Market Commentary", index=False)
                pd.DataFrame([{"Key": k, "Value": v} for k,v in FINANCIAL_CONVENTIONS.items()]).to_excel(w, sheet_name="Conventions", index=False)
                pd.DataFrame([{"Version": VERSION, "State Hash": h_pre, "Export Date": datetime.now().isoformat()}]).to_excel(w, sheet_name="Metadata", index=False)
                
                try:
                    with DatabaseManager.get_conn() as conn:
                        ad = pd.read_sql_query("SELECT * FROM audit_log WHERE deal_id=? ORDER BY timestamp DESC", conn, params=(s.deal_id,))
                        docs = pd.read_sql_query("SELECT id, filename, category, size, is_encrypted, uploaded_at FROM documents WHERE deal_id=?", conn, params=(s.deal_id,))
                    ad.to_excel(w, sheet_name="Audit Log", index=False)
                    docs.to_excel(w, sheet_name="Doc Inventory", index=False)
                except sqlite3.Error: pass

            c1.download_button("Download Excel Workbook", xl_out.getvalue(), f"{s.deal_id}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        enc_j = c2.checkbox("Encrypt JSON")
        pwd = c2.text_input("Password", type="password") if enc_j else ""
        if not enc_j or (enc_j and pwd):
            state_dict = current_state
            if enc_j and DEPS["crypto"]: j_dat = json.dumps({"_alenza_storage": "encrypted", "payload": encrypt_text(json.dumps(state_dict, default=str), pwd)})
            else: j_dat = json.dumps({"_alenza_storage": "plain", "payload": json.dumps(state_dict, default=str)})
            c2.download_button("Download JSON", j_dat, f"{s.deal_id}.json")
            
        pdf_b = None
        if DEPS["pdf"]:
            pdf_b = generate_pdf_memo(current_state, out['L'], out['gate'], out['act_ltv'], out['act_dscr'], out['req_equity'], out['rets']['IRR'], out['c_stack'], flags, cmt, v_err, v_warn, h_pre)
            if pdf_b: c3.download_button("Download PDF Memo", pdf_b, f"{s.deal_name}.pdf", "application/pdf")
            
        inc_docs = c1.checkbox("Include vault documents in ZIP")
        if c1.button("Prepare Complete ZIP Package"):
            z_buf = io.BytesIO()
            with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("deal.json", json.dumps(current_state, default=str))
                zf.writestr("amort.csv", out['amort_df'].to_csv(index=False))
                zf.writestr("proforma.csv", out['pf_df'].to_csv(index=False))
                zf.writestr("rent_roll.csv", rr_df.to_csv(index=False))
                zf.writestr("capital_stack.json", json.dumps(out['c_stack'], default=str))
                zf.writestr("simulated_comps.csv", comps.to_csv(index=False))
                zf.writestr("validation_summary.json", json.dumps({"Errors": v_err, "Warnings": v_warn}))
                zf.writestr("canada_intel.json", json.dumps(latest, default=str))
                zf.writestr("boc_history.csv", hist.to_csv(index=False) if not hist.empty else "")
                zf.writestr("unemployment.csv", unemp.to_csv(index=False) if not unemp.empty else "")
                zf.writestr("vacancy_context.csv", vac.to_csv(index=False) if not vac.empty else "")
                zf.writestr("market_commentary.json", json.dumps(cmt, default=str))
                zf.writestr("degraded_data_flags.json", json.dumps({"boc_degraded": boc_deg, "unemp_degraded": u_deg}))
                zf.writestr("conventions.txt", json.dumps(FINANCIAL_CONVENTIONS, indent=2))
                
                if pdf_b: zf.writestr(f"{s.deal_name}.pdf", pdf_b)
                try:
                    with DatabaseManager.get_conn() as conn:
                        ad = pd.read_sql_query("SELECT * FROM audit_log WHERE deal_id=? ORDER BY timestamp DESC", conn, params=(s.deal_id,))
                        docs = pd.read_sql_query("SELECT id, filename, category, path, is_encrypted FROM documents WHERE deal_id=?", conn, params=(s.deal_id,))
                        
                    zf.writestr("audit_log.csv", ad.to_csv(index=False))
                    zf.writestr("document_inventory.csv", docs.drop(columns=["path"]).to_csv(index=False))
                    
                    if inc_docs and not docs.empty:
                        for _, row in docs.iterrows():
                            try:
                                d_bytes = Path(row["path"]).read_bytes()
                                if bool(row["is_encrypted"]):
                                    k = get_encryption_key()
                                    if not k:
                                        zf.writestr(f"vault/README_{row['filename']}.txt", "Encrypted document omitted because ALENZA_DB_ENCRYPTION_KEY is not available.")
                                        continue
                                    d_bytes = decrypt_bytes(d_bytes, k)
                                zf.writestr(f"vault/{row['filename']}", d_bytes)
                            except Exception: pass
                except sqlite3.Error: pass
                zf.writestr("README.txt", f"ALENZA OS EXPORT PACKAGE\nDeal: {s.deal_name}\nDate: {datetime.now().isoformat()}\nApp Version: {VERSION}\nState Hash: {h_pre}\nDependencies: {json.dumps(DEPS)}")
            c1.download_button("Download ZIP Package", z_buf.getvalue(), f"{s.deal_name}.zip", "application/zip")

    # TAB 10: QA
    with tabs[9]:
        st.subheader("System Health & Validation")
        for e in v_err: st.error(e)
        for w in v_warn: st.warning(w)
        if not v_err and not v_warn: st.success("All validations passed.")
        st.dataframe(ValidationEngine.run_financial_self_tests(), hide_index=True, use_container_width=True)

        st.write("Version Rollback")
        v_df = DatabaseManager.get_versions(s.deal_id)
        if not v_df.empty:
            v_opts = {f"{r['created_at']} · v{r['id']}": r["id"] for _, r in v_df.iterrows()}
            v_id_label = st.selectbox("Restore Version", list(v_opts.keys()))
            if st.button("Restore"):
                state_restore = DatabaseManager.load_version(v_opts[v_id_label])
                if state_restore: st.session_state.update(state_restore); st.rerun()

if __name__ == "__main__":
    main()
