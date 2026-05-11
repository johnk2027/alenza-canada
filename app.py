"""
Alenza Capital OS v3.3.0 - The Ultimate Edition
"""

from __future__ import annotations

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
from dataclasses import dataclass, field
from functools import wraps
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Callable

import numpy as np
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

# =============================================================================
# OPTIONAL DEPENDENCIES
# =============================================================================
@st.cache_resource
def check_dependencies() -> Dict[str, bool]:
    deps = {
        "ocr": False, "pdf": False, "excel_write": False, "excel_read": False,
        "xls_read": False, "plotly": False, "crypto": False, "numpy_financial": False,
    }
    try: import fitz; from PIL import Image; import pytesseract; deps["ocr"] = True
    except ImportError: pass
    try: from reportlab.lib.pagesizes import letter; from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle; from reportlab.lib.styles import getSampleStyleSheet; from reportlab.lib import colors; deps["pdf"] = True
    except ImportError: pass
    try: import xlsxwriter; deps["excel_write"] = True
    except ImportError: pass
    try: import openpyxl; deps["excel_read"] = True
    except ImportError: pass
    try: import xlrd; deps["xls_read"] = True
    except ImportError: pass
    try: import plotly.express as px; import plotly.graph_objects as go; from plotly.subplots import make_subplots; deps["plotly"] = True
    except ImportError: pass
    try: from cryptography.fernet import Fernet; deps["crypto"] = True
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

# =============================================================================
# CONFIGURATION
# =============================================================================

@dataclass(frozen=True)
class AppConfig:
    VERSION: str = "3.3.0"
    SCHEMA_VERSION: int = 4
    MAX_UPLOAD_MB: int = 50
    DATA_DIR: Path = Path("alenza_data")
    DB_PATH: Path = Path("alenza_data/alenza_platform.db")
    DOC_DIR: Path = Path("alenza_data/documents")
    
    PROPERTY_TYPES: List[str] = field(default_factory=lambda: [
        "Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality", "Self-Storage"
    ])
    TRANSACTION_TYPES: List[str] = field(default_factory=lambda: [
        "Acquisition", "Refinance", "Construction", "Bridge", "Recapitalization"
    ])

CFG = AppConfig()
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DEFAULT_RENT_ROLL = [{"Tenant": "Main Anchor", "SF": 25000, "Remaining Term": 5.5, "Monthly Rent": 45000}]
DEFAULT_RENT_COLS = ["Tenant", "SF", "Remaining Term", "Monthly Rent"]

# =============================================================================
# UTILITIES & SECURITY
# =============================================================================

def safe_float(v: Any, default: float = 0.0) -> float:
    if v is None or v == "": return default
    try:
        if isinstance(v, (int, float, np.number)): return float(v)
        if isinstance(v, str):
            c = re.sub(r"[^\d.\-()]", "", v).strip()
            if c.startswith("(") and c.endswith(")"): c = "-" + c[1:-1]
            return float(c) if c else default
    except: pass
    return default

def get_encryption_key() -> Optional[str]:
    return st.secrets.get("ALENZA_DB_ENCRYPTION_KEY") or os.environ.get("ALENZA_DB_ENCRYPTION_KEY")

def encrypt_text(pt: str, key: str) -> str:
    import base64
    k = base64.urlsafe_b64encode(hashlib.sha256(key.encode()).digest())
    from cryptography.fernet import Fernet
    return base64.b64encode(Fernet(k).encrypt(pt.encode())).decode()

def decrypt_text(et: str, key: str) -> str:
    import base64
    k = base64.urlsafe_b64encode(hashlib.sha256(key.encode()).digest())
    from cryptography.fernet import Fernet
    return Fernet(k).decrypt(base64.b64decode(et)).decode()

def safe_ratio(num, den):
    num, den = safe_float(num), safe_float(den)
    return num/den if den > 0 else 0.0

# =============================================================================
# FINANCIAL ENGINES
# =============================================================================

class UnderwritingEngine:
    LENDER_PROFILES = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05},
    }

    @staticmethod
    def size_loan(noi, appraisal, purchase_price, closing_costs, reserves, fees_pct, 
                  rate, amort, term, is_io, target_ltv, target_ltc, target_dscr, target_dy):
        noi, appraisal = max(0, safe_float(noi)), max(0, safe_float(appraisal))
        hard_costs = safe_float(purchase_price) + safe_float(closing_costs) + safe_float(reserves)
        loan, total_uses = 0.0, hard_costs
        gates = {"LTV": 0.0, "LTC": 0.0, "DSCR": 0.0, "Debt Yield": 0.0}

        for _ in range(15):
            total_uses = hard_costs + (loan * fees_pct)
            gates["LTV"] = appraisal * target_ltv
            gates["LTC"] = total_uses * target_ltc
            gates["Debt Yield"] = noi / target_dy if target_dy > 0 else 0
            
            m_rate = rate / 12
            if m_rate > 0 and noi > 0:
                if is_io:
                    gates["DSCR"] = (noi / target_dscr) / rate if rate > 0 else 0
                else:
                    pmt_factor = (1 - (1 + m_rate) ** -(amort * 12)) / m_rate
                    gates["DSCR"] = (noi / target_dscr) / 12 * pmt_factor if pmt_factor > 0 else 0
            
            new_loan = max(0.0, min(gates.values()))
            if abs(new_loan - loan) < 0.01: break
            loan = new_loan

        return loan, min(gates, key=gates.get) if gates else "N/A", gates, total_uses, total_uses - loan

    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io):
        loan_amt, rate = max(0, safe_float(loan_amt)), max(0, safe_float(rate))
        if loan_amt <= 0: return pd.DataFrame(columns=["Period","Payment","Principal","Interest","Balance"]), 0.0, 0.0
        
        m_rate = rate / 12
        term_m = int(term_yrs * 12)
        m_pmt = loan_amt * m_rate if is_io else (loan_amt * m_rate) / (1 - (1 + m_rate)**-(amort_yrs * 12)) if m_rate > 0 else loan_amt/(amort_yrs*12)
        
        data, bal = [], loan_amt
        for m in range(1, term_m + 1):
            intr = bal * m_rate
            prin = 0.0 if is_io else min(m_pmt - intr, bal)
            bal = max(0, bal - prin)
            data.append({"Period": m, "Payment": m_pmt if bal>0 or prin>0 else 0, "Principal": prin, "Interest": intr, "Balance": bal})
            if bal <= 0: break
        return pd.DataFrame(data), m_pmt, bal

    @staticmethod
    def rent_roll_metrics(df):
        if df.empty: return 0, 0, 0, 0, 0, 0
        df["SF"] = pd.to_numeric(df["SF"], errors='coerce').fillna(0)
        df["Monthly Rent"] = pd.to_numeric(df["Monthly Rent"], errors='coerce').fillna(0)
        df["Remaining Term"] = pd.to_numeric(df["Remaining Term"], errors='coerce').fillna(0)
        tsf = df["SF"].sum()
        if tsf <= 0: return 0, 0, 0, 0, 0, 0
        occ_df = df[(~df["Tenant"].str.lower().isin(["vacant", "empty", "available", ""])) & (df["SF"] > 0)]
        osf = occ_df["SF"].sum()
        if osf <= 0: return tsf, 0, 0, 0, 0, 0
        ann_rent = occ_df["Monthly Rent"].sum() * 12
        walt = (occ_df["Remaining Term"] * occ_df["SF"]).sum() / osf
        exp1 = occ_df[occ_df["Remaining Term"] <= 1.0]["SF"].sum() / osf
        return tsf, osf/tsf, ann_rent, ann_rent/osf, walt, exp1

    @staticmethod
    def capital_stack(senior, mezz, pref, sponsor, noi, snr_rate, mezz_rate, pref_rate):
        fc = (senior * snr_rate) + (mezz * mezz_rate) + (pref * pref_rate)
        return {
            "Senior": senior, "Mezz": mezz, "Pref": pref, "Sponsor": sponsor,
            "Total": senior+mezz+pref+sponsor, "FixedCharges": fc, "FCC": noi/fc if fc > 0 else 0
        }

    @staticmethod
    def score_deal(ltv, ltc, dscr, dy, profile):
        lim = UnderwritingEngine.LENDER_PROFILES.get(profile, UnderwritingEngine.LENDER_PROFILES["Bank / Credit Union"])
        s_ltv = max(0, 300 * (1 - ltv / lim["max_ltv"])) if lim["max_ltv"] > 0 else 0
        s_dscr = max(0, 300 * min(1.5, (dscr - 1.0) / max(lim["min_dscr"] - 1.0, 0.01))) if dscr > 1.0 else 0
        s_dy = max(0, 200 * min(1.5, dy / lim["min_dy"])) if lim["min_dy"] > 0 else 0
        s_ltc = max(0, 200 * (1 - ltc))
        tot = min(1000, int(s_ltv + s_dscr + s_dy + s_ltc))
        tier = "Tier 1 | Core" if tot>=850 else "Tier 2 | Conventional" if tot>=700 else "Tier 3 | Debt Fund" if tot>=550 else "Tier 4 | Restructure"
        return tot, tier

class InvestmentEngine:
    @staticmethod
    def calculate_pro_forma(noi, r_grow, e_grow, e_ratio, yrs=10):
        rev = noi / max(1 - e_ratio, 0.01) if noi > 0 else 0
        exp = rev * e_ratio
        rows = []
        for y in range(1, yrs + 1):
            rows.append({"Year": y, "Revenue": rev, "Expenses": exp, "Projected NOI": rev - exp, "NOI Margin": (rev-exp)/rev if rev>0 else 0})
            rev *= (1 + r_grow)
            exp *= (1 + e_grow)
        return pd.DataFrame(rows)

    @staticmethod
    def solve_returns(pp, loan, pf_df, cap, sell_costs, ann_ds, balloon, t_growth):
        eq = max(0.0, pp - loan)
        if eq <= 0 or pf_df.empty: return {"IRR": 0.0, "EM": 0.0, "Exit NOI": 0.0, "Gross Exit": 0.0, "Net Exit": 0.0, "Total CF": 0.0}
        
        cfs = [-eq]
        for _, r in pf_df.iterrows(): cfs.append(max(0, r["Projected NOI"]) - ann_ds)
        
        exit_noi = pf_df.iloc[-1]["Projected NOI"] * (1 + t_growth)
        gx = exit_noi / cap if cap > 0 else 0
        nx = gx * (1 - sell_costs) - balloon
        cfs[-1] += nx
        
        irr = 0.0
        if DEPS["numpy_financial"]:
            try: irr = npf.irr(cfs) or 0.0
            except: pass
        
        total_dist = sum(cfs[1:])
        return {"IRR": irr, "EM": total_dist/eq if eq>0 else 0, "Exit NOI": exit_noi, "Gross Exit": gx, "Net Exit": nx, "Total CF": total_dist}

class SensitivityEngine:
    @staticmethod
    def proceeds_heatmap(state):
        rows = []
        for ns in [-0.10, -0.05, 0, 0.05, 0.10]:
            row = {"NOI Shock": f"{ns:+.0%}"}
            for rs in [-0.01, -0.005, 0, 0.005, 0.01]:
                L, _, _, _, _ = UnderwritingEngine.size_loan(
                    state['noi']*(1+ns), state['appraisal'], state['purchase_price'], state['closing_costs'], state['reserves'], state['fees'],
                    max(0.001, state['rate']+rs), state['amort'], state['term'], state['is_io'], state['target_ltv'], state['target_ltc'], state['target_dscr'], state['target_dy']
                )
                row[f"Rate {rs:+.1%}"] = L
            rows.append(row)
        return pd.DataFrame(rows)

# =============================================================================
# DATA PERSISTENCE & DB
# =============================================================================

class DatabaseManager:
    @staticmethod
    def get_conn():
        CFG.DATA_DIR.mkdir(parents=True, exist_ok=True)
        CFG.DOC_DIR.mkdir(parents=True, exist_ok=True)
        c = sqlite3.connect(CFG.DB_PATH, timeout=30)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys=ON")
        c.execute("PRAGMA journal_mode=WAL")
        return c

    @classmethod
    def init_db(cls):
        with cls.get_conn() as c:
            c.executescript("""
                CREATE TABLE IF NOT EXISTS deals (id TEXT PRIMARY KEY, name TEXT, state_json TEXT, updated_at TIMESTAMP, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE IF NOT EXISTS deal_versions (id INTEGER PRIMARY KEY AUTOINCREMENT, deal_id TEXT, state_json TEXT, created_at TIMESTAMP, FOREIGN KEY(deal_id) REFERENCES deals(id) ON DELETE CASCADE);
                CREATE TABLE IF NOT EXISTS documents (id TEXT PRIMARY KEY, deal_id TEXT, filename TEXT, category TEXT, path TEXT, size INT, uploaded_at TIMESTAMP);
            """)

    @classmethod
    def save_deal(cls, deal_id, name, state):
        payload = json.dumps(state, default=str)
        now = datetime.now(timezone.utc).isoformat()
        with cls.get_conn() as c:
            try:
                c.execute("BEGIN TRANSACTION")
                c.execute("INSERT OR REPLACE INTO deals (id, name, state_json, updated_at) VALUES (?, ?, ?, ?)", (deal_id, name, payload, now))
                c.execute("INSERT INTO deal_versions (deal_id, state_json, created_at) VALUES (?, ?, ?)", (deal_id, payload, now))
                c.commit()
                return True
            except Exception as e:
                c.execute("ROLLBACK")
                logger.error(f"DB Error: {e}")
                return False

    @classmethod
    def load_deal(cls, deal_id):
        with cls.get_conn() as c:
            r = c.execute("SELECT state_json FROM deals WHERE id=?", (deal_id,)).fetchone()
            if r:
                try: return json.loads(r["state_json"])
                except:
                    key = get_encryption_key()
                    if key: return json.loads(decrypt_text(json.loads(r["state_json"]).get("payload", ""), key))
        return None

    @classmethod
    def delete_deal(cls, deal_id):
        with cls.get_conn() as c:
            for r in c.execute("SELECT path FROM documents WHERE deal_id=?", (deal_id,)).fetchall():
                Path(r["path"]).unlink(missing_ok=True)
            c.execute("DELETE FROM deals WHERE id=?", (deal_id,))
            c.commit()
            return True

# =============================================================================
# MARKET INTELLIGENCE & APIs (RESTORED FULLY)
# =============================================================================

@st.cache_data(ttl=3600)
def fetch_boc_history(days=365):
    sm = {"FXUSDCAD": "USD/CAD", "BD.CDN.2YR.DQ.YLD": "2Y Yield", "BD.CDN.5YR.DQ.YLD": "5Y Yield", "BD.CDN.10YR.DQ.YLD": "10Y Yield", "V122514": "Overnight Target"}
    try:
        r = requests.get(f"https://www.bankofcanada.ca/valet/observations/{','.join(sm.keys())}/json?recent={days}", timeout=5).json()
        df = pd.DataFrame([{"Date": pd.to_datetime(o["d"])} | {l: safe_float(o.get(k, {}).get("v")) for k, l in sm.items()} for o in r.get("observations", [])])
        return {c: {"val": df[c].iloc[-1], "date": df["Date"].iloc[-1]} for c in df.columns if c!="Date"}, df
    except: return {}, pd.DataFrame()

@st.cache_data(ttl=86400)
def fetch_unemployment():
    fb = pd.DataFrame({"Date": pd.date_range(start="2023-01-01", periods=12, freq="ME"), "Unemployment": [5.5, 5.7, 5.8, 6.0, 6.2, 6.4, 6.5, 6.6, 6.5, 6.4, 6.2, 6.1]})
    try:
        df = pd.read_csv("https://www150.statcan.gc.ca/n1/en/tbl/csv/14100287-eng.csv", low_memory=False)
        m = (df["GEO"].str.lower()=="canada") & (df["Labour force characteristics"].str.lower().str.contains("unemployment rate")) & (df["Sex"].str.lower()=="both sexes") & (df["Age group"].str.lower()=="15 years and over")
        out = df.loc[m, ["REF_DATE", "VALUE"]].rename(columns={"REF_DATE": "Date", "VALUE": "Unemployment"}).dropna()
        out["Date"] = pd.to_datetime(out["Date"])
        return out.sort_values("Date").reset_index(drop=True)
    except: return fb

@st.cache_data(ttl=86400)
def fetch_vacancy_rates():
    """RESTORED: Simulated vacancy data matrix."""
    return pd.DataFrame({
        "Property Class": ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality", "Self-Storage"],
        "National Vacancy": [2.1, 1.8, 5.2, 12.4, 4.5, 7.5, 3.8],
        "Trend": ["Tightening", "Tightening", "Softening", "High Vacancy", "Stable", "Mixed", "Stable"],
        "YoY Change": [-0.3, -0.2, 0.8, 1.5, 0.1, 0.4, 0.0],
    })

@st.cache_data(ttl=86400)
def geocode_address(address: str):
    """RESTORED: NRCan Geocoding API integration."""
    try:
        if not address.strip(): return None
        r = requests.get(f"https://geogratis.gc.ca/services/geolocation/en/locate?q={requests.utils.quote(address)}", timeout=5)
        res = r.json()
        if res and len(res) > 0:
            coords = res[0].get("geometry", {}).get("coordinates", [0, 0])
            return {"lon": coords[0], "lat": coords[1]}
    except: pass
    return None

@st.cache_data(ttl=86400)
def verify_federal_corp(num_str):
    clean = re.sub(r"\D", "", str(num_str))
    if not clean: return None
    try:
        r = requests.get(f"https://www.ic.gc.ca/app/scr/cc/CorporationsCanada/api/corporations/{clean}.json?lang=eng", timeout=5)
        return r.json() if r.status_code == 200 else None
    except: return None

def build_market_commentary(boc_latest, unemp_data, vacancy_data, state):
    """RESTORED: Full automated market commentary generator."""
    cmt = []
    
    # 1. Yield Curve
    if boc_latest:
        y2 = boc_latest.get('2Y Yield', {}).get('val')
        y10 = boc_latest.get('10Y Yield', {}).get('val')
        if y2 and y10:
            spr = y10 - y2
            if spr < -0.50: cmt.append(("high", "Yield Curve Deeply Inverted", f"2s10s spread is {spr:.2f}%. Credit conditions should be underwritten conservatively."))
            elif spr < 0: cmt.append(("medium", "Yield Curve Inverted", f"2s10s spread is {spr:.2f}%. Stress rate and refinance assumptions."))
            else: cmt.append(("low", "Normal Yield Curve", f"2s10s spread is {spr:.2f}%. Term structure is positively sloped."))
            
        # Risk Premium
        y5 = boc_latest.get('5Y Yield', {}).get('val')
        if y5 and state.get('rate', 0) > 0:
            risk_spread = (state['rate'] * 100) - y5
            if risk_spread < 1.50: cmt.append(("high", "Thin Risk Premium", f"Deal spread is {risk_spread:.2f}% over the 5Y GoC. Verify asset quality justifies this margin."))
            elif risk_spread < 3.00: cmt.append(("low", "Standard Risk Premium", f"Deal spread is {risk_spread:.2f}% over the 5Y GoC."))

    # 2. Employment
    if not unemp_data.empty:
        u_rate = unemp_data["Unemployment"].iloc[-1]
        if u_rate > 7.0: cmt.append(("high", "Elevated Unemployment", f"Unemployment is {u_rate:.1f}%. Tenant demand risk may be elevated."))
        elif u_rate < 5.5: cmt.append(("low", "Strong Labour Market", f"Unemployment is {u_rate:.1f}%. Fundamentals are supportive."))

    # 3. Vacancy
    if not vacancy_data.empty:
        row = vacancy_data[vacancy_data["Property Class"] == state.get("property_type", "Multifamily")]
        if not row.empty:
            vac = float(row["National Vacancy"].iloc[0])
            trend = row["Trend"].iloc[0]
            if vac > 8: cmt.append(("high", f"High Vacancy ({trend})", f"National average for {state.get('property_type')} is {vac:.1f}%. Ensure local market offsets this macro risk."))
            else: cmt.append(("low", f"Stable Vacancy ({trend})", f"National average for {state.get('property_type')} is {vac:.1f}%."))

    return cmt

# =============================================================================
# EXPORT HELPERS
# =============================================================================

def generate_pdf_memo(s, loan, gate, ltv, dscr, req_eq, score, tier, irr, em, cmt_cache) -> Optional[bytes]:
    if not DEPS["pdf"]: return None
    try:
        b = io.BytesIO()
        doc = SimpleDocTemplate(b, pagesize=letter)
        sty = getSampleStyleSheet()
        s_norm = sty["BodyText"]
        story = [Paragraph("<font color='#CFB87C'><b>ALENZA CAPITAL OS</b></font>", sty["Title"]), Paragraph("Indicative Underwriting Memo", sty["Heading2"]), Spacer(1, 12)]
        story.append(Paragraph(f"<b>Deal:</b> {s.deal_name}", s_norm))
        story.append(Paragraph(f"<b>Sponsor:</b> {s.sponsor}", s_norm))
        story.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%Y-%m-%d')}", s_norm))
        story.append(Spacer(1, 12))
        
        t1 = Table([["Metric", "Value", "Target/Constraint"], ["Max Proceeds", f"${loan:,.0f}", gate], ["LTV", f"{ltv:.1%}", f"{s.target_ltv:.1%}"], ["DSCR", f"{dscr:.2f}x", f"{s.target_dscr:.2f}x"], ["Req. Equity", f"${req_eq:,.0f}", "N/A"], ["Score", f"{score}/1000", tier], ["IRR", f"{irr:.2%}", "N/A"]], colWidths=[150, 150, 200])
        t1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F172A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#CFB87C")), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey)]))
        story.append(t1)
        
        story.append(Spacer(1, 16))
        story.append(Paragraph("Market Intelligence", sty["Heading2"]))
        for c in cmt_cache: story.append(Paragraph(f"<b>{c[1]}:</b> {c[2]}", s_norm))
        
        doc.build(story)
        b.seek(0)
        return b.getvalue()
    except Exception as e:
        logger.error(f"PDF Gen Error: {e}")
        return None

# =============================================================================
# MAIN APP LOOP
# =============================================================================

def main():
    st.set_page_config(page_title="Alenza Capital OS", page_icon="🏛️", layout="wide")
    st.markdown("""<style>
        :root { --g: #CFB87C; --bg: #0B0F19; --fg: #0F172A; }
        .stApp { background: var(--bg); }
        [data-testid="stMetric"] { background: var(--fg); border: 1px solid rgba(207,184,124,0.3); padding: 15px; border-radius: 8px; }
        [data-testid="stMetricValue"] { color: var(--g); }
        .stTabs [aria-selected="true"] { border-bottom: 2px solid var(--g) !important; color: var(--g) !important; }
    </style>""", unsafe_allow_html=True)
    
    DatabaseManager.init_db()
    
    if "deal_id" not in st.session_state: 
        st.session_state.update({
            "deal_id": f"deal_{int(time.time())}_{uuid.uuid4().hex[:6]}", "deal_name": "Untitled Deal", "sponsor": "", "property_address": "",
            "property_type": "Multifamily", "transaction_type": "Acquisition", "lender_profile": "Bank / Credit Union",
            "purchase_price": 0.0, "appraisal": 0.0, "noi": 0.0, "rate": 0.055, "amort": 25, "term": 5, "is_io": False,
            "fees": 0.015, "closing_costs": 0.0, "reserves": 0.0, "target_ltv": 0.75, "target_ltc": 0.80, "target_dscr": 1.25, "target_dy": 0.08,
            "mezz_debt": 0.0, "pref_equity": 0.0, "mezz_rate": 0.11, "pref_rate": 0.09,
            "pf_rev_growth": 0.03, "pf_exp_growth": 0.02, "pf_exp_ratio": 0.40, "pf_exit_cap": 0.06, "pf_sell_costs": 0.015, "pf_term_growth": 0.02,
            "rent_roll_dict": DEFAULT_RENT_ROLL.copy(), "diligence_notes": "", "unsaved_changes": False
        })
    s = st.session_state

    h_pre = hashlib.sha256(json.dumps({k:v for k,v in s.items() if k not in ["unsaved_changes", "cmt_cache"]}, default=str).encode()).hexdigest()

    # --- Calculations ---
    L, gate, gates, uses, _ = UnderwritingEngine.size_loan(
        s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, 
        s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy
    )
    
    amort_df, m_pmt, balloon = UnderwritingEngine.amort_schedule(L, s.rate, s.amort, s.term, s.is_io)
    annual_ds = m_pmt * 12
    
    # Capital Stack Integration
    req_equity = uses - L - s.mezz_debt - s.pref_equity
    c_stack = UnderwritingEngine.capital_stack(L, s.mezz_debt, s.pref_equity, max(0, req_equity), s.noi, s.rate, s.mezz_rate, s.pref_rate)
    
    # Derived Actuals
    act_ltv = safe_ratio(L, s.appraisal)
    act_ltc = safe_ratio(L, uses)
    act_dscr = safe_ratio(s.noi, annual_ds)
    act_dy = safe_ratio(s.noi, L)
    
    score, tier = UnderwritingEngine.score_deal(act_ltv, act_ltc, act_dscr, act_dy, s.lender_profile)
    
    # Pro Forma
    pf_df = InvestmentEngine.calculate_pro_forma(s.noi, s.pf_rev_growth, s.pf_exp_growth, s.pf_exp_ratio)
    rets = InvestmentEngine.solve_returns(s.purchase_price, L, pf_df, s.pf_exit_cap, s.pf_sell_costs, annual_ds, balloon, s.pf_term_growth)

    # --- Sidebar ---
    with st.sidebar:
        st.title("ALENZA OS")
        st.caption(f"Enterprise v{CFG.VERSION}")
        s.deal_name = st.text_input("Deal Name", s.deal_name)
        s.property_type = st.selectbox("Asset Class", CFG.PROPERTY_TYPES, index=CFG.PROPERTY_TYPES.index(s.property_type) if s.property_type in CFG.PROPERTY_TYPES else 0)
        s.lender_profile = st.selectbox("Lender", CFG.LENDER_PROFILES, index=CFG.LENDER_PROFILES.index(s.lender_profile) if s.lender_profile in CFG.LENDER_PROFILES else 0)
        
        deals = DatabaseManager.get_all_deals()
        if not deals.empty:
            dl = st.selectbox("Load Deal", deals["name"].tolist() + ["-- Select --"], index=len(deals))
            col_l, col_d = st.columns(2)
            if dl != "-- Select --":
                d_id = deals[deals["name"]==dl].iloc[0]["id"]
                if col_l.button("Load"):
                    loaded = DatabaseManager.load_deal(d_id)
                    if loaded: st.session_state.update(loaded); st.rerun()
                if col_d.button("Del"):
                    DatabaseManager.delete_deal(d_id); st.rerun()

        if st.button("💾 Save Deal", use_container_width=True):
            if DatabaseManager.save_deal(s.deal_id, s.deal_name, {k:v for k,v in s.items()}):
                s.unsaved_changes = False
                st.toast("Deal Saved", icon="✅")

    # --- Header ---
    st.title(f"🏢 {s.deal_name}")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Max Proceeds", f"${L:,.0f}", help=f"Binding: {gate}")
    c2.metric("Projected IRR", f"{rets['IRR']:.2%}", help=f"Score: {score} | {tier}")
    c3.metric("Equity Multiple", f"{rets['EM']:.2f}x")
    c4.metric("Req. Sponsor Equity", f"${abs(req_equity):,.0f}", delta="CASH OUT" if req_equity<0 else None, delta_color="inverse" if req_equity<0 else "normal")

    # --- 10 Tabs ---
    tabs = st.tabs(["Sizing & Risk", "Sensitivity", "Rent Roll", "Amortization", "Pro Forma", 
                    "Canada Intel", "Market Comps", "Diligence Room", "Save & Export", "QA & Health"])

    # TAB 1: SIZING & RISK
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
            s.rate = k4.slider("Interest Rate", 0.01, 0.15, s.rate, 0.0025, format="%.4f")
            
            p1, p2, p3 = st.columns(3)
            s.target_ltv = p1.number_input("Max LTV", value=s.target_ltv, step=0.05)
            s.target_dscr = p2.number_input("Min DSCR", value=s.target_dscr, step=0.05)
            s.target_dy = p3.number_input("Min Debt Yield", value=s.target_dy, step=0.01)

            st.subheader("Capital Stack (Integrated)")
            m1, m2 = st.columns(2)
            s.mezz_debt = m1.number_input("Mezzanine Debt", value=s.mezz_debt, step=50000.0)
            s.pref_equity = m2.number_input("Preferred Equity", value=s.pref_equity, step=50000.0)
            st.dataframe(pd.DataFrame({
                "Tranche": ["Senior", "Mezzanine", "Preferred", "Sponsor", "Total"],
                "Amount": [f"${L:,.0f}", f"${s.mezz_debt:,.0f}", f"${s.pref_equity:,.0f}", f"${max(0, req_equity):,.0f}", f"${uses:,.0f}"],
                "Cost": [f"{s.rate:.2%}", f"{s.mezz_rate:.2%}", f"{s.pref_rate:.2%}", "N/A", "N/A"]
            }), hide_index=True, use_container_width=True)
            st.metric("Fixed Charge Coverage", f"{c_stack['FCC']:.2f}x")

        with c2:
            st.subheader("Risk Narrative")
            flags = []
            if act_ltv > 0.75: flags.append(("high", f"High Leverage: {act_ltv:.1%} LTV"))
            if act_dscr < 1.20 and L > 0: flags.append(("high", f"Tight DSCR: {act_dscr:.2f}x"))
            if req_equity < 0: flags.append(("medium", f"Cash-out implied: ${abs(req_equity):,.0f} surplus"))
            if not flags: flags.append(("low", "Standard profile. No active flags."))
            
            for f in flags:
                if f[0]=="high": st.error(f[1])
                elif f[0]=="medium": st.warning(f[1])
                else: st.success(f[1])
                
            occ_rr = safe_ratio(UnderwritingEngine.rent_roll_metrics(pd.DataFrame(s.rent_roll_dict))[1], 1.0)
            st.metric("Breakeven Occupancy", f"{UnderwritingEngine.breakeven_occupancy(s.noi, max(occ_rr, 0.01), c_stack['FixedCharges']):.1%}")

    # TAB 2: SENSITIVITY
    with tabs[1]:
        st.subheader("Custom Stress Test")
        c1, c2, c3 = st.columns(3)
        r_shk = c1.slider("Rate Shock (bps)", -200, 200, 0, 10) / 10000
        n_shk = c2.slider("NOI Shock (%)", -30, 30, 0, 5) / 100
        l_shk = c3.slider("LTV Target Adj (%)", -15, 15, 0, 5) / 100
        
        sl, sg, _, _, _ = UnderwritingEngine.size_loan(s.noi*(1+n_shk), s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, max(0.001, s.rate+r_shk), s.amort, s.term, s.is_io, max(0.01, s.target_ltv+l_shk), s.target_ltc, s.target_dscr, s.target_dy)
        _, sm_pmt, _ = UnderwritingEngine.amort_schedule(sl, max(0.001, s.rate+r_shk), s.amort, s.term, s.is_io)
        
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Stressed Proceeds", f"${sl:,.0f}", delta=f"${sl-L:,.0f}" if L>0 else None)
        k2.metric("Stressed Constraint", sg)
        k3.metric("Stressed DSCR", f"{safe_ratio(s.noi*(1+n_shk), sm_pmt*12):.2f}x")
        k4.metric("Stressed LTV", f"{safe_ratio(sl, s.appraisal):.1%}")

        st.subheader("Proceeds Heatmap (NOI vs Rate)")
        hm = SensitivityEngine.proceeds_heatmap(dict(s))
        if DEPS["plotly"]:
            fig = px.imshow(hm.set_index("NOI Shock")/1e6, text_auto=".1f", aspect="auto", title="Max Proceeds ($MM)")
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.dataframe(hm.style.background_gradient(cmap="YlOrBr", axis=None).format({c: "${:,.0f}" for c in hm.columns if c!="NOI Shock"}), hide_index=True, use_container_width=True)

    # TAB 3: RENT ROLL
    with tabs[2]:
        st.subheader("Rent Roll Normalization")
        up = st.file_uploader("Import Excel/CSV")
        if up:
            try:
                df = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
                s.rent_roll_dict = normalize_rr(df).to_dict("records")
                st.rerun()
            except Exception as e: st.error(f"Import err: {e}")
        
        rr_df = pd.DataFrame(s.rent_roll_dict)
        if st.button("Add Blank Row"):
            d = rr_df.to_dict("records"); d.append({"Tenant":"", "SF":0, "Remaining Term":0, "Monthly Rent":0}); s.rent_roll_dict = d; st.rerun()
            
        err = st.data_editor(rr_df, num_rows="dynamic", use_container_width=True)
        if not err.select_dtypes(include=[np.number]).apply(pd.to_numeric).equals(rr_df.select_dtypes(include=[np.number]).apply(pd.to_numeric)):
            s.rent_roll_dict = err.to_dict("records"); st.rerun()
            
        tsf, occ, ar, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(pd.DataFrame(s.rent_roll_dict))
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Total SF", f"{tsf:,.0f}"); k2.metric("Occupancy", f"{occ:.1%}")
        k3.metric("Annual Rent", f"${ar:,.0f}"); k4.metric("Rent PSF", f"${psf:.2f}")
        k5.metric("WALT", f"{walt:.1f}"); k6.metric("12M Rollover", f"{exp1:.1%}")

    # TAB 4: AMORTIZATION
    with tabs[3]:
        st.subheader("Amortization & Paydown")
        k1, k2, k3 = st.columns(3)
        k1.metric("Monthly P&I", f"${m_pmt:,.2f}")
        k2.metric("Annual DS", f"${annual_ds:,.0f}")
        k3.metric("Balloon Balance", f"${balloon:,.0f}", delta=f"{safe_ratio(balloon, L):.1%} of orig" if L>0 else None, delta_color="normal")
        
        c1, c2 = st.columns(2)
        with c1:
            st.write("Payment Structure")
            if not amort_df.empty and DEPS["plotly"]:
                fig = px.bar(amort_df, x="Period", y=["Principal", "Interest"], color_discrete_sequence=["#CFB87C", "#1E293B"])
                fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A", barmode="stack")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.write("Term vs Balloon Analysis")
            ts = [3,5,7,10,15]
            if s.term not in ts: ts.append(s.term)
            sd = []
            for t in sorted(ts):
                _, _, b = UnderwritingEngine.amort_schedule(L, s.rate, s.amort, t, s.is_io)
                sd.append({"Term": t, "Balloon": b, "Paydown": 1 - safe_ratio(b, L) if L>0 else 0})
            st.dataframe(pd.DataFrame(sd).style.format({"Balloon": "${:,.0f}", "Paydown": "{:.1%}"}), hide_index=True, use_container_width=True)
            
        with st.expander("View Annual Summary"):
            if not amort_df.empty:
                av = amort_df.copy(); av["Year"] = ((av["Period"]-1)//12)+1
                s_av = av.groupby("Year").agg({"Payment":"sum", "Principal":"sum", "Interest":"sum", "Balance":"last"}).reset_index()
                st.dataframe(s_av.style.format("${:,.2f}", subset=["Payment","Principal","Interest","Balance"]), hide_index=True, use_container_width=True)

    # TAB 5: PRO FORMA
    with tabs[4]:
        st.subheader("10-Year Pro Forma & Returns")
        c1, c2, c3, c4 = st.columns(4)
        s.pf_rev_growth = c1.slider("Rev Growth", 0.0, 0.1, s.pf_rev_growth, 0.005)
        s.pf_exp_growth = c2.slider("Exp Growth", 0.0, 0.1, s.pf_exp_growth, 0.005)
        s.pf_exit_cap = c3.slider("Exit Cap", 0.04, 0.12, s.pf_exit_cap, 0.0025)
        s.pf_term_growth = c4.slider("Terminal Growth", 0.0, 0.05, s.pf_term_growth, 0.005)
        
        st.dataframe(pf_df.style.format("${:,.0f}", subset=["Revenue","Expenses","Projected NOI"]).format("{:.1%}", subset=["NOI Margin"]), hide_index=True, use_container_width=True)
        
        md = pf_df.iloc[-1]["NOI Margin"] - pf_df.iloc[0]["NOI Margin"] if not pf_df.empty else 0
        if md < -0.02: st.warning(f"Margin compression: {md:.1%}")
        
        st.write("Returns Metrics")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Net Exit Proceeds", f"${rets['Net Exit']:,.0f}")
        k2.metric("Total CF", f"${rets['Total CF']:,.0f}")
        k3.metric("Levered IRR", f"{rets['IRR']:.2%}")
        k4.metric("Equity Multiple", f"{rets['EM']:.2f}x")
        
        if not pf_df.empty and DEPS["plotly"]:
            cd = pf_df.copy(); cd["CF"] = cd["Projected NOI"] - annual_ds
            fig = px.bar(cd, x="Year", y="CF", title="Annual Cash Flow (After DS)", color_discrete_sequence=["#CFB87C"])
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
            st.plotly_chart(fig, use_container_width=True)

    # TAB 6: CANADA INTEL
    with tabs[5]:
        st.subheader("Sovereign Intelligence")
        latest, hist = fetch_boc_history()
        vac = fetch_vacancy_rates()
        unemp = fetch_unemployment()
        cmt = build_market_commentary(latest, unemp, vac, dict(s))
        s.cmt_cache = cmt
        
        for sev, tit, txt in cmt:
            if sev=="high": st.error(f"**{tit}**: {txt}")
            elif sev=="medium": st.warning(f"**{tit}**: {txt}")
            else: st.success(f"**{tit}**: {txt}")
            
        if latest:
            st.write("Rate Locking")
            lock = st.toggle("Lock Deal Rate to 5Y GoC")
            spread_bps = st.number_input("Risk Spread (bps)", 50, 1000, 250, 5)
            if lock:
                y5 = latest['5Y Yield']['val']
                new_r = (y5 + spread_bps/100)/100
                if abs(s.rate - new_r) > 0.0001:
                    s.rate = new_r
                    st.rerun()

            c1, c2, c3 = st.columns(3)
            c1.metric("5Y GoC Yield", f"{latest['5Y Yield']['val']:.2f}%")
            c2.metric("Overnight Target", f"{latest['Overnight Target']['val']:.2f}%")
            c3.metric("USD/CAD", f"{latest['USD/CAD']['val']:.4f}")
            if not hist.empty and DEPS["plotly"]:
                fig = px.line(hist, x="Date", y=["2Y Yield", "5Y Yield", "10Y Yield"], title="Yield Curve Tracking")
                fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
                st.plotly_chart(fig, use_container_width=True)
                
        st.write("Federal Corp Registry Verification")
        cn = st.text_input("Corp Number (BN9)")
        if cn:
            res = verify_federal_corp(cn)
            if res: st.success("Found"); st.json(res)
            else: st.warning("Not Found")
            
        st.write("Simulated Vacancy (By Class)")
        st.dataframe(vac, hide_index=True, use_container_width=True)

    # TAB 7: COMPS
    with tabs[6]:
        st.subheader("Market Comparables (Simulated)")
        comps = generate_comps(s.property_type, s.noi, s.appraisal, s.deal_id)
        
        # NRCan Geocoding specific integration
        geo = geocode_address(s.property_address) if s.property_address else None
        c_lat, c_lon = (geo["lat"], geo["lon"]) if geo else (43.65, -79.38)
        
        comps["lat"] = c_lat + np.random.default_rng(abs(hash(s.deal_id))).uniform(-0.05, 0.05, 5)
        comps["lon"] = c_lon + np.random.default_rng(abs(hash(s.deal_id)+1)).uniform(-0.05, 0.05, 5)
        
        st.dataframe(comps.drop(columns=["lat","lon"]).style.format({"Cap": "{:.2%}", "NOI": "${:,.0f}"}), hide_index=True, use_container_width=True)
        if DEPS["plotly"]:
            fig = px.scatter_mapbox(comps, lat="lat", lon="lon", hover_name="Comp", size_max=15, zoom=10, height=400, mapbox_style="carto-darkmatter")
            if geo: fig.add_trace(go.Scattermapbox(lat=[c_lat], lon=[c_lon], mode='markers', marker=go.scattermapbox.Marker(size=14, color='gold'), text=["Subject Property"]))
            fig.update_layout(margin={"r":0,"t":0,"l":0,"b":0})
            st.plotly_chart(fig, use_container_width=True)

    # TAB 8: DILIGENCE
    with tabs[7]:
        st.subheader("Document Vault & OCR")
        c1, c2 = st.columns([1, 1])
        with c1:
            cat = st.selectbox("Category", ["Appraisal", "Phase I", "T12", "Rent Roll", "Other"])
            up = st.file_uploader("Upload File")
            if up and st.button("Save to Vault"):
                if DatabaseManager.save_doc(s.deal_id, up, cat): st.success("Uploaded!"); st.rerun()
        with c2:
            st.write("Gap Analysis")
            with DatabaseManager.get_conn() as conn:
                docs = pd.read_sql_query("SELECT id, filename, category FROM documents WHERE deal_id=?", conn, params=(s.deal_id,))
            uc = docs["category"].tolist() if not docs.empty else []
            st.dataframe(pd.DataFrame([{"Req": x, "Status": "✅" if x in uc else "❌"} for x in ["Appraisal", "Phase I", "T12", "Rent Roll"]]), hide_index=True, use_container_width=True)
            
        if not docs.empty:
            st.dataframe(docs, hide_index=True, use_container_width=True)
            dl_id = st.selectbox("Delete Doc", docs["id"].tolist() + ["-- Select --"], index=len(docs))
            if dl_id != "-- Select --" and st.button("Delete"): DatabaseManager.delete_doc(dl_id); st.rerun()
            
            if DEPS["ocr"]:
                st.write("OCR Scanner")
                sc_id = st.selectbox("Scan Doc", docs["id"].tolist())
                if st.button("Extract PDF Text"):
                    try:
                        with DatabaseManager.get_conn() as conn:
                            p = conn.execute("SELECT path FROM documents WHERE id=?", (sc_id,)).fetchone()["path"]
                        d = fitz.open(p)
                        t = "\n".join([page.get_text() for page in d])
                        st.text_area("Extracted Preview", t[:1000] + "...", height=200)
                        k_count = {k: t.lower().count(k) for k in ["noi", "environmental", "phase", "lease", "rent"]}
                        st.dataframe(pd.DataFrame([{"Keyword": k, "Mentions": v} for k,v in k_count.items()]), hide_index=True)
                    except Exception as e: st.error(f"Scan failed: Ensure file is a valid PDF. {e}")

    # TAB 9: SAVE & EXPORT (JSON Encrypt restored)
    with tabs[8]:
        st.subheader("Export Suite")
        c1, c2, c3 = st.columns(3)
        
        # Excel
        if DEPS["excel_write"]:
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="xlsxwriter") as w:
                pd.DataFrame([s]).to_excel(w, sheet_name="Inputs", index=False)
                amort_df.to_excel(w, sheet_name="Amortization", index=False)
                pf_df.to_excel(w, sheet_name="ProForma", index=False)
                rr_df.to_excel(w, sheet_name="RentRoll", index=False)
            c1.download_button("Excel (xlsx)", out.getvalue(), f"{s.deal_id}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else: c1.button("Excel", disabled=True)
            
        # JSON (Encrypted/Plain)
        enc_j = c2.checkbox("Encrypt JSON")
        pwd = c2.text_input("Password", type="password") if enc_j else ""
        if not enc_j or (enc_j and pwd):
            if enc_j and DEPS["crypto"]:
                j_dat = json.dumps({"_alenza_storage": "encrypted", "payload": encrypt_text(json.dumps(dict(s), default=str), pwd)})
                c2.download_button("JSON (Encrypted)", j_dat, f"{s.deal_id}_enc.json", "application/json")
            else:
                c2.download_button("JSON (Plain)", json.dumps(dict(s), default=str), f"{s.deal_id}.json", "application/json")
                
        # PDF / ZIP
        if DEPS["pdf"]:
            pdf_b = generate_pdf_memo(pd.Series(s), L, gate, act_ltv, act_dscr, req_equity, score, tier, rets['IRR'], rets['EM'], s.get("cmt_cache", []))
            if pdf_b: c3.download_button("PDF Memo", pdf_b, f"{s.deal_name}.pdf", "application/pdf")
        
        z_buf = io.BytesIO()
        with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("deal.json", json.dumps(dict(s), default=str))
            zf.writestr("amort.csv", amort_df.to_csv(index=False))
            zf.writestr("README.txt", f"ALENZA OS EXPORT\nDeal: {s.deal_name}\nDate: {datetime.now().isoformat()}")
        c3.download_button("ZIP Package", z_buf.getvalue(), f"{s.deal_name}.zip", "application/zip")

    # TAB 10: QA
    with tabs[9]:
        st.subheader("System Health")
        st.write(pd.DataFrame([{"Module": k, "Status": "Active" if v else "Missing"} for k,v in DEPS.items()]))
        with DatabaseManager.get_conn() as c:
            ad = pd.read_sql_query("SELECT user, action, details, timestamp FROM audit_log WHERE deal_id=? ORDER BY timestamp DESC LIMIT 20", c, params=(s.deal_id,))
        st.write("Audit Log"); st.dataframe(ad, hide_index=True, use_container_width=True)

    # Post-run hash check
    if SecurityToolkit.hash_payload({k: v for k, v in s.items() if k not in ["unsaved_changes", "cmt_cache"]}) != h_pre:
        s.unsaved_changes = True

if __name__ == "__main__":
    main()
