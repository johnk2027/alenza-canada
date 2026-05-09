"""
ALENZA CAPITAL OS - ENTERPRISE UNDERWRITING TERMINAL
Version: 7.0 (Production Master)
Theme: Midnight Slate & CU Gold (Unified Dark Mode)

Description:
This is the core application file for the Alenza Capital Underwriting OS. 
It features a modular, class-based architecture handling SQLite persistence, 
Canadian sovereign data integrations, OCR financial parsing, and comprehensive 
commercial real estate (CRE) debt sizing logic.
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import json
import requests
import zipfile
import sqlite3
import uuid
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional

# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================
st.set_page_config(page_title="Alenza Capital OS", page_icon="🏛️", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# DEPENDENCIES & GRACEFUL DEGRADATION
# ==========================================
try:
    from PIL import Image
    import pytesseract
    import fitz  # PyMuPDF
    OCR_AVAILABLE = True
except ImportError:
    Image = pytesseract = fitz = None
    OCR_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    PDF_AVAILABLE = True
except ImportError:
    SimpleDocTemplate = None
    PDF_AVAILABLE = False

# ==========================================
# 2. DIRECTORY & DATABASE ARCHITECTURE
# ==========================================
# Safely resolve app directory regardless of where the script is executed from
APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "alenza_data"
DB_PATH = DATA_DIR / "alenza_platform.db"
DOC_DIR = DATA_DIR / "documents"

# Ensure directories exist
for d in [DATA_DIR, DOC_DIR]:
    d.mkdir(parents=True, exist_ok=True)

def clean_filename(filename: str) -> str:
    """Sanitize filenames to prevent OS path issues."""
    return re.sub(r'[^a-zA-Z0-9_\-\.]', '_', filename.strip())

def extract_clean_state() -> dict:
    """Safely extract only JSON-serializable primitives from session state."""
    keys = ['deal_id', 'sponsor', 'property_address', 'property_type', 'transaction_type', 
            'lender_profile', 'purchase_price', 'appraisal', 'noi', 'target_ltv', 
            'target_ltc', 'target_dscr', 'target_dy', 'rate', 'amort', 'term', 
            'is_io', 'fees', 'closing_costs', 'reserves', 'rent_roll_dict']
    return {k: st.session_state[k] for k in keys if k in st.session_state}

class DatabaseManager:
    @staticmethod
    def init_db():
        with sqlite3.connect(DB_PATH) as conn:
            c = conn.cursor()
            c.execute('''CREATE TABLE IF NOT EXISTS deals 
                (id TEXT PRIMARY KEY, name TEXT, state_json TEXT, updated_at TIMESTAMP)''')
            c.execute('''CREATE TABLE IF NOT EXISTS audit_log 
                (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TIMESTAMP)''')
            c.execute('''CREATE TABLE IF NOT EXISTS documents 
                (id TEXT PRIMARY KEY, deal_id TEXT, filename TEXT, category TEXT, path TEXT, uploaded_at TIMESTAMP)''')
            conn.commit()

    @staticmethod
    def log_audit(action: str, details: str):
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute('INSERT INTO audit_log (user, action, details, timestamp) VALUES (?, ?, ?, ?)', 
                         ("Local User", action, details, datetime.now()))

    @staticmethod
    def save_deal(deal_id: str, name: str, state: dict):
        state_json = json.dumps(state)
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute('INSERT OR REPLACE INTO deals (id, name, state_json, updated_at) VALUES (?, ?, ?, ?)',
                         (deal_id, name, state_json, datetime.now()))
        DatabaseManager.log_audit("SAVE_DEAL", f"Saved Deal: {name}")

    @staticmethod
    def delete_deal(deal_id: str):
        """Deletes deal records AND purges associated physical documents to prevent storage leaks."""
        with sqlite3.connect(DB_PATH) as conn:
            c = conn.cursor()
            # 1. Find all physical files associated with this deal
            c.execute('SELECT path FROM documents WHERE deal_id = ?', (deal_id,))
            rows = c.fetchall()
            
            # 2. Delete physical files from the OS
            for row in rows:
                file_path = Path(row[0])
                if file_path.exists():
                    try:
                        file_path.unlink()
                    except OSError as e:
                        pass
            
            # 3. Delete database records
            conn.execute('DELETE FROM deals WHERE id = ?', (deal_id,))
            conn.execute('DELETE FROM documents WHERE deal_id = ?', (deal_id,))
            
        DatabaseManager.log_audit("DELETE_DEAL", f"Deleted Deal ID and associated files: {deal_id}")

    @staticmethod
    def load_deal(deal_id: str):
        with sqlite3.connect(DB_PATH) as conn:
            c = conn.cursor()
            c.execute('SELECT state_json FROM deals WHERE id = ?', (deal_id,))
            row = c.fetchone()
            return json.loads(row[0]) if row else None

    @staticmethod
    def get_all_deals():
        with sqlite3.connect(DB_PATH) as conn:
            return pd.read_sql_query("SELECT id, name, updated_at FROM deals ORDER BY updated_at DESC", conn)
            
    @staticmethod
    def save_document(deal_id: str, file, category: str):
        safe_name = clean_filename(file.name)
        doc_id = f"doc_{int(datetime.now().timestamp())}_{safe_name}"
        path = DOC_DIR / doc_id
        path.write_bytes(file.getbuffer())
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute('INSERT INTO documents (id, deal_id, filename, category, path, uploaded_at) VALUES (?, ?, ?, ?, ?, ?)',
                         (doc_id, deal_id, safe_name, category, str(path), datetime.now()))
        DatabaseManager.log_audit("DOC_UPLOAD", f"Uploaded {safe_name} to {category}")

    @staticmethod
    def delete_document(doc_id: str):
        with sqlite3.connect(DB_PATH) as conn:
            c = conn.cursor()
            c.execute('SELECT path, filename FROM documents WHERE id = ?', (doc_id,))
            row = c.fetchone()
            if row:
                file_path = Path(row[0])
                if file_path.exists():
                    file_path.unlink()
                conn.execute('DELETE FROM documents WHERE id = ?', (doc_id,))
                DatabaseManager.log_audit("DOC_DELETE", f"Deleted document: {row[1]}")

DatabaseManager.init_db()

# ==========================================
# 3. CANADIAN SOVEREIGN INTELLIGENCE
# ==========================================
class CanadianIntel:
    @staticmethod
    @st.cache_data(ttl=3600)
    def get_boc_rates():
        try:
            res = requests.get("https://www.bankofcanada.ca/valet/observations/FXUSDCAD,MCANR,MCANR5Y/json", timeout=5)
            if res.status_code == 200:
                obs = res.json().get("observations", [])[-1]
                return {
                    "overnight": float(obs.get("MCANR", {}).get("v", 0)),
                    "5yr_bond": float(obs.get("MCANR5Y", {}).get("v", 0)),
                    "usd_cad": float(obs.get("FXUSDCAD", {}).get("v", 0)),
                    "date": obs.get("d")
                }
        except: return None

    @staticmethod
    @st.cache_data(ttl=86400)
    def verify_corporation(name: str) -> dict:
        if not name: return {"status": "error", "message": "Empty query"}
        try:
            res = requests.get(f"https://ised-isde.canada.ca/opendata/corporations/corporations.json?q={name}", timeout=5)
            if res.status_code == 200:
                results = res.json().get('results')
                if results:
                    return {"status": "found", "data": results[0]}
                else:
                    return {"status": "not_found", "message": "Entity not found in active federal registry."}
            else:
                return {"status": "error", "message": f"Registry API unavailable (Status: {res.status_code})."}
        except: 
            return {"status": "error", "message": "Corporations Canada registry is currently unreachable."}
        
    @staticmethod
    @st.cache_data(ttl=86400)
    def geocode_nrcan(address: str):
        if not address: return None
        try:
            res = requests.get(f"https://geogratis.gc.ca/services/geolocation/en/locate?q={address}&limit=1", timeout=5)
            if res.status_code == 200 and res.json():
                data = res.json()[0]
                return f"{data.get('municipality')}, {data.get('provinceCode')}"
        except: return None

# ==========================================
# 4. OCR & DATA EXPORT ENGINES
# ==========================================
class OCREngine:
    @staticmethod
    def calculate_confidence(match: str, line: str, keyword: str) -> float:
        """Dynamically grades the probability that an extracted number is the right one."""
        confidence = 0.5
        k_pos = line.lower().find(keyword)
        m_pos = line.find(match)
        if k_pos >= 0 and m_pos >= 0:
            dist = abs(k_pos - m_pos)
            if dist < 20: confidence += 0.35
            elif dist < 50: confidence += 0.15
        if '$' in line or '%' in line: confidence += 0.05
        if ',' in match: confidence += 0.05
        return min(confidence, 0.98)

    @staticmethod
    def extract_and_parse(file) -> Tuple[str, dict]:
        text = ""
        if file.name.lower().endswith('.pdf') and fitz:
            doc = fitz.open(stream=file.read(), filetype="pdf")
            text = "\n".join([page.get_text() for page in doc])
        elif Image and pytesseract:
            text = pytesseract.image_to_string(Image.open(file))
        
        if not text: return "", {}

        results = {}
        fields = {
            "Purchase Price / Cost Basis": ["purchase price", "cost basis", "acquisition price", "contract price"],
            "Appraised Value": ["appraised value", "market value", "as-is value"],
            "Gross Income": ["gross potential", "total income", "effective gross", "revenue", "egi"],
            "Vacancy / Credit Loss": ["vacancy", "credit loss", "vacancy loss"],
            "Operating Expenses": ["operating expenses", "total expenses", "opex"],
            "Stabilized NOI": ["net operating income", "noi", "net income"],
            "Debt Service": ["debt service", "annual debt service", "mortgage payment"],
            "CapEx / Reserves": ["capex", "capital expenditures", "replacement reserves"]
        }
        
        for field, keywords in fields.items():
            for line in text.splitlines():
                line_lower = line.lower()
                for k in keywords:
                    if k in line_lower:
                        matches = re.findall(r'[\$\(]?\s*-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{2})?\)?', line)
                        if matches:
                            match_str = matches[-1]
                            clean_val = float(match_str.replace('$', '').replace(',', '').replace('(', '-').replace(')', ''))
                            conf = OCREngine.calculate_confidence(match_str, line, k)
                            results[field] = {"value": clean_val, "confidence": conf, "source_line": line.strip()}
                            break
                if field in results:
                    break
        return text, results

class ExportEngine:
    @staticmethod
    def generate_excel(state: dict, loan_amt: float, gate: str, amort_df: pd.DataFrame, score: int, tier: str) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            wb = writer.book
            fmt_header = wb.add_format({'bold': True, 'bg_color': '#CFB87C', 'color': '#000000', 'border': 1})
            
            exec_data = {
                "Metric": ["Sponsor", "Property", "Appraisal", "NOI", "Supportable Proceeds", "Constraint", "Deal Score", "Tier"],
                "Value": [state.get('sponsor'), state.get('property_address'), state.get('appraisal'), 
                          state.get('noi'), loan_amt, gate, score, tier]
            }
            df_exec = pd.DataFrame(exec_data)
            df_exec.to_excel(writer, sheet_name="Executive Summary", index=False)
            ws_exec = writer.sheets["Executive Summary"]
            ws_exec.set_column('A:A', 25)
            ws_exec.set_column('B:B', 35)
            for col_num, value in enumerate(df_exec.columns.values):
                ws_exec.write(0, col_num, value, fmt_header)
            
            if 'rent_roll_dict' in state:
                df_rr = pd.DataFrame(state['rent_roll_dict'])
                df_rr.to_excel(writer, sheet_name="Rent Roll", index=False)
                ws_rr = writer.sheets["Rent Roll"]
                ws_rr.set_column('A:Z', 18)
                for col_num, value in enumerate(df_rr.columns.values):
                    ws_rr.write(0, col_num, value, fmt_header)
            
            if amort_df is not None and not amort_df.empty:
                amort_df.to_excel(writer, sheet_name="Amortization", index=False)
                ws_amort = writer.sheets["Amortization"]
                ws_amort.set_column('A:Z', 15)
                for col_num, value in enumerate(amort_df.columns.values):
                    ws_amort.write(0, col_num, value, fmt_header)
                    
            with sqlite3.connect(DB_PATH) as conn:
                audit_df = pd.read_sql_query("SELECT user, action, details, timestamp FROM audit_log ORDER BY timestamp DESC LIMIT 100", conn)
                audit_df.to_excel(writer, sheet_name="Audit Log", index=False)
                
        return output.getvalue()
        
    @staticmethod
    def generate_pdf(state: dict, loan_amt: float, gate: str, score: int, tier: str, risk_flags: List[str]) -> bytes:
        if not PDF_AVAILABLE: return b""
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        story = []
        
        story.append(Paragraph("ALENZA CAPITAL - UNDERWRITING TEAR SHEET", styles["Title"]))
        story.append(Spacer(1, 12))
        
        meta = f"""
        <b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}<br/>
        <b>Sponsor:</b> {state.get('sponsor')}<br/>
        <b>Property:</b> {state.get('property_address')} ({state.get('property_type')})<br/>
        """
        story.append(Paragraph(meta, styles["Normal"]))
        story.append(Spacer(1, 16))
        
        exec_summary = f"""
        <b>Supportable Proceeds:</b> ${loan_amt:,.0f}<br/>
        <b>Binding Constraint:</b> {gate}<br/>
        <b>Appraised Value:</b> ${state.get('appraisal', 0):,.0f}<br/>
        <b>Stabilized NOI:</b> ${state.get('noi', 0):,.0f}<br/>
        <b>Deal Score:</b> {score}/1000<br/>
        <b>Classification:</b> {tier}<br/>
        """
        story.append(Paragraph("Executive Summary", styles["Heading2"]))
        story.append(Paragraph(exec_summary, styles["Normal"]))
        story.append(Spacer(1, 16))
        
        story.append(Paragraph("Risk & Structural Analysis", styles["Heading2"]))
        for flag in risk_flags:
            clean_flag = flag.replace("⚠️", "").replace("🚨", "").replace("✅", "").strip()
            story.append(Paragraph(f"• {clean_flag}", styles["Normal"]))
        story.append(Spacer(1, 16))
        
        disclaimer = "This document is generated by Alenza OS. It is indicative only and does not constitute a loan commitment, credit approval, or legal advice. Subject to final lender diligence."
        story.append(Paragraph("Disclaimer", styles["Heading2"]))
        story.append(Paragraph(disclaimer, styles["Normal"]))
        
        doc.build(story)
        return buffer.getvalue()

# ==========================================
# 5. ADVANCED UNDERWRITING ENGINE
# ==========================================
class UnderwritingEngine:
    LENDER_PROFILES = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05}
    }

    @staticmethod
    def size_loan(noi, appraisal, pp, closing, reserves, fees_pct, rate, amort, term, is_io, t_ltv, t_ltc, t_dscr, t_dy):
        base_cost = pp + closing + reserves
        loan = appraisal * t_ltv
        for _ in range(5):
            total_uses = base_cost + (loan * fees_pct)
            ltv_loan = appraisal * t_ltv
            ltc_loan = total_uses * t_ltc
            dy_loan = noi / t_dy if t_dy > 0 else 0
            m_rate = rate / 12
            if is_io: dscr_loan = ((noi / t_dscr) / 12) / m_rate if m_rate > 0 else 0
            else: dscr_loan = ((noi / t_dscr) / 12) * ((1 - (1 + m_rate)**-(amort*12)) / m_rate) if m_rate > 0 else 0
            gates = {"LTV": ltv_loan, "LTC": ltc_loan, "DSCR": dscr_loan, "Debt Yield": dy_loan}
            loan = min(gates.values())
            
        gate = min(gates, key=gates.get)
        req_equity = total_uses - loan
        return loan, gate, gates, total_uses, req_equity

    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io):
        m_rate = rate / 12
        pmts = amort_yrs * 12
        term_months = int(term_yrs * 12)
        pmt = loan_amt * m_rate if is_io else (loan_amt * m_rate) / (1 - (1 + m_rate)**-pmts)
        
        sched, bal = [], loan_amt
        for i in range(1, term_months + 1):
            int_pmt = bal * m_rate
            prin_pmt = 0 if is_io else pmt - int_pmt
            bal -= prin_pmt
            sched.append({"Period": i, "Payment": pmt, "Principal": prin_pmt, "Interest": int_pmt, "Balance": max(0, bal)})
            if bal <= 0: break
            
        df = pd.DataFrame(sched)
        balloon = bal if bal > 0 else 0
        return df, pmt, balloon

    @staticmethod
    def rent_roll_metrics(df):
        df['SF'] = pd.to_numeric(df['SF'], errors='coerce').fillna(0)
        df['Monthly Rent'] = pd.to_numeric(df['Monthly Rent'], errors='coerce').fillna(0)
        df['Remaining Term'] = pd.to_numeric(df['Remaining Term'], errors='coerce').fillna(0)
        
        total_sf = df['SF'].sum()
        occ_df = df[~df['Tenant'].str.lower().isin(['vacant', 'empty', 'available'])]
        occ_sf = occ_df['SF'].sum()
        
        ann_rent = occ_df['Monthly Rent'].sum() * 12
        walt = (occ_df['Remaining Term'] * occ_df['SF']).sum() / occ_sf if occ_sf > 0 else 0
        exp_1yr = occ_df[occ_df['Remaining Term'] <= 1.0]['SF'].sum() / occ_sf if occ_sf > 0 else 0
        
        return total_sf, occ_sf/total_sf if total_sf > 0 else 0, ann_rent, ann_rent/occ_sf if occ_sf > 0 else 0, walt, exp_1yr

    @staticmethod
    def score_deal(ltv, ltc, dscr, dy, profile_name) -> Tuple[int, str]:
        limits = UnderwritingEngine.LENDER_PROFILES.get(profile_name, UnderwritingEngine.LENDER_PROFILES["Bank / Credit Union"])
        ltv_score = max(0, 300 * (1 - (ltv / limits['max_ltv'])))
        dscr_score = max(0, 300 * ((dscr - 1.0) / (limits['min_dscr'] - 1.0))) if dscr > 1 else 0
        dy_score = max(0, 200 * (dy / limits['min_dy']))
        ltc_score = max(0, 200 * (1 - ltc))
        score = min(1000, int(ltv_score + dscr_score + dy_score + ltc_score))
        
        if score >= 850: tier = "Tier 1 | Institutional Core"
        elif score >= 700: tier = "Tier 2 | Conventional Bankable"
        elif score >= 550: tier = "Tier 3 | Alternative / Debt Fund"
        else: tier = "Tier 4 | Private / Restructure"
        return score, tier

class RiskAnalysisEngine:
    @staticmethod
    def generate_narrative(actual_ltv: float, actual_dscr: float, walt: float, exp_1yr: float, is_io: bool, req_equity: float) -> List[str]:
        flags = []
        if actual_ltv > 0.75: flags.append(f"⚠️ **High Leverage:** Transaction requires {actual_ltv*100:.1f}% LTV, pushing beyond conventional parameters.")
        elif actual_ltv < 0.60: flags.append("✅ **Conservative Capitalization:** Low LTV indicates strong sponsor equity commitment.")
        
        if actual_dscr < 1.20: flags.append(f"⚠️ **Tight Cash Flow:** DSCR is exceptionally thin at {actual_dscr:.2f}x.")
        
        if is_io and actual_dscr < 1.25: flags.append("⚠️ **Structural Masking:** Interest-Only structure may be masking amortizing weakness.")
            
        if walt > 0 and walt < 2.5: flags.append(f"⚠️ **Short WALT:** WALT is {walt:.1f} years. Lenders will require significant leasing reserves.")
        if exp_1yr > 0.30: flags.append(f"🚨 **Rollover Exposure:** {exp_1yr*100:.1f}% of occupied SF is expiring within 12 months.")
            
        if req_equity < 0: flags.append("🚨 **Capital Stack Inversion:** Required equity is negative (cash-out scenario).")
            
        if not flags: flags.append("✅ **Clean Profile:** No major automated structural or cash-flow risk flags detected.")
        return flags

class MarketCompsEngine:
    @staticmethod
    def generate_comps(property_type: str, noi: float) -> pd.DataFrame:
        np.random.seed(int(noi))
        base_cap = {"Multifamily": 0.045, "Industrial": 0.055, "Retail": 0.065, "Office": 0.075}.get(property_type, 0.06)
        comps = []
        for i in range(1, 6):
            cap_variance = np.random.uniform(-0.0075, 0.0075)
            comp_cap = base_cap + cap_variance
            comp_value = (noi * np.random.uniform(0.8, 1.2)) / comp_cap
            comps.append({
                "Comparable": f"{property_type} Asset {chr(64+i)}",
                "Distance (km)": round(np.random.uniform(0.5, 8.0), 1),
                "Sale Date": f"202{np.random.randint(4, 6)}-0{np.random.randint(1, 9)}",
                "Cap Rate": f"{comp_cap*100:.2f}%",
                "Est. Value": f"${comp_value:,.0f}"
            })
        return pd.DataFrame(comps)

# ==========================================
# 6. INITIALIZE SESSION STATE
# ==========================================
DEFAULT_STATE = {
    "deal_id": f"deal_{int(datetime.now().timestamp())}",
    "sponsor": "Alenza Client",
    "property_address": "100 King St W, Toronto, ON",
    "property_type": "Multifamily",
    "transaction_type": "Acquisition",
    "lender_profile": "Bank / Credit Union",
    "purchase_price": 12500000.0,
    "appraisal": 13750000.0,
    "noi": 1060322.0,
    "target_ltv": 0.75,
    "target_ltc": 0.80,
    "target_dscr": 1.25,
    "target_dy": 0.085,
    "rate": 0.0525,
    "amort": 25,
    "term": 5,
    "is_io": False,
    "fees": 0.02,
    "closing_costs": 50000.0,
    "reserves": 0.0,
    "rent_roll_dict": [
        {"Tenant": "Main Anchor", "SF": 25000, "Remaining Term": 5.5, "Monthly Rent": 45000},
        {"Tenant": "In-Line A", "SF": 3500, "Remaining Term": 1.2, "Monthly Rent": 8000},
        {"Tenant": "Vacant", "SF": 5000, "Remaining Term": 0, "Monthly Rent": 0}
    ]
}

for k, v in DEFAULT_STATE.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# 7. UI/UX: FULL DARK MODE CSS
# ==========================================
st.markdown("""
    <style>
    /* Base Backgrounds */
    .stApp { background-color: #0B0F19 !important; font-family: 'Inter', 'Helvetica Neue', sans-serif; }
    .main { background-color: #0B0F19 !important; color: #F3F4F6 !important; }
    
    /* Top Header Strip */
    header[data-testid="stHeader"] { background-color: #0B0F19 !important; border-bottom: 2px solid #CFB87C !important; }
    
    /* Sidebar */
    section[data-testid="stSidebar"] { background-color: #0F172A !important; border-right: 1px solid #1E293B !important; }
    section[data-testid="stSidebar"] * { color: #F3F4F6 !important; }
    section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { color: #CFB87C !important; }
    
    /* Fix for Sidebar Captions and Disclaimers */
    section[data-testid="stSidebar"] .stCaption, 
    section[data-testid="stSidebar"] small, 
    section[data-testid="stSidebar"] p[data-testid="stMarkdownContainer"] > em { 
        color: #9CA3AF !important; 
    }
    
    /* Form Inputs (Sidebar & Main) */
    .stTextInput input, .stNumberInput input, .stSelectbox select { 
        background-color: #111827 !important; 
        border: 1px solid #1E293B !important; 
        color: #F3F4F6 !important; 
    }
    
    /* Typography Overrides */
    .main * { color: #F3F4F6 !important; }
    h1 { border-bottom: 3px solid #CFB87C !important; padding-bottom: 10px !important; font-weight: 800 !important; color: #F3F4F6 !important; }
    h2, h3, h4 { font-weight: 700 !important; color: #E5E7EB !important; }
    
    /* HUD Metric Cards */
    div[data-testid="stMetric"] { 
        background-color: #111827 !important; 
        border: 1px solid #1E293B !important; 
        border-top: 4px solid #CFB87C !important; 
        border-radius: 6px !important; 
        padding: 15px !important; 
        box-shadow: 0 4px 6px rgba(0,0,0,0.3) !important; 
    }
    [data-testid="stMetricValue"] { font-size: 28px !important; font-weight: 800 !important; color: #F3F4F6 !important; }
    [data-testid="stMetricLabel"] { font-size: 11px !important; text-transform: uppercase !important; letter-spacing: 1.2px !important; color: #9CA3AF !important; font-weight: 700 !important;}
    
    /* Tab Navigation */
    .stTabs [data-baseweb="tab-list"] { border-bottom: 2px solid #1E293B !important; gap: 8px !important; }
    .stTabs [data-baseweb="tab"] { background-color: transparent !important; color: #9CA3AF !important; font-weight: 700 !important; text-transform: uppercase; font-size: 12px; }
    .stTabs [aria-selected="true"] { color: #F3F4F6 !important; border-bottom: 4px solid #CFB87C !important; }
    
    /* Interactive Tables (Dark Zebra Striping) */
    .stDataFrame { border: 1px solid #1E293B !important; border-radius: 6px !important; }
    th { background-color: #0F172A !important; color: #CFB87C !important; font-weight: 700 !important; padding: 12px !important; border-bottom: 1px solid #1E293B !important;}
    td { background-color: #111827 !important; padding: 10px !important; border-bottom: 1px solid #1E293B !important; }
    tbody tr:nth-child(even) td { background-color: #1F2937 !important; }
    
    /* Buttons */
    .stButton>button, .stDownloadButton>button { 
        background-color: #CFB87C !important; 
        color: #000000 !important; 
        font-weight: 800 !important; 
        border-radius: 4px !important; 
        text-transform: uppercase; 
        transition: 0.2s; 
        width: 100%; 
        border: none !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover { 
        background-color: #B09B65 !important; 
        transform: translateY(-1px);
    }
    
    /* Alerts / Notifications */
    .stAlert { border-left: 5px solid #CFB87C !important; background-color: #111827 !important; color: #F3F4F6 !important; box-shadow: 0 2px 5px rgba(0,0,0,0.3); }
    .stAlert * { color: #F3F4F6 !important; }
    
    /* Expanders */
    div[data-testid="stExpander"] { border: 1px solid #1E293B !important; background-color: #0F172A !important; border-radius: 6px !important;}
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 8. SIDEBAR COMMAND CENTER & STATE
# ==========================================
with st.sidebar:
    st.title("🏛️ ALENZA OS")
    
    with st.expander("📁 PIPELINE MANAGER", expanded=True):
        all_deals = DatabaseManager.get_all_deals()
        deal_options = ["-- Start New Deal --"] + all_deals['name'].tolist()
        selected = st.selectbox("Select Deal", deal_options)
        
        c_load, c_del, c_dup = st.columns(3)
        with c_load:
            if st.button("📂 Load") and selected != "-- Start New Deal --":
                deal_id_to_load = all_deals.loc[all_deals['name'] == selected, 'id'].values[0]
                loaded_state = DatabaseManager.load_deal(deal_id_to_load)
                if loaded_state:
                    for k, v in loaded_state.items(): st.session_state[k] = v
                    st.session_state.deal_id = loaded_state.get('deal_id', deal_id_to_load)
                    st.rerun()
                    
        with c_del:
            if st.button("🗑️ Delete") and selected != "-- Start New Deal --":
                deal_id_to_del = all_deals.loc[all_deals['name'] == selected, 'id'].values[0]
                DatabaseManager.delete_deal(deal_id_to_del)
                st.success("Wiped.")
                st.rerun()
                
        with c_dup:
            if st.button("📑 Dup."):
                st.session_state.deal_id = f"deal_{int(datetime.now().timestamp())}"
                st.warning("Duplicated in memory. Click 'Save Deal Record' to persist.")

    st.markdown("---")
    
    with st.expander("🏢 ASSET PROFILE", expanded=True):
        st.session_state.sponsor = st.text_input("Sponsor", st.session_state.sponsor)
        st.session_state.property_address = st.text_input("Address", st.session_state.property_address)
        st.session_state.property_type = st.selectbox("Type", ["Multifamily", "Industrial", "Retail", "Office"], index=["Multifamily", "Industrial", "Retail", "Office"].index(st.session_state.property_type))
        st.session_state.appraisal = st.number_input("Appraisal ($)", value=st.session_state.appraisal, step=100000.0)
        st.session_state.purchase_price = st.number_input("Cost Basis ($)", value=st.session_state.purchase_price, step=100000.0)
        st.session_state.noi = st.number_input("Stabilized NOI ($)", value=st.session_state.noi, step=10000.0)

    with st.expander("📊 CREDIT POLICY", expanded=True):
        st.session_state.lender_profile = st.selectbox("Policy Preset", list(UnderwritingEngine.LENDER_PROFILES.keys()), index=list(UnderwritingEngine.LENDER_PROFILES.keys()).index(st.session_state.lender_profile))
        preset = UnderwritingEngine.LENDER_PROFILES[st.session_state.lender_profile]
        st.session_state.target_ltv = st.slider("Max LTV %", 50.0, 95.0, preset['max_ltv']*100) / 100
        st.session_state.target_dscr = st.slider("Min DSCR x", 1.0, 1.75, preset['min_dscr'])
        st.session_state.target_dy = st.slider("Min DY %", 5.0, 15.0, preset['min_dy']*100) / 100
        st.session_state.target_ltc = st.slider("Max LTC %", 50.0, 100.0, 80.0) / 100

    with st.expander("💰 DEBT STRUCTURE", expanded=True):
        st.session_state.is_io = st.checkbox("Interest-Only", value=st.session_state.is_io)
        st.session_state.rate = st.slider("Rate %", 3.0, 12.0, st.session_state.rate*100) / 100
        st.session_state.amort = st.number_input("Amort (Yrs)", value=st.session_state.amort)
        st.session_state.term = st.number_input("Term (Yrs)", value=st.session_state.term)
        st.session_state.fees = st.slider("Fees %", 0.0, 5.0, st.session_state.fees*100) / 100
        st.session_state.closing_costs = st.number_input("Closing Costs", value=st.session_state.closing_costs)
        st.session_state.reserves = st.number_input("Reserves", value=st.session_state.reserves)

# ==========================================
# 9. EXECUTE MATH & ENGINE VALIDATION
# ==========================================
s = st.session_state

# Sizing Engine
loan_amt, gate, gates, total_uses, req_equity = UnderwritingEngine.size_loan(
    s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, 
    s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy)

# Amortization Engine
amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(loan_amt, s.rate, s.amort, s.term, s.is_io)
annual_ds = monthly_pmt * 12

# Actual Metrics
actual_ltv = loan_amt / s.appraisal if s.appraisal else 0
actual_ltc = loan_amt / total_uses if total_uses else 0
actual_dscr = s.noi / annual_ds if annual_ds else 0
actual_dy = s.noi / loan_amt if loan_amt else 0

# Rent Roll Engine
tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(pd.DataFrame(s.rent_roll_dict))

# Scoring Engine
score, classification = UnderwritingEngine.score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, s.lender_profile)

# Risk Engine
risk_flags = RiskAnalysisEngine.generate_narrative(actual_ltv, actual_dscr, walt, exp1, s.is_io, req_equity)


# ==========================================
# 10. MAIN DASHBOARD HUD
# ==========================================
st.title(f"{s.sponsor} | {s.property_address}")
st.caption(f"INSTITUTIONAL WORKSTATION | ACTIVE CONSTRAINT: {gate} | TIER: {classification.upper()}")

# HUD Metrics
m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
m2.metric("ACTUAL LTV", f"{actual_ltv*100:.1f}%")
m3.metric("ACTUAL LTC", f"{actual_ltc*100:.1f}%")
m4.metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")
m5.metric("BALLOON", f"${balloon:,.0f}")
m6.metric("DEAL SCORE", f"{score}/1000")

st.markdown("---")

# ==========================================
# 11. WORKFLOW TABS
# ==========================================
tabs = st.tabs([
    "📊 Sizing & Risk", "📝 Rent Roll", "📅 Amortization", "📈 Market Comps", 
    "📎 Diligence Room", "🤖 OCR Extract", "🇨🇦 Canada Intel", "💾 Save & Export"
])

# TAB 1: SIZING & RISK
with tabs[0]:
    c1, c2 = st.columns([1.5, 1], gap="large")
    with c1:
        st.subheader("Constraint Analysis")
        df_gates = pd.DataFrame({"Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
                                 "Threshold": [f"{s.target_ltv*100}%", f"{s.target_ltc*100}%", f"{s.target_dscr}x", f"{s.target_dy*100}%"],
                                 "Proceeds Limit": [f"${gates['LTV']:,.0f}", f"${gates['LTC']:,.0f}", f"${gates['DSCR']:,.0f}", f"${gates['Debt Yield']:,.0f}"],
                                 "Binding": ["✅ YES" if gate == k else "" for k in ["LTV", "LTC", "DSCR", "Debt Yield"]]})
        st.table(df_gates)
        
        st.subheader("Sources & Uses")
        df_su = pd.DataFrame({"Uses": ["Cost Basis", "Closing Costs", "Reserves", "Financing Fees", "Total"],
                              "U Amount": [s.purchase_price, s.closing_costs, s.reserves, loan_amt*s.fees, total_uses],
                              "Sources": ["Senior Debt", "Sponsor Equity", "", "", "Total"],
                              "S Amount": [loan_amt, req_equity, 0, 0, total_uses]})
        st.dataframe(df_su.style.format({"U Amount": "${:,.0f}", "S Amount": "${:,.0f}"}), hide_index=True, use_container_width=True)
        
    with c2:
        st.subheader("Executive Risk Narrative")
        for flag in risk_flags:
            if "⚠️" in flag or "🚨" in flag:
                st.warning(flag)
            else:
                st.info(flag)

# TAB 2: RENT ROLL
with tabs[1]:
    st.subheader("Interactive Rent Roll")
    edited_rr = st.data_editor(pd.DataFrame(s.rent_roll_dict), num_rows="dynamic", use_container_width=True)
    s.rent_roll_dict = edited_rr.to_dict('records')
    
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("WALT", f"{walt:.2f} Yrs")
    r2.metric("Occupancy", f"{occ*100:.1f}%")
    r3.metric("Annual Rent", f"${ann_rent:,.0f}")
    r4.metric("12-Mo Rollover Risk", f"{exp1*100:.1f}%")

# TAB 3: AMORTIZATION
with tabs[2]:
    st.subheader(f"Schedule: {s.term} Year Term")
    st.area_chart(amort_df[["Principal", "Interest"]].head(60))
    st.dataframe(amort_df.style.format("${:,.2f}"), use_container_width=True, height=300)

# TAB 4: MARKET COMPS
with tabs[3]:
    st.subheader("Simulated Market Comparables")
    st.caption("Auto-generated based on Property Type and NOI.")
    comps_df = MarketCompsEngine.generate_comps(s.property_type, s.noi)
    st.dataframe(comps_df, hide_index=True, use_container_width=True)

# TAB 5: DILIGENCE ROOM & GAP ANALYSIS
with tabs[4]:
    st.subheader("Diligence Vault & Gap Analysis")
    REQUIRED_DOCS = ["Appraisal", "Phase I ESA", "T12 Financials", "Rent Roll", "Sponsor Bio", "Purchase Agreement"]
    
    with sqlite3.connect(DB_PATH) as conn:
        docs = pd.read_sql_query('SELECT id, filename, category, uploaded_at FROM documents WHERE deal_id = ?', conn, params=(s.deal_id,))
    
    d1, d2 = st.columns(2)
    with d1:
        st.write("### Upload Document")
        cat = st.selectbox("Category", REQUIRED_DOCS + ["Other"])
        doc_file = st.file_uploader("Drop File Here")
        if st.button("Save to Vault") and doc_file:
            DatabaseManager.save_document(s.deal_id, doc_file, cat)
            st.success(f"Saved {doc_file.name}")
            st.rerun()
            
    with d2:
        st.write("### Package Gap Analysis")
        uploaded_cats = docs['category'].tolist() if not docs.empty else []
        gap_df = pd.DataFrame({"Requirement": REQUIRED_DOCS, 
                               "Status": ["✅ Uploaded" if c in uploaded_cats else "❌ Missing" for c in REQUIRED_DOCS]})
        st.dataframe(gap_df, hide_index=True, use_container_width=True)
    
    st.write("### Vault Inventory")
    if not docs.empty:
        st.dataframe(docs[['filename', 'category', 'uploaded_at']], use_container_width=True)
        doc_to_delete = st.selectbox("Select Document to Delete", ["-- None --"] + docs['filename'].tolist())
        if st.button("🗑️ Delete Selected Document") and doc_to_delete != "-- None --":
            doc_id_to_del = docs.loc[docs['filename'] == doc_to_delete, 'id'].values[0]
            DatabaseManager.delete_document(doc_id_to_del)
            st.success("Document deleted.")
            st.rerun()
    else:
        st.info("No documents uploaded yet.")

# TAB 6: OCR EXTRACT
with tabs[5]:
    st.subheader("AI Financial Extraction")
    if not OCR_AVAILABLE:
        st.warning("OCR dependencies (pytesseract, PyMuPDF) not found. Extraction disabled.")
    
    uploaded_fin = st.file_uploader("Upload Appraisal / T12 / Operating Statement (PDF/Image)", type=["pdf", "png", "jpg"])
    
    if uploaded_fin and OCR_AVAILABLE:
        with st.spinner("Extracting parameters with confidence scoring..."):
            text, extracted = OCREngine.extract_and_parse(uploaded_fin)
            if extracted:
                st.success("Extraction Complete.")
                st.dataframe(pd.DataFrame(extracted).T, use_container_width=True)
                if st.button("Apply Parameters to Underwriting Model"):
                    if "Stabilized NOI" in extracted: s.noi = extracted["Stabilized NOI"]["value"]
                    if "Purchase Price / Cost Basis" in extracted: s.purchase_price = extracted["Purchase Price / Cost Basis"]["value"]
                    if "Appraised Value" in extracted: s.appraisal = extracted["Appraised Value"]["value"]
                    DatabaseManager.log_audit("OCR_APPLY", "Applied Extracted Parameters")
                    st.success("Model updated. View Sidebar for changes.")
            else:
                st.warning("Could not identify high-confidence parameters.")

# TAB 7: CANADA INTEL
with tabs[6]:
    st.subheader("🇨🇦 Sovereign Intelligence")
    ca1, ca2 = st.columns(2)
    with ca1:
        st.write("### Live Bank of Canada Rates")
        boc = CanadianIntel.get_boc_rates()
        if boc:
            st.metric("5-Year Bond Yield", f"{boc['5yr_bond']:.2f}%")
            st.metric("Overnight Rate", f"{boc['overnight']:.2f}%")
            st.caption(f"Last Updated: {boc['date']}")
        else:
            st.info("BoC API currently unavailable.")
            
    with ca2:
        st.write("### Federal Corporation Registry")
        corp = st.text_input("Verify Federal Corporation Name")
        if corp:
            result = CanadianIntel.verify_corporation(corp)
            if result["status"] == "found":
                st.success(f"Verified: {result['data'].get('name')}")
                st.json(result['data'])
            elif result["status"] == "not_found":
                st.warning(result["message"])
            else:
                st.error(result["message"])
        
        st.write("### NRCan Address Validation")
        nrcan_query = st.text_input("Verify Address Coordinates")
        if nrcan_query:
            verified = CanadianIntel.geocode_nrcan(nrcan_query)
            if verified:
                st.success(f"📍 Standardized: {verified}")
            else:
                st.warning("Could not verify via Natural Resources Canada.")

# TAB 8: EXPORT & SAVE
with tabs[7]:
    st.subheader("Save & Package Export")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("💾 Save Deal Record to Database", use_container_width=True):
            clean_state = extract_clean_state()
            DatabaseManager.save_deal(s.deal_id, f"{s.sponsor} - {s.property_type}", clean_state)
            st.success("Deal permanently saved to SQLite.")
            
    with c2:
        excel_bytes = ExportEngine.generate_excel(extract_clean_state(), loan_amt, gate, amort_df, score, classification)
        st.download_button("📊 Download Excel Model", data=excel_bytes, file_name=f"{clean_filename(s.sponsor)}_Model.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
    with c3:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w') as z:
            # Inject Excel
            z.writestr("Underwriting_Model.xlsx", excel_bytes)
            
            # Inject PDF Summary
            if PDF_AVAILABLE:
                pdf_bytes = ExportEngine.generate_pdf(extract_clean_state(), loan_amt, gate, score, classification, risk_flags)
                z.writestr("Executive_Summary.pdf", pdf_bytes)
                
            # Inject SQLite DB
            if DB_PATH.exists(): z.write(DB_PATH, "Database_Backup.db")
            
            # Inject Physical Diligence Documents
            with sqlite3.connect(DB_PATH) as conn:
                vault_docs = pd.read_sql_query('SELECT filename, path FROM documents WHERE deal_id = ?', conn, params=(s.deal_id,))
                for _, row in vault_docs.iterrows():
                    p = Path(row['path'])
                    if p.exists():
                        z.write(p, f"Diligence_Vault/{row['filename']}")

        st.download_button("📦 Download Full Deal Package", buf.getvalue(), file_name=f"{clean_filename(s.sponsor)}_Package.zip", mime="application/zip", use_container_width=True)

# Footer Disclaimers
st.markdown("---")
st.caption("⚠️ **DISCLAIMER:** ALENZA CAPITAL OS is an indicative modeling tool. Outputs do not constitute a loan commitment, appraisal, or legal advice. Final terms are subject to formal credit committee approval and third-party diligence verification.")
