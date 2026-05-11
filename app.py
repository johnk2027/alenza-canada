# Alenza Capital OS
# Enterprise underwriting workspace
# Midnight Slate and CU Gold theme

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


st.set_page_config(
    page_title="Alenza Capital OS",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)


try:
    from PIL import Image
    import pytesseract
    import fitz  # PyMuPDF
    OCR_AVAILABLE = True
except Exception:
    Image = pytesseract = fitz = None
    OCR_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    PDF_AVAILABLE = True
except Exception:
    SimpleDocTemplate = Paragraph = Spacer = getSampleStyleSheet = None
    PDF_AVAILABLE = False

try:
    import xlsxwriter  # noqa: F401
    EXCEL_AVAILABLE = True
except Exception:
    EXCEL_AVAILABLE = False

try:
    import openpyxl  # noqa: F401
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

try:
    import xlrd  # noqa: F401
    XLRD_AVAILABLE = True
except Exception:
    XLRD_AVAILABLE = False


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "alenza_data"
DB_PATH = DATA_DIR / "alenza_platform.db"
DOC_DIR = DATA_DIR / "documents"

for d in [DATA_DIR, DOC_DIR]:
    d.mkdir(parents=True, exist_ok=True)


def clean_filename(filename: str) -> str:
    """Sanitize filenames to prevent OS path issues."""
    base = str(filename or "file").strip()
    base = re.sub(r"[^a-zA-Z0-9_\-\.]", "_", base)
    return base[:150] or "file"


def current_audit_user() -> str:
    try:
        if hasattr(st, "user") and getattr(st.user, "email", None):
            return str(st.user.email)
    except Exception:
        pass
    try:
        return str(st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User")))
    except Exception:
        return os.environ.get("APP_USER", "Local User")


def safe_float(value: Any, default: float = 0.0) -> float:
    """Convert a value to float."""
    try:
        if value is None or value == "":
            return default
        if isinstance(value, str):
            value = value.replace("$", "").replace(",", "").replace("%", "").strip()
        return float(value)
    except Exception:
        return default


def safe_int(value: Any, default: int = 0) -> int:
    """Convert a value to int."""
    try:
        return int(float(value))
    except Exception:
        return default


def normalize_percent(value: Any, default: float) -> float:
    """Accept 0.75 or 75 and return decimal 0.75."""
    x = safe_float(value, default)
    if x > 1.5:
        return x / 100
    return x


DEFAULT_RENT_ROLL = [
    {"Tenant": "Main Anchor", "SF": 25000, "Remaining Term": 5.5, "Monthly Rent": 45000},
    {"Tenant": "In-Line A", "SF": 3500, "Remaining Term": 1.2, "Monthly Rent": 8000},
    {"Tenant": "Vacant", "SF": 5000, "Remaining Term": 0, "Monthly Rent": 0},
]

DEFAULT_STATE = {
    "deal_id": f"deal_{int(datetime.now().timestamp())}",
    "deal_name": "Untitled Deal",
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
    "rent_roll_dict": DEFAULT_RENT_ROLL.copy(),
}


def normalize_rent_roll_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize any rent-roll-like file/dataframe into required schema."""
    required = ["Tenant", "SF", "Remaining Term", "Monthly Rent"]

    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame(columns=required)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    for col in df.columns:
        c = str(col).strip().lower()
        c_clean = re.sub(r"[^a-z0-9]+", " ", c).strip()

        if c_clean in ["tenant", "tenant name", "lessee", "occupant", "customer", "company", "unit tenant"]:
            col_map[col] = "Tenant"
        elif c_clean in ["sf", "sq ft", "sqft", "square feet", "area", "gla", "unit sf", "leased sf", "suite sf"]:
            col_map[col] = "SF"
        elif c_clean in [
            "remaining term", "term remaining", "lease term remaining", "years remaining",
            "remaining years", "walt", "lease term", "term yrs", "term years"
        ]:
            col_map[col] = "Remaining Term"
        elif c_clean in [
            "monthly rent", "rent month", "monthly base rent", "base rent monthly",
            "month rent", "rent per month", "monthly revenue"
        ]:
            col_map[col] = "Monthly Rent"
        elif c_clean in ["annual rent", "annual base rent", "yearly rent", "annual revenue", "base rent annual"]:
            col_map[col] = "Annual Rent"

    df = df.rename(columns=col_map)

    # If duplicate mapped columns exist, keep the first non-null value row-wise.
    for target in ["Tenant", "SF", "Remaining Term", "Monthly Rent", "Annual Rent"]:
        matching = [c for c in df.columns if c == target]
        if len(matching) > 1:
            temp = df.loc[:, matching]
            df = df.drop(columns=matching)
            df[target] = temp.bfill(axis=1).iloc[:, 0]

    if "Annual Rent" in df.columns and "Monthly Rent" not in df.columns:
        df["Monthly Rent"] = pd.to_numeric(df["Annual Rent"], errors="coerce").fillna(0) / 12

    for col in required:
        if col not in df.columns:
            df[col] = "" if col == "Tenant" else 0

    out = df[required].copy()
    out["Tenant"] = out["Tenant"].fillna("").astype(str).str.strip()
    out["SF"] = pd.to_numeric(out["SF"], errors="coerce").fillna(0)
    out["Remaining Term"] = pd.to_numeric(out["Remaining Term"], errors="coerce").fillna(0)
    out["Monthly Rent"] = pd.to_numeric(out["Monthly Rent"], errors="coerce").fillna(0)

    # Drop rows that are completely empty.
    out = out[~((out["Tenant"] == "") & (out["SF"] == 0) & (out["Monthly Rent"] == 0))]
    return out.reset_index(drop=True)


def normalize_loaded_state(state: dict) -> dict:
    """Make old saved deals compatible with current app schema."""
    normalized = DEFAULT_STATE.copy()
    if isinstance(state, dict):
        normalized.update(state)

    normalized["purchase_price"] = safe_float(normalized.get("purchase_price"), 0.0)
    normalized["appraisal"] = safe_float(normalized.get("appraisal"), 0.0)
    normalized["noi"] = safe_float(normalized.get("noi"), 0.0)
    normalized["target_ltv"] = normalize_percent(normalized.get("target_ltv"), 0.75)
    normalized["target_ltc"] = normalize_percent(normalized.get("target_ltc"), 0.80)
    normalized["target_dscr"] = safe_float(normalized.get("target_dscr"), 1.25)
    normalized["target_dy"] = normalize_percent(normalized.get("target_dy"), 0.085)
    normalized["rate"] = normalize_percent(normalized.get("rate"), 0.0525)
    normalized["amort"] = max(1, safe_int(normalized.get("amort"), 25))
    normalized["term"] = max(1, safe_int(normalized.get("term"), 5))
    normalized["fees"] = normalize_percent(normalized.get("fees"), 0.02)
    normalized["closing_costs"] = safe_float(normalized.get("closing_costs"), 0.0)
    normalized["reserves"] = safe_float(normalized.get("reserves"), 0.0)
    normalized["is_io"] = bool(normalized.get("is_io", False))

    prop_types = ["Multifamily", "Industrial", "Retail", "Office"]
    if normalized.get("property_type") not in prop_types:
        normalized["property_type"] = "Multifamily"

    if normalized.get("lender_profile") not in ["Bank / Credit Union", "LifeCo / Core", "Bridge / Private", "CMHC Multifamily"]:
        normalized["lender_profile"] = "Bank / Credit Union"

    rr = normalized.get("rent_roll_dict", [])
    if isinstance(rr, dict):
        rr = [rr]
    normalized["rent_roll_dict"] = normalize_rent_roll_columns(pd.DataFrame(rr)).to_dict("records")

    return normalized


def extract_clean_state() -> dict:
    """Collect deal fields for storage and export."""
    keys = [
        "deal_id", "deal_name", "sponsor", "property_address", "property_type", "transaction_type",
        "lender_profile", "purchase_price", "appraisal", "noi", "target_ltv",
        "target_ltc", "target_dscr", "target_dy", "rate", "amort", "term",
        "is_io", "fees", "closing_costs", "reserves", "rent_roll_dict", "last_saved_at"
    ]
    state = {}
    for k in keys:
        if k in st.session_state:
            val = st.session_state[k]
            if isinstance(val, np.generic):
                val = val.item()
            state[k] = val
    state["rent_roll_dict"] = normalize_rent_roll_columns(
        pd.DataFrame(state.get("rent_roll_dict", []))
    ).to_dict("records")
    return state


def reset_to_new_deal(name: str = "", rerun: bool = True):
    fresh = DEFAULT_STATE.copy()
    clean_name = str(name or "Untitled Deal").strip() or "Untitled Deal"
    fresh["deal_id"] = f"deal_{int(datetime.now().timestamp())}_{uuid.uuid4().hex[:8]}"
    fresh["deal_name"] = clean_name
    fresh["sponsor"] = ""
    fresh["property_address"] = ""
    fresh["purchase_price"] = 0.0
    fresh["appraisal"] = 0.0
    fresh["noi"] = 0.0
    fresh["closing_costs"] = 0.0
    fresh["reserves"] = 0.0
    fresh["rent_roll_dict"] = []

    for k, v in fresh.items():
        st.session_state[k] = v

    st.session_state.unsaved_changes = True

    if rerun:
        st.rerun()


for k, v in normalize_loaded_state(DEFAULT_STATE).items():
    if k not in st.session_state:
        st.session_state[k] = v


class DatabaseManager:
    @staticmethod
    def init_db():
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA busy_timeout=30000")
            c = conn.cursor()
            c.execute(
                """CREATE TABLE IF NOT EXISTS deals
                (id TEXT PRIMARY KEY, name TEXT, state_json TEXT, updated_at TIMESTAMP)"""
            )
            c.execute(
                """CREATE TABLE IF NOT EXISTS audit_log
                (id INTEGER PRIMARY KEY AUTOINCREMENT, user TEXT, action TEXT, details TEXT, timestamp TIMESTAMP)"""
            )
            c.execute(
                """CREATE TABLE IF NOT EXISTS documents
                (id TEXT PRIMARY KEY, deal_id TEXT, filename TEXT, category TEXT, path TEXT, uploaded_at TIMESTAMP)"""
            )
            conn.commit()

    @staticmethod
    def log_audit(action: str, details: str):
        try:
            with sqlite3.connect(DB_PATH, timeout=30) as conn:
                conn.execute(
                    "INSERT INTO audit_log (user, action, details, timestamp) VALUES (?, ?, ?, ?)",
                    (current_audit_user(), str(action)[:100], str(details)[:1000], datetime.now())
                )
        except Exception:
            pass

    @staticmethod
    def save_deal(deal_id: str, name: str, state: dict):
        state = normalize_loaded_state(state)
        state_json = json.dumps(state, default=str)
        safe_name = str(name or "Untitled Deal").strip() or "Untitled Deal"

        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            conn.execute(
                "INSERT OR REPLACE INTO deals (id, name, state_json, updated_at) VALUES (?, ?, ?, ?)",
                (deal_id, safe_name, state_json, datetime.now())
            )
        DatabaseManager.log_audit("SAVE_DEAL", f"Saved Deal: {safe_name}")

    @staticmethod
    def delete_deal(deal_id: str):
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            c = conn.cursor()
            c.execute("SELECT path FROM documents WHERE deal_id = ?", (deal_id,))
            rows = c.fetchall()

            for row in rows:
                try:
                    file_path = Path(row[0])
                    if file_path.exists() and file_path.is_file():
                        file_path.unlink()
                except Exception:
                    pass

            conn.execute("DELETE FROM deals WHERE id = ?", (deal_id,))
            conn.execute("DELETE FROM documents WHERE deal_id = ?", (deal_id,))

        DatabaseManager.log_audit("DELETE_DEAL", f"Deleted Deal ID and associated files: {deal_id}")

    @staticmethod
    def load_deal(deal_id: str):
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            c = conn.cursor()
            c.execute("SELECT state_json FROM deals WHERE id = ?", (deal_id,))
            row = c.fetchone()

        if not row:
            return None

        try:
            return normalize_loaded_state(json.loads(row[0]))
        except Exception:
            return None

    @staticmethod
    def get_all_deals():
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            return pd.read_sql_query(
                "SELECT id, name, updated_at FROM deals ORDER BY updated_at DESC",
                conn
            )

    @staticmethod
    def save_document(deal_id: str, file, category: str):
        safe_name = clean_filename(file.name)
        doc_id = f"doc_{int(datetime.now().timestamp())}_{uuid.uuid4().hex}_{safe_name}"[:200]
        path = DOC_DIR / doc_id
        path.write_bytes(file.getbuffer())

        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            conn.execute(
                "INSERT INTO documents (id, deal_id, filename, category, path, uploaded_at) VALUES (?, ?, ?, ?, ?, ?)",
                (doc_id, deal_id, safe_name, category, str(path), datetime.now())
            )
        DatabaseManager.log_audit("DOC_UPLOAD", f"Uploaded {safe_name} to {category}")

    @staticmethod
    def delete_document(doc_id: str):
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            c = conn.cursor()
            c.execute("SELECT path, filename FROM documents WHERE id = ?", (doc_id,))
            row = c.fetchone()

            if row:
                try:
                    file_path = Path(row[0])
                    if file_path.exists() and file_path.is_file():
                        file_path.unlink()
                except Exception:
                    pass

                conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
                DatabaseManager.log_audit("DOC_DELETE", f"Deleted document: {row[1]}")


DatabaseManager.init_db()


class CanadianIntel:
    @staticmethod
    @st.cache_data(ttl=3600)
    def get_boc_rates():
        """Fetch Canadian market rates from Bank of Canada Valet."""
        try:
            url = (
                "https://www.bankofcanada.ca/valet/observations/"
                "FXUSDCAD,BD.CDN.2YR.DQ.YLD,BD.CDN.5YR.DQ.YLD,BD.CDN.10YR.DQ.YLD/json"
            )
            res = requests.get(url, timeout=10)
            res.raise_for_status()
            observations = res.json().get("observations", [])

            if not observations:
                return {"error": "No observations returned."}

            latest = observations[-1]

            def get_series(series: str):
                try:
                    raw = latest.get(series, {}).get("v")
                    return float(raw) if raw not in [None, ""] else None
                except Exception:
                    return None

            return {
                "usd_cad": get_series("FXUSDCAD"),
                "2yr_bond": get_series("BD.CDN.2YR.DQ.YLD"),
                "5yr_bond": get_series("BD.CDN.5YR.DQ.YLD"),
                "10yr_bond": get_series("BD.CDN.10YR.DQ.YLD"),
                "date": latest.get("d"),
                "error": None,
            }
        except Exception as e:
            return {"error": str(e)}

    @staticmethod
    @st.cache_data(ttl=86400)
    def verify_corporation(identifier: str) -> dict:
        """
        Verify federal corporation by corporation number or 9-digit business number.
        Broad name-search APIs are inconsistent, so this expects an identifier.
        """
        cleaned = re.sub(r"\D", "", identifier or "")

        if not cleaned:
            return {
                "status": "error",
                "message": "Enter a federal corporation number or 9-digit business number."
            }

        candidate_urls = [
            f"https://www.ic.gc.ca/app/scr/cc/CorporationsCanada/api/corporations/{cleaned}.json?lang=eng",
            f"https://api.ised-isde.canada.ca/corporations/v1/corporations/{cleaned}"
        ]

        last_error = ""
        for url in candidate_urls:
            try:
                res = requests.get(url, timeout=10)
                if res.status_code == 404:
                    continue
                if res.status_code != 200:
                    last_error = f"Status {res.status_code}"
                    continue

                data = res.json()
                corp_data = data[0] if isinstance(data, list) and data else data

                if not isinstance(corp_data, dict) or not corp_data:
                    continue

                primary_name = None
                for key in ["name", "corporationName", "corporation_name"]:
                    if corp_data.get(key):
                        primary_name = corp_data.get(key)
                        break

                names = corp_data.get("corporationNames") or corp_data.get("corporation_names") or []
                if not primary_name and isinstance(names, list):
                    for item in names:
                        if isinstance(item, dict):
                            name_obj = item.get("CorporationName", item)
                            if isinstance(name_obj, dict) and name_obj.get("name"):
                                primary_name = name_obj.get("name")
                                break

                return {
                    "status": "found",
                    "name": primary_name or "Corporation found",
                    "data": corp_data,
                }
            except Exception as e:
                last_error = str(e)

        return {
            "status": "not_found",
            "message": f"Corporation not found or registry unavailable. {last_error}".strip()
        }

    @staticmethod
    @st.cache_data(ttl=86400)
    def geocode_nrcan(address: str):
        """Find an address using Canadian geolocation services."""
        if not address:
            return None

        try:
            res = requests.get(
                "https://geolocator.api.geo.ca/geolocation",
                params={"q": address, "lang": "en", "limit": 1},
                timeout=10,
            )
            if res.status_code == 200:
                data = res.json()
                features = data.get("features", []) if isinstance(data, dict) else []
                if features:
                    feature = features[0]
                    props = feature.get("properties", {})
                    geom = feature.get("geometry", {})
                    coords = geom.get("coordinates", [None, None])
                    return {
                        "label": props.get("title") or props.get("name") or address,
                        "longitude": coords[0] if len(coords) > 0 else None,
                        "latitude": coords[1] if len(coords) > 1 else None,
                        "raw": props,
                    }
        except Exception:
            pass

        try:
            res = requests.get(
                "https://geogratis.gc.ca/services/geolocation/en/locate",
                params={"q": address, "limit": 1},
                timeout=10,
            )
            if res.status_code == 200:
                data = res.json()
                if isinstance(data, list) and data:
                    item = data[0]
                    return {
                        "label": f"{item.get('municipality', '')}, {item.get('provinceCode', '')}".strip(", "),
                        "longitude": item.get("longitude"),
                        "latitude": item.get("latitude"),
                        "raw": item,
                    }
        except Exception:
            pass

        return None


class OCREngine:
    @staticmethod
    def calculate_confidence(match: str, line: str, keyword: str) -> float:
        confidence = 0.5
        k_pos = line.lower().find(keyword)
        m_pos = line.find(match)
        if k_pos >= 0 and m_pos >= 0:
            dist = abs(k_pos - m_pos)
            if dist < 20:
                confidence += 0.35
            elif dist < 50:
                confidence += 0.15
        if "$" in line or "%" in line:
            confidence += 0.05
        if "," in match:
            confidence += 0.05
        return min(confidence, 0.98)

    @staticmethod
    def extract_and_parse(file) -> Tuple[str, dict]:
        text = ""

        try:
            if file.name.lower().endswith(".pdf") and fitz:
                doc = fitz.open(stream=file.read(), filetype="pdf")
                text = "\n".join([page.get_text() for page in doc])
            elif Image and pytesseract:
                text = pytesseract.image_to_string(Image.open(file))
        except Exception:
            return "", {}

        if not text:
            return "", {}

        results = {}
        fields = {
            "Purchase Price / Cost Basis": ["purchase price", "cost basis", "acquisition price", "contract price"],
            "Appraised Value": ["appraised value", "market value", "as-is value"],
            "Gross Income": ["gross potential", "total income", "effective gross", "revenue", "egi"],
            "Vacancy / Credit Loss": ["vacancy", "credit loss", "vacancy loss"],
            "Operating Expenses": ["operating expenses", "total expenses", "opex"],
            "Stabilized NOI": ["net operating income", "noi", "net income"],
            "Debt Service": ["debt service", "annual debt service", "mortgage payment"],
            "CapEx / Reserves": ["capex", "capital expenditures", "replacement reserves"],
        }

        for field, keywords in fields.items():
            for line in text.splitlines():
                line_lower = line.lower()
                for k in keywords:
                    if k in line_lower:
                        matches = re.findall(
                            r"[\$\(]?\s*-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d{1,2})?\)?",
                            line,
                        )
                        if matches:
                            match_str = matches[-1]
                            clean = (
                                match_str.replace("$", "")
                                .replace(",", "")
                                .replace("(", "-")
                                .replace(")", "")
                                .strip()
                            )
                            clean_val = safe_float(clean, 0.0)
                            conf = OCREngine.calculate_confidence(match_str, line, k)
                            results[field] = {
                                "value": clean_val,
                                "confidence": conf,
                                "source_line": line.strip(),
                            }
                            break
                if field in results:
                    break

        return text, results


class ExportEngine:
    @staticmethod
    def generate_excel(state: dict, loan_amt: float, gate: str, amort_df: pd.DataFrame, score: int, tier: str) -> bytes:
        output = io.BytesIO()

        engine = "xlsxwriter" if EXCEL_AVAILABLE else "openpyxl"
        with pd.ExcelWriter(output, engine=engine) as writer:
            fmt_header = None
            if engine == "xlsxwriter":
                wb = writer.book
                fmt_header = wb.add_format({"bold": True, "bg_color": "#CFB87C", "color": "#000000", "border": 1})

            exec_data = {
                "Metric": [
                    "Sponsor", "Property", "Property Type", "Appraisal", "NOI",
                    "Supportable Proceeds", "Constraint", "Deal Score", "Tier"
                ],
                "Value": [
                    state.get("sponsor"), state.get("property_address"), state.get("property_type"),
                    state.get("appraisal"), state.get("noi"), loan_amt, gate, score, tier
                ],
            }
            df_exec = pd.DataFrame(exec_data)
            df_exec.to_excel(writer, sheet_name="Executive Summary", index=False)

            df_rr = normalize_rent_roll_columns(pd.DataFrame(state.get("rent_roll_dict", [])))
            df_rr.to_excel(writer, sheet_name="Rent Roll", index=False)

            if amort_df is not None and not amort_df.empty:
                amort_df.to_excel(writer, sheet_name="Amortization", index=False)

            try:
                with sqlite3.connect(DB_PATH, timeout=30) as conn:
                    audit_df = pd.read_sql_query(
                        "SELECT user, action, details, timestamp FROM audit_log ORDER BY timestamp DESC LIMIT 100",
                        conn,
                    )
                audit_df.to_excel(writer, sheet_name="Audit Log", index=False)
            except Exception:
                pass

            if engine == "xlsxwriter":
                for ws_name in writer.sheets:
                    ws = writer.sheets[ws_name]
                    ws.set_column("A:Z", 18)
                for sheet, df in [
                    ("Executive Summary", df_exec),
                    ("Rent Roll", df_rr),
                    ("Amortization", amort_df if amort_df is not None else pd.DataFrame()),
                ]:
                    if sheet in writer.sheets and fmt_header is not None and df is not None and not df.empty:
                        ws = writer.sheets[sheet]
                        for col_num, value in enumerate(df.columns.values):
                            ws.write(0, col_num, value, fmt_header)

        return output.getvalue()

    @staticmethod
    def generate_pdf(state: dict, loan_amt: float, gate: str, score: int, tier: str, risk_flags: List[str]) -> bytes:
        if not PDF_AVAILABLE:
            return b""

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
        styles = getSampleStyleSheet()
        story = []

        story.append(Paragraph("ALENZA CAPITAL - UNDERWRITING TEAR SHEET", styles["Title"]))
        story.append(Spacer(1, 12))

        meta = f"""
        <b>Prepared:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}<br/>
        <b>Sponsor:</b> {state.get('sponsor') or 'N/A'}<br/>
        <b>Property:</b> {state.get('property_address') or 'N/A'} ({state.get('property_type') or 'N/A'})<br/>
        """
        story.append(Paragraph(meta, styles["Normal"]))
        story.append(Spacer(1, 16))

        exec_summary = f"""
        <b>Supportable Proceeds:</b> ${loan_amt:,.0f}<br/>
        <b>Binding Constraint:</b> {gate}<br/>
        <b>Appraised Value:</b> ${safe_float(state.get('appraisal')):,.0f}<br/>
        <b>Stabilized NOI:</b> ${safe_float(state.get('noi')):,.0f}<br/>
        <b>Deal Score:</b> {score}/1000<br/>
        <b>Classification:</b> {tier}<br/>
        """
        story.append(Paragraph("Executive Summary", styles["Heading2"]))
        story.append(Paragraph(exec_summary, styles["Normal"]))
        story.append(Spacer(1, 16))

        story.append(Paragraph("Risk & Structural Analysis", styles["Heading2"]))
        for flag in risk_flags:
            clean_flag = re.sub(r"[^\x00-\x7F]+", "", str(flag)).strip()
            story.append(Paragraph(f"- {clean_flag}", styles["Normal"]))
        story.append(Spacer(1, 16))

        disclaimer = (
            "This document was prepared in Alenza OS. It is indicative only and does not constitute "
            "a loan commitment, credit approval, or legal advice. Subject to final lender diligence."
        )
        story.append(Paragraph("Disclaimer", styles["Heading2"]))
        story.append(Paragraph(disclaimer, styles["Normal"]))

        doc.build(story)
        return buffer.getvalue()


class UnderwritingEngine:
    LENDER_PROFILES = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05},
    }

    @staticmethod
    def size_loan(noi, appraisal, pp, closing, reserves, fees_pct, rate, amort, term, is_io, t_ltv, t_ltc, t_dscr, t_dy):
        noi = safe_float(noi)
        appraisal = max(0, safe_float(appraisal))
        pp = max(0, safe_float(pp))
        closing = max(0, safe_float(closing))
        reserves = max(0, safe_float(reserves))
        fees_pct = max(0, safe_float(fees_pct))
        rate = max(0, safe_float(rate))
        amort = max(1, safe_int(amort, 25))
        t_ltv = max(0, safe_float(t_ltv))
        t_ltc = max(0, safe_float(t_ltc))
        t_dscr = max(0.01, safe_float(t_dscr, 1.25))
        t_dy = max(0.0001, safe_float(t_dy, 0.085))

        base_cost = pp + closing + reserves
        loan = appraisal * t_ltv if appraisal > 0 else 0
        total_uses = base_cost
        gates = {"LTV": 0, "LTC": 0, "DSCR": 0, "Debt Yield": 0}

        for _ in range(7):
            total_uses = base_cost + (loan * fees_pct)
            ltv_loan = appraisal * t_ltv if appraisal > 0 else 0
            ltc_loan = total_uses * t_ltc if total_uses > 0 else 0
            dy_loan = noi / t_dy if t_dy > 0 else 0

            m_rate = rate / 12
            if is_io:
                dscr_loan = ((noi / t_dscr) / 12) / m_rate if m_rate > 0 else 0
            else:
                if m_rate > 0:
                    dscr_loan = ((noi / t_dscr) / 12) * ((1 - (1 + m_rate) ** -(amort * 12)) / m_rate)
                else:
                    dscr_loan = noi * amort / t_dscr

            gates = {
                "LTV": max(0, ltv_loan),
                "LTC": max(0, ltc_loan),
                "DSCR": max(0, dscr_loan),
                "Debt Yield": max(0, dy_loan),
            }
            loan = min(gates.values()) if gates else 0

        gate = min(gates, key=gates.get) if gates else "N/A"
        req_equity = total_uses - loan
        return loan, gate, gates, total_uses, req_equity

    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io):
        loan_amt = max(0, safe_float(loan_amt))
        rate = max(0, safe_float(rate))
        amort_yrs = max(1, safe_int(amort_yrs, 25))
        term_yrs = max(1, safe_int(term_yrs, 5))

        m_rate = rate / 12
        pmts = amort_yrs * 12
        term_months = int(term_yrs * 12)

        if loan_amt <= 0:
            return pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]), 0.0, 0.0

        if is_io:
            pmt = loan_amt * m_rate
        else:
            if m_rate > 0:
                pmt = (loan_amt * m_rate) / (1 - (1 + m_rate) ** -pmts)
            else:
                pmt = loan_amt / pmts

        sched, bal = [], loan_amt
        for i in range(1, term_months + 1):
            int_pmt = bal * m_rate
            prin_pmt = 0 if is_io else max(0, pmt - int_pmt)
            bal = max(0, bal - prin_pmt)
            sched.append({
                "Period": i,
                "Payment": pmt,
                "Principal": prin_pmt,
                "Interest": int_pmt,
                "Balance": bal,
            })
            if bal <= 0:
                break

        df = pd.DataFrame(sched)
        balloon = bal if bal > 0 else 0
        return df, pmt, balloon

    @staticmethod
    def rent_roll_metrics(df):
        df = normalize_rent_roll_columns(df)

        if df.empty:
            return 0, 0, 0, 0, 0, 0

        total_sf = df["SF"].sum()
        vacant_labels = ["vacant", "empty", "available", "vacancy", "n/a", "none", ""]
        occ_df = df[
            (~df["Tenant"].str.lower().isin(vacant_labels)) &
            (df["SF"] > 0)
        ].copy()

        occ_sf = occ_df["SF"].sum()
        ann_rent = occ_df["Monthly Rent"].sum() * 12
        occupancy = occ_sf / total_sf if total_sf > 0 else 0
        rent_psf = ann_rent / occ_sf if occ_sf > 0 else 0
        walt = (occ_df["Remaining Term"] * occ_df["SF"]).sum() / occ_sf if occ_sf > 0 else 0
        exp_1yr = occ_df[occ_df["Remaining Term"] <= 1.0]["SF"].sum() / occ_sf if occ_sf > 0 else 0

        return total_sf, occupancy, ann_rent, rent_psf, walt, exp_1yr

    @staticmethod
    def score_deal(ltv, ltc, dscr, dy, profile_name) -> Tuple[int, str]:
        limits = UnderwritingEngine.LENDER_PROFILES.get(
            profile_name,
            UnderwritingEngine.LENDER_PROFILES["Bank / Credit Union"]
        )

        ltv = max(0, safe_float(ltv))
        ltc = max(0, safe_float(ltc))
        dscr = max(0, safe_float(dscr))
        dy = max(0, safe_float(dy))

        ltv_score = max(0, 300 * (1 - (ltv / limits["max_ltv"]))) if limits["max_ltv"] > 0 else 0

        if limits["min_dscr"] > 1.0:
            dscr_score = max(0, 300 * ((dscr - 1.0) / (limits["min_dscr"] - 1.0))) if dscr > 1 else 0
        else:
            dscr_score = 300 if dscr >= 1.0 else 0

        dy_score = max(0, 200 * min(1.5, dy / limits["min_dy"])) if limits["min_dy"] > 0 else 0
        ltc_score = max(0, 200 * (1 - ltc))
        score = min(1000, int(ltv_score + dscr_score + dy_score + ltc_score))

        if score >= 850:
            tier = "Tier 1 | Institutional Core"
        elif score >= 700:
            tier = "Tier 2 | Conventional Bankable"
        elif score >= 550:
            tier = "Tier 3 | Alternative / Debt Fund"
        else:
            tier = "Tier 4 | Private / Restructure"
        return score, tier


class RiskAnalysisEngine:
    @staticmethod
    def generate_narrative(actual_ltv: float, actual_dscr: float, walt: float, exp_1yr: float, is_io: bool, req_equity: float) -> List[str]:
        flags = []

        if actual_ltv > 0.75:
            flags.append(f"⚠️ **High Leverage:** Transaction requires {actual_ltv * 100:.1f}% LTV, pushing beyond conventional parameters.")
        elif 0 < actual_ltv < 0.60:
            flags.append("✅ **Conservative Capitalization:** Low LTV indicates strong sponsor equity commitment.")

        if actual_dscr > 0 and actual_dscr < 1.20:
            flags.append(f"⚠️ **Tight Cash Flow:** DSCR is exceptionally thin at {actual_dscr:.2f}x.")

        if is_io and actual_dscr < 1.25:
            flags.append("⚠️ **Structural Masking:** Interest-only structure may be masking amortizing weakness.")

        if walt > 0 and walt < 2.5:
            flags.append(f"⚠️ **Short WALT:** WALT is {walt:.1f} years. Lenders will require significant leasing reserves.")

        if exp_1yr > 0.30:
            flags.append(f"🚨 **Rollover Exposure:** {exp_1yr * 100:.1f}% of occupied SF is expiring within 12 months.")

        if req_equity < 0:
            flags.append("🚨 **Capital Stack Inversion:** Required equity is negative, indicating a cash-out scenario.")

        if not flags:
            flags.append("✅ **Clean Profile:** No major automated structural or cash-flow risk flags detected.")

        return flags


class MarketCompsEngine:
    @staticmethod
    def generate_comps(property_type: str, noi: float) -> pd.DataFrame:
        noi = max(1, safe_float(noi, 1))
        np.random.seed(int(noi) % 1000000)

        base_cap = {
            "Multifamily": 0.045,
            "Industrial": 0.055,
            "Retail": 0.065,
            "Office": 0.075,
        }.get(property_type, 0.06)

        comps = []
        for i in range(1, 6):
            cap_variance = np.random.uniform(-0.0075, 0.0075)
            comp_cap = max(0.01, base_cap + cap_variance)
            comp_value = (noi * np.random.uniform(0.8, 1.2)) / comp_cap
            comps.append({
                "Comparable": f"{property_type} Asset {chr(64 + i)}",
                "Distance (km)": round(np.random.uniform(0.5, 8.0), 1),
                "Sale Date": f"202{np.random.randint(4, 6)}-{np.random.randint(1, 13):02d}",
                "Cap Rate": f"{comp_cap * 100:.2f}%",
                "Est. Value": f"${comp_value:,.0f}",
            })

        return pd.DataFrame(comps)


st.markdown(
    """
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

    /* Sidebar captions */
    section[data-testid="stSidebar"] .stCaption,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] p[data-testid="stMarkdownContainer"] > em {
        color: #9CA3AF !important;
    }

    /* Inputs */
    .stTextInput input, .stNumberInput input, .stSelectbox select {
        background-color: #111827 !important;
        border: 1px solid #1E293B !important;
        color: #F3F4F6 !important;
    }

    .stTextArea textarea {
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

    /* File uploader */
    section[data-testid="stFileUploaderDropzone"] {
        background-color: #111827 !important;
        border: 1px dashed #CFB87C !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


with st.sidebar:
    st.title("🏛️ ALENZA OS")

    with st.expander("📁 PIPELINE MANAGER", expanded=True):
        all_deals = DatabaseManager.get_all_deals()

        deal_lookup = {}
        deal_options = ["-- Start New Deal --"]
        if not all_deals.empty:
            for _, row in all_deals.iterrows():
                short_id = str(row["id"])[-8:]
                updated = pd.to_datetime(row["updated_at"], errors="coerce")
                updated_label = updated.strftime("%Y-%m-%d %H:%M") if pd.notna(updated) else "no date"
                label = f"{row['name']} · {updated_label} · {short_id}"
                deal_options.append(label)
                deal_lookup[label] = row["id"]

        selected = st.selectbox("Select Deal", deal_options)
        selected_deal_id = deal_lookup.get(selected)
        new_deal_name = st.text_input("New Deal Name", value="Untitled Deal")

        c_new, c_load = st.columns(2)
        c_del, c_dup = st.columns(2)

        with c_new:
            if st.button("➕ New"):
                reset_to_new_deal(new_deal_name)

        with c_load:
            if st.button("📂 Load") and selected_deal_id:
                loaded_state = DatabaseManager.load_deal(selected_deal_id)
                if loaded_state:
                    loaded_state = normalize_loaded_state(loaded_state)
                    for k, v in loaded_state.items():
                        st.session_state[k] = v
                    st.session_state.deal_id = loaded_state.get("deal_id", selected_deal_id)
                    st.session_state.unsaved_changes = False
                    st.rerun()
                else:
                    st.error("Could not load selected deal.")

        confirm_delete = st.checkbox("Confirm delete selected deal", value=False)

        with c_del:
            if st.button("🗑️ Del") and selected_deal_id:
                if confirm_delete:
                    DatabaseManager.delete_deal(selected_deal_id)
                    st.success("Deal deleted.")
                    st.rerun()
                else:
                    st.warning("Check confirm before deleting.")

        with c_dup:
            if st.button("📑 Dup.") and selected_deal_id:
                st.session_state.deal_id = f"deal_{int(datetime.now().timestamp())}_{uuid.uuid4().hex[:8]}"
                st.session_state.deal_name = f"Copy of {st.session_state.get('deal_name', 'Untitled Deal')}"
                st.session_state.unsaved_changes = True
                st.warning("Duplicated in memory. Save to persist.")

    st.markdown("---")

    s = st.session_state

    with st.expander("🏢 ASSET PROFILE", expanded=True):
        s.deal_name = st.text_input("Deal Name", value=s.get("deal_name", "Untitled Deal"))
        s.sponsor = st.text_input("Sponsor", value=s.get("sponsor", ""))
        s.property_address = st.text_input("Address", value=s.get("property_address", ""))

        property_types = ["Multifamily", "Industrial", "Retail", "Office"]
        prop_idx = property_types.index(s.get("property_type", "Multifamily")) if s.get("property_type") in property_types else 0
        s.property_type = st.selectbox("Type", property_types, index=prop_idx)

        s.appraisal = st.number_input("Appraisal ($)", value=safe_float(s.get("appraisal")), step=100000.0, min_value=0.0)
        s.purchase_price = st.number_input("Cost Basis ($)", value=safe_float(s.get("purchase_price")), step=100000.0, min_value=0.0)
        s.noi = st.number_input("Stabilized NOI ($)", value=safe_float(s.get("noi")), step=10000.0, min_value=0.0)

    with st.expander("📊 CREDIT POLICY", expanded=True):
        profiles = list(UnderwritingEngine.LENDER_PROFILES.keys())
        profile_idx = profiles.index(s.get("lender_profile")) if s.get("lender_profile") in profiles else 0
        s.lender_profile = st.selectbox("Policy Preset", profiles, index=profile_idx)

        preset = UnderwritingEngine.LENDER_PROFILES[s.lender_profile]
        s.target_ltv = st.slider("Max LTV %", 50.0, 95.0, float(normalize_percent(s.get("target_ltv"), preset["max_ltv"]) * 100), step=0.5) / 100
        s.target_dscr = st.slider("Min DSCR x", 1.0, 1.75, float(safe_float(s.get("target_dscr"), preset["min_dscr"])), step=0.05)
        s.target_dy = st.slider("Min DY %", 5.0, 15.0, float(normalize_percent(s.get("target_dy"), preset["min_dy"]) * 100), step=0.25) / 100
        s.target_ltc = st.slider("Max LTC %", 50.0, 100.0, float(normalize_percent(s.get("target_ltc"), 0.80) * 100), step=0.5) / 100

    with st.expander("💰 DEBT STRUCTURE", expanded=True):
        s.is_io = st.checkbox("Interest-Only", value=bool(s.get("is_io", False)))
        s.rate = st.slider("Rate %", 0.0, 15.0, float(normalize_percent(s.get("rate"), 0.0525) * 100), step=0.05) / 100
        s.amort = st.number_input("Amort (Yrs)", value=max(1, safe_int(s.get("amort"), 25)), step=1, min_value=1)
        s.term = st.number_input("Term (Yrs)", value=max(1, safe_int(s.get("term"), 5)), step=1, min_value=1)
        s.fees = st.slider("Fees %", 0.0, 5.0, float(normalize_percent(s.get("fees"), 0.02) * 100), step=0.05) / 100
        s.closing_costs = st.number_input("Closing Costs", value=safe_float(s.get("closing_costs")), step=1000.0, min_value=0.0)
        s.reserves = st.number_input("Reserves", value=safe_float(s.get("reserves")), step=1000.0, min_value=0.0)


s = st.session_state
s.rent_roll_dict = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", []))).to_dict("records")

loan_amt, gate, gates, total_uses, req_equity = UnderwritingEngine.size_loan(
    s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees,
    s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy
)

amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(
    loan_amt, s.rate, s.amort, s.term, s.is_io
)
annual_ds = monthly_pmt * 12

actual_ltv = loan_amt / s.appraisal if safe_float(s.appraisal) else 0
actual_ltc = loan_amt / total_uses if total_uses else 0
actual_dscr = s.noi / annual_ds if annual_ds else 0
actual_dy = s.noi / loan_amt if loan_amt else 0

tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(
    pd.DataFrame(s.rent_roll_dict)
)

score, classification = UnderwritingEngine.score_deal(
    actual_ltv, actual_ltc, actual_dscr, actual_dy, s.lender_profile
)
risk_flags = RiskAnalysisEngine.generate_narrative(
    actual_ltv, actual_dscr, walt, exp1, s.is_io, req_equity
)


headline_sponsor = s.sponsor or "New Deal"
headline_property = s.property_address or "Property Address Pending"

st.title(f"{headline_sponsor} | {headline_property}")
st.caption(f"INSTITUTIONAL WORKSTATION | ACTIVE CONSTRAINT: {gate} | TIER: {classification.upper()}")

m1, m2, m3, m4, m5, m6 = st.columns(6)
m1.metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
m2.metric("ACTUAL LTV", f"{actual_ltv * 100:.1f}%")
m3.metric("ACTUAL LTC", f"{actual_ltc * 100:.1f}%")
m4.metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")
m5.metric("BALLOON", f"${balloon:,.0f}")
m6.metric("DEAL SCORE", f"{score}/1000")

st.markdown("---")


tabs = st.tabs([
    "📊 Sizing & Risk",
    "📝 Rent Roll",
    "📅 Amortization",
    "📈 Market Comps",
    "📎 Diligence Room",
    "🤖 OCR Extract",
    "🇨🇦 Canada Intel",
    "💾 Save & Export",
])


with tabs[0]:
    c1, c2 = st.columns([1.5, 1], gap="large")

    with c1:
        st.subheader("Constraint Analysis")
        df_gates = pd.DataFrame({
            "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
            "Threshold": [
                f"{s.target_ltv * 100:.1f}%",
                f"{s.target_ltc * 100:.1f}%",
                f"{s.target_dscr:.2f}x",
                f"{s.target_dy * 100:.2f}%"
            ],
            "Proceeds Limit": [
                f"${gates.get('LTV', 0):,.0f}",
                f"${gates.get('LTC', 0):,.0f}",
                f"${gates.get('DSCR', 0):,.0f}",
                f"${gates.get('Debt Yield', 0):,.0f}"
            ],
            "Binding": ["✅ YES" if gate == k else "" for k in ["LTV", "LTC", "DSCR", "Debt Yield"]],
        })
        st.dataframe(df_gates, hide_index=True, use_container_width=True)

        st.subheader("Sources & Uses")
        df_su = pd.DataFrame({
            "Uses": ["Cost Basis", "Closing Costs", "Reserves", "Financing Fees", "Total"],
            "U Amount": [s.purchase_price, s.closing_costs, s.reserves, loan_amt * s.fees, total_uses],
            "Sources": ["Senior Debt", "Sponsor Equity", "", "", "Total"],
            "S Amount": [loan_amt, req_equity, 0, 0, total_uses],
        })
        st.dataframe(
            df_su.style.format({"U Amount": "${:,.0f}", "S Amount": "${:,.0f}"}),
            hide_index=True,
            use_container_width=True,
        )

    with c2:
        st.subheader("Executive Risk Narrative")
        for flag in risk_flags:
            if "⚠️" in flag or "🚨" in flag:
                st.warning(flag)
            else:
                st.info(flag)


with tabs[1]:
    st.subheader("Interactive Rent Roll")

    rr_upload = st.file_uploader(
        "Auto-import Rent Roll CSV or Excel",
        type=["csv", "xlsx", "xls"],
        key="rent_roll_upload",
    )

    if rr_upload:
        try:
            filename = rr_upload.name.lower()
            rr_upload.seek(0)
            if filename.endswith(".csv"):
                imported_rr = pd.read_csv(rr_upload)
            elif filename.endswith(".xls"):
                if not XLRD_AVAILABLE:
                    raise RuntimeError("Reading .xls files requires xlrd. Add xlrd to requirements.txt or upload .xlsx/.csv.")
                imported_rr = pd.read_excel(rr_upload, engine="xlrd")
            else:
                if not OPENPYXL_AVAILABLE:
                    raise RuntimeError("Reading .xlsx files requires openpyxl. Add openpyxl to requirements.txt or upload .csv.")
                imported_rr = pd.read_excel(rr_upload, engine="openpyxl")

            imported_rr = normalize_rent_roll_columns(imported_rr)
            st.write("Preview")
            st.dataframe(imported_rr, use_container_width=True, hide_index=True)

            if st.button("Apply Imported Rent Roll"):
                s.rent_roll_dict = imported_rr.to_dict("records")
                DatabaseManager.log_audit("RENT_ROLL_IMPORT", f"Imported rent roll: {rr_upload.name}")
                st.success("Rent roll imported.")
                st.rerun()

        except Exception as e:
            st.error(f"Could not import rent roll: {e}")

    rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

    edited_rr = st.data_editor(
        rr_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tenant": st.column_config.TextColumn("Tenant"),
            "SF": st.column_config.NumberColumn("SF", min_value=0, step=100),
            "Remaining Term": st.column_config.NumberColumn("Remaining Term", min_value=0.0, step=0.25),
            "Monthly Rent": st.column_config.NumberColumn("Monthly Rent", min_value=0.0, step=100.0, format="$%.2f"),
        },
    )

    s.rent_roll_dict = normalize_rent_roll_columns(edited_rr).to_dict("records")
    tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(pd.DataFrame(s.rent_roll_dict))

    r1, r2, r3, r4, r5, r6 = st.columns(6)
    r1.metric("Total SF", f"{tot_sf:,.0f}")
    r2.metric("Occupancy", f"{occ * 100:.1f}%")
    r3.metric("Annual Rent", f"${ann_rent:,.0f}")
    r4.metric("Rent PSF", f"${psf:,.2f}")
    r5.metric("WALT", f"{walt:.2f} Yrs")
    r6.metric("12-Mo Rollover", f"{exp1 * 100:.1f}%")


with tabs[2]:
    st.subheader(f"Amortization Schedule: {s.term} Year Term")

    if amort_df is None or amort_df.empty:
        st.warning("No amortization schedule available. Enter appraisal, NOI, and debt terms.")
    else:
        chart_df = amort_df.copy()
        chart_df["Year"] = ((chart_df["Period"] - 1) // 12) + 1

        annual_df = chart_df.groupby("Year", as_index=False).agg({
            "Principal": "sum",
            "Interest": "sum",
            "Payment": "sum",
            "Balance": "last",
        })

        c1, c2, c3 = st.columns(3)
        c1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
        c2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
        c3.metric("Balloon Balance", f"${balloon:,.0f}")

        st.write("### Monthly Principal vs Interest")
        monthly_pi = chart_df.set_index("Period")[["Principal", "Interest"]]
        st.bar_chart(monthly_pi)

        st.write("### Loan Balance Over Term")
        balance_chart = chart_df.set_index("Period")[["Balance"]]
        st.line_chart(balance_chart)

        st.write("### Annual Summary")
        st.dataframe(
            annual_df.style.format({
                "Principal": "${:,.2f}",
                "Interest": "${:,.2f}",
                "Payment": "${:,.2f}",
                "Balance": "${:,.2f}",
            }),
            use_container_width=True,
            hide_index=True,
        )

        st.write("### Full Monthly Schedule")
        st.dataframe(
            amort_df.style.format({
                "Payment": "${:,.2f}",
                "Principal": "${:,.2f}",
                "Interest": "${:,.2f}",
                "Balance": "${:,.2f}",
            }),
            use_container_width=True,
            height=350,
            hide_index=True,
        )


with tabs[3]:
    st.subheader("Simulated Market Comparables")
    st.caption("Comparison set based on property type and NOI. Replace with verified broker/valuation comps before committee.")
    comps_df = MarketCompsEngine.generate_comps(s.property_type, s.noi)
    st.dataframe(comps_df, hide_index=True, use_container_width=True)


with tabs[4]:
    st.subheader("Diligence Vault & Gap Analysis")
    REQUIRED_DOCS = ["Appraisal", "Phase I ESA", "T12 Financials", "Rent Roll", "Sponsor Bio", "Purchase Agreement"]

    try:
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            docs = pd.read_sql_query(
                "SELECT id, filename, category, uploaded_at FROM documents WHERE deal_id = ?",
                conn,
                params=(s.deal_id,),
            )
    except Exception:
        docs = pd.DataFrame(columns=["id", "filename", "category", "uploaded_at"])

    d1, d2 = st.columns(2)

    with d1:
        st.write("### Upload Document")
        cat = st.selectbox("Category", REQUIRED_DOCS + ["Other"])
        doc_file = st.file_uploader("Drop File Here", key="vault_upload")
        if st.button("Save to Vault") and doc_file:
            DatabaseManager.save_document(s.deal_id, doc_file, cat)
            st.success(f"Saved {doc_file.name}")
            st.rerun()

    with d2:
        st.write("### Package Gap Analysis")
        uploaded_cats = docs["category"].tolist() if not docs.empty and "category" in docs.columns else []
        gap_df = pd.DataFrame({
            "Requirement": REQUIRED_DOCS,
            "Status": ["✅ Uploaded" if c in uploaded_cats else "❌ Missing" for c in REQUIRED_DOCS],
        })
        st.dataframe(gap_df, hide_index=True, use_container_width=True)

    st.write("### Vault Inventory")
    if not docs.empty:
        st.dataframe(docs[["filename", "category", "uploaded_at"]], use_container_width=True, hide_index=True)
        doc_to_delete = st.selectbox("Select Document to Delete", ["-- None --"] + docs["filename"].tolist())
        if st.button("🗑️ Delete Selected Document") and doc_to_delete != "-- None --":
            doc_id_to_del = docs.loc[docs["filename"] == doc_to_delete, "id"].values[0]
            DatabaseManager.delete_document(doc_id_to_del)
            st.success("Document deleted.")
            st.rerun()
    else:
        st.info("No documents uploaded yet.")


with tabs[5]:
    st.subheader("Financial Document Extraction")

    if not OCR_AVAILABLE:
        st.warning("OCR dependencies not found. Install pytesseract, pillow, and pymupdf to enable extraction.")
    else:
        st.caption("Local OCR is best for clean statements. For scanned multi-column offering memorandums, use a structured document service such as Textract or Document AI before importing results.")

    uploaded_fin = st.file_uploader(
        "Upload Appraisal / T12 / Operating Statement (PDF/Image)",
        type=["pdf", "png", "jpg", "jpeg"],
        key="ocr_upload",
    )

    if uploaded_fin and OCR_AVAILABLE:
        with st.spinner("Extracting parameters with confidence scoring..."):
            text, extracted = OCREngine.extract_and_parse(uploaded_fin)

        if extracted:
            st.success("Extraction complete.")
            st.dataframe(pd.DataFrame(extracted).T, use_container_width=True)

            with st.expander("Raw Extracted Text"):
                st.text_area("OCR Text", text[:10000], height=250)

            if st.button("Apply Parameters to Underwriting Model"):
                if "Stabilized NOI" in extracted:
                    s.noi = extracted["Stabilized NOI"]["value"]
                if "Purchase Price / Cost Basis" in extracted:
                    s.purchase_price = extracted["Purchase Price / Cost Basis"]["value"]
                if "Appraised Value" in extracted:
                    s.appraisal = extracted["Appraised Value"]["value"]
                DatabaseManager.log_audit("OCR_APPLY", "Applied extracted parameters")
                st.success("Model updated.")
                st.rerun()
        else:
            st.warning("Could not identify high-confidence parameters.")


with tabs[6]:
    st.subheader("🇨🇦 Sovereign Intelligence")
    ca1, ca2 = st.columns(2)

    with ca1:
        st.write("### Live Bank of Canada Rates")
        boc = CanadianIntel.get_boc_rates()

        if boc and not boc.get("error"):
            if boc.get("5yr_bond") is not None:
                st.metric("5-Year Canada Yield", f"{boc['5yr_bond']:.2f}%")
            if boc.get("2yr_bond") is not None:
                st.metric("2-Year Canada Yield", f"{boc['2yr_bond']:.2f}%")
            if boc.get("10yr_bond") is not None:
                st.metric("10-Year Canada Yield", f"{boc['10yr_bond']:.2f}%")
            if boc.get("usd_cad") is not None:
                st.metric("USD/CAD", f"{boc['usd_cad']:.4f}")
            st.caption(f"Last Updated: {boc.get('date')}")
        else:
            st.info("Bank of Canada data is currently unavailable.")
            if boc and boc.get("error"):
                with st.expander("Connection detail"):
                    st.code(boc.get("error"))

    with ca2:
        st.write("### Federal Corporation Registry")
        corp = st.text_input("Verify Federal Corporation Number or BN9")
        if corp:
            result = CanadianIntel.verify_corporation(corp)
            if result["status"] == "found":
                st.success(f"Verified: {result.get('name')}")
                st.json(result.get("data", {}))
            elif result["status"] == "not_found":
                st.warning(result["message"])
            else:
                st.error(result["message"])

        st.write("### NRCan / Geo.ca Address Validation")
        nrcan_query = st.text_input("Verify Address Coordinates")
        if nrcan_query:
            verified = CanadianIntel.geocode_nrcan(nrcan_query)
            if verified:
                st.success(f"📍 Standardized: {verified.get('label')}")
                st.write(f"Latitude: `{verified.get('latitude')}`")
                st.write(f"Longitude: `{verified.get('longitude')}`")
            else:
                st.warning("Could not verify via NRCan / Geo.ca.")


with tabs[7]:
    st.subheader("Save & Package Export")
    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("💾 Save Deal Record to Database", use_container_width=True):
            st.session_state.last_saved_at = datetime.now().isoformat(timespec="seconds")
            clean_state = extract_clean_state()
            deal_name = clean_state.get("deal_name") or clean_state.get("sponsor") or "Untitled Deal"
            DatabaseManager.save_deal(s.deal_id, deal_name, clean_state)
            st.session_state.unsaved_changes = False
            st.success(f"Deal saved at {st.session_state.last_saved_at}.")

    clean_state = extract_clean_state()
    export_stamp = datetime.now().strftime("%Y%m%d_%H%M")

    try:
        excel_bytes = ExportEngine.generate_excel(clean_state, loan_amt, gate, amort_df, score, classification)
        excel_ready = True
    except Exception as e:
        excel_bytes = b""
        excel_ready = False
        st.error(f"Excel export failed: {e}")

    with c2:
        if excel_ready:
            st.download_button(
                "📊 Download Excel Model",
                data=excel_bytes,
                file_name=f"{clean_filename(s.get('deal_name') or s.sponsor or 'Alenza')}_{export_stamp}_Model.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📊 Excel Unavailable", disabled=True, use_container_width=True)

    with c3:
        buf = io.BytesIO()
        try:
            with zipfile.ZipFile(buf, "w") as z:
                if excel_ready:
                    z.writestr(f"Underwriting_Model_{export_stamp}.xlsx", excel_bytes)

                if PDF_AVAILABLE:
                    pdf_bytes = ExportEngine.generate_pdf(clean_state, loan_amt, gate, score, classification, risk_flags)
                    if pdf_bytes:
                        z.writestr(f"Executive_Summary_{export_stamp}.pdf", pdf_bytes)

                if DB_PATH.exists():
                    z.write(DB_PATH, f"Database_Backup_{export_stamp}.db")

                try:
                    with sqlite3.connect(DB_PATH, timeout=30) as conn:
                        vault_docs = pd.read_sql_query(
                            "SELECT filename, path FROM documents WHERE deal_id = ?",
                            conn,
                            params=(s.deal_id,),
                        )
                    for _, row in vault_docs.iterrows():
                        p = Path(row["path"])
                        if p.exists() and p.is_file():
                            z.write(p, f"Diligence_Vault/{row['filename']}")
                except Exception:
                    pass

            st.download_button(
                "📦 Download Full Deal Package",
                buf.getvalue(),
                file_name=f"{clean_filename(s.get('deal_name') or s.sponsor or 'Alenza')}_{export_stamp}_Package.zip",
                mime="application/zip",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Package export failed: {e}")


with st.expander("Deployment Notes"):
    st.write("SQLite is suitable for local or low-concurrency use. For a shared team deployment, move the persistence layer to PostgreSQL.")
    st.write("Keep alenza_data/ out of source control. Store API keys in Streamlit secrets or environment variables.")
    if not OPENPYXL_AVAILABLE:
        st.warning("openpyxl is not installed. .xlsx rent-roll uploads may fail.")
    if not XLRD_AVAILABLE:
        st.info("xlrd is not installed. Legacy .xls rent-roll uploads are disabled until xlrd is added.")

st.markdown("---")
st.caption(
    "⚠️ **DISCLAIMER:** ALENZA CAPITAL OS is an indicative modeling tool. Outputs do not constitute "
    "a loan commitment, appraisal, or legal advice. Final terms are subject to formal credit committee "
    "approval and third-party diligence verification."
)
