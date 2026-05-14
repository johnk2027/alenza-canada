"""
Alenza Capital OS - Single-file Streamlit CRE Underwriting Workstation.

GitHub / reviewer note
----------------------
This file intentionally remains single-file because the deployment target is simple
Streamlit hosting and the owner requested paste-friendly distribution. To keep the
code humane and reviewable, the file is divided into numbered sections and includes
plain-English notes near the financial and data-science logic.

Model-governance note
---------------------
This is an underwriting decision-support tool, not a credit approval engine. It
shows deterministic sizing gates, simulated market context, API-backed macro data
when available, and Monte Carlo stress analytics. Any final credit decision should
be supported by source documents, lender policy, and professional judgment.
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
import threading
import time
import uuid
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple, TypedDict

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
        "plotly": False, "crypto": False, "numpy_financial": False, "pydantic": False,
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
    try: import pydantic; deps["pydantic"] = True
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
if DEPS.get("pydantic"):
    try:
        from pydantic import BaseModel, Field, ValidationError
    except Exception:
        DEPS["pydantic"] = False

# =============================================================================
# 3. CONSTANTS, CONVENTIONS & EXPLICIT STATE CONTRACT
# =============================================================================
VERSION = "5.1.1-final"
MAX_UPLOAD_MB = 50
KDF_ITERATIONS = 600_000  # 2026 hardening: PBKDF2-SHA256 work factor for password-derived document/deal encryption.
EPS = 1e-9
DATA_DIR = Path("alenza_data")
DB_PATH = DATA_DIR / "alenza_platform.db"
DOC_DIR = DATA_DIR / "documents"
DB_WRITE_LOCK = threading.RLock()  # Process-local guard; SQLite WAL/busy_timeout handle file-level contention.
FERNET_CACHE: Dict[str, Any] = {}  # Per-process cache prevents repeating the expensive PBKDF2 derivation on every save.
MASTER_KDF_SALT = b"AlenzaCapitalOS:v9:master-fernet"

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

MODEL_GOVERNANCE_NOTES = {
    "Purpose": "Decision-support only; not a final credit approval model.",
    "Data Provenance": "BoC and StatsCan are fetched live when available; 2026 fallback data is explicitly labeled simulated.",
    "Risk Method": "Monte Carlo engine applies vectorized stochastic shocks and vectorized Newton IRR solves; convergence rate is disclosed.",
    "Human Review": "Final conclusions should be checked against lender policy, source documents, and market evidence.",
}


class ReturnMetrics(TypedDict):
    """Explicit output contract for investment-return calculations."""
    IRR: float
    EM: float
    Exit_NOI: float
    Gross_Exit: float
    Net_Exit: float
    Total_CF: float


class LoanSizingResult(TypedDict):
    """Documented shape for deterministic loan-sizing results."""
    loan: float
    binding_gate: str
    gates: Dict[str, Any]
    total_uses: float
    required_equity: float

DEAL_STATE_KEYS = [
    "deal_id", "deal_name", "sponsor", "property_address", "property_type", "transaction_type", "lender_profile",
    "purchase_price", "appraisal", "noi", "rate", "amort", "term", "refi_amort", "is_io", "fees", 
    "closing_costs", "reserves", "target_ltv", "target_ltc", "target_dscr", "target_dy", 
    "mezz_debt", "pref_equity", "mezz_rate", "pref_rate", "pf_rev_growth", "pf_exp_growth", 
    "pf_exp_ratio", "pf_exit_cap", "pf_sell_costs", "pf_term_growth", "rate_lock_enabled", 
    "rate_lock_spread_bps", "mc_sims", "mc_noi_vol", "mc_rate_vol_bps", "mc_exit_cap_vol_bps", "rent_roll_dict", "diligence_notes"
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
        "refi_amort": 25,
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
        # Monte Carlo defaults: deliberately conservative but editable in the UI.
        "mc_sims": 2000,
        "mc_noi_vol": 0.075,
        "mc_rate_vol_bps": 75,
        "mc_exit_cap_vol_bps": 50,
        "rent_roll_dict": [r.copy() for r in DEFAULT_RENT_ROLL],
        "diligence_notes": "",
    }

class SessionManager:
    """Single source of truth for Streamlit session state.

    Streamlit reruns the file on every interaction. This manager centralizes
    initialization, backfilling, and dirty-state detection so widgets and export
    logic do not drift into separate shadow states.
    """

    @staticmethod
    def initialize() -> None:
        """Hydrate state and run the blocking math-integrity gate once.

        The integrity gate uses ``ValidationEngine.run_financial_self_test_bools()``
        rather than the QA display DataFrame. That prevents the old bug where
        truthy strings such as "❌ FAIL" allowed a broken engine to boot.
        
        This method also backfills missing keys, including newly introduced
        values such as ``refi_amort``, before any sidebar widget reads them.
        UI code should still prefer bracket access or ``st.session_state.get``.
        """
        if not st.session_state.get("boot_self_tests_passed"):
            checks = ValidationEngine.run_financial_self_test_bools()
            if not checks or not all(bool(v) for v in checks.values()):
                st.error("FATAL: Financial Engine Integrity Check Failed. System halted before underwriting outputs.")
                st.dataframe(ValidationEngine.run_financial_self_tests(), hide_index=True, use_container_width=True)
                st.stop()
            st.session_state["boot_self_tests_passed"] = True

        defaults = default_state()
        for key, val in defaults.items():
            if key not in st.session_state or st.session_state[key] is None:
                st.session_state[key] = val
        st.session_state.setdefault("_last_saved_hash", "")
        st.session_state.setdefault("unsaved_changes", False)

    @staticmethod
    def mark_dirty() -> None:
        """Widget callback: mark the in-memory deal as changed immediately.

        Hash comparison is still used as a backstop, but this callback provides
        proactive UI feedback during Streamlit reruns. Widgets that matter to
        saved deal state should call this via ``on_change=SessionManager.mark_dirty``.
        """
        st.session_state["unsaved_changes"] = True

    @staticmethod
    def current_state() -> dict:
        defaults = default_state()
        return {key: st.session_state.get(key, defaults.get(key)) for key in DEAL_STATE_KEYS}

    @staticmethod
    def state_hash(state: dict) -> str:
        return hashlib.sha256(json.dumps(state, default=str, sort_keys=True).encode()).hexdigest()

    @staticmethod
    def refresh_dirty_flag(state: dict) -> str:
        h = SessionManager.state_hash(state)
        st.session_state.unsaved_changes = bool(st.session_state.get("_last_saved_hash") and h != st.session_state.get("_last_saved_hash"))
        return h

    @staticmethod
    def mark_saved(state: dict) -> str:
        h = SessionManager.state_hash(state)
        st.session_state["_last_saved_hash"] = h
        st.session_state["unsaved_changes"] = False
        return h


def get_current_state() -> dict:
    return SessionManager.current_state()

# =============================================================================
# 5. UTILITIES & CRYPTOGRAPHY
# =============================================================================
def safe_float(v: Any, default: float = 0.0) -> float:
    """Minimal internal numeric fallback; never regex-deletes user text.

    This helper is intentionally boring: it accepts numeric objects and plain
    strings that Python can parse directly. It does not remove letters, units,
    hashtags, or currency prose. User-facing financial inputs must use
    parse_financial_float(), which raises NumericParseError instead of silently
    substituting a default.
    """
    if v is None or v == "":
        return default
    try:
        x = float(v)
        return x if np.isfinite(x) else default
    except (TypeError, ValueError, OverflowError):
        return default

class NumericParseError(ValueError):
    """Raised when a financial input cannot be parsed without data loss."""


class UnderwritingError(ValueError):
    """Raised by pure finance engines when an input set is not underwritable.

    The deterministic underwriting engine does not return magic loan amounts for
    invalid input. The UI/service layer catches this exception and turns it into
    an explicit red diagnostic.
    """


def parse_financial_float(v: Any, field: str, default: Optional[float] = None) -> float:
    """Parse a scalar financial value and fail loudly on malformed input.

    A finance workstation must not convert garbage such as ``"1O000"`` to 0.
    This parser accepts common currency/accounting formats but raises a typed
    error if alphabetic characters remain after removing known formatting marks.
    """
    if v is None or v == "":
        if default is not None:
            return float(default)
        raise NumericParseError(f"{field} is blank")
    if isinstance(v, (int, float, np.number)):
        if pd.isna(v) or not np.isfinite(float(v)):
            raise NumericParseError(f"{field} is not finite")
        return float(v)
    text = str(v).strip()
    accounting_negative = text.startswith("(") and text.endswith(")")
    cleaned = text.replace("$", "").replace(",", "").replace("%", "").replace(" ", "")
    if accounting_negative:
        cleaned = "-" + cleaned[1:-1]
    # Fail on accidental letters instead of regex-deleting them.
    if re.search(r"[A-Za-z]", cleaned):
        raise NumericParseError(f"{field} contains letters and was not parsed: {v!r}")
    try:
        out = float(cleaned)
    except Exception as exc:
        raise NumericParseError(f"{field} is not a valid number: {v!r}") from exc
    if not np.isfinite(out):
        raise NumericParseError(f"{field} is not finite")
    return out


def clean_numeric_series(s: pd.Series, number_format: str = "Canadian/US") -> pd.Series:
    """Vectorized rent-roll numeric cleanup with regional separator support.

    What happens here
    -----------------
    1. We capture accounting negatives such as ``(1,000.00)`` before stripping
       formatting.
    2. We normalize either Canadian/US numbers (``1,234.56``) or European
       decimal-comma numbers (``1.234,56`` / ``1 234,56``).
    3. We extract the first valid float-like token with a strict pattern that
       accepts ``-1.5``, ``.5`` and ``5.`` but rejects a solo ``.``.

    Reference note
    --------------
    This function is intentionally for *rent-roll columns*, not free-form deal
    inputs. Rent-roll imports often contain values like ``"€ 500,00 extra"``.
    The extraction allows the row to be surfaced in diagnostics rather than
    crashing the whole app. User-entered deal-level inputs still use
    ``parse_financial_float()``, which fails loudly on malformed text.
    """
    txt = s.astype(str).str.strip()
    is_negative = txt.str.contains(r"^\(.*\)$", regex=True, na=False)

    # Strip common currency symbols and percent signs, but keep separators until
    # we know which locale the user selected.
    clean_txt = txt.str.replace(r"[$€£%]", "", regex=True)

    fmt = str(number_format or "Canadian/US").lower()
    if "euro" in fmt:
        # European: 1.234,56 or 1 234,56 -> 1234.56.
        # Dots/spaces are thousands separators; comma is decimal.
        processed = clean_txt.str.replace(r"[\s.]", "", regex=True).str.replace(",", ".", regex=False)
    else:
        # Canadian/US: 1,234.56 -> 1234.56.
        # Commas/spaces are thousands separators; dot is decimal.
        processed = clean_txt.str.replace(r"[,\s]", "", regex=True)

    # Robust float extraction: matches -1.5, .5, 5., 500 but not a bare dot.
    extracted = processed.str.extract(r"([-+]?(?:\d+\.?\d*|\.\d+))", expand=False)
    out = pd.to_numeric(extracted, errors="coerce")
    return out.mask(is_negative, -out.abs())


def calculate_mortgage_payment(principal: float, annual_rate: float, years: int, is_io: bool = False) -> float:
    """Return the single source-of-truth monthly mortgage payment.

    Reference implemented here
    --------------------------
    This function intentionally mirrors the compact reference block supplied in
    the latest review note: it is the only scalar monthly-payment routine used
    by self-tests, refinance stress, and explanatory UI copy.

    Formula: M = P * i * (1+i)^n / ((1+i)^n - 1)
    where P = principal, i = monthly rate, n = total monthly payments.

    Engineering note
    ----------------
    If the debt is explicitly interest-only, amortization is zero, or the
    amortizing formula overflows during real-time slider changes, the function
    falls back to IO payment. The vectorized amortization schedule below uses
    the same assumptions for consistency.
    """
    principal = max(0.0, safe_float(principal))
    annual_rate = max(0.0, safe_float(annual_rate))
    if principal <= EPS:
        return 0.0
    monthly_rate = annual_rate / 12.0
    months = int(max(0, safe_float(years, 0)) * 12)
    if is_io or months <= 0:
        return float(principal * monthly_rate)
    if monthly_rate <= EPS:
        return float(principal / months) if months > 0 else 0.0
    try:
        compounding = (1.0 + monthly_rate) ** months
        return float(principal * (monthly_rate * compounding) / (compounding - 1.0))
    except (ZeroDivisionError, OverflowError, FloatingPointError):
        return float(principal * monthly_rate)

def safe_ratio(num: float, den: float) -> float:
    return float(num) / float(den) if float(den) > 0 else 0.0

def get_encryption_key() -> Optional[str]:
    """Read the optional encryption key and warn if it is weakly sized."""
    key = None
    try:
        key = st.secrets.get("ALENZA_DB_ENCRYPTION_KEY")
    except Exception:
        key = None
    if not key:
        key = os.environ.get("ALENZA_DB_ENCRYPTION_KEY")
    if key and len(str(key)) < 44:
        logger.warning("ALENZA_DB_ENCRYPTION_KEY should be a high-entropy 32-byte random secret encoded as base64 or a similarly strong value.")
    return str(key) if key else None

def _cached_master_fernet(secret: str):
    """Return a cached Fernet instance derived once per process/session.

    PBKDF2 is deliberately expensive (600k iterations). Performing it on every
    save makes Streamlit look frozen. New encrypted payloads use one cached
    master Fernet key; legacy v1 payloads with per-message salt are still
    decrypted by make_fernet_legacy().
    """
    if not DEPS["crypto"]:
        raise RuntimeError("Cryptography package missing.")
    cache_key = hashlib.sha256(str(secret).encode()).hexdigest()
    # Prefer session-state cache when Streamlit is active; fall back to process
    # cache for tests or non-Streamlit execution. This keeps the 600k-iteration
    # KDF off the hot save path after first use.
    try:
        session_cache = st.session_state.setdefault("_fernet_cache", {})
        if cache_key in session_cache:
            return session_cache[cache_key]
    except Exception:
        session_cache = None
    if cache_key not in FERNET_CACHE:
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=MASTER_KDF_SALT, iterations=KDF_ITERATIONS)
        key = base64.urlsafe_b64encode(kdf.derive(str(secret).encode()))
        FERNET_CACHE[cache_key] = Fernet(key)
    if session_cache is not None:
        session_cache[cache_key] = FERNET_CACHE[cache_key]
    return FERNET_CACHE[cache_key]

def make_fernet(secret: str, salt: bytes):
    """Legacy compatibility helper for v1 payloads that stored salt+ciphertext."""
    if not DEPS["crypto"]: raise RuntimeError("Cryptography package missing.")
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=KDF_ITERATIONS)
    key = base64.urlsafe_b64encode(kdf.derive(secret.encode()))
    return Fernet(key)

def encrypt_text(text: str, secret: str) -> str:
    token = _cached_master_fernet(secret).encrypt(text.encode())
    return "v2:" + token.decode()

def decrypt_text(token_b64: str, secret: str) -> str:
    if str(token_b64).startswith("v2:"):
        return _cached_master_fernet(secret).decrypt(str(token_b64)[3:].encode()).decode()
    raw = base64.b64decode(token_b64.encode())
    salt, token = raw[:16], raw[16:]
    return make_fernet(secret, salt).decrypt(token).decode()

def encrypt_bytes(data: bytes, secret: str) -> bytes:
    return b"v2:" + _cached_master_fernet(secret).encrypt(data)

def decrypt_bytes(raw: bytes, secret: str) -> bytes:
    if raw.startswith(b"v2:"):
        return _cached_master_fernet(secret).decrypt(raw[3:])
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
def normalize_rr_with_diagnostics(df: pd.DataFrame, number_format: str = "Canadian/US") -> Tuple[pd.DataFrame, List[str], Dict[str, str]]:
    """Normalize a rent roll and return warnings plus the column map.

    The prior version silently dropped unmapped lease fields. This version still
    returns the canonical four columns required by the underwriting engine, but
    it also reports dropped/unmapped columns so the UI can warn the user before
    they rely on incomplete lease information.
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=DEFAULT_RENT_COLS), ["Rent roll is empty."], {}
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    col_map: Dict[str, str] = {}
    warnings: List[str] = []
    # Numeric parser assumes Canadian/US number formatting (1,234.56).
    # European decimal comma values such as 1.234,56 are not silently converted.
    for col in df.columns:
        key = re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip()
        if key in ["tenant", "tenant name", "lessee", "occupant", "company", "name"]:
            col_map[col] = "Tenant"
        elif key in ["sf", "sq ft", "sqft", "square feet", "square footage", "gla", "area", "leased sf", "lease area"]:
            col_map[col] = "SF"
        elif key in ["remaining term", "term remaining", "lease term", "years remaining", "walt", "remaining years"]:
            col_map[col] = "Remaining Term"
        elif key in ["monthly rent", "rent month", "monthly base rent", "rent", "base rent monthly"]:
            col_map[col] = "Monthly Rent"
        elif key in ["annual rent", "annual base rent", "yearly rent", "annualized rent"]:
            col_map[col] = "Annual Rent"

    unmapped = [c for c in df.columns if c not in col_map]
    if unmapped:
        warnings.append("Unmapped rent-roll columns preserved outside the sizing engine: " + ", ".join(unmapped[:8]) + ("..." if len(unmapped) > 8 else ""))

    df = df.rename(columns=col_map)
    if "Annual Rent" in df.columns and "Monthly Rent" not in df.columns:
        df["Monthly Rent"] = clean_numeric_series(df["Annual Rent"], number_format).fillna(0) / 12
        warnings.append("Monthly Rent was derived from Annual Rent / 12.")

    for col in DEFAULT_RENT_COLS:
        if col not in df.columns:
            df[col] = "" if col == "Tenant" else 0.0
            warnings.append(f"Missing required rent-roll column '{col}'; defaulted for sizing.")

    out = df[DEFAULT_RENT_COLS].copy()
    out["Tenant"] = out["Tenant"].fillna("").astype(str).str.strip()

    # Diagnostic watchdog: catch common decimal-comma formats *before* numeric
    # conversion. This is deliberately broad: 1.234,56, 1 234,56 and 1234,56.
    # Reference: uploaded numeric-engine spec asked for whitespace separators and
    # malformed accounting-string coverage.
    numeric_cols = ["SF", "Remaining Term", "Monthly Rent"]
    if str(number_format).lower().startswith("canadian"):
        euro_pattern = r"(?:\d+[ .]\d{3},\d+|\d+,\d+)"
        if any(out[c].astype(str).str.contains(euro_pattern, regex=True, na=False).any() for c in numeric_cols):
            warnings.append("🚨 Formatting alert: decimal-comma values detected (examples: 1 234,56 or 1.234,56). Switch Rent Roll Number Format to European before relying on metrics.")

    # Convert each numeric column through the same vectorized parser and report
    # rows that could not be parsed. We fill NaN with zero only after warning, so
    # the UI can show that a financial metric was affected by bad source data.
    for c in numeric_cols:
        raw = out[c].copy()
        parsed = clean_numeric_series(raw, number_format)
        bad_mask = raw.astype(str).str.strip().ne("") & parsed.isna()
        if bool(bad_mask.any()):
            warnings.append(f"{int(bad_mask.sum())} value(s) in '{c}' could not be parsed and were defaulted to 0 for underwriting metrics.")
        text_with_letters = raw.astype(str).str.contains(r"[A-Za-z]", regex=True, na=False) & parsed.notna()
        if bool(text_with_letters.any()):
            warnings.append(f"{int(text_with_letters.sum())} value(s) in '{c}' contained extra text; the first numeric token was extracted for underwriting metrics.")
        out[c] = parsed.fillna(0).clip(lower=0)

    empty_named = int((out["Tenant"] == "").sum())
    if empty_named:
        warnings.append(f"{empty_named} row(s) have blank tenant names; they remain in SF totals but are treated as vacant/non-income rows.")
    return out.reset_index(drop=True), warnings, col_map


def normalize_rr(df: pd.DataFrame) -> pd.DataFrame:
    out, _, _ = normalize_rr_with_diagnostics(df)
    return out

def generate_comps(property_type: str, noi: float, appraisal: float, seed_text: str) -> pd.DataFrame:
    seed_raw = hashlib.sha256(f"{property_type}-{safe_float(noi)}-{safe_float(appraisal)}-{seed_text}".encode()).hexdigest()
    seed = int(seed_raw[:16], 16) % (2**32)
    rng = np.random.default_rng(seed)
    base_cap = {"Multifamily": 0.045, "Industrial": 0.055, "Retail": 0.065, "Office": 0.075, "Mixed-Use": 0.060, "Hospitality": 0.080, "Self-Storage": 0.058}.get(property_type, 0.06)
    
    rows = []
    for i in range(5):
        cap = max(0.03, base_cap + rng.uniform(-0.006, 0.006))
        comp_noi = safe_float(noi) * rng.uniform(0.70, 1.30) if safe_float(noi) > 0 else (safe_float(appraisal) * cap * rng.uniform(0.75, 1.25) if safe_float(appraisal) > 0 else rng.uniform(350_000, 1_500_000))
        rows.append({
            "Comparable": f"[SIMULATED] {property_type} Comp {i + 1}",
            "Distance (km)": f"{rng.uniform(0.5, 8.0):.1f}",
            # Simulated comp dates are intentionally constrained to late 2025
            # through May 14, 2026 so the demo reflects current-market timing
            # rather than stale historical comps. These remain simulated comps.
            "Sale Date": (datetime(2025, 10, 1) + timedelta(days=int(rng.integers(0, (datetime(2026, 5, 14) - datetime(2025, 10, 1)).days + 1)))).strftime("%Y-%m-%d"),
            "Cap Rate": cap, "NOI": comp_noi, "Value": comp_noi / cap if cap > 0 else 0,
            "lat": 43.6532 + rng.uniform(-0.08, 0.08), "lon": -79.3832 + rng.uniform(-0.08, 0.08),
        })
    return pd.DataFrame(rows)

# =============================================================================
# 7. CACHED FINANCIAL ENGINES
# =============================================================================

if DEPS.get("pydantic"):
    class PydanticDealState(BaseModel):
        """Optional fast schema validation when Pydantic is installed.

        The app still runs without Pydantic, but GitHub / production installs can
        add it for stronger typed coercion and field bounds inside this single file.
        """
        purchase_price: float = Field(default=0.0, ge=0)
        appraisal: float = Field(default=0.0, ge=0)
        noi: float = Field(default=0.0, ge=0)
        rate: float = Field(default=0.055, ge=0, le=0.30)
        amort: int = Field(default=25, ge=1)
        term: int = Field(default=5, ge=1)
        refi_amort: int = Field(default=25, ge=1, le=40)
        fees: float = Field(default=0.015, ge=0, le=0.10)
        closing_costs: float = Field(default=0.0, ge=0)
        reserves: float = Field(default=0.0, ge=0)
        target_ltv: float = Field(default=0.75, ge=0, le=1.25)
        target_ltc: float = Field(default=0.80, ge=0, le=1.50)
        target_dscr: float = Field(default=1.25, ge=0, le=5.0)
        target_dy: float = Field(default=0.08, ge=0, le=0.30)
        mezz_debt: float = Field(default=0.0, ge=0)
        pref_equity: float = Field(default=0.0, ge=0)
        mezz_rate: float = Field(default=0.11, ge=0, le=0.50)
        pref_rate: float = Field(default=0.09, ge=0, le=0.50)
        pf_rev_growth: float = Field(default=0.03, ge=-0.10, le=0.25)
        pf_exp_growth: float = Field(default=0.02, ge=-0.10, le=0.25)
        pf_exp_ratio: float = Field(default=0.40, ge=0.01, le=0.95)
        pf_exit_cap: float = Field(default=0.06, ge=0.02, le=0.20)
        pf_sell_costs: float = Field(default=0.015, ge=0.0, le=0.20)
        pf_term_growth: float = Field(default=0.02, ge=-0.05, le=0.10)
        mc_sims: int = Field(default=2000, ge=100, le=25000)
        mc_noi_vol: float = Field(default=0.075, ge=0.0, le=0.50)
        mc_rate_vol_bps: float = Field(default=75.0, ge=0.0, le=500.0)
        mc_exit_cap_vol_bps: float = Field(default=50.0, ge=0.0, le=500.0)

        class Config:
            extra = "ignore"


class DealStateValidator:
    """Single-file schema layer with optional Pydantic acceleration.

    When Pydantic is installed, the validator uses a BaseModel for typed coercion
    and field bounds. When it is not installed, the manual fallback keeps the app
    runnable on minimal Streamlit deployments.
    """

    NUMERIC_DEFAULTS = {
        "purchase_price": 0.0, "appraisal": 0.0, "noi": 0.0, "rate": 0.055,
        "amort": 25, "term": 5, "refi_amort": 25, "fees": 0.015, "closing_costs": 0.0, "reserves": 0.0,
        "target_ltv": 0.75, "target_ltc": 0.80, "target_dscr": 1.25, "target_dy": 0.08,
        "mezz_debt": 0.0, "pref_equity": 0.0, "mezz_rate": 0.11, "pref_rate": 0.09,
        "pf_rev_growth": 0.03, "pf_exp_growth": 0.02, "pf_exp_ratio": 0.40,
        "pf_exit_cap": 0.06, "pf_sell_costs": 0.015, "pf_term_growth": 0.02,
        "mc_sims": 2000, "mc_noi_vol": 0.075, "mc_rate_vol_bps": 75, "mc_exit_cap_vol_bps": 50,
    }

    RATIO_BOUNDS = {
        "rate": (0.0, 0.30), "fees": (0.0, 0.10), "target_ltv": (0.0, 1.25),
        "target_ltc": (0.0, 1.50), "target_dscr": (0.0, 5.0), "target_dy": (0.0, 0.30),
        "mezz_rate": (0.0, 0.50), "pref_rate": (0.0, 0.50),
        "pf_rev_growth": (-0.10, 0.25), "pf_exp_growth": (-0.10, 0.25),
        "pf_exp_ratio": (0.01, 0.95), "pf_exit_cap": (0.02, 0.20),
        "pf_sell_costs": (0.0, 0.20), "pf_term_growth": (-0.05, 0.10),
        "mc_noi_vol": (0.0, 0.50), "mc_rate_vol_bps": (0.0, 500.0), "mc_exit_cap_vol_bps": (0.0, 500.0),
    }

    @classmethod
    def coerce(cls, state: dict) -> Tuple[dict, List[str]]:
        clean = dict(state)
        notes: List[str] = []
        if DEPS.get("pydantic") and "PydanticDealState" in globals():
            try:
                model = PydanticDealState(**{k: clean.get(k) for k in cls.NUMERIC_DEFAULTS})
                values = model.model_dump() if hasattr(model, "model_dump") else model.dict()
                clean.update(values)
                return clean, notes
            except Exception as e:
                notes.append(f"Pydantic validation fallback used: {e}")
        for key, default in cls.NUMERIC_DEFAULTS.items():
            try:
                val = parse_financial_float(clean.get(key), key, default)
            except NumericParseError as e:
                # Finance inputs must fail loudly. We do not silently replace a
                # malformed amount or rate with zero because that can misprice a deal.
                raise UnderwritingError(f"Input validation failed for {key}: {e}") from e
            if key in ["amort", "term", "mc_sims"]:
                val = int(max(1, round(val)))
            if key in cls.RATIO_BOUNDS:
                lo, hi = cls.RATIO_BOUNDS[key]
                clipped = min(max(val, lo), hi)
                if clipped != val:
                    notes.append(f"{key} clipped from {val} to {clipped}")
                val = clipped
            clean[key] = val
        clean["mc_sims"] = int(min(max(clean.get("mc_sims", 2000), 100), 25000))
        return clean, notes

class UnderwritingEngine:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def size_loan(noi: float, appraisal: float, purchase_price: float, closing_costs: float, reserves: float, fees_pct: float, rate: float, amort: int, term: int, is_io: bool, target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float) -> Tuple[float, str, dict, float, float]:
        """Closed-form senior mortgage sizing.

        Mathematical assumptions
        ------------------------
        * Income-property proceeds require positive forward 12-month NOI.
        * LTV uses appraisal when available. If appraisal is blank but purchase
          price is positive, purchase price is used as value with an explicit
          diagnostic note so the UI can warn the user.
        * LTC includes circular financing fees using the closed-form solution:
          L = LTC * hard_costs / (1 - LTC * fee_pct). No fixed-point loop.
        * DSCR is sized from maximum permitted annual debt service. For IO,
          debt service is interest-only. For amortizing debt, the formula uses
          the inverse of the standard mortgage payment factor.
        * Debt yield is NOI / loan.

        Returns
        -------
        (loan, binding_gate, gates, total_uses, required_equity)
        where gates contains numeric active gates plus diagnostic metadata.
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
        collateral_value = appraisal if appraisal > EPS else purchase_price
        gates: Dict[str, Any] = {}
        skipped: List[str] = []
        messages: List[str] = []

        if noi <= EPS:
            raise UnderwritingError("NOI is required for income-property mortgage sizing; deterministic mortgage proceeds were not calculated.")
        if appraisal <= EPS and purchase_price > EPS:
            messages.append("Appraisal missing; using purchase price as value for LTV. Confirm this is acceptable before credit submission.")
        elif appraisal <= EPS:
            raise UnderwritingError("Appraisal or purchase price is required for LTV sizing.")

        if collateral_value > 0 and target_ltv > 0:
            gates["LTV"] = collateral_value * target_ltv
        else:
            skipped.append("LTV skipped: appraisal or target LTV missing")

        if hard_costs > 0 and target_ltc > 0:
            denom = 1.0 - target_ltc * fees_pct
            if denom > EPS:
                gates["LTC"] = (target_ltc * hard_costs) / denom
            else:
                raise UnderwritingError("Invalid LTC/fee combination: target_ltc * fees_pct makes the fee-adjusted denominator non-positive.")
        else:
            skipped.append("LTC skipped: purchase price/uses or target LTC missing")

        if target_dy > 0:
            gates["Debt Yield"] = noi / target_dy
        else:
            skipped.append("Debt Yield skipped: target debt yield is zero")

        if target_dscr > 0:
            max_annual_ds = noi / target_dscr
            if is_io:
                if rate > EPS:
                    gates["DSCR"] = max_annual_ds / rate
                else:
                    skipped.append("DSCR skipped: IO loan requires positive interest rate")
            else:
                months = amort * 12
                max_monthly_ds = max_annual_ds / 12.0
                if rate <= EPS:
                    gates["DSCR"] = max_monthly_ds * months
                else:
                    m_rate = rate / 12.0
                    payment_factor = m_rate / (1.0 - (1.0 + m_rate) ** (-months))
                    gates["DSCR"] = max_monthly_ds / payment_factor
        else:
            skipped.append("DSCR skipped: target DSCR is zero")

        numeric_gates = {
            k: float(v) for k, v in gates.items()
            if isinstance(v, (int, float, np.number)) and np.isfinite(v) and float(v) >= 0
        }
        if not numeric_gates:
            raise UnderwritingError("No active underwriting gates were available: " + "; ".join(skipped))

        loan = max(0.0, min(numeric_gates.values()))
        total_uses = hard_costs + loan * fees_pct
        gates.update({
            "_messages": messages,
            "_used_gates": sorted(numeric_gates.keys()),
            "_skipped_gates": skipped,
            "_method": "Closed-form sizing: min(LTV, fee-adjusted LTC, DSCR, Debt Yield)",
        })
        binding_gate = min(numeric_gates, key=numeric_gates.get)
        return round(loan, 2), binding_gate, gates, round(total_uses, 2), round(total_uses - loan, 2)

    @staticmethod
    @st.cache_data(show_spinner=False)
    def amort_schedule(loan_amt: float, rate: float, amort_yrs: int, term_yrs: int, is_io: bool) -> Tuple[pd.DataFrame, float, float]:
        """Vectorized amortization schedule for the displayed loan term.

        Complexity is O(T) in NumPy array operations over months, rather than a
        Python loop that builds one dictionary per period.
        """
        loan_amt = max(0.0, safe_float(loan_amt))
        rate = max(0.0, safe_float(rate))
        amort_yrs = max(1, int(safe_float(amort_yrs, 1)))
        term_yrs = max(1, int(safe_float(term_yrs, 1)))
        if loan_amt <= EPS:
            return pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]), 0.0, 0.0
        months_amort = amort_yrs * 12
        months_term = min(term_yrs * 12, months_amort)
        periods = np.arange(1, months_term + 1, dtype=float)
        m_rate = rate / 12.0
        # This scalar payment is deliberately sourced from calculate_mortgage_payment()
        # so amortization, refinance stress, and QA checks cannot drift apart.
        m_pmt = calculate_mortgage_payment(loan_amt, rate, amort_yrs, is_io)
        if is_io:
            interest = np.full_like(periods, m_pmt, dtype=float)
            principal = np.zeros_like(periods, dtype=float)
            balance = np.full_like(periods, loan_amt, dtype=float)
        elif m_rate <= EPS:
            interest = np.zeros_like(periods, dtype=float)
            principal = np.full_like(periods, m_pmt, dtype=float)
            balance = np.maximum(0.0, loan_amt - m_pmt * periods)
        else:
            prev_periods = periods - 1.0
            prev_balance = loan_amt * np.power(1.0 + m_rate, prev_periods) - m_pmt * ((np.power(1.0 + m_rate, prev_periods) - 1.0) / m_rate)
            interest = prev_balance * m_rate
            principal = np.minimum(np.maximum(m_pmt - interest, 0.0), prev_balance)
            balance = np.maximum(0.0, prev_balance - principal)
        df = pd.DataFrame({
            "Period": periods.astype(int),
            "Payment": np.round(np.full_like(periods, m_pmt, dtype=float), 2),
            "Principal": np.round(principal, 2),
            "Interest": np.round(interest, 2),
            "Balance": np.round(balance, 2),
        })
        return df, round(float(m_pmt), 2), round(float(balance[-1]) if len(balance) else 0.0, 2)

    @staticmethod
    @st.cache_data(show_spinner=False)
    def run_refi_stress(noi: float, balloon_bal: float, refi_amort: int, base_rate: float) -> pd.DataFrame:
        """Assess refinance viability across rate-expansion scenarios.

        Reference / rationale
        ---------------------
        This integrates the uploaded reference block's refinance stress test into
        the institutional app. It answers a different question than initial
        proceeds sizing: *can the balloon be refinanced at maturity if rates are
        wider by 100-500 bps?*

        The payment source is calculate_mortgage_payment(), so refi DSCR uses
        the same mortgage formula as sizing and amortization.
        """
        noi = max(0.0, safe_float(noi))
        balloon_bal = max(0.0, safe_float(balloon_bal))
        refi_amort = max(1, int(safe_float(refi_amort, 25)))
        base_rate = max(0.0, safe_float(base_rate))
        scenarios = np.array([0, 100, 200, 300, 400, 500], dtype=float)
        rates = base_rate + scenarios / 10000.0
        monthly_ds = np.array([calculate_mortgage_payment(balloon_bal, r, refi_amort, False) for r in rates], dtype=float)
        annual_ds = monthly_ds * 12.0
        dscr = np.divide(noi, annual_ds, out=np.zeros_like(annual_ds), where=annual_ds > EPS)
        return pd.DataFrame({
            "Spread": [f"+{int(bps)} bps" for bps in scenarios],
            "Rate": rates,
            "Refi DS": annual_ds,
            "Refi DSCR": dscr,
        })

    @staticmethod
    def capital_stack(senior_debt: float, mezz_debt: float, pref_equity: float, sponsor_equity: float, noi: float, senior_rate: float, mezz_rate: float, pref_rate: float) -> dict:
        senior_debt, mezz_debt, pref_equity, sponsor_equity = max(0.0, safe_float(senior_debt)), max(0.0, safe_float(mezz_debt)), max(0.0, safe_float(pref_equity)), safe_float(sponsor_equity)
        senior_cost, mezz_cost, pref_cost = senior_debt * safe_float(senior_rate), mezz_debt * safe_float(mezz_rate), pref_equity * safe_float(pref_rate)
        fc = senior_cost + mezz_cost + pref_cost
        tot = senior_debt + mezz_debt + pref_equity + sponsor_equity
        return {"Senior": senior_debt, "Mezz": mezz_debt, "Pref": pref_equity, "Sponsor": sponsor_equity, "Total": tot, "FixedCharges": fc, "FCC": safe_ratio(max(0, safe_float(noi)), fc)}

    @staticmethod
    def rent_roll_metrics_with_diagnostics(df: pd.DataFrame) -> Tuple[Tuple[float, float, float, float, float, float], List[str]]:
        df, warnings, _ = normalize_rr_with_diagnostics(df)
        if df.empty:
            return (0, 0, 0, 0, 0, 0), warnings + ["Rent roll is empty; occupancy and WALT are not meaningful."]
        total_sf = df["SF"].sum()
        if total_sf <= 0:
            return (0, 0, 0, 0, 0, 0), warnings + ["Rent roll has no positive SF."]
        occ = df[(df["SF"] > 0) & (~df["Tenant"].str.lower().isin(["", "vacant", "available", "empty"]))]
        occ_sf = occ["SF"].sum()
        if occ_sf <= 0:
            return (total_sf, 0, 0, 0, 0, 0), warnings + ["Rent roll has no occupied/income rows after normalization."]
        ar = occ["Monthly Rent"].sum() * 12
        walt = (occ["Remaining Term"] * occ["SF"]).sum() / occ_sf
        exp1 = occ[occ["Remaining Term"] <= 1.0]["SF"].sum() / occ_sf
        return (total_sf, occ_sf/total_sf, ar, ar/occ_sf if occ_sf > 0 else 0, walt, exp1), warnings

    @staticmethod
    def rent_roll_metrics(df: pd.DataFrame) -> Tuple[float, float, float, float, float, float]:
        metrics, _ = UnderwritingEngine.rent_roll_metrics_with_diagnostics(df)
        return metrics

    @staticmethod
    def breakeven_occupancy(noi: float, occ: float, ann_ds: float) -> float:
        noi, occ, ann_ds = safe_float(noi), safe_float(occ), safe_float(ann_ds)
        if noi <= 0 or occ <= 0 or ann_ds <= 0: return 0.0
        return min(1.5, max(0.0, ann_ds / noi * occ))

class InvestmentEngine:
    @staticmethod
    @st.cache_data(show_spinner=False)
    def calculate_pro_forma(noi: float, r_grow: float, e_grow: float, e_ratio: float, yrs: int = 10) -> pd.DataFrame:
        """Build a simple annual pro-forma from base NOI.

        Tiny or negative NOI is treated as infeasible/zero instead of letting
        floating-point noise create meaningless margins or IRRs.
        """
        noi = max(0.0, safe_float(noi))
        r_grow = safe_float(r_grow)
        e_grow = safe_float(e_grow)
        e_ratio = min(max(safe_float(e_ratio), 0.0), 0.95)
        yrs = max(1, int(safe_float(yrs, 10)))
        if noi <= EPS:
            return pd.DataFrame([{
                "Year": y, "Revenue": 0.0, "Expenses": 0.0,
                "Projected NOI": 0.0, "NOI Margin": 0.0
            } for y in range(1, yrs + 1)])
        rev = noi / max(1 - e_ratio, 0.01)
        exp = rev * e_ratio
        rows = []
        for y in range(1, yrs + 1):
            rows.append({"Year": y, "Revenue": rev, "Expenses": exp, "Projected NOI": rev - exp, "NOI Margin": safe_ratio(rev-exp, rev)})
            rev *= (1 + r_grow); exp *= (1 + e_grow)
        return pd.DataFrame(rows)

    @staticmethod
    def solve_returns(pp: float, loan: float, pf_df: pd.DataFrame, cap: float, sell_costs: float, ann_ds: float, balloon: float, t_growth: float) -> dict:
        eq = max(0.0, pp - loan)
        if eq <= 1e-6 or pf_df.empty:
            return {"IRR": 0.0, "EM": 0.0, "Exit NOI": 0.0, "Gross Exit": 0.0, "Net Exit": 0.0, "Total CF": 0.0}
        
        cfs = [-eq]
        for _, r in pf_df.iterrows(): cfs.append(max(0.0, r["Projected NOI"]) - ann_ds)
        exit_noi = pf_df.iloc[-1]["Projected NOI"] * (1 + t_growth)
        gx = exit_noi / cap if cap > 0 else 0.0
        nx = gx * (1 - sell_costs) - balloon
        cfs[-1] += nx
        
        irr = 0.0
        if DEPS["numpy_financial"]:
            try:
                irr = npf.irr(cfs) or 0.0
            except ValueError:
                # IRR can fail on non-conventional cash flows with multiple sign changes
                # or near-total equity loss. For deterministic UI display, default to 0.0
                # and rely on Monte Carlo downside bands for risk interpretation.
                pass
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
                shocked_noi = noi_base * (1 + n)
                shocked_rate = max(0.001, rate_base + r)
                try:
                    L, _, _, _, _ = UnderwritingEngine.size_loan(
                        shocked_noi, appraisal, pp, costs, res, fees, shocked_rate,
                        amort, term, is_io, ltv, ltc, dscr, dy
                    )
                except UnderwritingError:
                    L = np.nan
                row[f"{r*100:+.1f}%"] = L
            data.append(row)
        return pd.DataFrame(data)


class RiskAnalyticsEngine:
    """Vectorized Monte Carlo downside analytics for underwriting risk.

    Stochastic process notes
    ------------------------
    * NOI follows a one-period lognormal shock. This is a simplified geometric
      Brownian motion assumption that keeps NOI non-negative.
    * Interest rates and exit cap rates follow clipped normal shocks around the
      base case.
    * Shock arrays are generated once and all underwriting gates are evaluated
      with NumPy array operations. No scenario-level Python loop is used.
    * The stable seed uses fixed-precision numeric formatting so the same deal
      inputs produce the same random draws across platforms.
    """

    @staticmethod
    def _stable_seed(seed_text: str, noi: float, purchase_price: float, rate: float) -> int:
        payload = f"mc|{seed_text}|{noi:.10f}|{purchase_price:.10f}|{rate:.10f}"
        return int(hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16], 16) % (2**32)

    @staticmethod
    def _vectorized_loan_size(
        noi_s: np.ndarray, appraisal: float, purchase_price: float, closing_costs: float, reserves: float,
        fees_pct: float, rate_s: np.ndarray, amort: int, is_io: bool,
        target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float,
    ) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
        """Vectorized closed-form loan sizing for Monte Carlo scenarios."""
        sims = len(noi_s)
        hard_costs = max(0.0, purchase_price + closing_costs + reserves)
        collateral = appraisal if appraisal > EPS else purchase_price
        inf = np.full(sims, np.inf, dtype=float)

        ltv_gate = np.full(sims, collateral * target_ltv, dtype=float) if collateral > EPS and target_ltv > 0 else np.zeros(sims, dtype=float)
        invalid_ltc_mask = np.zeros(sims, dtype=bool)
        if hard_costs > 0 and target_ltc > 0:
            ltc_denom = 1.0 - target_ltc * fees_pct
            if ltc_denom > EPS:
                ltc_gate = np.full(sims, (target_ltc * hard_costs) / ltc_denom, dtype=float)
            else:
                # Conservative hard stop: do not silently ignore an invalid LTC formula.
                ltc_gate = np.zeros(sims, dtype=float)
                invalid_ltc_mask[:] = True
        else:
            ltc_gate = inf.copy()
        dy_gate = np.where(target_dy > 0, noi_s / target_dy, np.inf)

        if target_dscr > 0:
            max_annual_ds = noi_s / target_dscr
            if is_io:
                dscr_gate = np.where(rate_s > EPS, max_annual_ds / rate_s, np.inf)
            else:
                months = max(1, int(amort)) * 12
                m_rate = rate_s / 12.0
                payment_factor = np.where(
                    m_rate > EPS,
                    m_rate / (1.0 - np.power(1.0 + m_rate, -months)),
                    1.0 / months,
                )
                dscr_gate = (max_annual_ds / 12.0) / payment_factor
        else:
            dscr_gate = inf.copy()

        gate_matrix = np.vstack([ltv_gate, ltc_gate, dy_gate, dscr_gate])
        loan = np.nanmin(gate_matrix, axis=0)
        loan = np.where(np.isfinite(loan) & (noi_s > EPS), np.maximum(loan, 0.0), 0.0)
        gate_names = np.array(["LTV", "LTC", "Debt Yield", "DSCR"])
        binding_idx = np.nanargmin(gate_matrix, axis=0)
        binding = np.where(loan > 0, gate_names[binding_idx], "Missing/Zero")
        uses = hard_costs + loan * fees_pct
        return loan, uses, binding, invalid_ltc_mask

    @staticmethod
    def _annual_debt_service_vectorized(loan: np.ndarray, rate_s: np.ndarray, amort: int, is_io: bool) -> Tuple[np.ndarray, np.ndarray]:
        months = max(1, int(amort)) * 12
        if is_io:
            monthly = loan * rate_s / 12.0
            balloon = loan.copy()
        else:
            m_rate = rate_s / 12.0
            payment_factor = np.where(
                m_rate > EPS,
                m_rate / (1.0 - np.power(1.0 + m_rate, -months)),
                1.0 / months,
            )
            monthly = loan * payment_factor
            term_months = max(1, int(amort)) * 12
            # Term is handled outside by passing actual term to closed form in the main function.
            balloon = loan.copy()
        return monthly * 12.0, balloon

    @staticmethod
    def _irr_newton_vectorized(equity: np.ndarray, annual_cf: np.ndarray, max_iter: int = 60) -> Tuple[np.ndarray, np.ndarray]:
        """Vectorized Newton IRR for conventional scenario cash flows.

        This is not a nearest-neighbour grid. It solves NPV(rate)=0 for every
        scenario using NumPy array operations, and returns a convergence mask so
        downstream displays can disclose how many scenarios failed to converge.
        """
        years = np.arange(1, annual_cf.shape[1] + 1, dtype=float)
        r = np.full(equity.shape, 0.10, dtype=float)
        valid = equity > EPS
        converged = np.zeros_like(valid, dtype=bool)
        for _ in range(max_iter):
            base = np.clip(1.0 + r, 0.05, 2.50)
            disc = np.power(base[:, None], years[None, :])
            npv = -equity + np.sum(annual_cf / disc, axis=1)
            deriv = -np.sum(years[None, :] * annual_cf / np.power(base[:, None], years[None, :] + 1.0), axis=1)
            step = np.divide(npv, deriv, out=np.zeros_like(npv), where=np.abs(deriv) > EPS)
            r_next = np.clip(r - step, -0.95, 1.00)
            rel_tol = 1e-6 * (np.abs(equity) + 1.0)
            step_tol = np.abs(step) < 1e-8
            converged_now = (np.abs(npv) <= rel_tol) | step_tol
            converged |= converged_now & valid
            r = np.where(valid & ~converged_now, r_next, r)
        raw_irr = np.where(valid & np.isfinite(r) & converged, r, np.nan)
        extreme_loss = raw_irr <= -0.949999
        irr_valid = converged & valid & np.isfinite(raw_irr) & ~extreme_loss
        irr = np.where(irr_valid, raw_irr, np.nan)
        return irr, irr_valid

    @staticmethod
    @st.cache_data(show_spinner=False)
    def monte_carlo_deal(
        noi: float, appraisal: float, purchase_price: float, closing_costs: float, reserves: float,
        fees_pct: float, rate: float, amort: int, term: int, is_io: bool,
        target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float,
        pf_rev_growth: float, pf_exp_growth: float, pf_exp_ratio: float, pf_exit_cap: float,
        pf_sell_costs: float, pf_term_growth: float, sims: int, noi_vol: float,
        rate_vol_bps: float, exit_cap_vol_bps: float, seed_text: str,
    ) -> Tuple[pd.DataFrame, dict]:
        sims = int(min(max(safe_float(sims, 2000), 100), 25000))
        noi = max(0.0, safe_float(noi))
        appraisal = max(0.0, safe_float(appraisal))
        purchase_price = max(0.0, safe_float(purchase_price))
        closing_costs = max(0.0, safe_float(closing_costs))
        reserves = max(0.0, safe_float(reserves))
        fees_pct = max(0.0, safe_float(fees_pct))
        rate = max(0.0, safe_float(rate))
        amort = max(1, int(safe_float(amort, 1)))
        term = max(1, int(safe_float(term, 1)))

        if noi <= EPS or purchase_price <= EPS:
            empty = pd.DataFrame(columns=["Mortgage", "DSCR", "IRR", "Exit Cap", "Rate", "NOI"])
            return empty, {"Message": "Monte Carlo requires positive NOI and purchase price.", "Simulations": sims}
        appraisal_proxy_used = bool(appraisal <= EPS and purchase_price > EPS)

        rng = np.random.default_rng(RiskAnalyticsEngine._stable_seed(seed_text, noi, purchase_price, rate))
        noi_vol = min(max(safe_float(noi_vol, 0.075), 0.0), 0.50)
        noi_s = noi * rng.lognormal(mean=-0.5 * noi_vol**2, sigma=noi_vol, size=sims)
        rate_s = np.clip(rate + rng.normal(0.0, safe_float(rate_vol_bps, 75) / 10000.0, sims), 0.001, 0.30)
        exit_cap_s = np.clip(pf_exit_cap + rng.normal(0.0, safe_float(exit_cap_vol_bps, 50) / 10000.0, sims), 0.02, 0.20)

        loan, uses, binding, invalid_ltc_mask = RiskAnalyticsEngine._vectorized_loan_size(
            noi_s, appraisal, purchase_price, closing_costs, reserves, fees_pct, rate_s,
            amort, is_io, target_ltv, target_ltc, target_dscr, target_dy
        )

        months_amort = amort * 12
        months_term = term * 12
        m_rate = rate_s / 12.0
        if is_io:
            monthly_ds = loan * m_rate
            balloon = loan
        else:
            payment_factor = np.where(
                m_rate > EPS,
                m_rate / (1.0 - np.power(1.0 + m_rate, -months_amort)),
                1.0 / months_amort,
            )
            monthly_ds = loan * payment_factor
            # Closed-form remaining balance after term_months payments.
            bal_amort = np.where(
                m_rate > EPS,
                loan * np.power(1.0 + m_rate, months_term) - monthly_ds * ((np.power(1.0 + m_rate, months_term) - 1.0) / m_rate),
                np.maximum(0.0, loan - monthly_ds * months_term),
            )
            balloon = np.maximum(0.0, bal_amort)
        annual_ds = monthly_ds * 12.0

        years = np.arange(1, 11, dtype=float)
        e_ratio = min(max(safe_float(pf_exp_ratio), 0.0), 0.95)
        rev0 = noi_s / max(1.0 - e_ratio, 0.01)
        exp0 = rev0 * e_ratio
        rev_y = rev0[:, None] * np.power(1.0 + safe_float(pf_rev_growth), years[None, :] - 1.0)
        exp_y = exp0[:, None] * np.power(1.0 + safe_float(pf_exp_growth), years[None, :] - 1.0)
        noi_y = rev_y - exp_y
        annual_cf = noi_y - annual_ds[:, None]
        exit_noi = noi_y[:, -1] * (1.0 + safe_float(pf_term_growth))
        gross_exit = exit_noi / exit_cap_s
        net_exit = gross_exit * (1.0 - safe_float(pf_sell_costs)) - balloon
        annual_cf[:, -1] += net_exit

        equity = np.maximum(0.0, uses - loan)
        total_cf = annual_cf.sum(axis=1)
        em = np.where(equity > EPS, total_cf / equity, 0.0)

        # Vectorized Newton IRR. Scenarios that fail convergence or imply an
        # extreme-loss IRR below -95% are marked NaN and disclosed, not forced to 0%.
        irr, irr_valid = RiskAnalyticsEngine._irr_newton_vectorized(equity, annual_cf)
        no_equity_or_bad = (equity <= EPS) | ~np.isfinite(total_cf) | ~irr_valid
        irr = np.where(no_equity_or_bad, np.nan, irr)

        dscr = np.divide(noi_s, annual_ds, out=np.zeros_like(noi_s), where=annual_ds > EPS)
        dy = np.divide(noi_s, loan, out=np.zeros_like(noi_s), where=loan > EPS)

        df = pd.DataFrame({
            "Mortgage": loan,
            "Binding Gate": binding,
            "DSCR": dscr,
            "Debt Yield": dy,
            "Yield on Cost": np.divide(noi_s, uses, out=np.zeros_like(noi_s), where=uses > EPS),
            "IRR": irr,
            "Equity Multiple": em,
            "Exit Cap": exit_cap_s,
            "Rate": rate_s,
            "NOI": noi_s,
        })
        hard_costs_summary = max(0.0, purchase_price + closing_costs + reserves)
        invalid_ltc = bool(hard_costs_summary > 0 and target_ltc > 0 and (1.0 - target_ltc * fees_pct) <= EPS)
        summary = {
            "Simulations": sims,
            "Invalid LTC/Fee Combination": invalid_ltc,
            "Invalid LTC/Fee Scenarios": int(invalid_ltc_mask.sum()),
            "IRR Invalid/Undefined Probability": float((~irr_valid).mean()),
            "Mortgage P5": float(df["Mortgage"].quantile(0.05)),
            "Mortgage P50": float(df["Mortgage"].quantile(0.50)),
            "Mortgage P95": float(df["Mortgage"].quantile(0.95)),
            "DSCR P5": float(df["DSCR"].quantile(0.05)),
            "DSCR VaR 95%": float(df["DSCR"].quantile(0.05)),
            "Survival Rate": float((df["DSCR"] >= 1.00).mean()),
            "Success Rate": float((df["DSCR"] >= safe_float(target_dscr, 1.25)).mean()),
            "Risk of Default": float((df["DSCR"] < 1.00).mean()),
            "Avg Yield on Cost": float(df["Yield on Cost"].mean()),
            "IRR P5": float(df["IRR"].quantile(0.05)),
            "IRR P50": float(df["IRR"].quantile(0.50)),
            "Probability DSCR < 1.20x": float((df["DSCR"] < 1.20).mean()),
            "Probability Negative IRR": float((df["IRR"].dropna() < 0.0).mean()) if df["IRR"].notna().any() else 0.0,
            "Probability Zero Mortgage": float((df["Mortgage"] <= 0.0).mean()),
            "IRR Valid Rate": float(np.mean(irr_valid)),
            "IRR Method": "Vectorized Newton solve with convergence disclosure",
            "IRR Percentile Basis": "IRR percentiles exclude undefined/extreme-loss scenarios; review IRR Valid Rate.",
            "Appraisal Proxy Used": bool(appraisal_proxy_used),
        }
        return df, summary


@st.cache_data(show_spinner="Simulating market backtest...")
def run_monte_carlo_backtest(state: Dict[str, Any], iterations: int = 100) -> pd.DataFrame:
    """Convenience wrapper for a 100-point/quick Monte Carlo backtest.

    This function delegates to the vectorized RiskAnalyticsEngine rather than
    looping through scenarios in Python. It exists to make the GitHub code easier
    to read: the UI can request a compact backtest while the mathematical engine
    remains the same source of truth as the full Monte Carlo panel.

    References in-code:
    * Four-gate loan sizing uses UnderwritingEngine.size_loan.
    * Debt service uses calculate_mortgage_payment.
    * DSCR VaR is the empirical 5th percentile of DSCR observations.
    """
    s, _notes = DealStateValidator.coerce(dict(state))
    df, _summary = RiskAnalyticsEngine.monte_carlo_deal(
        s["noi"], s["appraisal"], s["purchase_price"], s["closing_costs"], s["reserves"], s["fees"], s["rate"],
        s["amort"], s["term"], s["is_io"], s["target_ltv"], s["target_ltc"], s["target_dscr"], s["target_dy"],
        s["pf_rev_growth"], s["pf_exp_growth"], s["pf_exp_ratio"], s["pf_exit_cap"], s["pf_sell_costs"],
        s["pf_term_growth"], int(iterations), s["mc_noi_vol"], s["mc_rate_vol_bps"], s["mc_exit_cap_vol_bps"],
        str(s.get("deal_id", "quick_backtest")),
    )
    if not df.empty:
        df = df.rename(columns={"Mortgage": "loan_amount", "Debt Yield": "debt_yield", "Yield on Cost": "yield_on_cost", "Exit Cap": "exit_cap", "Rate": "rate", "NOI": "noi", "DSCR": "dscr"})
        df.insert(0, "iteration", np.arange(1, len(df) + 1))
        df["is_viable"] = df["dscr"] >= float(s.get("target_dscr", 1.25))
    return df


def get_simulation_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """Summarize Monte Carlo results into lender-facing risk metrics."""
    if df is None or df.empty:
        return {"Success Rate": "0.0%", "Mean DSCR": 0.0, "VaR (Value at Risk) 95%": 0.0, "Risk of Default": "0.0%", "Avg Yield on Cost": 0.0}
    dscr_col = "dscr" if "dscr" in df.columns else "DSCR"
    yoc_col = "yield_on_cost" if "yield_on_cost" in df.columns else "Yield on Cost"
    valid = df[dscr_col].replace([np.inf, -np.inf], np.nan).dropna()
    success_rate = float((df.get("is_viable", df[dscr_col] >= 1.20)).mean()) if len(df) else 0.0
    return {
        "Mean DSCR": float(valid.mean()) if not valid.empty else 0.0,
        "VaR (Value at Risk) 95%": float(valid.quantile(0.05)) if not valid.empty else 0.0,
        "Success Rate": f"{success_rate * 100:.1f}%",
        "Avg Yield on Cost": float(df[yoc_col].mean()) if yoc_col in df.columns else 0.0,
        "Risk of Default": f"{float((df[dscr_col] < 1.0).mean()) * 100:.1f}%" if dscr_col in df.columns else "0.0%",
    }

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
    def run_financial_self_test_bools() -> Dict[str, bool]:
        """Raw boolean startup checks for the financial kernel.

        This function is intentionally separate from the DataFrame QA report.
        It uses explicit golden constants for the payment engine and a direct
        closed-form sizing smoke test so startup cannot be fooled by display
        strings or DataFrame truthiness.
        """
        checks: Dict[str, bool] = {}
        try:
            checks["IO payment golden value"] = abs(calculate_mortgage_payment(1_000_000, 0.06, 25, True) - 5000.0) < 0.01
            checks["Amortizing payment golden value"] = abs(calculate_mortgage_payment(1_000_000, 0.06, 25, False) - 6443.01) < 0.10
            checks["Zero-rate boundary"] = abs(calculate_mortgage_payment(120_000, 0.0, 10, False) - 1000.0) < 0.01
            L, gate, gates, _, _ = UnderwritingEngine.size_loan(1_000_000, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
            numeric_gates = {k: v for k, v in gates.items() if not str(k).startswith("_") and isinstance(v, (int, float, np.number))}
            checks["Closed-form loan positive"] = L > 0
            checks["Binding gate equals min numeric gate"] = bool(numeric_gates) and abs(L - min(numeric_gates.values())) < 1.0
            try:
                UnderwritingEngine.size_loan(0, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
                checks["Zero NOI blocks sizing"] = False
            except UnderwritingError:
                checks["Zero NOI blocks sizing"] = True
        except Exception as exc:
            logger.exception("Startup financial self-test failed: %s", exc)
            checks["Self-test exception"] = False
        return checks

    @staticmethod
    def run_financial_self_tests() -> pd.DataFrame:
        rows = []
        def check(name, passed, detail): rows.append({"Test": name, "Status": "✅ PASS" if passed else "❌ FAIL", "Detail": detail})

        try:
            # Startup reference tests: these mirror the compact reference block
            # provided in the latest note. They verify the scalar payment source
            # before any larger underwriting logic is trusted.
            check("Unified IO payment source", abs(calculate_mortgage_payment(1_000_000, 0.06, 25, True) - 5000.0) < 0.01, "$5,000 expected")
            check("Unified amortizing payment source", abs(calculate_mortgage_payment(1_000_000, 0.06, 25, False) - 6443.01) < 0.10, "$6,443.01 expected")
            check("Unified zero-rate boundary", abs(calculate_mortgage_payment(120_000, 0.0, 10, False) - 1000.0) < 0.01, "$1,000 expected")

            L, gate, gates, uses, eq = UnderwritingEngine.size_loan(1_000_000, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
            check("Loan proceeds positive", L > 0, f"${L:,.0f}")
            numeric_gates = {k: v for k, v in gates.items() if not str(k).startswith("_") and isinstance(v, (int, float, np.number))}
            check("Binding gate exact match", abs(L - min(numeric_gates.values())) < 1, gate)

            amort, pmt, balloon = UnderwritingEngine.amort_schedule(1_000_000, 0.06, 25, 5, False)
            ref_pmt = -npf.pmt(0.06/12, 25*12, 1_000_000) if DEPS.get("numpy_financial") else (1_000_000*(0.06/12))/(1-(1+0.06/12)**-(25*12))
            ref_balloon = 1_000_000*(1+0.06/12)**60 - ref_pmt*((1+0.06/12)**60 - 1)/(0.06/12)
            check("Payment vs independent reference", abs(pmt - ref_pmt) < 0.02, f"${pmt:,.2f} vs ${ref_pmt:,.2f}")
            check("Balloon vs independent reference", abs(balloon - ref_balloon) < 0.02, f"${balloon:,.2f} vs ${ref_balloon:,.2f}")
            refi_test = UnderwritingEngine.run_refi_stress(750_000, balloon, 25, 0.06)
            check("Refi stress returns scenarios", len(refi_test) == 6 and refi_test["Refi DSCR"].notna().all(), f"{len(refi_test)} rows")

            L_io, _, gates_io, _, _ = UnderwritingEngine.size_loan(1_000_000, 100_000_000, 100_000_000, 0, 0, 0.0, 0.06, 25, 5, True, 1.0, 1.0, 1.25, 0.01)
            check("Exact IO DSCR Sizing", abs(gates_io["DSCR"] - 13333333.33) < 1, f"${gates_io['DSCR']:,.0f}")

            try:
                UnderwritingEngine.size_loan(0, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
                check("Zero NOI raises underwriting error", False, "No exception raised")
            except UnderwritingError as exc:
                check("Zero NOI raises underwriting error", "NOI is required" in str(exc), str(exc))

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

            mc_df, mc_summary = RiskAnalyticsEngine.monte_carlo_deal(1_000_000, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09, 0.03, 0.02, 0.40, 0.06, 0.015, 0.02, 300, 0.075, 75, 50, "self-test")
            check("Monte Carlo engine returns rows", len(mc_df) == 300 and mc_summary.get("Simulations") == 300, f"{len(mc_df)} rows")

            if DEPS["crypto"]:
                pt = b"AlenzaTestDoc"
                key = "AlenzaTestKey-32plus-chars-for-local-self-test-only"
                check("Round-Trip Crypto", decrypt_bytes(encrypt_bytes(pt, key), key) == pt, f"PBKDF2HMAC Salted AES, {KDF_ITERATIONS:,} iterations")
                
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
        c = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None, check_same_thread=False)
        c.row_factory = sqlite3.Row
        c.execute("PRAGMA foreign_keys=ON")
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA busy_timeout=30000")
        c.execute("PRAGMA synchronous=NORMAL")
        return c

    @staticmethod
    @contextmanager
    def db_session(write: bool = False, retries: int = 5, backoff: float = 0.08):
        """Context-managed SQLite connection with retrying write transactions.

        SQLite remains a local/single-file persistence layer, not a multi-user
        database server. This wrapper uses WAL, busy_timeout, BEGIN IMMEDIATE,
        and bounded retries so local Streamlit sessions fail gracefully rather
        than crashing on transient ``database is locked`` errors. A thread lock
        protects only this process; WAL/retry handles the SQLite side.
        """
        last_exc = None
        for attempt in range(max(1, retries)):
            lock = DB_WRITE_LOCK if write else threading.RLock()
            with lock:
                conn = DatabaseManager.get_conn()
                try:
                    if write:
                        conn.execute("BEGIN IMMEDIATE")
                    yield conn
                    if write:
                        conn.commit()
                    return
                except sqlite3.OperationalError as exc:
                    last_exc = exc
                    if write:
                        try: conn.rollback()
                        except Exception: pass
                    locked = "locked" in str(exc).lower() or "busy" in str(exc).lower()
                    conn.close()
                    if not locked or attempt >= retries - 1:
                        raise
                    time.sleep(backoff * (2 ** attempt))
                    continue
                except Exception:
                    if write:
                        try: conn.rollback()
                        except Exception: pass
                    conn.close()
                    raise
                finally:
                    try: conn.close()
                    except Exception: pass
        if last_exc:
            raise last_exc

    @classmethod
    def init_db(cls):
        try:
            with cls.db_session(write=True) as c:
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
                    if "event_hash" not in cols_aud:
                        c.execute("ALTER TABLE audit_log ADD COLUMN event_hash TEXT DEFAULT ''")
        except sqlite3.Error as e: logger.error(f"DB Init failed: {e}")

    @classmethod
    def log_audit(cls, deal_id: str, action: str, details: str = ""):
        try: user = st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User"))
        except Exception: user = os.environ.get("APP_USER", "Local User")
        ts = datetime.now(timezone.utc).isoformat()
        try:
            with cls.db_session(write=True) as c:
                last_event = c.execute("SELECT event_hash FROM audit_log ORDER BY id DESC LIMIT 1").fetchone()
                prev_hash = last_event["event_hash"] if last_event and last_event["event_hash"] else "GENESIS"
                payload = f"{deal_id}|{user}|{action}|{details}|{ts}|{prev_hash}"
                event_hash = hashlib.sha256(payload.encode()).hexdigest()
                c.execute("INSERT INTO audit_log (deal_id, user, action, details, timestamp, prev_hash, event_hash) VALUES (?, ?, ?, ?, ?, ?, ?)", 
                          (deal_id, user, action, details, ts, prev_hash, event_hash))
        except sqlite3.Error as e: logger.error(f"Audit log failed: {e}")

    @classmethod
    def save_deal(cls, deal_id: str, name: str, state: dict) -> bool:
        clean_state = {k: state.get(k) for k in DEAL_STATE_KEYS}
        key = get_encryption_key()
        
        try:
            if key and DEPS["crypto"]:
                payload = json.dumps({"_alenza_storage": "encrypted", "payload": encrypt_text(json.dumps(clean_state), key)})
            elif os.environ.get("ALENZA_ALLOW_PLAINTEXT", "0") == "1":
                payload = json.dumps({"_alenza_storage": "plain", "payload": json.dumps(clean_state)})
            else:
                # Session-only demo mode: no disk persistence, no false security claim.
                if "_session_only_deals" not in st.session_state:
                    st.session_state["_session_only_deals"] = {}
                saved_at = datetime.now(timezone.utc).isoformat()
                st.session_state["_session_only_deals"][deal_id] = {"name": name, "state": clean_state, "saved_at": saved_at}
                st.session_state.setdefault("_session_only_versions", {}).setdefault(deal_id, []).append({"state": clean_state.copy(), "saved_at": saved_at})
                st.session_state["_last_save_mode"] = "Session-only; configure ALENZA_DB_ENCRYPTION_KEY or ALENZA_ALLOW_PLAINTEXT=1 for disk persistence. Version history is kept only for this browser session."
                return True
                
            now = datetime.now(timezone.utc).isoformat()
            with cls.db_session(write=True) as c:
                c.execute("INSERT OR REPLACE INTO deals (id, name, state_json, updated_at) VALUES (?, ?, ?, ?)", (deal_id, name, payload, now))
                c.execute("INSERT INTO deal_versions (deal_id, state_json, created_at) VALUES (?, ?, ?)", (deal_id, payload, now))
            cls.log_audit(deal_id, "SAVE_DEAL", f"Saved: {name}")
            st.session_state["_last_save_mode"] = "Database"
            return True
        except (sqlite3.Error, ValueError, TypeError, RuntimeError) as e:
            logger.error(f"Save failed: {e}")
            return False

    @classmethod
    def load_deal(cls, deal_id: str) -> Optional[dict]:
        try:
            with cls.db_session(write=False) as c:
                r = c.execute("SELECT state_json FROM deals WHERE id=?", (deal_id,)).fetchone()
                if not r: return None
                
                raw = json.loads(r["state_json"])
                if raw.get("_alenza_storage") == "encrypted":
                    key = get_encryption_key()
                    if not key: raise ValueError("Encrypted deal requires ALENZA_DB_ENCRYPTION_KEY")
                    state = json.loads(decrypt_text(raw["payload"], key))
                elif raw.get("_alenza_storage") == "plain":
                    if os.environ.get("ALENZA_ALLOW_PLAINTEXT", "0") != "1":
                        raise ValueError("Plaintext deal exists but ALENZA_ALLOW_PLAINTEXT=1 is not enabled for reads.")
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
            with cls.db_session(write=False) as c:
                return pd.read_sql_query("SELECT id, name, created_at, updated_at FROM deals ORDER BY updated_at DESC", c)
        except sqlite3.Error: return pd.DataFrame()

    @classmethod
    def delete_deal(cls, deal_id: str) -> bool:
        try:
            with cls.db_session(write=True) as c:
                d_name = c.execute("SELECT name FROM deals WHERE id=?", (deal_id,)).fetchone()
                deal_name = d_name["name"] if d_name else "Unknown"
                docs = c.execute("SELECT path FROM documents WHERE deal_id=?", (deal_id,)).fetchall()
                for row in docs:
                    try: Path(row["path"]).unlink(missing_ok=True)
                    except OSError: pass
                c.execute("DELETE FROM documents WHERE deal_id=?", (deal_id,))
                c.execute("DELETE FROM deal_versions WHERE deal_id=?", (deal_id,))
                c.execute("DELETE FROM deals WHERE id=?", (deal_id,))
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
        if not (key and DEPS["crypto"]) and os.environ.get("ALENZA_ALLOW_PLAINTEXT", "0") != "1":
            raise RuntimeError("Document vault persistence requires ALENZA_DB_ENCRYPTION_KEY or ALENZA_ALLOW_PLAINTEXT=1 for local demo mode.")
        is_enc = bool(key and DEPS["crypto"])
        
        if is_enc:
            content = encrypt_bytes(content, key)
            safe_name += ".enc"
            
        doc_root = Path(os.path.realpath(DOC_DIR)).resolve()
        path = Path(os.path.realpath(DOC_DIR / f"{doc_id}_{safe_name}")).resolve()
        try:
            path.relative_to(doc_root)
        except ValueError:
            raise ValueError("Unsafe document path rejected")

        try:
            with cls.db_session(write=True) as c:
                path.write_bytes(content)
                c.execute("INSERT INTO documents (id, deal_id, filename, category, path, size, is_encrypted, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", 
                          (doc_id, deal_id, safe_name.replace(".enc", ""), category, str(path), original_size, is_enc, datetime.now().isoformat()))
            cls.log_audit(deal_id, "DOC_UPLOAD", f"Uploaded {safe_name}")
            return True
        except (sqlite3.Error, OSError) as e:
            logger.error(f"Doc save failed: {e}")
            return False

    @classmethod
    def delete_doc(cls, doc_id: str) -> bool:
        try:
            with cls.db_session(write=True) as c:
                r = c.execute("SELECT deal_id, path, filename FROM documents WHERE id=?", (doc_id,)).fetchone()
                if not r: return False
                try: Path(r["path"]).unlink(missing_ok=True)
                except OSError: pass
                c.execute("DELETE FROM documents WHERE id=?", (doc_id,))
            cls.log_audit(r["deal_id"], "DOC_DELETE", f"Deleted {r['filename']}")
            return True
        except sqlite3.Error: return False

    @classmethod
    def get_versions(cls, deal_id: str) -> pd.DataFrame:
        session_versions = st.session_state.get("_session_only_versions", {}).get(deal_id, [])
        if session_versions and st.session_state.get("_last_save_mode", "").startswith("Session-only"):
            return pd.DataFrame([{"id": i, "created_at": v["saved_at"]} for i, v in enumerate(reversed(session_versions), 1)])
        try:
            with cls.db_session(write=False) as c:
                return pd.read_sql_query("SELECT id, created_at FROM deal_versions WHERE deal_id=? ORDER BY created_at DESC", c, params=(deal_id,))
        except sqlite3.Error: return pd.DataFrame()

    @classmethod
    def load_version(cls, version_id: int) -> Optional[dict]:
        try:
            # Session-only demo mode stores versions newest-last in memory. get_versions()
            # displays them newest-first with 1-based IDs, so translate that UI ID back
            # into the underlying list index before falling through to SQLite.
            if st.session_state.get("_last_save_mode", "").startswith("Session-only"):
                deal_id = st.session_state.get("deal_id")
                versions = st.session_state.get("_session_only_versions", {}).get(deal_id, [])
                idx = len(versions) - int(version_id)
                if 0 <= idx < len(versions):
                    return dict(versions[idx].get("state", {}))
                return None
            with cls.db_session(write=False) as c:
                r = c.execute("SELECT deal_id, state_json FROM deal_versions WHERE id=?", (version_id,)).fetchone()
                if not r: return None
                
                raw = json.loads(r["state_json"])
                if raw.get("_alenza_storage") == "encrypted":
                    key = get_encryption_key()
                    if not key: raise ValueError("Encrypted deal requires ALENZA_DB_ENCRYPTION_KEY")
                    state = json.loads(decrypt_text(raw["payload"], key))
                elif raw.get("_alenza_storage") == "plain":
                    if os.environ.get("ALENZA_ALLOW_PLAINTEXT", "0") != "1":
                        raise ValueError("Plaintext version exists but ALENZA_ALLOW_PLAINTEXT=1 is not enabled for reads.")
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
    # 2026 fallback context. These are conservative simulated placeholders used only
    # when StatsCan cannot be reached or its schema changes.
    fallback_year = max(2026, datetime.now(timezone.utc).year)
    fb = pd.DataFrame({
        "Date": pd.date_range(start=f"{fallback_year}-01-01", periods=12, freq="ME"),
        "Unemployment": [6.7, 6.8, 6.8, 6.9, 6.9, 6.8, 6.8, 6.7, 6.7, 6.6, 6.6, 6.5],
        "Source": [f"{fallback_year} fallback / simulated"] * 12,
    })
    headers = {"User-Agent": f"AlenzaCapitalOS/{VERSION}"}
    try:
        df = pd.read_csv("https://www150.statcan.gc.ca/n1/en/tbl/csv/14100287-eng.csv", storage_options=headers, low_memory=False)
        expected_cols = ["GEO", "Labour force characteristics", "Sex", "Age group", "REF_DATE", "VALUE"]
        missing_cols = [c for c in expected_cols if c not in df.columns]
        if missing_cols:
            logger.warning("StatsCan CSV format changed. Missing columns: %s", missing_cols)
            return fb, True
        m = (
            (df["GEO"].astype(str).str.lower() == "canada")
            & (df["Labour force characteristics"].astype(str).str.lower().str.contains("unemployment rate", na=False))
            & (df["Sex"].astype(str).str.lower() == "both sexes")
            & (df["Age group"].astype(str).str.lower() == "15 years and over")
        )
        out = df.loc[m, ["REF_DATE", "VALUE"]].rename(columns={"REF_DATE": "Date", "VALUE": "Unemployment"}).dropna()
        out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
        out["Unemployment"] = pd.to_numeric(out["Unemployment"], errors="coerce")
        out = out.dropna().sort_values("Date").reset_index(drop=True)
        if out.empty:
            logger.warning("StatsCan returned no usable unemployment rows. Using fallback.")
            return fb, True
        out["Source"] = "StatsCan"
        return out, False
    except Exception as e:
        logger.warning("StatsCan fetch failed: %s", e)
        return fb, True

@st.cache_data(ttl=86400)
def fetch_vacancy_rates():
    return pd.DataFrame({
        "Property Class": ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use", "Hospitality", "Self-Storage"],
        "National Vacancy": [2.4, 2.1, 5.6, 13.2, 4.9, 7.9, 4.1],
        "Trend": ["2026 fallback: Tight", "2026 fallback: Tight", "2026 fallback: Soft", "2026 fallback: High Vacancy", "2026 fallback: Stable", "2026 fallback: Mixed", "2026 fallback: Stable"],
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
    if not address:
        return None
    try:
        r = requests.get("https://geogratis.gc.ca/services/geolocation/en/locate", params={"q": address}, timeout=8)
        r.raise_for_status()
        payload = r.json()
        coords = []
        if isinstance(payload, dict) and payload.get("features"):
            coords = payload["features"][0].get("geometry", {}).get("coordinates", [])
        elif isinstance(payload, list) and payload:
            coords = payload[0].get("geometry", {}).get("coordinates", [])
        if len(coords) >= 2:
            lon, lat = safe_float(coords[0]), safe_float(coords[1])
            if -180 <= lon <= 180 and -90 <= lat <= 90:
                return {"lon": lon, "lat": lat, "source": "address geocode"}
    except Exception as e:
        logger.warning("Geocode failed: %s", e)
    return None

@st.cache_data(ttl=3600)
def fetch_ip_city_anchor() -> Optional[dict]:
    """Best-effort city anchor from public IP geolocation.

    Streamlit generally cannot access the browser user's true IP address from
    app code. This server-side lookup may reflect the hosting provider's egress
    city rather than the end user's city, so the UI labels it clearly. If the
    lookup fails, the Market Comps tab falls back to Toronto.
    """
    try:
        r = requests.get("https://ipapi.co/json/", timeout=4, headers={"User-Agent": f"AlenzaCapitalOS/{VERSION}"})
        r.raise_for_status()
        payload = r.json()
        lat_raw = payload.get("latitude") if isinstance(payload, dict) else None
        lon_raw = payload.get("longitude") if isinstance(payload, dict) else None
        lat = float(lat_raw)
        lon = float(lon_raw)
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            city = str(payload.get("city") or "IP-derived city")
            region = str(payload.get("region") or payload.get("country_name") or "")
            label = city if not region else f"{city}, {region}"
            return {"lat": lat, "lon": lon, "label": label, "source": "IP-derived/server egress city"}
    except Exception as e:
        logger.warning("IP city anchor lookup failed: %s", e)
    return None

# =============================================================================
# 10. EXPORT HELPERS (PDF)
# =============================================================================
def generate_pdf_memo(s: dict, loan: float, gate: str, ltv: float, dscr: float, req_eq: float, irr: float, c_stack: dict, flags: list, cmt: list, v_err: list, v_warn: list, h_pre: str) -> Optional[bytes]:
    if not DEPS["pdf"]: return None
    try:
        b = io.BytesIO()
        doc = SimpleDocTemplate(b, pagesize=letter)
        try:
            sty = getSampleStyleSheet()
            s_norm = sty.get("BodyText") or sty.get("Normal")
        except Exception as style_error:
            # ReportLab style discovery can fail in broken local installs.
            # Rather than crashing the app, create a tiny fallback style set.
            logger.warning("ReportLab styles unavailable; using fallback PDF styles: %s", style_error)
            from reportlab.lib.styles import ParagraphStyle
            sty = {
                "Title": ParagraphStyle("Title", fontSize=18, leading=22, spaceAfter=12),
                "Heading2": ParagraphStyle("Heading2", fontSize=14, leading=18, spaceAfter=8),
                "Heading3": ParagraphStyle("Heading3", fontSize=12, leading=16, spaceAfter=6),
                "BodyText": ParagraphStyle("BodyText", fontSize=9, leading=12),
            }
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
def calculate_outputs(state: dict) -> dict:
    # Human note: every run starts by normalizing the deal state. This prevents
    # odd UI or JSON-imported values from silently contaminating the model.
    try:
        state, schema_notes = DealStateValidator.coerce(state)
        L, gate, gates, uses, _ = UnderwritingEngine.size_loan(state['noi'], state['appraisal'], state['purchase_price'], state['closing_costs'], state['reserves'], state['fees'], state['rate'], state['amort'], state['term'], state['is_io'], state['target_ltv'], state['target_ltc'], state['target_dscr'], state['target_dy'])
    except UnderwritingError as exc:
        state = dict(state)
        schema_notes = [str(exc)]
        L, gate, gates = 0.0, "Underwriting Error", {"_messages": [str(exc)], "_used_gates": [], "_skipped_gates": ["All"]}
        uses = max(0.0, safe_float(state.get('purchase_price')) + safe_float(state.get('closing_costs')) + safe_float(state.get('reserves')))
    amort_df, m_pmt, balloon = UnderwritingEngine.amort_schedule(L, state['rate'], state['amort'], state['term'], state['is_io'])
    annual_ds = m_pmt * 12
    req_equity = uses - L - state['mezz_debt'] - state['pref_equity']
    c_stack = UnderwritingEngine.capital_stack(L, state['mezz_debt'], state['pref_equity'], max(0.0, req_equity), state['noi'], state['rate'], state['mezz_rate'], state['pref_rate'])
    collateral_value = safe_float(state.get('appraisal')) if safe_float(state.get('appraisal')) > EPS else safe_float(state.get('purchase_price'))
    act_ltv = safe_ratio(L, collateral_value)
    act_ltc = safe_ratio(L, uses)
    act_dscr = safe_ratio(state['noi'], annual_ds)
    act_dy = safe_ratio(state['noi'], L)
    pf_df = InvestmentEngine.calculate_pro_forma(state['noi'], state['pf_rev_growth'], state['pf_exp_growth'], state['pf_exp_ratio'])
    rets = InvestmentEngine.solve_returns(state['purchase_price'], L, pf_df, state['pf_exit_cap'], state['pf_sell_costs'], annual_ds, balloon, state['pf_term_growth'])
    
    return {
        "L": L, "gate": gate, "gates": gates, "uses": uses, "req_equity": req_equity, "amort_df": amort_df, 
        "m_pmt": m_pmt, "balloon": balloon, "annual_ds": annual_ds, "c_stack": c_stack, "act_ltv": act_ltv, 
        "act_ltc": act_ltc, "act_dscr": act_dscr, "act_dy": act_dy, "pf_df": pf_df, "rets": rets, "schema_notes": schema_notes
    }

def main():
    apply_theme()
    DatabaseManager.init_db()

    # Hydrate session state and run the blocking math-integrity gate before
    # aliasing ``s``. This guarantees refi_amort and future state keys exist
    # in old Streamlit browser sessions before any widget reads them.
    SessionManager.initialize()
    s = st.session_state

    # --- Sidebar ---
    with st.sidebar:
        st.title("ALENZA OS")
        if get_encryption_key() and DEPS.get("crypto"):
            st.caption("Persistence security: encrypted at rest")
        elif os.environ.get("ALENZA_ALLOW_PLAINTEXT", "0") == "1":
            st.warning("Local demo mode: plaintext persistence allowed by ALENZA_ALLOW_PLAINTEXT=1")
        else:
            st.error("Persistence disabled until ALENZA_DB_ENCRYPTION_KEY is configured.")
        s.deal_name = st.text_input("Deal Name", s.deal_name, on_change=SessionManager.mark_dirty)
        s.sponsor = st.text_input("Sponsor", s.sponsor, on_change=SessionManager.mark_dirty)
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
    if st.session_state.get("unsaved_changes"):
        st.caption("Unsaved changes detected since the last successful save.")
    header_placeholder = st.empty()

    # --- Tabs ---
    tabs = st.tabs(["Sizing & Risk", "Sensitivity", "Rent Roll", "Amortization", "Pro Forma", 
                    "Canada Intel", "Market Comps", "Diligence Room", "Save & Export", "Refi Stress", "QA & Health"])

    # TAB 1: Sizing & Risk
    with tabs[0]:
        c1, c2 = st.columns([1.5, 1])
        with c1:
            st.subheader("Asset & Debt Setup")
            k1, k2 = st.columns(2)
            s.purchase_price = k1.number_input("Purchase Price", value=s.purchase_price, step=50000.0, on_change=SessionManager.mark_dirty)
            s.appraisal = k2.number_input("Appraisal", value=s.appraisal, step=50000.0, on_change=SessionManager.mark_dirty)
            s.property_address = st.text_input("Property Address (For Geocoding)", value=s.property_address, on_change=SessionManager.mark_dirty)
            k3, k4 = st.columns(2)
            s.noi = k3.number_input("Stabilized NOI", value=s.noi, step=5000.0, on_change=SessionManager.mark_dirty)
            
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
        st.subheader("Sensitivity & Quant Risk")
        st.caption("Human note: the heatmap shows deterministic gate movement; the Monte Carlo block below shows probabilistic downside bands.")
        q1, q2, q3, q4 = st.columns(4)
        s.mc_sims = q1.number_input("Monte Carlo Runs", min_value=100, max_value=10000, value=int(s.mc_sims), step=100)
        s.mc_noi_vol = q2.slider("NOI Volatility (%)", 0.0, 50.0, float(s.mc_noi_vol * 100), 0.5) / 100.0
        s.mc_rate_vol_bps = q3.slider("Rate Shock Vol (bps)", 0, 500, int(s.mc_rate_vol_bps), 5)
        s.mc_exit_cap_vol_bps = q4.slider("Exit Cap Vol (bps)", 0, 500, int(s.mc_exit_cap_vol_bps), 5)
        sens_placeholder = st.empty()

    # TAB 3: Rent Roll
    with tabs[2]:
        st.subheader("Rent Roll Normalization")
        rr_number_format = st.selectbox("Rent Roll Number Format", ["Canadian/US", "European"], help="Canadian/US parses 1,234.56. European parses 1.234,56 or 1 234,56.")
        st.caption("Only Tenant, SF, Remaining Term, and Monthly Rent are used in underwriting metrics. Other columns are retained in the deal file but ignored by calculations.")
        rr_df, rr_warnings, rr_map = normalize_rr_with_diagnostics(pd.DataFrame(s.rent_roll_dict), rr_number_format)
        err = rr_df.copy()
        if rr_map:
            with st.expander("Column mapping preview", expanded=False):
                st.dataframe(pd.DataFrame([{"Source Column": k, "Mapped To": v} for k, v in rr_map.items()]), hide_index=True, use_container_width=True)
        for msg in rr_warnings:
            st.warning(msg)
        if len(rr_df) > 1000:
            st.error("Large rent roll detected. Dynamic editing is disabled above 1,000 rows to avoid freezing the browser. Upload a smaller edit batch or use the normalized static view.")
            st.dataframe(rr_df.head(1000), hide_index=True, use_container_width=True)
            err = rr_df.copy()
        
        if st.button("Add Blank Row"):
            d = rr_df.to_dict("records"); d.append({"Tenant":"", "SF":0, "Remaining Term":0, "Monthly Rent":0}); s.rent_roll_dict = d; st.rerun()
            
        if len(rr_df) <= 1000:
            err = st.data_editor(rr_df, num_rows="dynamic", use_container_width=True)
        if not normalize_rr_with_diagnostics(err, rr_number_format)[0].equals(rr_df):
            st.warning("Rent roll changes are staged. Click below to apply to the deal model.")
            if st.button("Apply Rent Roll Changes", type="primary"):
                s.rent_roll_dict = normalize_rr_with_diagnostics(err, rr_number_format)[0].to_dict("records")
                st.rerun()
            
        (tsf, occ, ar, psf, walt, exp1), rr_metric_warnings = UnderwritingEngine.rent_roll_metrics_with_diagnostics(pd.DataFrame(s.rent_roll_dict))
        for msg in rr_metric_warnings:
            if "empty" in msg.lower() or "no occupied" in msg.lower() or "no positive" in msg.lower():
                st.error(msg)
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Total SF", f"{tsf:,.0f}"); k2.metric("Occupancy", f"{occ:.1%}")
        k3.metric("Annual Rent", f"${ar:,.0f}"); k4.metric("Rent PSF", f"${psf:,.2f}")
        k5.metric("WALT", f"{walt:.1f}"); k6.metric("Exp. <= 1Y", f"{exp1:.1%}")

    # TAB 4: Amortization
    with tabs[3]:
        st.subheader("Amortization & Paydown")
        amort_placeholder = st.empty()

    # TAB 5: Pro Forma
    with tabs[4]:
        st.subheader("10-Year Pro Forma & Returns")
        with st.expander("Financial Conventions & Model Governance"):
            st.write(FINANCIAL_CONVENTIONS)
            st.write(MODEL_GOVERNANCE_NOTES)
        
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
        st.info("Simulated market comp sale dates are constrained to late 2025 through May 14, 2026 for current-market context. These are not real transaction records.")
        comps = generate_comps(s.property_type, s.noi, s.appraisal, s.deal_id)
        geo = geocode_address(st.session_state.get("property_address", "")) if st.session_state.get("property_address") else None
        if geo:
            st.caption("Map anchor: subject property address geocode.")
        else:
            geo = fetch_ip_city_anchor()
            if geo:
                st.warning(f"Property address geocode unavailable. Using {geo.get('label', 'IP-derived city')} as a best-effort city anchor. In hosted Streamlit deployments this may reflect the server egress city, not the end user's exact location.")
            else:
                geo = {"lat": 43.65, "lon": -79.38, "label": "Toronto fallback", "source": "fallback"}
                st.warning("Property address and IP-city lookup unavailable. Using Toronto fallback anchor; comps remain simulated and location is approximate.")
        c_lat, c_lon = (geo["lat"], geo["lon"])
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
                try:
                    if DatabaseManager.save_doc(s.deal_id, up, cat): st.success("Uploaded!"); st.rerun()
                except Exception as exc:
                    st.error(f"Vault save blocked: {exc}")
        
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
    h_pre = SessionManager.refresh_dirty_flag(current_state)

    # Fill Top Header
    with header_placeholder.container():
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Mortgage Amount", f"${out['L']:,.0f}")
        c1.caption(f"Binding Constraint: **{out['gate']}**")
        for _msg in out.get("gates", {}).get("_messages", []):
            st.warning(_msg)
        if out.get("schema_notes"):
            st.info("Input normalization notes: " + "; ".join(out["schema_notes"]))
        if out["L"] <= 0:
            missing_msgs = out.get("gates", {}).get("_messages", [])
            if missing_msgs:
                st.error("Mortgage not calculated: " + "; ".join(missing_msgs))
            else:
                st.error("Mortgage not calculated because one or more underwriting constraints failed.")
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
            
        rr_metrics_for_risk, rr_risk_warnings = UnderwritingEngine.rent_roll_metrics_with_diagnostics(pd.DataFrame(s.rent_roll_dict))
        occ_rr = safe_ratio(rr_metrics_for_risk[1], 1.0)
        if rr_risk_warnings:
            st.caption("Rent-roll diagnostic path used for risk metrics: " + "; ".join(rr_risk_warnings[:2]))
        st.metric("Breakeven Occupancy", f"{UnderwritingEngine.breakeven_occupancy(s.noi, max(occ_rr, 0.01), out['c_stack']['FixedCharges']):.1%}")

    # Fill Tab 2 Sensitivity
    with sens_placeholder.container():
        st.markdown("#### Deterministic Gate Heatmap")
        hm = SensitivityEngine.proceeds_heatmap(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy)
        hm_numeric = hm.drop(columns=["NOI Shock"], errors="ignore")
        if DEPS["plotly"]:
            fig = px.imshow(hm_numeric/1e6, text_auto=".1f", aspect="auto", title="Mortgage Amount / Senior Loan Proceeds ($MM; N/A = underwriting error)")
            fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
            st.plotly_chart(fig, use_container_width=True)
            if hm_numeric.isna().any().any(): st.caption("N/A cells indicate an underwriting error, not a valid zero-dollar gate result.")
        else:
            st.dataframe(hm.style.background_gradient(subset=hm_numeric.columns, cmap="YlOrBr").format({c: "${:,.0f}" for c in hm_numeric.columns}), hide_index=True, use_container_width=True)

        st.markdown("#### Monte Carlo Risk Lens")
        st.caption(
            "Monte Carlo uses a stable seed so assumption changes can be compared on the same random draw sequence."
        )
        with st.spinner(f"Running {int(s.mc_sims):,} Monte Carlo simulations..."):
            mc_df, mc_summary = RiskAnalyticsEngine.monte_carlo_deal(
                s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate,
                s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy,
                s.pf_rev_growth, s.pf_exp_growth, s.pf_exp_ratio, s.pf_exit_cap, s.pf_sell_costs,
                s.pf_term_growth, s.mc_sims, s.mc_noi_vol, s.mc_rate_vol_bps, s.mc_exit_cap_vol_bps, s.deal_id,
            )
        if mc_df.empty:
            st.info(mc_summary.get("Message", "Monte Carlo unavailable."))
        else:
            r1, r2, r3, r4 = st.columns(4)
            r1.metric("Mortgage P50", f"${mc_summary['Mortgage P50']:,.0f}")
            r2.metric("Mortgage P5", f"${mc_summary['Mortgage P5']:,.0f}")
            r3.metric("DSCR P5 / VaR 95%", f"{mc_summary['DSCR VaR 95%']:.2f}x")
            r4.metric("Prob. DSCR < 1.20x", f"{mc_summary['Probability DSCR < 1.20x']:.1%}")

            b1, b2, b3, b4 = st.columns(4)
            b1.metric("Survival Rate", f"{mc_summary['Survival Rate']:.1%}", help="Share of simulated scenarios with DSCR >= 1.00x.")
            b2.metric("Success Rate", f"{mc_summary['Success Rate']:.1%}", help="Share of scenarios meeting the target DSCR input.")
            b3.metric("Risk of Default", f"{mc_summary['Risk of Default']:.1%}", help="Share of scenarios with DSCR below 1.00x.")
            b4.metric("Avg Yield on Cost", f"{mc_summary['Avg Yield on Cost']:.2%}")

            st.caption("Backtest reference: DSCR VaR 95% is the empirical 5th percentile of simulated DSCR, not a parametric normal approximation. Survival Rate = DSCR >= 1.00x.")
            if DEPS["plotly"]:
                fig_mc = px.histogram(mc_df, x="Mortgage", nbins=40, title="Mortgage Amount Distribution")
                fig_mc.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
                st.plotly_chart(fig_mc, use_container_width=True)
                fig_dscr = px.histogram(mc_df, x="DSCR", nbins=40, title="DSCR Distribution / Survival Backtest")
                fig_dscr.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A")
                st.plotly_chart(fig_dscr, use_container_width=True)
            st.caption("IRR P5/P50 are conditional on valid IRRs. Undefined/extreme-loss IRRs below -95% are excluded and disclosed through IRR Valid Rate.")
            if mc_summary.get("Appraisal Proxy Used"):
                st.warning("Monte Carlo used purchase price as the LTV value because appraisal is blank.")
            st.dataframe(mc_df.describe(percentiles=[0.05, 0.50, 0.95]).T, use_container_width=True)

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
                    SessionManager.mark_saved(current_state)
                    mode = st.session_state.get("_last_save_mode", "Unknown")
                    if str(mode).startswith("Session-only"):
                        st.warning(mode)
                    st.toast("Deal Saved", icon="✅")
                else:
                    st.error("Save failed. Check encryption/key settings and server logs.")

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
                mc_export, mc_summary_export = RiskAnalyticsEngine.monte_carlo_deal(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy, s.pf_rev_growth, s.pf_exp_growth, s.pf_exp_ratio, s.pf_exit_cap, s.pf_sell_costs, s.pf_term_growth, s.mc_sims, s.mc_noi_vol, s.mc_rate_vol_bps, s.mc_exit_cap_vol_bps, s.deal_id)
                if not mc_export.empty:
                    mc_export.to_excel(w, sheet_name="Monte Carlo", index=False)
                    pd.DataFrame([mc_summary_export]).to_excel(w, sheet_name="MC Summary", index=False)
                comps.to_excel(w, sheet_name="Market Comps", index=False)
                
                # FLATTENED VALIDATION LIST FOR PROPER EXCEL EXPORT
                val_data = [{"Type": "Error", "Message": e} for e in v_err] + [{"Type": "Warning", "Message": w} for w in v_warn]
                if not val_data:
                    val_data = [{"Type": "Status", "Message": "All validations passed."}]
                pd.DataFrame(val_data).to_excel(w, sheet_name="Validation", index=False)
                
                if latest:
                    latest_rows = [{"Series": k, "Value": v.get("val"), "Date": v.get("date")} for k, v in latest.items()]
                    pd.DataFrame(latest_rows).to_excel(w, sheet_name="Canada Intel", index=False)
                pd.DataFrame(cmt, columns=["Severity", "Topic", "Note"]).to_excel(w, sheet_name="Market Commentary", index=False)
                pd.DataFrame([{"Key": k, "Value": v} for k,v in FINANCIAL_CONVENTIONS.items()]).to_excel(w, sheet_name="Conventions", index=False)
                pd.DataFrame([{"Key": k, "Value": v} for k,v in MODEL_GOVERNANCE_NOTES.items()]).to_excel(w, sheet_name="Model Governance", index=False)
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
                mc_zip, mc_summary_zip = RiskAnalyticsEngine.monte_carlo_deal(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy, s.pf_rev_growth, s.pf_exp_growth, s.pf_exp_ratio, s.pf_exit_cap, s.pf_sell_costs, s.pf_term_growth, s.mc_sims, s.mc_noi_vol, s.mc_rate_vol_bps, s.mc_exit_cap_vol_bps, s.deal_id)
                zf.writestr("monte_carlo.csv", mc_zip.to_csv(index=False) if not mc_zip.empty else "")
                zf.writestr("monte_carlo_summary.json", json.dumps(mc_summary_zip, default=str))
                zf.writestr("degraded_data_flags.json", json.dumps({"boc_degraded": boc_deg, "unemp_degraded": u_deg}))
                zf.writestr("conventions.txt", json.dumps(FINANCIAL_CONVENTIONS, indent=2))
                zf.writestr("model_governance.json", json.dumps(MODEL_GOVERNANCE_NOTES, indent=2))
                
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


    # TAB 10: Refinance Stress
    with tabs[9]:
        st.subheader("Refinance Stress at Balloon")
        st.caption("Reference: this tab incorporates the uploaded unified refi-stress block. It uses the same calculate_mortgage_payment() source as amortization and debt sizing, so Annual DS and Refi DSCR cannot drift across modules.")
        current_refi_amort = int(st.session_state.get("refi_amort", default_state().get("refi_amort", 25)))
        st.session_state["refi_amort"] = st.number_input(
            "Refi Underwriting Amortization (Yrs)",
            min_value=1,
            max_value=40,
            value=current_refi_amort,
            step=1,
            on_change=SessionManager.mark_dirty,
        )
        refi_amort = int(st.session_state.get("refi_amort", current_refi_amort))
        balloon_for_refi = float(out.get("balloon", 0.0))
        if balloon_for_refi <= EPS:
            st.info("No positive balloon balance is available. Refi stress requires a calculated loan and amortization schedule.")
        else:
            stress_df = UnderwritingEngine.run_refi_stress(float(current_state.get("noi", 0.0)), balloon_for_refi, refi_amort, float(current_state.get("rate", 0.0)))
            st.warning(f"Year {int(current_state.get('term', 0))} balloon balance: ${balloon_for_refi:,.2f}")
            st.dataframe(stress_df.style.format({"Rate": "{:.2%}", "Refi DS": "${:,.2f}", "Refi DSCR": "{:.2f}x"}), hide_index=True, use_container_width=True)
            min_dscr = float(stress_df["Refi DSCR"].min()) if not stress_df.empty else 0.0
            if min_dscr < 1.20:
                st.error(f"Refi stress warning: minimum scenario DSCR is {min_dscr:.2f}x, below common refinance comfort thresholds.")
            else:
                st.success(f"Refi stress passes current scenario set. Minimum DSCR: {min_dscr:.2f}x.")

    # TAB 11: QA
    with tabs[10]:
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
