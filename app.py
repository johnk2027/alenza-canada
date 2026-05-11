"""
Alenza Capital OS v3.0
Enterprise underwriting workspace - Production Grade
Midnight Slate and CU Gold theme
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
import time
import hashlib
import logging
import threading
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional, Callable
from functools import lru_cache, wraps
from contextlib import contextmanager
from dataclasses import dataclass, asdict
from enum import Enum

# ==========================================
# CONFIGURATION & INITIALIZATION
# ==========================================

# Setup logging first
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('alenza_os.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Page configuration
st.set_page_config(
    page_title="Alenza Capital OS",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Constants
APP_VERSION = "3.0.0"
SCHEMA_VERSION = 4
EXPORT_ENCRYPTION = True
AUTO_SAVE_INTERVAL = 300  # 5 minutes
MAX_UPLOAD_SIZE_MB = 50
SESSION_TIMEOUT_MINUTES = 120

# ==========================================
# TYPE DEFINITIONS & ENUMS
# ==========================================

class PropertyType(Enum):
    MULTIFAMILY = "Multifamily"
    INDUSTRIAL = "Industrial"
    RETAIL = "Retail"
    OFFICE = "Office"
    MIXED_USE = "Mixed-Use"
    HOSPITALITY = "Hospitality"
    SELF_STORAGE = "Self-Storage"

class LenderProfile(Enum):
    BANK = "Bank / Credit Union"
    LIFECO = "LifeCo / Core"
    BRIDGE = "Bridge / Private"
    CMHC = "CMHC Multifamily"

class TransactionType(Enum):
    ACQUISITION = "Acquisition"
    REFINANCE = "Refinance"
    CONSTRUCTION = "Construction"
    BRIDGE = "Bridge"
    RECAP = "Recapitalization"

class DealTier(Enum):
    TIER_1 = "Tier 1 | Institutional Core"
    TIER_2 = "Tier 2 | Conventional Bankable"
    TIER_3 = "Tier 3 | Alternative / Debt Fund"
    TIER_4 = "Tier 4 | Private / Restructure"

@dataclass
class DealState:
    """Type-safe deal state container"""
    deal_id: str = ""
    deal_name: str = "Untitled Deal"
    sponsor: str = ""
    property_address: str = ""
    property_type: str = "Multifamily"
    transaction_type: str = "Acquisition"
    lender_profile: str = "Bank / Credit Union"
    purchase_price: float = 0.0
    appraisal: float = 0.0
    noi: float = 0.0
    target_ltv: float = 0.75
    target_ltc: float = 0.80
    target_dscr: float = 1.25
    target_dy: float = 0.085
    rate: float = 0.0525
    amort: int = 25
    term: int = 5
    is_io: bool = False
    fees: float = 0.02
    closing_costs: float = 0.0
    reserves: float = 0.0
    rent_roll_dict: List[Dict] = None
    
    def __post_init__(self):
        if self.rent_roll_dict is None:
            self.rent_roll_dict = []

# ==========================================
# DEPENDENCY MANAGEMENT
# ==========================================

@st.cache_resource
def check_dependencies() -> Dict[str, bool]:
    """Check and cache available dependencies"""
    deps = {
        "ocr": False,
        "pdf": False,
        "excel_write": False,
        "excel_read": False,
        "plotly": False,
        "crypto": False,
    }
    
    try:
        from PIL import Image
        import pytesseract
        import fitz
        deps["ocr"] = True
    except ImportError:
        pass
    
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate
        deps["pdf"] = True
    except ImportError:
        pass
    
    try:
        import xlsxwriter
        deps["excel_write"] = True
    except ImportError:
        pass
    
    try:
        import openpyxl
        deps["excel_read"] = True
    except ImportError:
        pass
    
    try:
        import plotly.express as px
        deps["plotly"] = True
    except ImportError:
        pass
    
    try:
        from cryptography.fernet import Fernet
        deps["crypto"] = True
    except ImportError:
        pass
    
    return deps

DEPENDENCIES = check_dependencies()

# ==========================================
# ERROR HANDLING DECORATORS
# ==========================================

def handle_errors(default_return=None, show_ui_error=True):
    """Decorator for consistent error handling"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"Error in {func.__name__}: {e}", exc_info=True)
                if show_ui_error:
                    st.error(f"Operation failed: {str(e)[:200]}")
                return default_return
        return wrapper
    return decorator

def retry_on_failure(max_retries=3, backoff=0.5):
    """Decorator for retrying operations"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_error = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_error = e
                    if attempt < max_retries - 1:
                        time.sleep(backoff * (2 ** attempt))
            raise last_error
        return wrapper
    return decorator

# ==========================================
# UTILITY FUNCTIONS
# ==========================================

def safe_float(value: Any, default: float = 0.0) -> float:
    """Safely convert value to float with comprehensive cleaning"""
    if value is None or value == "" or (isinstance(value, float) and np.isnan(value)):
        return default
    try:
        if isinstance(value, (int, float, np.number)):
            return float(value)
        if isinstance(value, str):
            # Remove currency symbols and formatting
            cleaned = re.sub(r'[^\d.\-()]', '', value)
            cleaned = cleaned.strip()
            # Handle accounting notation (parentheses for negative)
            if cleaned.startswith('(') and cleaned.endswith(')'):
                cleaned = '-' + cleaned[1:-1]
            return float(cleaned) if cleaned else default
        return default
    except (ValueError, TypeError):
        return default

def safe_int(value: Any, default: int = 0) -> int:
    """Safely convert value to integer"""
    try:
        return int(safe_float(value, default))
    except (ValueError, TypeError):
        return default

def normalize_percent(value: Any, default: float = 0.0) -> float:
    """Accept percentage as 0.75 or 75, return decimal"""
    x = safe_float(value, default)
    return x / 100 if x > 1.5 else max(0.0, min(2.0, x))

def generate_id(prefix: str = "deal") -> str:
    """Generate unique identifier"""
    timestamp = int(datetime.now(timezone.utc).timestamp())
    random_hex = uuid.uuid4().hex[:8]
    return f"{prefix}_{timestamp}_{random_hex}"

def hash_state(state: Dict) -> str:
    """Create hash of state for change detection"""
    state_str = json.dumps(state, sort_keys=True, default=str)
    return hashlib.sha256(state_str.encode()).hexdigest()

# ==========================================
# AUTO-SAVE MANAGER
# ==========================================

class AutoSaveManager:
    """Manages automatic saving of deal state"""
    
    def __init__(self):
        self.last_save_hash = None
        self.last_save_time = None
        self._timer = None
    
    def start_auto_save(self):
        """Start auto-save timer"""
        if self._timer:
            self._timer.cancel()
        self._timer = threading.Timer(AUTO_SAVE_INTERVAL, self._auto_save)
        self._timer.daemon = True
        self._timer.start()
    
    def _auto_save(self):
        """Perform auto-save"""
        try:
            current_state = extract_clean_state()
            current_hash = hash_state(current_state)
            
            if current_hash != self.last_save_hash:
                DatabaseManager.save_deal(
                    current_state['deal_id'],
                    current_state.get('deal_name', 'Auto-saved Deal'),
                    current_state
                )
                self.last_save_hash = current_hash
                self.last_save_time = datetime.now(timezone.utc)
                logger.info("Auto-save completed")
        except Exception as e:
            logger.error(f"Auto-save failed: {e}")
        finally:
            self.start_auto_save()  # Reschedule

# Global auto-save manager
auto_save = AutoSaveManager()

# ==========================================
# STATE MANAGEMENT
# ==========================================

def initialize_session_state():
    """Initialize all session state variables"""
    defaults = {
        'deal_id': generate_id('deal'),
        'deal_name': 'Untitled Deal',
        'sponsor': '',
        'property_address': '',
        'property_type': PropertyType.MULTIFAMILY.value,
        'transaction_type': TransactionType.ACQUISITION.value,
        'lender_profile': LenderProfile.BANK.value,
        'purchase_price': 0.0,
        'appraisal': 0.0,
        'noi': 0.0,
        'target_ltv': 0.75,
        'target_ltc': 0.80,
        'target_dscr': 1.25,
        'target_dy': 0.085,
        'rate': 0.0525,
        'amort': 25,
        'term': 5,
        'is_io': False,
        'fees': 0.02,
        'closing_costs': 0.0,
        'reserves': 0.0,
        'rent_roll_dict': [],
        'last_saved_at': None,
        'unsaved_changes': False,
        'current_tab': 0,
        'user_preferences': {
            'show_advanced': False,
            'currency_format': 'USD',
            'dark_mode': True,
        }
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def extract_clean_state() -> Dict:
    """Extract current deal state for saving"""
    keys = [
        'deal_id', 'deal_name', 'sponsor', 'property_address',
        'property_type', 'transaction_type', 'lender_profile',
        'purchase_price', 'appraisal', 'noi', 'target_ltv',
        'target_ltc', 'target_dscr', 'target_dy', 'rate',
        'amort', 'term', 'is_io', 'fees', 'closing_costs',
        'reserves', 'rent_roll_dict', 'last_saved_at'
    ]
    
    state = {}
    for key in keys:
        if key in st.session_state:
            val = st.session_state[key]
            # Convert numpy types
            if isinstance(val, (np.integer,)):
                val = int(val)
            elif isinstance(val, (np.floating,)):
                val = float(val)
            elif isinstance(val, np.ndarray):
                val = val.tolist()
            state[key] = val
    
    state['schema_version'] = SCHEMA_VERSION
    state['app_version'] = APP_VERSION
    return state

def on_state_change():
    """Callback when any input changes"""
    st.session_state.unsaved_changes = True

# ==========================================
# DATABASE MANAGER (Simplified for brevity - full version from database.py)
# ==========================================

class DatabaseManager:
    """Database operations with connection pooling"""
    
    _pool = []
    
    @classmethod
    @contextmanager
    def get_connection(cls):
        conn = None
        try:
            if cls._pool:
                conn = cls._pool.pop()
            else:
                conn = sqlite3.connect('alenza_data/alenza_platform.db', timeout=30)
                conn.execute("PRAGMA journal_mode=WAL")
                conn.execute("PRAGMA foreign_keys=ON")
                conn.row_factory = sqlite3.Row
            yield conn
        finally:
            if conn and len(cls._pool) < 5:
                cls._pool.append(conn)
            elif conn:
                conn.close()
    
    @classmethod
    def init_db(cls):
        """Initialize database schema"""
        os.makedirs('alenza_data', exist_ok=True)
        os.makedirs('alenza_data/documents', exist_ok=True)
        
        with cls.get_connection() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS deals (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    state_json TEXT,
                    state_hash TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user TEXT NOT NULL,
                    action TEXT NOT NULL,
                    details TEXT,
                    deal_id TEXT REFERENCES deals(id) ON DELETE CASCADE,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS deal_versions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    deal_id TEXT REFERENCES deals(id) ON DELETE CASCADE,
                    state_json TEXT,
                    changed_by TEXT,
                    change_summary TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY,
                    deal_id TEXT REFERENCES deals(id) ON DELETE CASCADE,
                    filename TEXT NOT NULL,
                    original_filename TEXT,
                    category TEXT NOT NULL,
                    path TEXT NOT NULL,
                    file_size INTEGER,
                    file_hash TEXT,
                    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE INDEX IF NOT EXISTS idx_deals_updated ON deals(updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp DESC);
                CREATE INDEX IF NOT EXISTS idx_versions_deal ON deal_versions(deal_id, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_docs_deal ON documents(deal_id);
            """)
            conn.commit()
    
    @classmethod
    @handle_errors(default_return=False)
    def save_deal(cls, deal_id: str, name: str, state: Dict) -> bool:
        """Save deal with versioning"""
        state_json = json.dumps(state, default=str)
        state_hash = hash_state(state)
        safe_name = str(name or "Untitled Deal").strip()[:200]
        
        with cls.get_connection() as conn:
            # Check for changes
            row = conn.execute("SELECT state_hash FROM deals WHERE id = ?", (deal_id,)).fetchone()
            if row and row['state_hash'] == state_hash:
                return True  # No changes
            
            # Save deal
            conn.execute(
                """INSERT OR REPLACE INTO deals (id, name, state_json, state_hash, updated_at)
                VALUES (?, ?, ?, ?, ?)""",
                (deal_id, safe_name, state_json, state_hash, datetime.now(timezone.utc))
            )
            
            # Save version
            old_state = {}
            if row:
                old_row = conn.execute("SELECT state_json FROM deals WHERE id = ?", (deal_id,)).fetchone()
                if old_row:
                    try:
                        old_state = json.loads(old_row['state_json'])
                    except:
                        pass
            
            change_summary = cls._summarize_changes(old_state, state)
            conn.execute(
                """INSERT INTO deal_versions (deal_id, state_json, changed_by, change_summary)
                VALUES (?, ?, ?, ?)""",
                (deal_id, state_json, cls._get_user(), change_summary)
            )
            
            conn.commit()
            return True
    
    @classmethod
    @handle_errors(default_return=None)
    def load_deal(cls, deal_id: str) -> Optional[Dict]:
        """Load deal state"""
        with cls.get_connection() as conn:
            row = conn.execute("SELECT state_json FROM deals WHERE id = ?", (deal_id,)).fetchone()
            if row:
                return json.loads(row['state_json'])
        return None
    
    @classmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_all_deals(cls) -> pd.DataFrame:
        """Get all deals"""
        with cls.get_connection() as conn:
            return pd.read_sql_query(
                "SELECT id, name, created_at, updated_at FROM deals ORDER BY updated_at DESC",
                conn
            )
    
    @staticmethod
    def _get_user() -> str:
        """Get current user"""
        return os.getenv('ALENZA_USER', 'Local User')
    
    @staticmethod
    def _summarize_changes(old: Dict, new: Dict) -> str:
        """Create change summary"""
        tracked = ['deal_name', 'purchase_price', 'appraisal', 'noi', 'rate']
        changes = []
        for field in tracked:
            old_val = old.get(field) if old else None
            new_val = new.get(field) if new else None
            if str(old_val) != str(new_val):
                changes.append(f"{field}: {old_val} → {new_val}")
        return "; ".join(changes) if changes else "No material changes"

# ==========================================
# FINANCIAL ENGINE (Core Underwriting Logic)
# ==========================================

class UnderwritingEngine:
    """Core financial calculations"""
    
    LENDER_LIMITS = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05},
    }
    
    @staticmethod
    @handle_errors(default_return=(0, "Error", {}, 0, 0))
    def size_loan(
        noi: float, appraisal: float, purchase_price: float,
        closing_costs: float, reserves: float, fees_pct: float,
        rate: float, amort: int, term: int, is_io: bool,
        target_ltv: float, target_ltc: float, target_dscr: float, target_dy: float
    ) -> Tuple[float, str, Dict[str, float], float, float]:
        """Calculate maximum supportable loan"""
        
        # Input validation and sanitization
        noi = max(0, safe_float(noi))
        appraisal = max(0, safe_float(appraisal))
        purchase_price = max(0, safe_float(purchase_price))
        closing_costs = max(0, safe_float(closing_costs))
        reserves = max(0, safe_float(reserves))
        fees_pct = max(0, min(0.10, safe_float(fees_pct)))
        rate = max(0.001, min(0.30, safe_float(rate)))
        amort = max(1, min(40, safe_int(amort)))
        target_ltv = max(0.01, min(1.25, safe_float(target_ltv)))
        target_ltc = max(0.01, min(1.25, safe_float(target_ltc)))
        target_dscr = max(0.01, safe_float(target_dscr))
        target_dy = max(0.0001, min(0.25, safe_float(target_dy)))
        
        base_cost = purchase_price + closing_costs + reserves
        gates = {"LTV": 0, "LTC": 0, "DSCR": 0, "Debt Yield": 0}
        loan = 0
        
        # Iterate to convergence (typically 2-3 iterations needed)
        for _ in range(10):
            total_uses = base_cost + (loan * fees_pct)
            
            # Calculate each constraint
            if appraisal > 0:
                gates["LTV"] = appraisal * target_ltv
            
            if total_uses > 0:
                gates["LTC"] = total_uses * target_ltc
            
            if target_dy > 0:
                gates["Debt Yield"] = noi / target_dy
            
            # DSCR constraint
            monthly_rate = rate / 12
            if monthly_rate > 0:
                if is_io:
                    gates["DSCR"] = (noi / target_dscr) / 12 / monthly_rate
                else:
                    total_payments = amort * 12
                    if total_payments > 0:
                        pmt_factor = (1 - (1 + monthly_rate) ** -total_payments) / monthly_rate
                        if pmt_factor > 0:
                            gates["DSCR"] = (noi / target_dscr) / 12 * pmt_factor
                        else:
                            gates["DSCR"] = float('inf')
                    else:
                        gates["DSCR"] = float('inf')
            else:
                gates["DSCR"] = float('inf')
            
            new_loan = min(gates.values())
            
            # Check convergence
            if abs(new_loan - loan) < 1.0:
                break
            
            loan = new_loan
        
        # Determine binding constraint
        gate = min(gates, key=gates.get) if gates else "N/A"
        req_equity = total_uses - loan
        
        return loan, gate, gates, total_uses, req_equity
    
    @staticmethod
    @handle_errors(default_return=(pd.DataFrame(), 0, 0))
    def amort_schedule(
        loan_amt: float, rate: float, amort_yrs: int, term_yrs: int, is_io: bool
    ) -> Tuple[pd.DataFrame, float, float]:
        """Generate amortization schedule"""
        
        loan_amt = max(0, safe_float(loan_amt))
        rate = max(0, min(0.30, safe_float(rate)))
        amort_yrs = max(1, min(40, safe_int(amort_yrs)))
        term_yrs = max(1, min(40, safe_int(term_yrs)))
        
        if loan_amt <= 0:
            return pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]), 0, 0
        
        monthly_rate = rate / 12
        total_payments = amort_yrs * 12
        term_months = term_yrs * 12
        
        # Calculate monthly payment
        if is_io:
            monthly_pmt = loan_amt * monthly_rate
        elif monthly_rate > 0:
            monthly_pmt = (loan_amt * monthly_rate) / (1 - (1 + monthly_rate) ** -total_payments)
        else:
            monthly_pmt = loan_amt / total_payments
        
        # Generate schedule
        schedule = []
        balance = loan_amt
        
        for period in range(1, term_months + 1):
            interest = balance * monthly_rate
            
            if is_io:
                principal_paid = 0
            else:
                principal_paid = min(monthly_pmt - interest, balance)
            
            balance -= principal_paid
            balance = max(0, balance)
            
            schedule.append({
                "Period": period,
                "Payment": monthly_pmt if balance > 0 else 0,
                "Principal": principal_paid,
                "Interest": interest,
                "Balance": balance
            })
            
            if balance <= 0:
                break
        
        df = pd.DataFrame(schedule)
        balloon = balance
        
        return df, monthly_pmt, balloon
    
    @staticmethod
    def rent_roll_metrics(df: pd.DataFrame) -> Tuple[float, float, float, float, float, float]:
        """Calculate rent roll metrics"""
        
        if df is None or df.empty:
            return 0, 0, 0, 0, 0, 0
        
        # Normalize the dataframe
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        
        # Map columns
        col_map = {}
        for col in df.columns:
            c = str(col).lower().strip()
            if 'tenant' in c or 'name' in c:
                col_map[col] = 'Tenant'
            elif 'sf' in c or 'sq' in c or 'area' in c:
                col_map[col] = 'SF'
            elif 'term' in c or 'lease' in c:
                col_map[col] = 'Remaining Term'
            elif 'rent' in c or 'revenue' in c:
                col_map[col] = 'Monthly Rent'
        
        df = df.rename(columns=col_map)
        
        # Ensure required columns exist
        for col in ['Tenant', 'SF', 'Remaining Term', 'Monthly Rent']:
            if col not in df.columns:
                df[col] = '' if col == 'Tenant' else 0
        
        # Clean data
        df['SF'] = pd.to_numeric(df['SF'], errors='coerce').fillna(0)
        df['Remaining Term'] = pd.to_numeric(df['Remaining Term'], errors='coerce').fillna(0)
        df['Monthly Rent'] = pd.to_numeric(df['Monthly Rent'], errors='coerce').fillna(0)
        df['Tenant'] = df['Tenant'].fillna('').astype(str).str.strip()
        
        total_sf = df['SF'].sum()
        if total_sf <= 0:
            return 0, 0, 0, 0, 0, 0
        
        # Identify occupied units
        vacant_keywords = ['vacant', 'empty', 'available', 'vacancy', 'n/a', 'none', '']
        occupied = df[~df['Tenant'].str.lower().isin(vacant_keywords) & (df['SF'] > 0)]
        
        occupied_sf = occupied['SF'].sum()
        if occupied_sf <= 0:
            return total_sf, 0, 0, 0, 0, 0
        
        occupancy = occupied_sf / total_sf
        annual_rent = occupied['Monthly Rent'].sum() * 12
        rent_psf = annual_rent / occupied_sf if occupied_sf > 0 else 0
        
        # WALT calculation
        walt = (occupied['Remaining Term'] * occupied['SF']).sum() / occupied_sf
        
        # Rollover exposure
        expiring_1yr_sf = occupied[occupied['Remaining Term'] <= 1.0]['SF'].sum()
        rollover_1yr = expiring_1yr_sf / occupied_sf if occupied_sf > 0 else 0
        
        return total_sf, occupancy, annual_rent, rent_psf, walt, rollover_1yr
    
    @staticmethod
    def score_deal(
        actual_ltv: float, actual_ltc: float, 
        actual_dscr: float, actual_dy: float, 
        profile: str
    ) -> Tuple[int, str]:
        """Score the deal based on underwriting metrics"""
        
        limits = UnderwritingEngine.LENDER_LIMITS.get(
            profile, 
            UnderwritingEngine.LENDER_LIMITS["Bank / Credit Union"]
        )
        
        # Calculate component scores
        ltv_score = 0
        if limits['max_ltv'] > 0:
            ltv_score = max(0, 300 * (1 - actual_ltv / limits['max_ltv']))
        
        dscr_score = 0
        if limits['min_dscr'] > 1.0 and actual_dscr > 1.0:
            dscr_score = max(0, 300 * (actual_dscr - 1.0) / (limits['min_dscr'] - 1.0))
        elif actual_dscr >= 1.0:
            dscr_score = 300
        
        dy_score = 0
        if limits['min_dy'] > 0:
            dy_score = max(0, 200 * min(1.5, actual_dy / limits['min_dy']))
        
        ltc_score = max(0, 200 * (1 - actual_ltc))
        
        total_score = min(1000, int(ltv_score + dscr_score + dy_score + ltc_score))
        
        # Determine tier
        if total_score >= 850:
            tier = DealTier.TIER_1.value
        elif total_score >= 700:
            tier = DealTier.TIER_2.value
        elif total_score >= 550:
            tier = DealTier.TIER_3.value
        else:
            tier = DealTier.TIER_4.value
        
        return total_score, tier

# ==========================================
# SENSITIVITY ANALYSIS ENGINE
# ==========================================

class SensitivityEngine:
    """Advanced sensitivity and stress testing"""
    
    @staticmethod
    def generate_matrix(state: Dict, scenarios: List[str] = None) -> pd.DataFrame:
        """Generate sensitivity matrix"""
        
        if scenarios is None:
            scenarios = ['Base', 'Rate +1%', 'Rate -1%', 'NOI -10%', 'NOI +10%', 'Combined Stress']
        
        results = []
        base_noi = safe_float(state.get('noi'))
        base_rate = safe_float(state.get('rate'))
        
        for scenario in scenarios:
            noi_adj = base_noi
            rate_adj = base_rate
            
            if 'Rate +1%' in scenario:
                rate_adj += 0.01
            elif 'Rate -1%' in scenario:
                rate_adj -= 0.01
            
            if 'NOI -10%' in scenario:
                noi_adj *= 0.90
            elif 'NOI +10%' in scenario:
                noi_adj *= 1.10
            
            if 'Combined' in scenario:
                rate_adj += 0.01
                noi_adj *= 0.90
            
            # Calculate proceeds for this scenario
            loan, gate, _, _, _ = UnderwritingEngine.size_loan(
                noi=noi_adj,
                appraisal=state.get('appraisal'),
                purchase_price=state.get('purchase_price'),
                closing_costs=state.get('closing_costs'),
                reserves=state.get('reserves'),
                fees_pct=state.get('fees'),
                rate=rate_adj,
                amort=state.get('amort'),
                term=state.get('term'),
                is_io=state.get('is_io'),
                target_ltv=state.get('target_ltv'),
                target_ltc=state.get('target_ltc'),
                target_dscr=state.get('target_dscr'),
                target_dy=state.get('target_dy')
            )
            
            results.append({
                'Scenario': scenario,
                'Rate': f"{rate_adj*100:.2f}%",
                'NOI': f"${noi_adj:,.0f}",
                'Max Proceeds': f"${loan:,.0f}",
                'Constraint': gate,
                'Change vs Base': f"${loan - results[0]['Max Proceeds'].replace('$','').replace(',','') if results else 0:,.0f}"
            })
        
        return pd.DataFrame(results)

# ==========================================
# UI COMPONENTS
# ==========================================

class UIComponents:
    """Reusable UI components"""
    
    @staticmethod
    def metric_card(label: str, value: str, delta: str = None, help_text: str = None):
        """Display a styled metric card"""
        st.metric(
            label=label,
            value=value,
            delta=delta,
            help=help_text
        )
    
    @staticmethod
    def validation_banner(errors: List[str], warnings: List[str]):
        """Display validation errors and warnings"""
        if errors:
            for error in errors:
                st.error(f"❌ {error}")
        
        if warnings:
            for warning in warnings:
                st.warning(f"⚠️ {warning}")
    
    @staticmethod
    def deal_header(sponsor: str, property_name: str, constraint: str, tier: str):
        """Display deal header"""
        st.title(f"{sponsor or 'New Deal'} | {property_name or 'Property Address Pending'}")
        st.caption(f"INSTITUTIONAL WORKSTATION | CONSTRAINT: {constraint} | {tier}")
    
    @staticmethod
    def tab_navigation() -> int:
        """Create tab navigation and return selected index"""
        tabs = [
            "📊 Sizing & Risk",
            "🧪 Sensitivity",
            "📝 Rent Roll",
            "📅 Amortization",
            "💾 Save & Export",
            "⚙️ Settings"
        ]
        return st.tabs(tabs)

# ==========================================
# VALIDATION ENGINE
# ==========================================

class ValidationEngine:
    """Comprehensive deal validation"""
    
    @staticmethod
    def validate_deal(state: Dict) -> Tuple[List[str], List[str]]:
        """Validate deal state and return errors and warnings"""
        errors = []
        warnings = []
        
        # Required fields
        if not state.get('deal_name', '').strip():
            errors.append("Deal name is required")
        
        if not state.get('property_address', '').strip():
            errors.append("Property address is required")
        
        # Financial validation
        purchase_price = safe_float(state.get('purchase_price'))
        appraisal = safe_float(state.get('appraisal'))
        noi = safe_float(state.get('noi'))
        
        if purchase_price < 0:
            errors.append("Purchase price cannot be negative")
        
        if appraisal < 0:
            errors.append("Appraisal cannot be negative")
        
        if noi < 0:
            errors.append("NOI cannot be negative")
        
        # Reasonability checks
        if noi > 0 and appraisal > 0:
            cap_rate = noi / appraisal
            if cap_rate < 0.01:
                warnings.append(f"Implied cap rate {cap_rate:.2%} is unusually low")
            elif cap_rate > 0.20:
                warnings.append(f"Implied cap rate {cap_rate:.2%} is unusually high")
        
        if purchase_price > 0 and appraisal > 0:
            premium = (purchase_price - appraisal) / appraisal
            if premium > 0.30:
                warnings.append(f"Purchase price is {premium:.1%} above appraisal")
        
        # Rate validation
        rate = safe_float(state.get('rate'))
        if rate <= 0:
            errors.append("Interest rate must be positive")
        elif rate > 0.30:
            errors.append("Interest rate cannot exceed 30%")
        
        # Amortization validation
        amort = safe_int(state.get('amort'))
        if amort < 1:
            errors.append("Amortization must be at least 1 year")
        elif amort > 40:
            errors.append("Amortization cannot exceed 40 years")
        
        # LTV/DSCR validation
        target_ltv = safe_float(state.get('target_ltv'))
        if target_ltv <= 0:
            errors.append("Target LTV must be positive")
        
        target_dscr = safe_float(state.get('target_dscr'))
        if target_dscr <= 0:
            errors.append("Target DSCR must be positive")
        
        return errors, warnings

# ==========================================
# MAIN APPLICATION
# ==========================================

def main():
    """Main application entry point"""
    
    # Initialize
    initialize_session_state()
    DatabaseManager.init_db()
    
    # Start auto-save
    if 'auto_save_started' not in st.session_state:
        auto_save.start_auto_save()
        st.session_state.auto_save_started = True
    
    # Sidebar
    with st.sidebar:
        render_sidebar()
    
    # Main content
    s = st.session_state
    
    # Execute core calculations
    loan_amt, gate, gates, total_uses, req_equity = UnderwritingEngine.size_loan(
        noi=s.noi,
        appraisal=s.appraisal,
        purchase_price=s.purchase_price,
        closing_costs=s.closing_costs,
        reserves=s.reserves,
        fees_pct=s.fees,
        rate=s.rate,
        amort=s.amort,
        term=s.term,
        is_io=s.is_io,
        target_ltv=s.target_ltv,
        target_ltc=s.target_ltc,
        target_dscr=s.target_dscr,
        target_dy=s.target_dy
    )
    
    amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(
        loan_amt, s.rate, s.amort, s.term, s.is_io
    )
    
    annual_ds = monthly_pmt * 12
    
    # Calculate actual metrics
    actual_ltv = loan_amt / s.appraisal if safe_float(s.appraisal) > 0 else 0
    actual_ltc = loan_amt / total_uses if total_uses > 0 else 0
    actual_dscr = s.noi / annual_ds if annual_ds > 0 else 0
    actual_dy = s.noi / loan_amt if loan_amt > 0 else 0
    
    # Score deal
    score, tier = UnderwritingEngine.score_deal(
        actual_ltv, actual_ltc, actual_dscr, actual_dy, s.lender_profile
    )
    
    # Validate
    errors, warnings = ValidationEngine.validate_deal(extract_clean_state())
    
    # Display header
    UIComponents.deal_header(
        sponsor=s.sponsor,
        property_name=s.property_address,
        constraint=gate,
        tier=tier
    )
    
    # Display validation
    UIComponents.validation_banner(errors, warnings)
    
    # KPI Dashboard
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        UIComponents.metric_card("MAX PROCEEDS", f"${loan_amt:,.0f}")
    with col2:
        UIComponents.metric_card("ACTUAL LTV", f"{actual_ltv*100:.1f}%")
    with col3:
        UIComponents.metric_card("ACTUAL LTC", f"{actual_ltc*100:.1f}%")
    with col4:
        UIComponents.metric_card("ACTUAL DSCR", f"{actual_dscr:.2f}x")
    with col5:
        UIComponents.metric_card("BALLOON", f"${balloon:,.0f}")
    with col6:
        UIComponents.metric_card("DEAL SCORE", f"{score}/1000", 
                                help_text=f"{tier}")
    
    st.markdown("---")
    
    # Main tabs
    tabs = st.tabs([
        "📊 Sizing & Risk",
        "🧪 Sensitivity",
        "📝 Rent Roll",
        "📅 Amortization",
        "💾 Save & Export",
        "⚙️ Settings & Health"
    ])
    
    # Tab 1: Sizing & Risk
    with tabs[0]:
        render_sizing_tab(
            gates=gates,
            gate=gate,
            total_uses=total_uses,
            loan_amt=loan_amt,
            req_equity=req_equity,
            actual_ltv=actual_ltv,
            actual_dscr=actual_dscr
        )
    
    # Tab 2: Sensitivity
    with tabs[1]:
        render_sensitivity_tab()
    
    # Tab 3: Rent Roll
    with tabs[2]:
        render_rent_roll_tab()
    
    # Tab 4: Amortization
    with tabs[3]:
        render_amortization_tab(amort_df, monthly_pmt, annual_ds, balloon)
    
    # Tab 5: Save & Export
    with tabs[4]:
        render_export_tab(
            loan_amt=loan_amt,
            gate=gate,
            amort_df=amort_df,
            score=score,
            tier=tier
        )
    
    # Tab 6: Settings & Health
    with tabs[5]:
        render_settings_tab()
    
    # Footer
    st.markdown("---")
    st.caption(
        "⚠️ **DISCLAIMER:** ALENZA CAPITAL OS is an indicative modeling tool. "
        "Outputs do not constitute a loan commitment, appraisal, or legal advice. "
        "Final terms are subject to formal credit committee approval and third-party diligence verification."
    )

# ==========================================
# TAB RENDERERS
# ==========================================

def render_sidebar():
    """Render sidebar with deal management"""
    st.title("🏛️ ALENZA OS")
    
    # Deal management
    with st.expander("📁 DEAL MANAGER", expanded=True):
        # New deal
        new_name = st.text_input("New Deal Name", value="Untitled Deal")
        if st.button("➕ New Deal", use_container_width=True):
            new_id = generate_id('deal')
            st.session_state.deal_id = new_id
            st.session_state.deal_name = new_name
            st.session_state.unsaved_changes = True
            st.rerun()
        
        # Load existing deal
        st.markdown("---")
        deals_df = DatabaseManager.get_all_deals()
        if not deals_df.empty:
            deal_options = {
                f"{row['name']} ({row['id'][-8:]})": row['id']
                for _, row in deals_df.iterrows()
            }
            selected = st.selectbox("Load Deal", list(deal_options.keys()))
            if st.button("📂 Load", use_container_width=True):
                deal_id = deal_options[selected]
                state = DatabaseManager.load_deal(deal_id)
                if state:
                    for k, v in state.items():
                        st.session_state[k] = v
                    st.session_state.unsaved_changes = False
                    st.rerun()
        else:
            st.info("No saved deals")
    
    # Quick save
    if st.button("💾 Quick Save", use_container_width=True):
        state = extract_clean_state()
        errors, _ = ValidationEngine.validate_deal(state)
        if errors:
            st.error("Fix errors before saving")
        else:
            DatabaseManager.save_deal(
                state['deal_id'],
                state.get('deal_name', 'Untitled'),
                state
            )
            st.session_state.unsaved_changes = False
            st.success("Saved!")
    
    # Auto-save status
    if st.session_state.unsaved_changes:
        st.warning("⚠️ Unsaved changes")
    
    # Deal parameters
    st.markdown("---")
    
    with st.expander("🏢 ASSET PROFILE", expanded=True):
        s = st.session_state
        s.deal_name = st.text_input("Deal Name", value=s.deal_name, on_change=on_state_change)
        s.sponsor = st.text_input("Sponsor", value=s.sponsor, on_change=on_state_change)
        s.property_address = st.text_input("Address", value=s.property_address, on_change=on_state_change)
        
        property_types = [pt.value for pt in PropertyType]
        current_idx = property_types.index(s.property_type) if s.property_type in property_types else 0
        s.property_type = st.selectbox("Type", property_types, index=current_idx, on_change=on_state_change)
        
        s.appraisal = st.number_input(
            "Appraisal ($)", 
            value=safe_float(s.appraisal),
            step=100000.0, min_value=0.0,
            on_change=on_state_change
        )
        s.purchase_price = st.number_input(
            "Cost Basis ($)",
            value=safe_float(s.purchase_price),
            step=100000.0, min_value=0.0,
            on_change=on_state_change
        )
        s.noi = st.number_input(
            "Stabilized NOI ($)",
            value=safe_float(s.noi),
            step=10000.0, min_value=0.0,
            on_change=on_state_change
        )
    
    with st.expander("📊 CREDIT POLICY", expanded=True):
        s = st.session_state
        profiles = list(UnderwritingEngine.LENDER_LIMITS.keys())
        current_profile = profiles.index(s.lender_profile) if s.lender_profile in profiles else 0
        s.lender_profile = st.selectbox("Policy Preset", profiles, index=current_profile, on_change=on_state_change)
        
        limits = UnderwritingEngine.LENDER_LIMITS[s.lender_profile]
        s.target_ltv = st.slider(
            "Max LTV %",
            50.0, 95.0,
            float(normalize_percent(s.target_ltv, limits['max_ltv']) * 100),
            step=0.5,
            on_change=on_state_change
        ) / 100
        s.target_dscr = st.slider(
            "Min DSCR",
            1.0, 1.75,
            float(safe_float(s.target_dscr, limits['min_dscr'])),
            step=0.05,
            on_change=on_state_change
        )
        s.target_dy = st.slider(
            "Min DY %",
            5.0, 15.0,
            float(normalize_percent(s.target_dy, limits['min_dy']) * 100),
            step=0.25,
            on_change=on_state_change
        ) / 100
    
    with st.expander("💰 DEBT STRUCTURE", expanded=True):
        s = st.session_state
        s.is_io = st.checkbox("Interest-Only", value=bool(s.is_io), on_change=on_state_change)
        s.rate = st.slider(
            "Rate %",
            0.0, 15.0,
            float(normalize_percent(s.rate, 0.0525) * 100),
            step=0.05,
            on_change=on_state_change
        ) / 100
        s.amort = st.number_input(
            "Amort (Yrs)",
            value=max(1, safe_int(s.amort, 25)),
            step=1, min_value=1, max_value=40,
            on_change=on_state_change
        )
        s.term = st.number_input(
            "Term (Yrs)",
            value=max(1, safe_int(s.term, 5)),
            step=1, min_value=1, max_value=40,
            on_change=on_state_change
        )
        s.fees = st.slider(
            "Fees %",
            0.0, 5.0,
            float(normalize_percent(s.fees, 0.02) * 100),
            step=0.05,
            on_change=on_state_change
        ) / 100

def render_sizing_tab(gates, gate, total_uses, loan_amt, req_equity, actual_ltv, actual_dscr):
    """Render sizing and risk tab"""
    s = st.session_state
    
    col1, col2 = st.columns([1.5, 1])
    
    with col1:
        st.subheader("Constraint Analysis")
        df_gates = pd.DataFrame({
            "Constraint": list(gates.keys()),
            "Threshold": [
                f"{s.target_ltv*100:.1f}%",
                f"{s.target_ltc*100:.1f}%",
                f"{s.target_dscr:.2f}x",
                f"{s.target_dy*100:.2f}%"
            ],
            "Proceeds Limit": [f"${v:,.0f}" for v in gates.values()],
            "Binding": ["✅" if g == gate else "" for g in gates.keys()]
        })
        st.dataframe(df_gates, hide_index=True, use_container_width=True)
        
        st.subheader("Sources & Uses")
        total_fees = loan_amt * s.fees
        su_data = {
            "Uses": ["Cost Basis", "Closing Costs", "Reserves", "Financing Fees", "Total"],
            "Amount": [
                s.purchase_price, s.closing_costs, s.reserves,
                total_fees, total_uses
            ],
            "% of Total": [
                f"{s.purchase_price/total_uses*100:.1f}%" if total_uses > 0 else "0%",
                f"{s.closing_costs/total_uses*100:.1f}%" if total_uses > 0 else "0%",
                f"{s.reserves/total_uses*100:.1f}%" if total_uses > 0 else "0%",
                f"{total_fees/total_uses*100:.1f}%" if total_uses > 0 else "0%",
                "100%"
            ]
        }
        df_su = pd.DataFrame(su_data)
        st.dataframe(
            df_su.style.format({"Amount": "${:,.0f}"}),
            hide_index=True,
            use_container_width=True
        )
    
    with col2:
        st.subheader("Risk Assessment")
        
        # Generate risk flags
        flags = []
        if actual_ltv > 0.75:
            flags.append(f"⚠️ High Leverage: {actual_ltv*100:.1f}% LTV")
        elif actual_ltv < 0.55:
            flags.append("✅ Conservative Capital Structure")
        
        if actual_dscr < 1.20:
            flags.append(f"⚠️ Tight Coverage: {actual_dscr:.2f}x DSCR")
        elif actual_dscr > 1.50:
            flags.append("✅ Strong Debt Service Coverage")
        
        if s.is_io:
            flags.append("ℹ️ Interest-Only Structure")
        
        if req_equity < 0:
            flags.append("🚨 Negative Equity Required")
        
        if not flags:
            flags.append("✅ No Major Risk Flags")
        
        for flag in flags:
            if "🚨" in flag:
                st.error(flag)
            elif "⚠️" in flag:
                st.warning(flag)
            else:
                st.info(flag)

def render_sensitivity_tab():
    """Render sensitivity analysis tab"""
    st.subheader("Sensitivity Analysis")
    
    sensitivity_df = SensitivityEngine.generate_matrix(extract_clean_state())
    st.dataframe(sensitivity_df, hide_index=True, use_container_width=True)
    
    # Additional stress scenarios
    st.subheader("Custom Stress Test")
    col1, col2 = st.columns(2)
    with col1:
        rate_shock = st.slider("Rate Shock (bps)", -200, 200, 0, 25) / 10000
    with col2:
        noi_shock = st.slider("NOI Shock (%)", -30, 30, 0, 5) / 100
    
    s = st.session_state
    shocked_loan, shocked_gate, _, _, _ = UnderwritingEngine.size_loan(
        noi=s.noi * (1 + noi_shock),
        appraisal=s.appraisal,
        purchase_price=s.purchase_price,
        closing_costs=s.closing_costs,
        reserves=s.reserves,
        fees_pct=s.fees,
        rate=s.rate + rate_shock,
        amort=s.amort,
        term=s.term,
        is_io=s.is_io,
        target_ltv=s.target_ltv,
        target_ltc=s.target_ltc,
        target_dscr=s.target_dscr,
        target_dy=s.target_dy
    )
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Stressed Proceeds", f"${shocked_loan:,.0f}")
    col2.metric("Binding Constraint", shocked_gate)
    col3.metric("Proceeds Change", f"${shocked_loan - (loan_amt if 'loan_amt' in dir() else 0):,.0f}")

def render_rent_roll_tab():
    """Render rent roll editor"""
    st.subheader("Rent Roll Management")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Import Rent Roll (CSV, Excel)",
        type=['csv', 'xlsx', 'xls'],
        key='rent_roll_upload'
    )
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.write("Preview:")
            st.dataframe(df.head(), use_container_width=True)
            
            if st.button("Apply Imported Data"):
                # Convert to standard format
                normalized = []
                for _, row in df.iterrows():
                    normalized.append({
                        'Tenant': str(row.get('Tenant', row.get('tenant', ''))),
                        'SF': safe_float(row.get('SF', row.get('sf', 0))),
                        'Remaining Term': safe_float(row.get('Remaining Term', row.get('term', 0))),
                        'Monthly Rent': safe_float(row.get('Monthly Rent', row.get('rent', 0)))
                    })
                st.session_state.rent_roll_dict = normalized
                st.session_state.unsaved_changes = True
                st.rerun()
        except Exception as e:
            st.error(f"Import failed: {e}")
    
    # Interactive editor
    if st.session_state.rent_roll_dict:
        df = pd.DataFrame(st.session_state.rent_roll_dict)
    else:
        df = pd.DataFrame(columns=['Tenant', 'SF', 'Remaining Term', 'Monthly Rent'])
    
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tenant": st.column_config.TextColumn("Tenant", required=True),
            "SF": st.column_config.NumberColumn("SF", min_value=0, step=100),
            "Remaining Term": st.column_config.NumberColumn("Remaining Term (Yrs)", min_value=0.0, step=0.5),
            "Monthly Rent": st.column_config.NumberColumn("Monthly Rent ($)", min_value=0.0, step=100.0, format="$%.2f")
        }
    )
    
    if st.button("Update Rent Roll"):
        st.session_state.rent_roll_dict = edited_df.to_dict('records')
        st.session_state.unsaved_changes = True
        st.rerun()
    
    # Metrics
    if not edited_df.empty:
        total_sf, occ, annual_rent, rent_psf, walt, rollover = UnderwritingEngine.rent_roll_metrics(edited_df)
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("Total SF", f"{total_sf:,.0f}")
        col2.metric("Occupancy", f"{occ*100:.1f}%")
        col3.metric("Annual Rent", f"${annual_rent:,.0f}")
        col4.metric("Rent PSF", f"${rent_psf:.2f}")
        col5.metric("WALT", f"{walt:.2f} yrs")
        col6.metric("1-Yr Rollover", f"{rollover*100:.1f}%")

def render_amortization_tab(amort_df, monthly_pmt, annual_ds, balloon):
    """Render amortization schedule"""
    st.subheader("Amortization Schedule")
    
    if amort_df is None or amort_df.empty:
        st.warning("No amortization data available")
        return
    
    # Summary metrics
    col1, col2, col3 = st.columns(3)
    col1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
    col2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
    col3.metric("Balloon Balance", f"${balloon:,.0f}")
    
    # Charts
    st.subheader("Payment Breakdown")
    chart_data = amort_df.set_index('Period')[['Principal', 'Interest']]
    st.bar_chart(chart_data)
    
    st.subheader("Loan Balance")
    balance_chart = amort_df.set_index('Period')[['Balance']]
    st.line_chart(balance_chart)
    
    # Full schedule
    st.subheader("Full Schedule")
    st.dataframe(
        amort_df.style.format({
            "Payment": "${:,.2f}",
            "Principal": "${:,.2f}",
            "Interest": "${:,.2f}",
            "Balance": "${:,.2f}"
        }),
        use_container_width=True,
        height=300,
        hide_index=True
    )

def render_export_tab(loan_amt, gate, amort_df, score, tier):
    """Render export and save options"""
    st.subheader("Save & Export")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("### Save Deal")
        if st.button("💾 Save to Database", use_container_width=True):
            state = extract_clean_state()
            errors, warnings = ValidationEngine.validate_deal(state)
            
            if errors:
                st.error("Cannot save: Fix validation errors first")
                for error in errors:
                    st.error(f"• {error}")
            else:
                if warnings:
                    st.warning("Saving with warnings:")
                    for warning in warnings:
                        st.warning(f"• {warning}")
                
                success = DatabaseManager.save_deal(
                    state['deal_id'],
                    state.get('deal_name', 'Untitled'),
                    state
                )
                
                if success:
                    st.session_state.unsaved_changes = False
                    st.session_state.last_save_time = datetime.now().isoformat()
                    st.success(f"Deal saved at {st.session_state.last_save_time}")
                else:
                    st.error("Save failed")
    
    with col2:
        st.write("### Export Options")
        
        # Excel export
        if DEPENDENCIES['excel_write']:
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Summary sheet
                    pd.DataFrame({
                        'Metric': ['Sponsor', 'Property', 'Max Proceeds', 'Constraint', 'Score', 'Tier'],
                        'Value': [
                            st.session_state.sponsor,
                            st.session_state.property_address,
                            f"${loan_amt:,.0f}",
                            gate,
                            score,
                            tier
                        ]
                    }).to_excel(writer, sheet_name='Summary', index=False)
                    
                    # Rent roll
                    if st.session_state.rent_roll_dict:
                        pd.DataFrame(st.session_state.rent_roll_dict).to_excel(
                            writer, sheet_name='Rent Roll', index=False
                        )
                    
                    # Amortization
                    if amort_df is not None and not amort_df.empty:
                        amort_df.to_excel(writer, sheet_name='Amortization', index=False)
                
                st.download_button(
                    "📊 Download Excel Model",
                    data=output.getvalue(),
                    file_name=f"alenza_model_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Excel export failed: {e}")
        else:
            st.warning("Excel export requires xlsxwriter")
        
        # JSON export
        state_json = json.dumps(extract_clean_state(), indent=2, default=str)
        st.download_button(
            "📄 Download Deal JSON",
            data=state_json,
            file_name=f"alenza_deal_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True
        )

def render_settings_tab():
    """Render settings and health tab"""
    st.subheader("Settings & System Health")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("### User Preferences")
        
        dark_mode = st.toggle("Dark Mode", value=True)
        auto_save = st.toggle("Auto-Save", value=True)
        show_advanced = st.toggle("Show Advanced Options", value=False)
        
        if st.button("Save Preferences"):
            st.session_state.user_preferences = {
                'dark_mode': dark_mode,
                'auto_save': auto_save,
                'show_advanced': show_advanced
            }
            st.success("Preferences saved")
    
    with col2:
        st.write("### System Information")
        
        # Version info
        st.info(f"Alenza OS Version: {APP_VERSION}")
        
        # Database stats
        try:
            deals_df = DatabaseManager.get_all_deals()
            st.info(f"Saved Deals: {len(deals_df) if not deals_df.empty else 0}")
        except:
            st.error("Database connection failed")
        
        # Dependencies
        st.write("#### Available Features")
        for dep, available in DEPENDENCIES.items():
            icon = "✅" if available else "❌"
            st.write(f"{icon} {dep}")
    
    # Health checks
    st.markdown("---")
    st.subheader("System Health")
    
    if st.button("Run Health Check"):
        with st.spinner("Checking system health..."):
            health_results = {
                "Database": "✅ Connected" if check_database_health() else "❌ Failed",
                "API Access": "✅ Available" if check_api_health() else "⚠️ Limited",
                "File System": "✅ Writable" if check_filesystem() else "❌ Read-only",
                "Memory": f"✅ {check_memory_usage():.1f}% used"
            }
            
            for component, status in health_results.items():
                st.write(f"{component}: {status}")

def check_database_health() -> bool:
    """Check if database is accessible"""
    try:
        deals_df = DatabaseManager.get_all_deals()
        return True
    except:
        return False

def check_api_health() -> bool:
    """Check if external APIs are reachable"""
    try:
        response = requests.get("https://www.bankofcanada.ca", timeout=5)
        return response.status_code == 200
    except:
        return False

def check_filesystem() -> bool:
    """Check if filesystem is writable"""
    try:
        test_file = Path("alenza_data/test_write.tmp")
        test_file.write_text("test")
        test_file.unlink()
        return True
    except:
        return False

def check_memory_usage() -> float:
    """Get memory usage percentage"""
    try:
        import psutil
        return psutil.virtual_memory().percent
    except:
        return 0.0

# ==========================================
# APPLICATION ENTRY POINT
# ==========================================

if __name__ == "__main__":
    main()
