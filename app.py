"""
Alenza Capital OS v3.0.1
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
from functools import wraps
from enum import Enum

# ==========================================
# LOGGING SETUP
# ==========================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ==========================================
# PAGE CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="Alenza Capital OS",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# CONSTANTS
# ==========================================

APP_VERSION = "3.0.1"
SCHEMA_VERSION = 4
MAX_UPLOAD_SIZE_MB = 50

# ==========================================
# DEPENDENCY CHECKING
# ==========================================

@st.cache_resource
def check_dependencies() -> Dict[str, bool]:
    """Check available dependencies"""
    deps = {
        "ocr": False,
        "pdf": False,
        "excel_write": False,
        "excel_read": False,
        "plotly": False,
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
    
    return deps

DEPENDENCIES = check_dependencies()

# Import optional dependencies based on availability
if DEPENDENCIES["ocr"]:
    from PIL import Image
    import pytesseract
    import fitz

if DEPENDENCIES["pdf"]:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

if DEPENDENCIES["excel_write"]:
    import xlsxwriter

# ==========================================
# UTILITY FUNCTIONS
# ==========================================

def safe_float(value: Any, default: float = 0.0) -> float:
    """Safely convert value to float"""
    if value is None or value == "":
        return default
    try:
        if isinstance(value, (int, float, np.number)):
            return float(value)
        if isinstance(value, str):
            # Remove currency symbols, commas, percentages
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
    timestamp = int(datetime.now().timestamp())
    random_hex = uuid.uuid4().hex[:8]
    return f"{prefix}_{timestamp}_{random_hex}"

def clean_currency_string(value: str) -> float:
    """Convert formatted currency string to float"""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        return safe_float(value.replace('$', '').replace(',', '').strip())
    return 0.0

def hash_state(state: Dict) -> str:
    """Create SHA-256 hash of state for change detection"""
    state_str = json.dumps(state, sort_keys=True, default=str)
    return hashlib.sha256(state_str.encode()).hexdigest()

def sanitize_filename(filename: str) -> str:
    """Sanitize filename to prevent path traversal"""
    # Get name and extension
    path = Path(str(filename))
    name = path.stem or "file"
    ext = path.suffix[:10]  # Limit extension length
    
    # Remove potentially dangerous characters
    name = re.sub(r'[^a-zA-Z0-9_.-]', '_', name)
    name = name[:100]  # Limit length
    
    # Add unique suffix to prevent collisions
    unique_id = uuid.uuid4().hex[:8]
    
    return f"{name}_{unique_id}{ext}"

def get_current_user() -> str:
    """Get current user from environment or session"""
    try:
        # Try Streamlit secrets first
        return str(st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User")))
    except:
        return os.environ.get("APP_USER", "Local User")

def summarize_changes(old_state: dict, new_state: dict) -> str:
    """Create human-readable change summary for audit log"""
    tracked_fields = [
        'deal_name', 'sponsor', 'property_address', 'property_type',
        'purchase_price', 'appraisal', 'noi', 'rate', 'amort', 'term',
        'target_ltv', 'target_ltc', 'target_dscr', 'target_dy'
    ]
    
    changes = []
    for field in tracked_fields:
        old_val = old_state.get(field) if old_state else None
        new_val = new_state.get(field) if new_state else None
        if str(old_val) != str(new_val):
            changes.append(f"{field}: {old_val} → {new_val}")
    
    # Check rent roll changes
    old_rr_len = len(old_state.get('rent_roll_dict', [])) if old_state else 0
    new_rr_len = len(new_state.get('rent_roll_dict', [])) if new_state else 0
    if old_rr_len != new_rr_len:
        changes.append(f"rent_roll_rows: {old_rr_len} → {new_rr_len}")
    
    return "; ".join(changes[:20]) if changes else "No material changes"

# ==========================================
# ENCRYPTION UTILITIES
# ==========================================

def encrypt_deal_state(state: dict, secret_key: str) -> str:
    """
    Encrypt deal state for secure export.
    Uses AES-128 via Fernet (cryptography library).
    
    Args:
        state: Deal state dictionary to encrypt
        secret_key: Password/passphrase for encryption
        
    Returns:
        Base64-encoded encrypted string, or plain JSON if encryption unavailable
    """
    try:
        import base64
        from cryptography.fernet import Fernet
        
        # Derive a 32-byte key from the password using SHA-256
        key = base64.urlsafe_b64encode(
            hashlib.sha256(secret_key.encode()).digest()
        )
        fernet = Fernet(key)
        
        # Convert state to JSON and encrypt
        state_json = json.dumps(state, default=str)
        encrypted = fernet.encrypt(state_json.encode())
        
        # Return as base64 for safe storage/transmission
        return base64.b64encode(encrypted).decode()
        
    except ImportError:
        logger.warning("cryptography package not installed. Exporting without encryption.")
        return json.dumps(state, indent=2, default=str)
    except Exception as e:
        logger.error(f"Encryption failed: {e}")
        st.warning(f"Encryption unavailable: {e}")
        return json.dumps(state, indent=2, default=str)

def decrypt_deal_state(encrypted_state: str, secret_key: str) -> dict:
    """
    Decrypt an encrypted deal state.
    
    Args:
        encrypted_state: Base64-encoded encrypted string
        secret_key: Password/passphrase for decryption
        
    Returns:
        Decrypted deal state dictionary, or empty dict if decryption fails
    """
    try:
        import base64
        from cryptography.fernet import Fernet
        
        # Derive the same key from the password
        key = base64.urlsafe_b64encode(
            hashlib.sha256(secret_key.encode()).digest()
        )
        fernet = Fernet(key)
        
        # Decode from base64 and decrypt
        encrypted = base64.b64decode(encrypted_state)
        state_json = fernet.decrypt(encrypted)
        
        return json.loads(state_json)
        
    except ImportError:
        logger.warning("cryptography package not installed. Cannot decrypt.")
        st.error("Encryption package not available. Install with: pip install cryptography")
        return {}
    except Exception as e:
        logger.error(f"Decryption failed: {e}")
        st.error(f"Decryption failed. Wrong password or corrupted data.")
        return {}

def is_encrypted_state(data: str) -> bool:
    """
    Check if a string is an encrypted deal state.
    Encrypted data is base64 encoded Fernet output (starts with gAAAAA).
    """
    try:
        import base64
        decoded = base64.b64decode(data)
        # Fernet tokens start with version byte 0x80
        return len(decoded) > 0 and decoded[0] == 0x80
    except:
        return False

def hash_file_content(content: bytes) -> str:
    """Calculate SHA-256 hash of file content for integrity verification"""
    return hashlib.sha256(content).hexdigest()

# ==========================================
# ERROR DECORATOR
# ==========================================

def handle_errors(default_return=None, show_error=True):
    """Decorator for consistent error handling"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                logger.error(f"Error in {func.__name__}: {e}", exc_info=True)
                if show_error:
                    st.error(f"Operation failed: {str(e)[:200]}")
                return default_return
        return wrapper
    return decorator

# ==========================================
# DATABASE MANAGER
# ==========================================

class DatabaseManager:
    """Database operations manager"""
    
    @staticmethod
    @handle_errors(default_return=False, show_error=False)
    def init_db():
        """Initialize database with proper schema"""
        os.makedirs('alenza_data', exist_ok=True)
        os.makedirs('alenza_data/documents', exist_ok=True)
        
        with sqlite3.connect('alenza_data/alenza_platform.db', timeout=30) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA foreign_keys=ON")
            
            # Create tables if they don't exist
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS deals (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    state_json TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS deal_versions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    deal_id TEXT,
                    state_json TEXT,
                    changed_by TEXT,
                    change_summary TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (deal_id) REFERENCES deals(id) ON DELETE CASCADE
                );
                
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY,
                    deal_id TEXT,
                    filename TEXT NOT NULL,
                    category TEXT NOT NULL,
                    path TEXT NOT NULL,
                    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (deal_id) REFERENCES deals(id) ON DELETE CASCADE
                );
                
                CREATE INDEX IF NOT EXISTS idx_deals_updated ON deals(updated_at DESC);
                CREATE INDEX IF NOT EXISTS idx_versions_deal ON deal_versions(deal_id);
                CREATE INDEX IF NOT EXISTS idx_docs_deal ON documents(deal_id);
            """)
            
            # Check if created_at column exists, add if not
            cursor = conn.execute("PRAGMA table_info(deals)")
            columns = [row[1] for row in cursor.fetchall()]
            
            if 'created_at' not in columns:
                try:
                    conn.execute("ALTER TABLE deals ADD COLUMN created_at TIMESTAMP")
                    conn.execute("UPDATE deals SET created_at = updated_at WHERE created_at IS NULL")
                    conn.commit()
                    logger.info("Added created_at column to deals table")
                except Exception as e:
                    logger.warning(f"Could not add created_at column: {e}")
            
            conn.commit()
    
    @staticmethod
    @handle_errors(default_return=False)
    def save_deal(deal_id: str, name: str, state: dict) -> bool:
        """Save deal to database"""
        state_json = json.dumps(state, default=str)
        safe_name = str(name or "Untitled Deal").strip()[:200]
        now = datetime.now().isoformat()
        
        with sqlite3.connect('alenza_data/alenza_platform.db', timeout=30) as conn:
            # Get old state for change summary
            old_state = {}
            row = conn.execute("SELECT state_json FROM deals WHERE id = ?", (deal_id,)).fetchone()
            if row:
                try:
                    old_state = json.loads(row['state_json'])
                except:
                    pass
            
            # Save deal
            conn.execute("""
                INSERT OR REPLACE INTO deals (id, name, state_json, updated_at, created_at)
                VALUES (?, ?, ?, ?, COALESCE((SELECT created_at FROM deals WHERE id = ?), ?))
            """, (deal_id, safe_name, state_json, now, deal_id, now))
            
            # Save version
            change_summary = summarize_changes(old_state, state)
            conn.execute("""
                INSERT INTO deal_versions (deal_id, state_json, changed_by, change_summary)
                VALUES (?, ?, ?, ?)
            """, (deal_id, state_json, get_current_user(), change_summary))
            
            conn.commit()
            logger.info(f"Deal saved: {safe_name}")
            return True
    
    @staticmethod
    @handle_errors(default_return=None)
    def load_deal(deal_id: str) -> Optional[dict]:
        """Load deal from database"""
        with sqlite3.connect('alenza_data/alenza_platform.db', timeout=30) as conn:
            row = conn.execute("SELECT state_json FROM deals WHERE id = ?", (deal_id,)).fetchone()
            if row:
                return json.loads(row['state_json'])
        return None
    
    @staticmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_all_deals() -> pd.DataFrame:
        """Get all deals"""
        with sqlite3.connect('alenza_data/alenza_platform.db', timeout=30) as conn:
            # Check which columns exist
            cursor = conn.execute("PRAGMA table_info(deals)")
            columns = [row[1] for row in cursor.fetchall()]
            
            if 'created_at' in columns:
                query = """
                    SELECT id, name, created_at, updated_at 
                    FROM deals 
                    ORDER BY updated_at DESC
                """
            else:
                query = """
                    SELECT id, name, updated_at as created_at, updated_at 
                    FROM deals 
                    ORDER BY updated_at DESC
                """
            
            return pd.read_sql_query(query, conn)
    
    @staticmethod
    @handle_errors(default_return=False)
    def delete_deal(deal_id: str) -> bool:
        """Delete deal and associated data"""
        with sqlite3.connect('alenza_data/alenza_platform.db', timeout=30) as conn:
            # Delete documents first
            cursor = conn.execute("SELECT path FROM documents WHERE deal_id = ?", (deal_id,))
            for row in cursor.fetchall():
                try:
                    Path(row['path']).unlink(missing_ok=True)
                except:
                    pass
            
            conn.execute("DELETE FROM documents WHERE deal_id = ?", (deal_id,))
            conn.execute("DELETE FROM deal_versions WHERE deal_id = ?", (deal_id,))
            conn.execute("DELETE FROM deals WHERE id = ?", (deal_id,))
            conn.commit()
            return True


def get_current_user() -> str:
    """Get current user"""
    try:
        return str(st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User")))
    except:
        return "Local User"

def summarize_changes(old_state: dict, new_state: dict) -> str:
    """Create human-readable change summary"""
    tracked_fields = [
        'deal_name', 'purchase_price', 'appraisal', 'noi',
        'rate', 'amort', 'term', 'target_ltv', 'target_dscr'
    ]
    
    changes = []
    for field in tracked_fields:
        old_val = old_state.get(field) if old_state else None
        new_val = new_state.get(field) if new_state else None
        if str(old_val) != str(new_val):
            changes.append(f"{field}: {old_val} → {new_val}")
    
    return "; ".join(changes) if changes else "No material changes"

# ==========================================
# SESSION STATE MANAGEMENT
# ==========================================

def initialize_session_state():
    """Initialize all session state variables"""
    defaults = {
        'deal_id': generate_id('deal'),
        'deal_name': 'Untitled Deal',
        'sponsor': '',
        'property_address': '',
        'property_type': 'Multifamily',
        'transaction_type': 'Acquisition',
        'lender_profile': 'Bank / Credit Union',
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
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def extract_clean_state() -> dict:
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
            if isinstance(val, np.integer):
                val = int(val)
            elif isinstance(val, np.floating):
                val = float(val)
            elif isinstance(val, np.ndarray):
                val = val.tolist()
            state[key] = val
    
    state['schema_version'] = SCHEMA_VERSION
    state['app_version'] = APP_VERSION
    return state


# ==========================================
# FINANCIAL ENGINE
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
    def size_loan(
        noi, appraisal, purchase_price, closing_costs, reserves,
        fees_pct, rate, amort, term, is_io,
        target_ltv, target_ltc, target_dscr, target_dy
    ) -> tuple:
        """Calculate maximum supportable loan amount"""
        
        # Sanitize inputs
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
        
        # Iterative convergence
        for _ in range(10):
            total_uses = base_cost + (loan * fees_pct)
            
            # LTV constraint
            if appraisal > 0:
                gates["LTV"] = appraisal * target_ltv
            
            # LTC constraint
            if total_uses > 0:
                gates["LTC"] = total_uses * target_ltc
            
            # Debt yield constraint
            if target_dy > 0:
                gates["Debt Yield"] = noi / target_dy
            
            # DSCR constraint
            monthly_rate = rate / 12
            if monthly_rate > 0:
                if is_io:
                    gates["DSCR"] = (noi / target_dscr) / 12 / monthly_rate
                else:
                    total_payments = amort * 12
                    pmt_factor = (1 - (1 + monthly_rate) ** -total_payments) / monthly_rate
                    if pmt_factor > 0:
                        gates["DSCR"] = (noi / target_dscr) / 12 * pmt_factor
                    else:
                        gates["DSCR"] = float('inf')
            else:
                gates["DSCR"] = float('inf')
            
            new_loan = min(gates.values())
            
            # Check convergence
            if abs(new_loan - loan) < 1.0:
                break
            
            loan = max(0, new_loan)
        
        gate = min(gates, key=gates.get) if gates else "N/A"
        req_equity = total_uses - loan
        
        return loan, gate, gates, total_uses, req_equity
    
    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io) -> tuple:
        """Generate amortization schedule"""
        loan_amt = max(0, safe_float(loan_amt))
        rate = max(0, min(0.30, safe_float(rate)))
        amort_yrs = max(1, min(40, safe_int(amort_yrs)))
        term_yrs = max(1, min(40, safe_int(term_yrs)))
        
        if loan_amt <= 0:
            return pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]), 0.0, 0.0
        
        monthly_rate = rate / 12
        total_payments = amort_yrs * 12
        term_months = int(term_yrs * 12)
        
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
            
            balance = max(0, balance - principal_paid)
            
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
    def rent_roll_metrics(df) -> tuple:
        """Calculate rent roll metrics"""
        if df is None or df.empty:
            return 0, 0, 0, 0, 0, 0
        
        df = df.copy()
        
        # Ensure required columns
        for col in ['Tenant', 'SF', 'Remaining Term', 'Monthly Rent']:
            if col not in df.columns:
                df[col] = '' if col == 'Tenant' else 0
        
        # Clean data
        df['SF'] = pd.to_numeric(df['SF'], errors='coerce').fillna(0).clip(lower=0)
        df['Remaining Term'] = pd.to_numeric(df['Remaining Term'], errors='coerce').fillna(0).clip(lower=0)
        df['Monthly Rent'] = pd.to_numeric(df['Monthly Rent'], errors='coerce').fillna(0).clip(lower=0)
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
        
        occupancy = min(1.0, occupied_sf / total_sf)
        annual_rent = occupied['Monthly Rent'].sum() * 12
        rent_psf = annual_rent / occupied_sf if occupied_sf > 0 else 0
        
        # WALT
        walt = (occupied['Remaining Term'] * occupied['SF']).sum() / occupied_sf if occupied_sf > 0 else 0
        
        # Rollover
        expiring_sf = occupied[occupied['Remaining Term'] <= 1.0]['SF'].sum()
        rollover = min(1.0, expiring_sf / occupied_sf) if occupied_sf > 0 else 0
        
        return total_sf, occupancy, annual_rent, rent_psf, walt, rollover
    
    @staticmethod
    def score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, profile) -> tuple:
        """Score deal based on metrics"""
        limits = UnderwritingEngine.LENDER_LIMITS.get(
            profile,
            UnderwritingEngine.LENDER_LIMITS["Bank / Credit Union"]
        )
        
        # LTV score
        ltv_score = max(0, 300 * (1 - actual_ltv / limits['max_ltv'])) if limits['max_ltv'] > 0 else 0
        
        # DSCR score
        if limits['min_dscr'] > 1.0 and actual_dscr > 1.0:
            dscr_score = max(0, 300 * (actual_dscr - 1.0) / (limits['min_dscr'] - 1.0))
        else:
            dscr_score = 300 if actual_dscr >= 1.0 else 0
        
        # Debt yield score
        dy_score = max(0, 200 * min(1.5, actual_dy / limits['min_dy'])) if limits['min_dy'] > 0 else 0
        
        # LTC score
        ltc_score = max(0, 200 * (1 - actual_ltc))
        
        total = min(1000, int(ltv_score + dscr_score + dy_score + ltc_score))
        
        if total >= 850:
            tier = "Tier 1 | Institutional Core"
        elif total >= 700:
            tier = "Tier 2 | Conventional Bankable"
        elif total >= 550:
            tier = "Tier 3 | Alternative / Debt Fund"
        else:
            tier = "Tier 4 | Private / Restructure"
        
        return total, tier


class SensitivityEngine:
    """Sensitivity analysis"""
    
    @staticmethod
    def generate_matrix(state: dict) -> pd.DataFrame:
        """Generate sensitivity matrix with FIXED calculation"""
        scenarios = [
            ('Base', 0, 0),
            ('Rate +1%', 0.01, 0),
            ('Rate -1%', -0.01, 0),
            ('NOI -10%', 0, -0.10),
            ('NOI +10%', 0, 0.10),
            ('Combined Stress', 0.01, -0.10),
        ]
        
        results = []
        base_proceeds = None
        
        for scenario_name, rate_adj, noi_adj in scenarios:
            adjusted_rate = max(0.001, safe_float(state.get('rate', 0.0525)) + rate_adj)
            adjusted_noi = max(0, safe_float(state.get('noi', 0)) * (1 + noi_adj))
            
            loan, gate, _, _, _ = UnderwritingEngine.size_loan(
                noi=adjusted_noi,
                appraisal=safe_float(state.get('appraisal', 0)),
                purchase_price=safe_float(state.get('purchase_price', 0)),
                closing_costs=safe_float(state.get('closing_costs', 0)),
                reserves=safe_float(state.get('reserves', 0)),
                fees_pct=safe_float(state.get('fees', 0)),
                rate=adjusted_rate,
                amort=safe_int(state.get('amort', 25)),
                term=safe_int(state.get('term', 5)),
                is_io=bool(state.get('is_io', False)),
                target_ltv=safe_float(state.get('target_ltv', 0.75)),
                target_ltc=safe_float(state.get('target_ltc', 0.80)),
                target_dscr=safe_float(state.get('target_dscr', 1.25)),
                target_dy=safe_float(state.get('target_dy', 0.085))
            )
            
            if scenario_name == 'Base':
                base_proceeds = loan
            
            # Calculate change from base properly
            if base_proceeds and base_proceeds > 0:
                change_pct = ((loan - base_proceeds) / base_proceeds) * 100
                change_str = f"{change_pct:+.1f}%"
            else:
                change_str = "N/A"
            
            results.append({
                'Scenario': scenario_name,
                'Rate': f"{adjusted_rate*100:.2f}%",
                'NOI': f"${adjusted_noi:,.0f}",
                'Max Proceeds': f"${loan:,.0f}",
                'Constraint': gate,
                'Change from Base': change_str
            })
        
        return pd.DataFrame(results)


class ValidationEngine:
    """Deal validation"""
    
    @staticmethod
    def validate(state: dict) -> tuple:
        """Validate deal state, return (errors, warnings)"""
        errors = []
        warnings = []
        
        # Required text fields (lenient - only check for completely empty)
        deal_name = str(state.get('deal_name', '')).strip()
        property_address = str(state.get('property_address', '')).strip()
        
        if not deal_name:
            errors.append("Deal name is required")
        
        # Property address is optional in initial data entry
        # Only warn if other fields suggest this is a real deal
        if safe_float(state.get('purchase_price')) > 0 and not property_address:
            warnings.append("Property address is recommended for complete deal profile")
        
        # Financial validation
        purchase_price = safe_float(state.get('purchase_price'))
        appraisal = safe_float(state.get('appraisal'))
        noi = safe_float(state.get('noi'))
        rate = safe_float(state.get('rate'))
        
        if purchase_price < 0:
            errors.append("Purchase price cannot be negative")
        
        if appraisal < 0:
            errors.append("Appraisal cannot be negative")
        
        if noi < 0:
            errors.append("NOI cannot be negative")
        
        if rate <= 0 and (purchase_price > 0 or noi > 0):
            errors.append("Interest rate must be positive for active deals")
        
        if rate > 0.30:
            errors.append("Interest rate cannot exceed 30%")
        
        # Cap rate reasonableness
        if noi > 0 and appraisal > 0:
            cap_rate = noi / appraisal
            if 0 < cap_rate < 0.01:
                warnings.append(f"Implied cap rate {cap_rate:.2%} is unusually low - verify NOI and appraisal")
            elif cap_rate > 0.20:
                warnings.append(f"Implied cap rate {cap_rate:.2%} is unusually high - verify NOI and appraisal")
        
        # Purchase vs appraisal
        if purchase_price > 0 and appraisal > 0:
            premium = (purchase_price - appraisal) / appraisal
            if premium > 0.30:
                warnings.append(f"Purchase price is {premium:.1%} above appraisal")
        
        # Target metrics validation
        target_ltv = safe_float(state.get('target_ltv'))
        target_dscr = safe_float(state.get('target_dscr'))
        
        if target_ltv <= 0:
            errors.append("Target LTV must be positive")
        
        if target_ltv > 1.25:
            errors.append("Target LTV cannot exceed 125%")
        
        if target_dscr <= 0:
            errors.append("Target DSCR must be positive")
        
        # Amortization validation
        amort = safe_int(state.get('amort'))
        if amort < 1:
            errors.append("Amortization must be at least 1 year")
        elif amort > 40:
            errors.append("Amortization cannot exceed 40 years")
        
        return errors, warnings


# ==========================================
# MAIN APPLICATION
# ==========================================

def main():
    """Main application entry point"""
    
    # Initialize
    DatabaseManager.init_db()
    initialize_session_state()
    
    s = st.session_state
    
    # ==========================================
    # SIDEBAR
    # ==========================================
    
    with st.sidebar:
        st.title("🏛️ ALENZA OS")
        st.caption(f"v{APP_VERSION}")
        
        # Deal management
        with st.expander("📁 DEAL MANAGER", expanded=True):
            # New deal
            new_name = st.text_input("New Deal Name", value="Untitled Deal", key="new_deal_name")
            if st.button("➕ New Deal", use_container_width=True, key="btn_new_deal"):
                new_id = generate_id('deal')
                for key in ['deal_id', 'deal_name', 'sponsor', 'property_address',
                           'purchase_price', 'appraisal', 'noi', 'closing_costs', 'reserves']:
                    if key == 'deal_id':
                        s[key] = new_id
                    elif key == 'deal_name':
                        s[key] = new_name
                    elif key in ['sponsor', 'property_address']:
                        s[key] = ''
                    else:
                        s[key] = 0.0
                s.rent_roll_dict = []
                s.unsaved_changes = True
                st.rerun()
            
            st.markdown("---")
            
            # Load existing
            deals_df = DatabaseManager.get_all_deals()
            if not deals_df.empty:
                deal_options = {}
                for _, row in deals_df.iterrows():
                    deal_id = str(row['id'])
                    name = str(row.get('name', 'Untitled'))[:50]
                    updated = str(row.get('updated_at', ''))[:19]
                    label = f"{name} ({deal_id[-8:]})"
                    deal_options[label] = deal_id
                
                selected_label = st.selectbox("Load Deal", list(deal_options.keys()), key="load_deal_select")
                
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("📂 Load", use_container_width=True, key="btn_load"):
                        deal_id = deal_options[selected_label]
                        state = DatabaseManager.load_deal(deal_id)
                        if state:
                            for k, v in state.items():
                                s[k] = v
                            s.unsaved_changes = False
                            st.rerun()
                
                with col2:
                    if st.button("🗑️ Delete", use_container_width=True, key="btn_delete"):
                        confirm = st.checkbox("Confirm delete?", key="confirm_delete")
                        if confirm:
                            deal_id = deal_options[selected_label]
                            DatabaseManager.delete_deal(deal_id)
                            st.success("Deleted")
                            st.rerun()
            else:
                st.info("No saved deals yet")
            
            # Save button
            if st.button("💾 Save Deal", use_container_width=True, key="btn_save"):
                state = extract_clean_state()
                errors, warnings = ValidationEngine.validate(state)
                
                if errors:
                    for error in errors:
                        st.error(f"• {error}")
                else:
                    if warnings:
                        for warning in warnings:
                            st.warning(f"• {warning}")
                    
                    deal_name = state.get('deal_name', 'Untitled')
                    if DatabaseManager.save_deal(state['deal_id'], deal_name, state):
                        s.unsaved_changes = False
                        s.last_saved_at = datetime.now().isoformat()
                        st.success("✅ Deal saved!")
                    else:
                        st.error("Save failed")
        
        # Unsaved warning
        if s.unsaved_changes:
            st.warning("⚠️ Unsaved changes")
        
        st.markdown("---")
        
        # Asset profile
        with st.expander("🏢 ASSET PROFILE", expanded=True):
            s.deal_name = st.text_input("Deal Name", value=s.get('deal_name', ''))
            s.sponsor = st.text_input("Sponsor", value=s.get('sponsor', ''))
            s.property_address = st.text_input("Property Address", value=s.get('property_address', ''))
            
            property_types = ["Multifamily", "Industrial", "Retail", "Office", "Mixed-Use"]
            current_pt = property_types.index(s.get('property_type', 'Multifamily')) if s.get('property_type') in property_types else 0
            s.property_type = st.selectbox("Property Type", property_types, index=current_pt)
            
            s.appraisal = st.number_input("Appraisal ($)", value=safe_float(s.get('appraisal', 0)), step=100000.0, min_value=0.0, format="%.0f")
            s.purchase_price = st.number_input("Cost Basis ($)", value=safe_float(s.get('purchase_price', 0)), step=100000.0, min_value=0.0, format="%.0f")
            s.noi = st.number_input("Stabilized NOI ($)", value=safe_float(s.get('noi', 0)), step=10000.0, min_value=0.0, format="%.0f")
        
        # Credit policy
        with st.expander("📊 CREDIT POLICY", expanded=True):
            profiles = list(UnderwritingEngine.LENDER_LIMITS.keys())
            current_lp = profiles.index(s.get('lender_profile', 'Bank / Credit Union')) if s.get('lender_profile') in profiles else 0
            s.lender_profile = st.selectbox("Lender Profile", profiles, index=current_lp)
            
            limits = UnderwritingEngine.LENDER_LIMITS[s.lender_profile]
            s.target_ltv = st.slider("Max LTV %", 50.0, 95.0, float(normalize_percent(s.get('target_ltv', limits['max_ltv'])) * 100), step=0.5) / 100
            s.target_dscr = st.slider("Min DSCR", 1.0, 1.75, float(safe_float(s.get('target_dscr', limits['min_dscr']))), step=0.05)
            s.target_dy = st.slider("Min DY %", 5.0, 15.0, float(normalize_percent(s.get('target_dy', limits['min_dy'])) * 100), step=0.25) / 100
            s.target_ltc = st.slider("Max LTC %", 50.0, 100.0, float(normalize_percent(s.get('target_ltc', 0.80)) * 100), step=0.5) / 100
        
        # Debt structure
        with st.expander("💰 DEBT STRUCTURE", expanded=True):
            s.is_io = st.checkbox("Interest-Only Period", value=bool(s.get('is_io', False)))
            s.rate = st.slider("Interest Rate %", 0.0, 15.0, float(normalize_percent(s.get('rate', 0.0525)) * 100), step=0.05) / 100
            s.amort = st.number_input("Amortization (Yrs)", value=max(1, safe_int(s.get('amort', 25))), step=1, min_value=1, max_value=40)
            s.term = st.number_input("Term (Yrs)", value=max(1, safe_int(s.get('term', 5))), step=1, min_value=1, max_value=40)
            s.fees = st.slider("Financing Fees %", 0.0, 5.0, float(normalize_percent(s.get('fees', 0.02)) * 100), step=0.05) / 100
            s.closing_costs = st.number_input("Closing Costs ($)", value=safe_float(s.get('closing_costs', 0)), step=1000.0, min_value=0.0, format="%.0f")
            s.reserves = st.number_input("Required Reserves ($)", value=safe_float(s.get('reserves', 0)), step=1000.0, min_value=0.0, format="%.0f")

    # ==========================================
    # MAIN CONTENT - CALCULATIONS
    # ==========================================
    
    # Normalize rent roll
    if s.rent_roll_dict:
        rr_df = pd.DataFrame(s.rent_roll_dict)
    else:
        rr_df = pd.DataFrame(columns=['Tenant', 'SF', 'Remaining Term', 'Monthly Rent'])
    
    # Core calculations
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
    
    # Actual metrics
    actual_ltv = loan_amt / s.appraisal if safe_float(s.appraisal) > 0 else 0
    actual_ltc = loan_amt / total_uses if total_uses > 0 else 0
    actual_dscr = s.noi / annual_ds if annual_ds > 0 else 0
    actual_dy = s.noi / loan_amt if loan_amt > 0 else 0
    
    # Rent roll metrics
    tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr_df)
    
    # Score deal
    score, tier = UnderwritingEngine.score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, s.lender_profile)
    
    # Validate
    errors, warnings = ValidationEngine.validate(extract_clean_state())
    
    # ==========================================
    # HEADER
    # ==========================================
    
    headline_sponsor = s.sponsor or "New Deal"
    headline_property = s.property_address or "Property Profile"
    
    st.title(f"{headline_sponsor} | {headline_property}")
    st.caption(f"ALENZA CAPITAL OS | CONSTRAINT: {gate} | {tier}")
    
    # Validation messages
    if errors:
        for error in errors[:3]:  # Show first 3 errors
            st.error(f"❌ {error}")
        if len(errors) > 3:
            st.error(f"... and {len(errors)-3} more errors")
    
    if warnings:
        for warning in warnings[:2]:  # Show first 2 warnings
            st.warning(f"⚠️ {warning}")
    
    # KPI Dashboard
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
    col2.metric("ACTUAL LTV", f"{actual_ltv*100:.1f}%")
    col3.metric("ACTUAL LTC", f"{actual_ltc*100:.1f}%")
    col4.metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")
    col5.metric("BALLOON", f"${balloon:,.0f}")
    col6.metric("DEAL SCORE", f"{score}/1000", help=tier)
    
    st.markdown("---")
    
    # ==========================================
    # MAIN TABS
    # ==========================================
    
    tabs = st.tabs([
        "📊 Sizing & Risk",
        "🧪 Sensitivity",
        "📝 Rent Roll",
        "📅 Amortization",
        "🇨🇦 Canada Intel",
        "📈 Market Comps",
        "💾 Save & Export"
    ])
    
    # ==========================================
    # TAB 0: SIZING & RISK
    # ==========================================
    
    with tabs[0]:
        col1, col2 = st.columns([1.5, 1])
        
        with col1:
            st.subheader("📐 Constraint Analysis")
            
            constraints_df = pd.DataFrame({
                "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
                "Threshold": [
                    f"{s.target_ltv*100:.1f}%",
                    f"{s.target_ltc*100:.1f}%",
                    f"{s.target_dscr:.2f}x",
                    f"{s.target_dy*100:.2f}%"
                ],
                "Max Proceeds": [
                    f"${gates.get('LTV', 0):,.0f}",
                    f"${gates.get('LTC', 0):,.0f}",
                    f"${gates.get('DSCR', 0):,.0f}",
                    f"${gates.get('Debt Yield', 0):,.0f}"
                ],
                "Binding": ["✅ ACTIVE" if gate == g else "" for g in ["LTV", "LTC", "DSCR", "Debt Yield"]]
            })
            
            st.dataframe(constraints_df, hide_index=True, use_container_width=True)
            
            st.subheader("💰 Sources & Uses")
            
            total_fees = loan_amt * s.fees
            
            su_df = pd.DataFrame({
                "Category": ["Cost Basis", "Closing Costs", "Reserves", "Financing Fees", "TOTAL USES"],
                "Uses": [
                    f"${s.purchase_price:,.0f}",
                    f"${s.closing_costs:,.0f}",
                    f"${s.reserves:,.0f}",
                    f"${total_fees:,.0f}",
                    f"${total_uses:,.0f}"
                ],
                "Sources": [
                    "Senior Debt",
                    "Sponsor Equity",
                    "",
                    "",
                    "TOTAL SOURCES"
                ],
                "Amount": [
                    f"${loan_amt:,.0f}",
                    f"${req_equity:,.0f}",
                    "",
                    "",
                    f"${total_uses:,.0f}"
                ]
            })
            
            st.dataframe(su_df, hide_index=True, use_container_width=True)
        
        with col2:
            st.subheader("🔍 Risk Assessment")
            
            # Generate risk flags
            flags = []
            
            if actual_ltv > 0.75:
                flags.append(("high", f"⚠️ High Leverage: {actual_ltv*100:.1f}% LTV exceeds conventional 75% threshold"))
            elif actual_ltv < 0.55:
                flags.append(("low", f"✅ Conservative Leverage: {actual_ltv*100:.1f}% LTV indicates strong equity commitment"))
            
            if actual_dscr < 1.20:
                flags.append(("high", f"⚠️ Tight Coverage: {actual_dscr:.2f}x DSCR provides minimal cushion"))
            elif actual_dscr > 1.50:
                flags.append(("low", f"✅ Strong Coverage: {actual_dscr:.2f}x DSCR exceeds typical requirements"))
            
            if s.is_io:
                flags.append(("medium", "ℹ️ Interest-Only Structure: Amortization risk during IO period"))
            
            if req_equity < 0:
                flags.append(("high", f"🚨 Negative Equity: ${abs(req_equity):,.0f} cash-out scenario"))
            
            if walt > 0 and walt < 3:
                flags.append(("high", f"⚠️ Short WALT: {walt:.1f} years - leasing risk elevated"))
            
            if exp1 > 0.30:
                flags.append(("high", f"🚨 High Rollover: {exp1*100:.1f}% of SF expiring within 12 months"))
            
            if not flags:
                flags.append(("low", "✅ No Significant Risk Flags Detected"))
            
            for severity, message in flags:
                if severity == "high":
                    st.error(message)
                elif severity == "medium":
                    st.warning(message)
                else:
                    st.success(message)
            
            # Key metrics
            st.markdown("---")
            st.subheader("📊 Key Metrics")
            
            breakeven_occ = (s.noi / annual_ds) * occ if annual_ds > 0 and occ > 0 else 0
            
            st.metric("Breakeven Occupancy", f"{breakeven_occ*100:.1f}%", 
                     delta=f"Current: {occ*100:.1f}%" if occ > 0 else None)
            st.metric("Required Equity", f"${req_equity:,.0f}")
            st.metric("Implied Cap Rate", f"{(s.noi/s.appraisal*100):.2f}%" if s.appraisal > 0 else "N/A")
    
    # ==========================================
    # TAB 1: SENSITIVITY
    # ==========================================
    
    with tabs[1]:
        st.subheader("🧪 Sensitivity Analysis")
        
        sensitivity_df = SensitivityEngine.generate_matrix(extract_clean_state())
        st.dataframe(sensitivity_df, hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🎯 Custom Stress Test")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            rate_shock = st.slider("Rate Shock (bps)", -200, 200, 0, 25, key="custom_rate_shock") / 10000
        with col2:
            noi_shock = st.slider("NOI Shock (%)", -30, 30, 0, 5, key="custom_noi_shock") / 100
        with col3:
            ltv_shock = st.slider("LTV Adjustment (%)", -10, 10, 0, 1, key="custom_ltv_shock") / 100
        
        stressed_loan, stressed_gate, _, _, _ = UnderwritingEngine.size_loan(
            noi=max(0, s.noi * (1 + noi_shock)),
            appraisal=s.appraisal,
            purchase_price=s.purchase_price,
            closing_costs=s.closing_costs,
            reserves=s.reserves,
            fees_pct=s.fees,
            rate=max(0.001, s.rate + rate_shock),
            amort=s.amort,
            term=s.term,
            is_io=s.is_io,
            target_ltv=min(1.25, max(0.01, s.target_ltv + ltv_shock)),
            target_ltc=s.target_ltc,
            target_dscr=s.target_dscr,
            target_dy=s.target_dy
        )
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Stressed Proceeds", f"${stressed_loan:,.0f}",
                   delta=f"${stressed_loan - loan_amt:,.0f}" if loan_amt > 0 else None)
        col2.metric("Constraint", stressed_gate)
        col3.metric("Stressed LTV", f"{(stressed_loan/s.appraisal*100):.1f}%" if s.appraisal > 0 else "N/A")
        col4.metric("Stressed DSCR", f"{(s.noi*(1+noi_shock)/(monthly_pmt*12)):.2f}x" if monthly_pmt > 0 else "N/A")
    
    # ==========================================
    # TAB 2: RENT ROLL
    # ==========================================
    
    with tabs[2]:
        st.subheader("📝 Rent Roll Management")
        
        # File upload
        uploaded_file = st.file_uploader(
            "Import Rent Roll (CSV or Excel)", 
            type=['csv', 'xlsx', 'xls'],
            key="rent_roll_upload"
        )
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    imported_df = pd.read_csv(uploaded_file)
                else:
                    imported_df = pd.read_excel(uploaded_file)
                
                # Auto-map columns
                col_mapping = {}
                for col in imported_df.columns:
                    col_lower = str(col).lower()
                    if 'tenant' in col_lower or 'name' in col_lower:
                        col_mapping[col] = 'Tenant'
                    elif 'sf' in col_lower or 'sq' in col_lower or 'area' in col_lower:
                        col_mapping[col] = 'SF'
                    elif 'term' in col_lower or 'lease' in col_lower:
                        col_mapping[col] = 'Remaining Term'
                    elif 'rent' in col_lower or 'revenue' in col_lower:
                        col_mapping[col] = 'Monthly Rent'
                
                st.info(f"Detected columns: {list(col_mapping.keys())}")
                
                if st.button("✅ Apply Imported Rent Roll"):
                    imported_df = imported_df.rename(columns=col_mapping)
                    
                    # Ensure required columns
                    for req_col in ['Tenant', 'SF', 'Remaining Term', 'Monthly Rent']:
                        if req_col not in imported_df.columns:
                            imported_df[req_col] = '' if req_col == 'Tenant' else 0
                    
                    # Clean data
                    imported_df['Tenant'] = imported_df['Tenant'].fillna('Vacant').astype(str)
                    imported_df['SF'] = pd.to_numeric(imported_df['SF'], errors='coerce').fillna(0)
                    imported_df['Remaining Term'] = pd.to_numeric(imported_df['Remaining Term'], errors='coerce').fillna(0)
                    imported_df['Monthly Rent'] = pd.to_numeric(imported_df['Monthly Rent'], errors='coerce').fillna(0)
                    
                    s.rent_roll_dict = imported_df[['Tenant', 'SF', 'Remaining Term', 'Monthly Rent']].to_dict('records')
                    s.unsaved_changes = True
                    st.success(f"✅ Imported {len(imported_df)} tenant records")
                    st.rerun()
                    
            except Exception as e:
                st.error(f"Import failed: {str(e)[:200]}")
        
        # Interactive editor
        if s.rent_roll_dict:
            edit_df = pd.DataFrame(s.rent_roll_dict)
        else:
            edit_df = pd.DataFrame(columns=['Tenant', 'SF', 'Remaining Term', 'Monthly Rent'])
        
        edited_df = st.data_editor(
            edit_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Tenant": st.column_config.TextColumn("Tenant Name", width="large"),
                "SF": st.column_config.NumberColumn("Square Feet", min_value=0, step=100, format="%d"),
                "Remaining Term": st.column_config.NumberColumn("Lease Term (Yrs)", min_value=0.0, step=0.5, format="%.1f"),
                "Monthly Rent": st.column_config.NumberColumn("Monthly Rent ($)", min_value=0.0, step=100.0, format="$%.2f")
            },
            key="rent_roll_editor"
        )
        
        if st.button("💾 Update Rent Roll", use_container_width=True):
            s.rent_roll_dict = edited_df.to_dict('records')
            s.unsaved_changes = True
            st.rerun()
        
        # Metrics
        if not edited_df.empty:
            st.markdown("---")
            st.subheader("📊 Rent Roll Analytics")
            
            met_total_sf, met_occ, met_ann_rent, met_psf, met_walt, met_exp1 = UnderwritingEngine.rent_roll_metrics(edited_df)
            
            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Total SF", f"{met_total_sf:,.0f}")
            col2.metric("Occupancy", f"{met_occ*100:.1f}%")
            col3.metric("Annual Rent", f"${met_ann_rent:,.0f}")
            col4.metric("Rent PSF", f"${met_psf:.2f}")
            col5.metric("WALT (Yrs)", f"{met_walt:.2f}")
            col6.metric("12-Mo Rollover", f"{met_exp1*100:.1f}%")
    
    # ==========================================
    # TAB 3: AMORTIZATION
    # ==========================================
    
    with tabs[3]:
        st.subheader(f"📅 Amortization Schedule - {s.term} Year Term")
        
        if amort_df is None or amort_df.empty:
            st.warning("Enter deal parameters to generate amortization schedule")
        else:
            col1, col2, col3 = st.columns(3)
            col1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
            col2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
            col3.metric("Balloon Balance", f"${balloon:,.0f}")
            
            st.markdown("---")
            
            # Charts
            st.write("### Payment Structure")
            chart_data = amort_df.set_index('Period')[['Principal', 'Interest']]
            st.bar_chart(chart_data, use_container_width=True)
            
            st.write("### Outstanding Balance")
            balance_chart = amort_df.set_index('Period')[['Balance']]
            st.line_chart(balance_chart, use_container_width=True)
            
            # Annual summary
            amort_df['Year'] = ((amort_df['Period'] - 1) // 12) + 1
            annual_summary = amort_df.groupby('Year').agg({
                'Payment': 'sum',
                'Principal': 'sum',
                'Interest': 'sum',
                'Balance': 'last'
            }).reset_index()
            
            st.write("### Annual Summary")
            st.dataframe(
                annual_summary.style.format({
                    'Payment': '${:,.2f}',
                    'Principal': '${:,.2f}',
                    'Interest': '${:,.2f}',
                    'Balance': '${:,.2f}'
                }),
                use_container_width=True,
                hide_index=True
            )
            
            # Full schedule
            with st.expander("View Full Monthly Schedule"):
                st.dataframe(
                    amort_df[['Period', 'Payment', 'Principal', 'Interest', 'Balance']].style.format({
                        'Payment': '${:,.2f}',
                        'Principal': '${:,.2f}',
                        'Interest': '${:,.2f}',
                        'Balance': '${:,.2f}'
                    }),
                    use_container_width=True,
                    height=400,
                    hide_index=True
                )
    
   # ==========================================
    # TAB 4: CANADA INTEL (SOVEREIGN DATA)
    # ==========================================
    
    with tabs[4]:
        st.subheader("🇨🇦 Canadian Sovereign Intelligence")
        st.caption("Real-time market data from Bank of Canada and Statistics Canada")
        
        # ==========================================
        # BANK OF CANADA RATES - ENHANCED WITH GRAPHS
        # ==========================================
        
        st.write("### 🏦 Bank of Canada - Key Rates & Yields")
        
        @st.cache_data(ttl=3600)
        def fetch_boc_rates_history():
            """Fetch Bank of Canada rates with history for charts"""
            try:
                # Fetch last 90 days of data for charts
                url = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD,BD.CDN.2YR.DQ.YLD,BD.CDN.5YR.DQ.YLD,BD.CDN.10YR.DQ.YLD/json?recent=90"
                response = requests.get(url, timeout=10)
                response.raise_for_status()
                data = response.json()
                
                observations = data.get('observations', [])
                if not observations:
                    return None, None
                
                # Build history dataframe
                history = []
                for obs in observations:
                    date = obs.get('d')
                    row = {'Date': pd.to_datetime(date)}
                    
                    for key, name in [
                        ('FXUSDCAD', 'USD/CAD'),
                        ('BD.CDN.2YR.DQ.YLD', '2-Year Yield'),
                        ('BD.CDN.5YR.DQ.YLD', '5-Year Yield'),
                        ('BD.CDN.10YR.DQ.YLD', '10-Year Yield')
                    ]:
                        try:
                            val = obs.get(key, {}).get('v')
                            row[name] = float(val) if val else None
                        except:
                            row[name] = None
                    
                    history.append(row)
                
                history_df = pd.DataFrame(history).sort_values('Date')
                
                # Latest values
                latest = history_df.iloc[-1] if not history_df.empty else None
                
                return latest, history_df
            except Exception as e:
                logger.error(f"BOC fetch failed: {e}")
                return None, None
        
        boc_latest, boc_history = fetch_boc_rates_history()
        
        if boc_latest is not None:
            # Metric cards
            col1, col2, col3, col4 = st.columns(4)
            
            usd_cad = boc_latest.get('USD/CAD')
            yield_2y = boc_latest.get('2-Year Yield')
            yield_5y = boc_latest.get('5-Year Yield')
            yield_10y = boc_latest.get('10-Year Yield')
            
            col1.metric(
                "USD / CAD", 
                f"{usd_cad:.4f}" if usd_cad else "N/A",
                help="US Dollar to Canadian Dollar exchange rate"
            )
            col2.metric(
                "Canada 2-Year", 
                f"{yield_2y:.2f}%" if yield_2y else "N/A",
                help="Government of Canada 2-year bond yield"
            )
            col3.metric(
                "Canada 5-Year", 
                f"{yield_5y:.2f}%" if yield_5y else "N/A",
                help="Government of Canada 5-year bond yield - CRE benchmark"
            )
            col4.metric(
                "Canada 10-Year", 
                f"{yield_10y:.2f}%" if yield_10y else "N/A",
                help="Government of Canada 10-year bond yield"
            )
            
            # Date stamp
            if not boc_history.empty:
                st.caption(f"📅 Latest data: {boc_history['Date'].iloc[-1].strftime('%B %d, %Y')}")
            
            st.markdown("---")
            
            # ==========================================
            # GRAPH 1: YIELD CURVE (Current Snapshot)
            # ==========================================
            st.write("#### 📈 Current Yield Curve")
            
            curve_data = pd.DataFrame({
                'Tenor': ['2-Year', '5-Year', '10-Year'],
                'Yield': [yield_2y, yield_5y, yield_10y]
            })
            
            # Create yield curve chart
            if DEPENDENCIES['plotly']:
                import plotly.express as px
                import plotly.graph_objects as go
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=curve_data['Tenor'],
                    y=curve_data['Yield'],
                    mode='lines+markers',
                    line=dict(color='#CFB87C', width=3),
                    marker=dict(size=12, color='#CFB87C'),
                    name='Current Yield Curve'
                ))
                
                fig.update_layout(
                    title='Government of Canada Yield Curve',
                    xaxis_title='Maturity',
                    yaxis_title='Yield (%)',
                    template='plotly_dark',
                    paper_bgcolor='#0B0F19',
                    plot_bgcolor='#0F172A',
                    height=350,
                    margin=dict(l=20, r=20, t=40, b=20),
                    yaxis=dict(tickformat='.2f', ticksuffix='%')
                )
                
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.bar_chart(curve_data.set_index('Tenor'), use_container_width=True)
            
            # Yield spread
            if yield_2y and yield_10y:
                spread = yield_10y - yield_2y
                spread_color = "normal" if spread > 0 else "inverted"
                
                col_a, col_b = st.columns(2)
                col_a.metric(
                    "2s10s Spread", 
                    f"{spread:.2f}%",
                    help="10Y minus 2Y yield. Negative = inverted curve (recession signal)"
                )
                col_b.metric(
                    "Curve Status",
                    "📈 Normal (Steepening)" if spread > 0.25 else 
                    "📉 Inverted (Warning)" if spread < 0 else 
                    "➡️ Flat (Transition)",
                    help="Yield curve shape indicates market growth expectations"
                )
            
            st.markdown("---")
            
            # ==========================================
            # GRAPH 2: YIELD HISTORY (90-Day Trend)
            # ==========================================
            st.write("#### 📊 90-Day Yield Trend")
            
            if not boc_history.empty and DEPENDENCIES['plotly']:
                fig2 = go.Figure()
                
                colors = {'2-Year Yield': '#CFB87C', '5-Year Yield': '#F59E0B', '10-Year Yield': '#EF4444'}
                
                for col_name in ['2-Year Yield', '5-Year Yield', '10-Year Yield']:
                    if col_name in boc_history.columns:
                        fig2.add_trace(go.Scatter(
                            x=boc_history['Date'],
                            y=boc_history[col_name],
                            mode='lines',
                            name=col_name,
                            line=dict(color=colors.get(col_name, '#FFFFFF'), width=2),
                            hovertemplate=f'{col_name}: %{{y:.2f}}%<br>Date: %{{x}}<extra></extra>'
                        ))
                
                fig2.update_layout(
                    title='Government of Canada Bond Yields - Last 90 Days',
                    xaxis_title='Date',
                    yaxis_title='Yield (%)',
                    template='plotly_dark',
                    paper_bgcolor='#0B0F19',
                    plot_bgcolor='#0F172A',
                    height=400,
                    margin=dict(l=20, r=20, t=40, b=20),
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                    yaxis=dict(tickformat='.2f', ticksuffix='%'),
                    hovermode='x unified'
                )
                
                st.plotly_chart(fig2, use_container_width=True)
            elif not boc_history.empty:
                # Fallback to Streamlit native chart
                chart_cols = [c for c in ['2-Year Yield', '5-Year Yield', '10-Year Yield'] if c in boc_history.columns]
                if chart_cols:
                    st.line_chart(
                        boc_history.set_index('Date')[chart_cols],
                        use_container_width=True
                    )
            
            st.markdown("---")
            
            # ==========================================
            # GRAPH 3: FX RATE HISTORY
            # ==========================================
            st.write("#### 💱 USD/CAD Exchange Rate - 90-Day Trend")
            
            if not boc_history.empty and 'USD/CAD' in boc_history.columns:
                if DEPENDENCIES['plotly']:
                    fig3 = go.Figure()
                    
                    fig3.add_trace(go.Scatter(
                        x=boc_history['Date'],
                        y=boc_history['USD/CAD'],
                        mode='lines',
                        name='USD/CAD',
                        fill='tozeroy',
                        fillcolor='rgba(207, 184, 124, 0.1)',
                        line=dict(color='#CFB87C', width=2),
                        hovertemplate='USD/CAD: %{y:.4f}<br>Date: %{x}<extra></extra>'
                    ))
                    
                    # Add average line
                    avg_rate = boc_history['USD/CAD'].mean()
                    fig3.add_hline(
                        y=avg_rate, 
                        line_dash="dash", 
                        line_color="#9CA3AF",
                        annotation_text=f"Avg: {avg_rate:.4f}",
                        annotation_position="bottom right"
                    )
                    
                    fig3.update_layout(
                        title='USD/CAD Exchange Rate - Last 90 Days',
                        xaxis_title='Date',
                        yaxis_title='USD/CAD',
                        template='plotly_dark',
                        paper_bgcolor='#0B0F19',
                        plot_bgcolor='#0F172A',
                        height=350,
                        margin=dict(l=20, r=20, t=40, b=20),
                        hovermode='x unified'
                    )
                    
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.line_chart(
                        boc_history.set_index('Date')['USD/CAD'],
                        use_container_width=True
                    )
                
                # FX stats
                fx_min = boc_history['USD/CAD'].min()
                fx_max = boc_history['USD/CAD'].max()
                fx_change = boc_history['USD/CAD'].iloc[-1] - boc_history['USD/CAD'].iloc[0]
                
                col_a, col_b, col_c = st.columns(3)
                col_a.metric("90-Day Low", f"{fx_min:.4f}")
                col_b.metric("90-Day High", f"{fx_max:.4f}")
                col_c.metric("90-Day Change", f"{fx_change:.4f}", 
                           delta=f"{fx_change:+.4f}")
            
            st.markdown("---")
            
            # ==========================================
            # GRAPH 4: RATE SPREAD TO DEAL
            # ==========================================
            st.write("#### 🎯 Deal Rate vs Market Benchmark")
            
            current_rate = safe_float(s.rate) * 100 if s.rate else 5.25
            
            if yield_5y and DEPENDENCIES['plotly']:
                comparison_data = pd.DataFrame({
                    'Rate Type': ['Your Deal Rate', '5Y GoC Benchmark', 'Spread'],
                    'Rate (%)': [current_rate, yield_5y, current_rate - yield_5y]
                })
                
                colors_bar = ['#CFB87C', '#3B82F6', '#10B981' if current_rate - yield_5y < 3 else '#F59E0B']
                
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(
                    x=comparison_data['Rate Type'],
                    y=comparison_data['Rate (%)'],
                    marker_color=colors_bar,
                    text=[f"{v:.2f}%" for v in comparison_data['Rate (%)']],
                    textposition='outside',
                    textfont=dict(color='white', size=14)
                ))
                
                fig4.update_layout(
                    title='Deal Pricing vs Government Benchmark',
                    template='plotly_dark',
                    paper_bgcolor='#0B0F19',
                    plot_bgcolor='#0F172A',
                    height=350,
                    margin=dict(l=20, r=20, t=40, b=20),
                    showlegend=False,
                    yaxis=dict(tickformat='.2f', ticksuffix='%')
                )
                
                st.plotly_chart(fig4, use_container_width=True)
                
                # Spread commentary
                spread_to_goc = current_rate - yield_5y
                
                if spread_to_goc < 1.5:
                    st.success(f"✅ **Tight Pricing:** {spread_to_goc:.2f}% spread indicates competitive institutional terms")
                elif spread_to_goc < 3.0:
                    st.info(f"ℹ️ **Standard Pricing:** {spread_to_goc:.2f}% spread is within typical CRE lending range")
                else:
                    st.warning(f"⚠️ **Wide Spread:** {spread_to_goc:.2f}% spread may reflect higher risk premium or transitional asset")
        
        else:
            st.warning("⚠️ Unable to fetch Bank of Canada data. Please check your internet connection.")
            st.info("💡 Charts and rate comparisons will appear here when data is available.")
    
    # ==========================================
    # TAB 5: MARKET COMPS
    # ==========================================
    
    with tabs[5]:
        st.subheader("📈 Market Comparables")
        st.caption("Simulated comparables based on property type - replace with verified broker data")
        
        @st.cache_data(ttl=3600)
        def generate_mock_comps(property_type, noi):
            """Generate mock comparables for demonstration"""
            np.random.seed(int(safe_float(noi)) % 10000)
            
            base_caps = {
                'Multifamily': 0.045,
                'Industrial': 0.055,
                'Retail': 0.065,
                'Office': 0.075,
                'Mixed-Use': 0.060
            }
            
            base_cap = base_caps.get(property_type, 0.06)
            comps = []
            
            for i in range(5):
                cap_var = np.random.uniform(-0.005, 0.005)
                comp_cap = max(0.03, base_cap + cap_var)
                comp_noi = safe_float(noi) * np.random.uniform(0.7, 1.3)
                comp_value = comp_noi / comp_cap if comp_cap > 0 else 0
                
                comps.append({
                    'Comparable': f"{property_type} Property {i+1}",
                    'Type': property_type,
                    'Distance': f"{np.random.uniform(0.5, 8):.1f} km",
                    'Sale Date': f"202{np.random.randint(3,6)}-{np.random.randint(1,13):02d}",
                    'Cap Rate': f"{comp_cap*100:.2f}%",
                    'Estimated Value': f"${comp_value:,.0f}",
                    'Price/Unit': f"${comp_value/max(1,np.random.randint(20,200)):,.0f}"
                })
            
            return pd.DataFrame(comps)
        
        comps_df = generate_mock_comps(s.property_type, s.noi)
        st.dataframe(comps_df, hide_index=True, use_container_width=True)
        
        # Cap rate comparison
        if s.noi > 0 and s.appraisal > 0:
            implied_cap = s.noi / s.appraisal * 100
            market_caps = [float(c.replace('%', '')) for c in comps_df['Cap Rate']]
            avg_market_cap = np.mean(market_caps) if market_caps else 0
            
            col1, col2 = st.columns(2)
            col1.metric("Your Implied Cap Rate", f"{implied_cap:.2f}%")
            col2.metric("Market Average Cap Rate", f"{avg_market_cap:.2f}%",
                       delta=f"{implied_cap - avg_market_cap:+.2f}%")
    
    # ==========================================
    # TAB 6: SAVE & EXPORT
    # ==========================================
    
    with tabs[6]:
        st.subheader("💾 Save & Export Deal")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.write("### Save to Database")
            if st.button("💾 Save Deal", use_container_width=True, key="export_save"):
                state = extract_clean_state()
                errors, warnings = ValidationEngine.validate(state)
                
                if errors:
                    for error in errors:
                        st.error(f"• {error}")
                else:
                    if DatabaseManager.save_deal(state['deal_id'], state.get('deal_name', 'Untitled'), state):
                        s.unsaved_changes = False
                        s.last_saved_at = datetime.now().isoformat()
                        st.success(f"✅ Saved at {s.last_saved_at[:19]}")
                    else:
                        st.error("Save failed")
            
            if s.get('last_saved_at'):
                st.caption(f"Last saved: {s.last_saved_at[:19]}")
        
        with col2:
            st.write("### Export Excel")
            
            if DEPENDENCIES['excel_write']:
                try:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Summary
                        pd.DataFrame({
                            'Metric': ['Deal Name', 'Sponsor', 'Property', 'Type', 'Max Proceeds', 
                                      'Constraint', 'LTV', 'DSCR', 'Score', 'Tier'],
                            'Value': [
                                s.deal_name, s.sponsor, s.property_address, s.property_type,
                                f"${loan_amt:,.0f}", gate, f"{actual_ltv*100:.1f}%",
                                f"{actual_dscr:.2f}x", score, tier
                            ]
                        }).to_excel(writer, sheet_name='Summary', index=False)
                        
                        # Rent Roll
                        if s.rent_roll_dict:
                            pd.DataFrame(s.rent_roll_dict).to_excel(
                                writer, sheet_name='Rent Roll', index=False
                            )
                        
                        # Amortization
                        if amort_df is not None and not amort_df.empty:
                            amort_df.to_excel(writer, sheet_name='Amortization', index=False)
                    
                    st.download_button(
                        "📊 Download Excel",
                        data=output.getvalue(),
                        file_name=f"alenza_deal_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Excel export failed: {e}")
            else:
                st.warning("Install xlsxwriter for Excel export")
        
        with col3:
            st.write("### Export JSON")
            
            state_json = json.dumps(extract_clean_state(), indent=2, default=str)
            st.download_button(
                "📄 Download JSON",
                data=state_json,
                file_name=f"alenza_deal_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True
            )
            
            # Full package
            if DEPENDENCIES['excel_write']:
                try:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zf:
                        zf.writestr('deal_state.json', state_json)
                        
                        excel_output = io.BytesIO()
                        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                            pd.DataFrame(s.rent_roll_dict).to_excel(writer, sheet_name='Rent Roll', index=False)
                            if amort_df is not None and not amort_df.empty:
                                amort_df.to_excel(writer, sheet_name='Amortization', index=False)
                        zf.writestr('model.xlsx', excel_output.getvalue())
                    
                    st.download_button(
                        "📦 Download Full Package",
                        data=zip_buffer.getvalue(),
                        file_name=f"alenza_package_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Package export failed: {e}")
    
    # ==========================================
    # FOOTER
    # ==========================================
    
    st.markdown("---")
    st.caption(
        "⚠️ **DISCLAIMER:** ALENZA CAPITAL OS is an indicative modeling tool for commercial real estate "
        "underwriting. Outputs do not constitute a loan commitment, appraisal, or legal advice. "
        "All final terms are subject to formal credit committee approval, third-party due diligence, "
        "and definitive documentation. © 2024 Alenza Capital. All rights reserved."
    )
    
    # Auto-save reminder
    if s.unsaved_changes:
        st.sidebar.warning(f"⏰ Unsaved changes - last auto-save check: {datetime.now().strftime('%H:%M:%S')}")


# ==========================================
# APP ENTRY POINT
# ==========================================

if __name__ == "__main__":
    main()
            
