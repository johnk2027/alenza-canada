"""
Alenza Capital OS v3.0.2
Enterprise underwriting workspace
Midnight Slate and CU Gold theme

Finalized Part 1 + Part 2.
Part 3 should continue below the Amortization tab block.
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
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional, Callable
from functools import wraps


# ==========================================
# LOGGING SETUP
# ==========================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()],
)

logger = logging.getLogger(__name__)


# ==========================================
# PAGE CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="Alenza Capital OS",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==========================================
# CONSTANTS
# ==========================================

APP_VERSION = "3.0.2"
SCHEMA_VERSION = 4
MAX_UPLOAD_SIZE_MB = 50

DATA_DIR = Path("alenza_data")
DOC_DIR = DATA_DIR / "documents"
DB_PATH = DATA_DIR / "alenza_platform.db"

PROPERTY_TYPES = [
    "Multifamily",
    "Industrial",
    "Retail",
    "Office",
    "Mixed-Use",
    "Hospitality",
    "Self-Storage",
]

TRANSACTION_TYPES = [
    "Acquisition",
    "Refinance",
    "Construction",
    "Bridge",
    "Recapitalization",
]

LENDER_PROFILES_LIST = [
    "Bank / Credit Union",
    "LifeCo / Core",
    "Bridge / Private",
    "CMHC Multifamily",
]

DEFAULT_RENT_ROLL_COLUMNS = [
    "Tenant",
    "SF",
    "Remaining Term",
    "Monthly Rent",
]

DEFAULT_RENT_ROLL = [
    {"Tenant": "Main Anchor", "SF": 25000, "Remaining Term": 5.5, "Monthly Rent": 45000},
    {"Tenant": "In-Line A", "SF": 3500, "Remaining Term": 1.2, "Monthly Rent": 8000},
]


# ==========================================
# DEPENDENCY CHECKING
# ==========================================

@st.cache_resource
def check_dependencies() -> Dict[str, bool]:
    deps = {
        "ocr": False,
        "pdf": False,
        "excel_write": False,
        "excel_read": False,
        "xls_read": False,
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
        import xlrd
        deps["xls_read"] = True
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

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def safe_float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default

    try:
        if isinstance(value, (int, float, np.number)):
            return float(value)

        if isinstance(value, str):
            cleaned = re.sub(r"[^\d.\-()]", "", value).strip()

            if cleaned.startswith("(") and cleaned.endswith(")"):
                cleaned = "-" + cleaned[1:-1]

            return float(cleaned) if cleaned else default

        return default

    except (ValueError, TypeError):
        return default


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(safe_float(value, default))
    except (ValueError, TypeError):
        return default


def normalize_percent(value: Any, default: float = 0.0, max_value: float = 1.25) -> float:
    x = safe_float(value, default)

    if x > 1.5:
        x = x / 100

    return max(0.0, min(max_value, x))


def generate_id(prefix: str = "deal") -> str:
    timestamp = int(datetime.now(timezone.utc).timestamp())
    return f"{prefix}_{timestamp}_{uuid.uuid4().hex[:12]}"


def clean_currency_string(value: Any) -> float:
    if isinstance(value, (int, float, np.number)):
        return float(value)

    if isinstance(value, str):
        return safe_float(value.replace("$", "").replace(",", "").strip())

    return 0.0


def hash_state(state: Dict[str, Any]) -> str:
    state_json = json.dumps(state, sort_keys=True, default=str)
    return hashlib.sha256(state_json.encode()).hexdigest()


def sanitize_filename(filename: str) -> str:
    path = Path(str(filename))
    name = path.stem or "file"
    suffix = path.suffix[:10]

    name = re.sub(r"[^a-zA-Z0-9_.-]", "_", name)
    name = re.sub(r"_+", "_", name).strip("._-")
    name = name[:100] or "file"

    return f"{name}_{uuid.uuid4().hex[:8]}{suffix}"


def safe_document_path(doc_id: str, filename: str) -> Path:
    DOC_DIR.mkdir(parents=True, exist_ok=True)

    safe_name = sanitize_filename(filename)
    candidate = (DOC_DIR / f"{doc_id}_{safe_name}").resolve()
    base = DOC_DIR.resolve()

    if not str(candidate).startswith(str(base)):
        raise ValueError("Unsafe document path rejected")

    return candidate


def get_current_user() -> str:
    try:
        if hasattr(st, "user") and getattr(st.user, "email", None):
            return str(st.user.email)
    except Exception:
        pass

    try:
        return str(st.secrets.get("APP_USER", os.environ.get("APP_USER", "Local User")))
    except Exception:
        return os.environ.get("APP_USER", "Local User")


def get_secret_value(name: str, default: Optional[str] = None) -> Optional[str]:
    try:
        value = st.secrets.get(name)
        if value:
            return str(value)
    except Exception:
        pass

    return os.environ.get(name, default)


def get_db_encryption_key() -> Optional[str]:
    return get_secret_value("ALENZA_DB_ENCRYPTION_KEY")


def summarize_changes(old_state: dict, new_state: dict) -> str:
    tracked_fields = [
        "deal_name",
        "sponsor",
        "property_address",
        "property_type",
        "transaction_type",
        "lender_profile",
        "purchase_price",
        "appraisal",
        "noi",
        "rate",
        "amort",
        "term",
        "is_io",
        "target_ltv",
        "target_ltc",
        "target_dscr",
        "target_dy",
        "fees",
        "closing_costs",
        "reserves",
        "mezz_debt",
        "pref_equity",
        "mezz_rate",
        "pref_rate",
    ]

    changes = []

    for field in tracked_fields:
        old_val = old_state.get(field) if old_state else None
        new_val = new_state.get(field) if new_state else None

        if str(old_val) != str(new_val):
            changes.append(f"{field}: {old_val} → {new_val}")

    old_rr_len = len(old_state.get("rent_roll_dict", [])) if old_state else 0
    new_rr_len = len(new_state.get("rent_roll_dict", [])) if new_state else 0

    if old_rr_len != new_rr_len:
        changes.append(f"rent_roll_rows: {old_rr_len} → {new_rr_len}")

    return "; ".join(changes[:30]) if changes else "No material changes"


def hash_file_content(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def normalize_rent_roll_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame(columns=DEFAULT_RENT_ROLL_COLUMNS)

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}

    for col in df.columns:
        c = re.sub(r"[^a-z0-9]+", " ", str(col).lower()).strip()

        if c in ["tenant", "tenant name", "lessee", "occupant", "customer", "company", "name"]:
            col_map[col] = "Tenant"
        elif c in ["sf", "sq ft", "sqft", "square feet", "area", "gla", "unit sf", "leased sf"]:
            col_map[col] = "SF"
        elif c in ["remaining term", "term remaining", "lease term remaining", "years remaining", "term yrs", "term years"]:
            col_map[col] = "Remaining Term"
        elif c in ["monthly rent", "rent month", "monthly base rent", "base rent monthly", "rent per month", "monthly revenue"]:
            col_map[col] = "Monthly Rent"
        elif c in ["annual rent", "annual base rent", "yearly rent", "annual revenue", "base rent annual"]:
            col_map[col] = "Annual Rent"

    df = df.rename(columns=col_map)

    if "Annual Rent" in df.columns and "Monthly Rent" not in df.columns:
        df["Monthly Rent"] = pd.to_numeric(df["Annual Rent"], errors="coerce").fillna(0) / 12

    for col in DEFAULT_RENT_ROLL_COLUMNS:
        if col not in df.columns:
            df[col] = "" if col == "Tenant" else 0

    out = df[DEFAULT_RENT_ROLL_COLUMNS].copy()

    out["Tenant"] = out["Tenant"].fillna("").astype(str).str.strip()
    out["SF"] = pd.to_numeric(out["SF"], errors="coerce").fillna(0).clip(lower=0)
    out["Remaining Term"] = pd.to_numeric(out["Remaining Term"], errors="coerce").fillna(0).clip(lower=0)
    out["Monthly Rent"] = pd.to_numeric(out["Monthly Rent"], errors="coerce").fillna(0).clip(lower=0)

    out = out[
        ~(
            (out["Tenant"] == "")
            & (out["SF"] == 0)
            & (out["Monthly Rent"] == 0)
        )
    ]

    return out.reset_index(drop=True)


# ==========================================
# OPTIONAL ENCRYPTION HELPERS
# ==========================================

def encrypt_text(plain_text: str, secret_key: str) -> str:
    import base64
    from cryptography.fernet import Fernet

    key = base64.urlsafe_b64encode(hashlib.sha256(secret_key.encode()).digest())
    fernet = Fernet(key)

    encrypted = fernet.encrypt(plain_text.encode())
    return base64.b64encode(encrypted).decode()


def decrypt_text(encrypted_text: str, secret_key: str) -> str:
    import base64
    from cryptography.fernet import Fernet

    key = base64.urlsafe_b64encode(hashlib.sha256(secret_key.encode()).digest())
    fernet = Fernet(key)

    encrypted = base64.b64decode(encrypted_text)
    return fernet.decrypt(encrypted).decode()


def serialize_state_for_storage(state: dict) -> str:
    state_json = json.dumps(state, default=str)
    encryption_key = get_db_encryption_key()

    if encryption_key and DEPENDENCIES["crypto"]:
        encrypted_payload = encrypt_text(state_json, encryption_key)
        return json.dumps(
            {
                "_alenza_storage": "encrypted",
                "version": 1,
                "payload": encrypted_payload,
            }
        )

    return json.dumps(
        {
            "_alenza_storage": "plain",
            "version": 1,
            "payload": state_json,
        }
    )


def deserialize_state_from_storage(raw: str) -> dict:
    if not raw:
        return {}

    try:
        wrapper = json.loads(raw)

        if isinstance(wrapper, dict) and wrapper.get("_alenza_storage") == "encrypted":
            encryption_key = get_db_encryption_key()

            if not encryption_key:
                raise ValueError("Encrypted deal state requires ALENZA_DB_ENCRYPTION_KEY")

            if not DEPENDENCIES["crypto"]:
                raise ValueError("Encrypted deal state requires cryptography package")

            decrypted_json = decrypt_text(wrapper.get("payload", ""), encryption_key)
            return json.loads(decrypted_json)

        if isinstance(wrapper, dict) and wrapper.get("_alenza_storage") == "plain":
            return json.loads(wrapper.get("payload", "{}"))

        if isinstance(wrapper, dict):
            return wrapper

    except json.JSONDecodeError:
        logger.warning("Legacy non-wrapper JSON detected; attempting direct load")
        return json.loads(raw)

    except Exception as e:
        logger.error(f"Could not deserialize deal state: {e}")
        raise

    return {}


def encrypt_deal_state(state: dict, secret_key: str) -> str:
    try:
        state_json = json.dumps(state, default=str)
        return encrypt_text(state_json, secret_key)
    except ImportError:
        logger.warning("cryptography is not installed; returning unencrypted JSON")
        return json.dumps(state, indent=2, default=str)
    except Exception as e:
        logger.error(f"Encryption failed: {e}")
        return json.dumps(state, indent=2, default=str)


def decrypt_deal_state(encrypted_state: str, secret_key: str) -> dict:
    try:
        return json.loads(decrypt_text(encrypted_state, secret_key))
    except ImportError:
        st.error("Install cryptography to decrypt protected exports.")
        return {}
    except Exception as e:
        logger.error(f"Decryption failed: {e}")
        return {}


def is_encrypted_state(data: str) -> bool:
    try:
        import base64
        decoded = base64.b64decode(data)
        return len(decoded) > 0 and decoded[0] == 0x80
    except Exception:
        return False


# ==========================================
# ERROR HANDLING
# ==========================================

def handle_errors(default_return=None, show_error=True):
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
# AUTOSAVE HELPERS
# ==========================================

def parse_iso_datetime(value: Any) -> Optional[datetime]:
    if not value:
        return None

    try:
        parsed = datetime.fromisoformat(str(value))

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)

        return parsed

    except Exception:
        return None


def seconds_since_iso(value: Any) -> Optional[float]:
    parsed = parse_iso_datetime(value)

    if parsed is None:
        return None

    return (datetime.now(timezone.utc) - parsed).total_seconds()


def save_current_deal(reason: str = "MANUAL_SAVE", silent: bool = False) -> bool:
    """
    Centralized save path used by manual save, autosave, and keyboard shortcut save.
    Returns True only when the deal was actually saved.
    """
    state = extract_clean_state()
    errors, warnings = ValidationEngine.validate(state)

    if errors:
        st.session_state.autosave_status = f"Save blocked: {len(errors)} validation error(s)"

        if not silent:
            st.error("Deal has blocking validation errors.")
            for error in errors:
                st.error(f"• {error}")

        return False

    deal_id = state.get("deal_id")
    deal_name = state.get("deal_name", "Untitled Deal")

    if not deal_id:
        st.session_state.autosave_status = "Save blocked: missing deal ID"
        return False

    saved = DatabaseManager.save_deal(deal_id, deal_name, state)

    if saved:
        now = utc_now_iso()

        st.session_state.unsaved_changes = False
        st.session_state.last_saved_at = now

        if reason == "AUTOSAVE":
            st.session_state.last_autosaved_at = now
            st.session_state.autosave_status = f"Autosaved at {now[:19]}"
        else:
            st.session_state.autosave_status = f"Saved at {now[:19]}"

        DatabaseManager.log_audit(
            action=reason,
            details=f"{reason} completed for {deal_name}",
            deal_id=deal_id,
        )

        return True

    st.session_state.autosave_status = "Save failed"

    if not silent:
        st.error("Save failed.")

    return False


def maybe_autosave_current_deal() -> bool:
    """
    Saves the current deal when:
    - autosave is enabled
    - there are unsaved changes
    - at least AUTOSAVE_INTERVAL_SECONDS has passed since last autosave/save
    """
    if not st.session_state.get("autosave_enabled", True):
        return False

    if not st.session_state.get("unsaved_changes", False):
        return False

    last_checkpoint = (
        st.session_state.get("last_autosaved_at")
        or st.session_state.get("last_saved_at")
    )

    elapsed = seconds_since_iso(last_checkpoint)

    if elapsed is not None and elapsed < AUTOSAVE_INTERVAL_SECONDS:
        return False

    return save_current_deal(reason="AUTOSAVE", silent=True)


# ==========================================
# DATABASE
# ==========================================

def get_db_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DOC_DIR.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row

    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")

    return conn


class DatabaseManager:
    @staticmethod
    @handle_errors(default_return=False, show_error=False)
    def init_db() -> bool:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        DOC_DIR.mkdir(parents=True, exist_ok=True)

        with get_db_connection() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS deals (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    state_json TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
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
                    original_filename TEXT,
                    category TEXT NOT NULL,
                    path TEXT NOT NULL,
                    file_size INTEGER,
                    file_hash TEXT,
                    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (deal_id) REFERENCES deals(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    deal_id TEXT,
                    user TEXT,
                    action TEXT NOT NULL,
                    details TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS app_metadata (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );

                CREATE INDEX IF NOT EXISTS idx_deals_updated
                    ON deals(updated_at DESC);

                CREATE INDEX IF NOT EXISTS idx_versions_deal
                    ON deal_versions(deal_id, created_at DESC);

                CREATE INDEX IF NOT EXISTS idx_docs_deal
                    ON documents(deal_id);

                CREATE INDEX IF NOT EXISTS idx_audit_log_time
                    ON audit_log(timestamp DESC);

                CREATE INDEX IF NOT EXISTS idx_audit_log_deal
                    ON audit_log(deal_id);
            """)

            conn.execute(
                """
                INSERT OR REPLACE INTO app_metadata (key, value, updated_at)
                VALUES (?, ?, ?)
                """,
                ("schema_version", str(SCHEMA_VERSION), utc_now_iso()),
            )

            conn.execute(
                """
                INSERT OR REPLACE INTO app_metadata (key, value, updated_at)
                VALUES (?, ?, ?)
                """,
                ("app_version", APP_VERSION, utc_now_iso()),
            )

            conn.commit()

        return True

    @staticmethod
    @handle_errors(default_return=False, show_error=False)
    def log_audit(action: str, details: str = "", deal_id: Optional[str] = None) -> bool:
        with get_db_connection() as conn:
            conn.execute(
                """
                INSERT INTO audit_log (deal_id, user, action, details, timestamp)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    deal_id,
                    get_current_user(),
                    str(action)[:100],
                    str(details)[:1500],
                    utc_now_iso(),
                ),
            )

            conn.commit()

        return True

    @staticmethod
    @handle_errors(default_return=False)
    def save_deal(deal_id: str, name: str, state: dict) -> bool:
        safe_name = str(name or "Untitled Deal").strip()[:200] or "Untitled Deal"
        now = utc_now_iso()

        state = normalize_loaded_state(state)
        state["last_saved_at"] = now
        state["schema_version"] = SCHEMA_VERSION
        state["app_version"] = APP_VERSION

        state_storage = serialize_state_for_storage(state)

        with get_db_connection() as conn:
            old_state = {}

            row = conn.execute(
                "SELECT state_json FROM deals WHERE id = ?",
                (deal_id,),
            ).fetchone()

            if row:
                try:
                    old_state = deserialize_state_from_storage(row["state_json"])
                except Exception as e:
                    logger.warning(f"Could not parse old state for change summary: {e}")

            change_summary = summarize_changes(old_state, state)

            conn.execute(
                """
                INSERT OR REPLACE INTO deals (id, name, state_json, updated_at, created_at)
                VALUES (?, ?, ?, ?, COALESCE((SELECT created_at FROM deals WHERE id = ?), ?))
                """,
                (deal_id, safe_name, state_storage, now, deal_id, now),
            )

            conn.execute(
                """
                INSERT INTO deal_versions (deal_id, state_json, changed_by, change_summary, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    deal_id,
                    state_storage,
                    get_current_user(),
                    change_summary,
                    now,
                ),
            )

            conn.commit()

        DatabaseManager.log_audit(
            action="SAVE_DEAL",
            details=f"Saved deal: {safe_name}. Changes: {change_summary}",
            deal_id=deal_id,
        )

        logger.info(f"Deal saved: {safe_name}")
        return True

    @staticmethod
    @handle_errors(default_return=None)
    def load_deal(deal_id: str) -> Optional[dict]:
        with get_db_connection() as conn:
            row = conn.execute(
                "SELECT state_json FROM deals WHERE id = ?",
                (deal_id,),
            ).fetchone()

            if not row:
                return None

            state = deserialize_state_from_storage(row["state_json"])
            state = normalize_loaded_state(state)

        DatabaseManager.log_audit(
            action="LOAD_DEAL",
            details=f"Loaded deal ID: {deal_id}",
            deal_id=deal_id,
        )

        return state

    @staticmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_all_deals() -> pd.DataFrame:
        with get_db_connection() as conn:
            return pd.read_sql_query(
                """
                SELECT id, name, created_at, updated_at
                FROM deals
                ORDER BY updated_at DESC
                """,
                conn,
            )

    @staticmethod
    @handle_errors(default_return=False)
    def delete_deal(deal_id: str) -> bool:
        DatabaseManager.log_audit(
            action="DELETE_DEAL",
            details=f"Deleted deal ID: {deal_id}",
            deal_id=deal_id,
        )

        with get_db_connection() as conn:
            cursor = conn.execute(
                "SELECT path FROM documents WHERE deal_id = ?",
                (deal_id,),
            )

            for row in cursor.fetchall():
                try:
                    Path(row["path"]).unlink(missing_ok=True)
                except Exception as e:
                    logger.warning(f"Could not delete document file {row['path']}: {e}")

            conn.execute("DELETE FROM documents WHERE deal_id = ?", (deal_id,))
            conn.execute("DELETE FROM deal_versions WHERE deal_id = ?", (deal_id,))
            conn.execute("DELETE FROM deals WHERE id = ?", (deal_id,))
            conn.commit()

        return True

    @staticmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_deal_versions(deal_id: str, limit: int = 20) -> pd.DataFrame:
        with get_db_connection() as conn:
            return pd.read_sql_query(
                """
                SELECT changed_by, change_summary, created_at
                FROM deal_versions
                WHERE deal_id = ?
                ORDER BY created_at DESC
                LIMIT ?
                """,
                conn,
                params=(deal_id, limit),
            )

    @staticmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_audit_log(deal_id: Optional[str] = None, limit: int = 100) -> pd.DataFrame:
        with get_db_connection() as conn:
            if deal_id:
                return pd.read_sql_query(
                    """
                    SELECT user, action, details, timestamp
                    FROM audit_log
                    WHERE deal_id = ?
                    ORDER BY timestamp DESC
                    LIMIT ?
                    """,
                    conn,
                    params=(deal_id, limit),
                )

            return pd.read_sql_query(
                """
                SELECT deal_id, user, action, details, timestamp
                FROM audit_log
                ORDER BY timestamp DESC
                LIMIT ?
                """,
                conn,
                params=(limit,),
            )

    @staticmethod
    @handle_errors(default_return=pd.DataFrame())
    def get_deal_documents(deal_id: str) -> pd.DataFrame:
        with get_db_connection() as conn:
            return pd.read_sql_query(
                """
                SELECT id, filename, original_filename, category, file_size, uploaded_at
                FROM documents
                WHERE deal_id = ?
                ORDER BY uploaded_at DESC
                """,
                conn,
                params=(deal_id,),
            )

    @staticmethod
    @handle_errors(default_return=False)
    def save_document(deal_id: str, file, category: str) -> bool:
        file_content = file.getvalue()

        if len(file_content) > MAX_UPLOAD_SIZE_MB * 1024 * 1024:
            raise ValueError(f"File exceeds {MAX_UPLOAD_SIZE_MB}MB limit")

        doc_id = generate_id("doc")
        file_path = safe_document_path(doc_id, file.name)

        file_path.write_bytes(file_content)

        safe_name = Path(file_path).name.replace(f"{doc_id}_", "", 1)
        file_hash = hash_file_content(file_content)

        with get_db_connection() as conn:
            conn.execute(
                """
                INSERT INTO documents (
                    id,
                    deal_id,
                    filename,
                    original_filename,
                    category,
                    path,
                    file_size,
                    file_hash,
                    uploaded_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    doc_id,
                    deal_id,
                    safe_name,
                    file.name,
                    str(category)[:100],
                    str(file_path),
                    len(file_content),
                    file_hash,
                    utc_now_iso(),
                ),
            )

            conn.commit()

        DatabaseManager.log_audit(
            action="DOC_UPLOAD",
            details=f"Uploaded {safe_name} to {category}",
            deal_id=deal_id,
        )

        return True

    @staticmethod
    @handle_errors(default_return=False)
    def delete_document(doc_id: str) -> bool:
        deal_id = None
        filename = None

        with get_db_connection() as conn:
            row = conn.execute(
                "SELECT deal_id, path, filename FROM documents WHERE id = ?",
                (doc_id,),
            ).fetchone()

            if not row:
                return False

            deal_id = row["deal_id"]
            filename = row["filename"]

            try:
                Path(row["path"]).unlink(missing_ok=True)
            except Exception as e:
                logger.warning(f"Could not delete document file {row['path']}: {e}")

            conn.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
            conn.commit()

        DatabaseManager.log_audit(
            action="DOC_DELETE",
            details=f"Deleted document: {filename}",
            deal_id=deal_id,
        )

        return True


# ==========================================
# FINANCIAL ENGINE
# ==========================================

class UnderwritingEngine:
    LENDER_LIMITS = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05},
    }

    @staticmethod
    def size_loan(noi, appraisal, purchase_price, closing_costs, reserves, fees_pct, rate, amort, term, is_io, target_ltv, target_ltc, target_dscr, target_dy) -> tuple:
        noi = max(0, safe_float(noi))
        appraisal = max(0, safe_float(appraisal))
        purchase_price = max(0, safe_float(purchase_price))
        closing_costs = max(0, safe_float(closing_costs))
        reserves = max(0, safe_float(reserves))
        fees_pct = normalize_percent(fees_pct, 0.0, 0.10)
        rate = max(0.001, normalize_percent(rate, 0.0525, 0.30))
        amort = max(1, min(40, safe_int(amort)))
        target_ltv = max(0.01, normalize_percent(target_ltv, 0.75, 1.25))
        target_ltc = max(0.01, normalize_percent(target_ltc, 0.80, 1.25))
        target_dscr = max(0.01, safe_float(target_dscr))
        target_dy = max(0.0001, normalize_percent(target_dy, 0.085, 0.25))

        base_cost = purchase_price + closing_costs + reserves
        gates = {"LTV": 0.0, "LTC": 0.0, "DSCR": 0.0, "Debt Yield": 0.0}
        loan = 0.0
        total_uses = base_cost

        for _ in range(12):
            total_uses = base_cost + (loan * fees_pct)
            gates["LTV"] = appraisal * target_ltv if appraisal > 0 else 0.0
            gates["LTC"] = total_uses * target_ltc if total_uses > 0 else 0.0
            gates["Debt Yield"] = noi / target_dy if target_dy > 0 else 0.0
            monthly_rate = rate / 12

            if monthly_rate > 0 and noi > 0:
                if is_io:
                    gates["DSCR"] = (noi / target_dscr) / 12 / monthly_rate
                else:
                    total_payments = amort * 12
                    pmt_factor = (1 - (1 + monthly_rate) ** -total_payments) / monthly_rate
                    gates["DSCR"] = (noi / target_dscr) / 12 * pmt_factor if pmt_factor > 0 else 0.0
            else:
                gates["DSCR"] = 0.0

            new_loan = min(gates.values())
            if abs(new_loan - loan) < 1.0:
                loan = max(0.0, new_loan)
                total_uses = base_cost + (loan * fees_pct)
                break
            loan = max(0.0, new_loan)

        gate = min(gates, key=gates.get) if gates else "N/A"
        return loan, gate, gates, total_uses, total_uses - loan

    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io) -> tuple:
        loan_amt = max(0, safe_float(loan_amt))
        rate = normalize_percent(rate, 0.0525, 0.30)
        amort_yrs = max(1, min(40, safe_int(amort_yrs)))
        term_yrs = max(1, min(40, safe_int(term_yrs)))

        if loan_amt <= 0:
            return pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]), 0.0, 0.0

        monthly_rate = rate / 12
        total_payments = amort_yrs * 12
        term_months = int(term_yrs * 12)

        if is_io:
            monthly_pmt = loan_amt * monthly_rate
        elif monthly_rate > 0:
            monthly_pmt = (loan_amt * monthly_rate) / (1 - (1 + monthly_rate) ** -total_payments)
        else:
            monthly_pmt = loan_amt / total_payments

        schedule = []
        balance = loan_amt

        for period in range(1, term_months + 1):
            interest = balance * monthly_rate
            principal_paid = 0.0 if is_io else min(max(monthly_pmt - interest, 0.0), balance)
            balance = max(0.0, balance - principal_paid)
            schedule.append({"Period": period, "Payment": monthly_pmt if balance > 0 else 0.0, "Principal": principal_paid, "Interest": interest, "Balance": balance})
            if balance <= 0:
                break

        return pd.DataFrame(schedule), monthly_pmt, balance

    @staticmethod
    def rent_roll_metrics(df) -> tuple:
        df = normalize_rent_roll_columns(df)
        if df.empty:
            return 0, 0, 0, 0, 0, 0

        total_sf = df["SF"].sum()
        if total_sf <= 0:
            return 0, 0, 0, 0, 0, 0

        vacant_keywords = ["vacant", "empty", "available", "vacancy", "n/a", "none", ""]
        occupied = df[(~df["Tenant"].str.lower().isin(vacant_keywords)) & (df["SF"] > 0)]
        occupied_sf = occupied["SF"].sum()
        if occupied_sf <= 0:
            return total_sf, 0, 0, 0, 0, 0

        occupancy = min(1.0, occupied_sf / total_sf)
        annual_rent = occupied["Monthly Rent"].sum() * 12
        rent_psf = annual_rent / occupied_sf if occupied_sf > 0 else 0
        walt = (occupied["Remaining Term"] * occupied["SF"]).sum() / occupied_sf if occupied_sf > 0 else 0
        expiring_sf = occupied[occupied["Remaining Term"] <= 1.0]["SF"].sum()
        rollover = min(1.0, expiring_sf / occupied_sf) if occupied_sf > 0 else 0
        return total_sf, occupancy, annual_rent, rent_psf, walt, rollover

    @staticmethod
    def breakeven_occupancy(current_noi: float, current_occupancy: float, annual_debt_service: float) -> float:
        current_noi = safe_float(current_noi)
        current_occupancy = safe_float(current_occupancy)
        annual_debt_service = safe_float(annual_debt_service)
        if current_noi <= 0 or current_occupancy <= 0 or annual_debt_service <= 0:
            return 0.0
        return max(0.0, min(1.5, annual_debt_service / current_noi * current_occupancy))

    @staticmethod
    def capital_stack(senior_debt: float, mezz_debt: float, pref_equity: float, sponsor_equity: float, noi: float, senior_rate: float, mezz_rate: float, pref_rate: float) -> dict:
        senior_debt = max(0, safe_float(senior_debt))
        mezz_debt = max(0, safe_float(mezz_debt))
        pref_equity = max(0, safe_float(pref_equity))
        sponsor_equity = max(0, safe_float(sponsor_equity))
        noi = max(0, safe_float(noi))
        senior_rate = normalize_percent(senior_rate, 0.0525, 0.30)
        mezz_rate = normalize_percent(mezz_rate, 0.10, 0.40)
        pref_rate = normalize_percent(pref_rate, 0.09, 0.40)
        senior_cost = senior_debt * senior_rate
        mezz_cost = mezz_debt * mezz_rate
        pref_cost = pref_equity * pref_rate
        fixed_charges = senior_cost + mezz_cost + pref_cost
        total_capital = senior_debt + mezz_debt + pref_equity + sponsor_equity
        return {
            "Senior Debt": senior_debt,
            "Mezzanine Debt": mezz_debt,
            "Preferred Equity": pref_equity,
            "Sponsor Equity": sponsor_equity,
            "Total Capital": total_capital,
            "Senior Cost": senior_cost,
            "Mezzanine Cost": mezz_cost,
            "Preferred Cost": pref_cost,
            "Fixed Charges": fixed_charges,
            "Fixed Charge Coverage": noi / fixed_charges if fixed_charges > 0 else 0,
        }

    @staticmethod
    def score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, profile) -> tuple:
        limits = UnderwritingEngine.LENDER_LIMITS.get(profile, UnderwritingEngine.LENDER_LIMITS["Bank / Credit Union"])
        actual_ltv = max(0, safe_float(actual_ltv))
        actual_ltc = max(0, safe_float(actual_ltc))
        actual_dscr = max(0, safe_float(actual_dscr))
        actual_dy = max(0, safe_float(actual_dy))
        ltv_score = max(0, 300 * (1 - actual_ltv / limits["max_ltv"])) if limits["max_ltv"] > 0 else 0
        dscr_score = max(0, 300 * (actual_dscr - 1.0) / (limits["min_dscr"] - 1.0)) if limits["min_dscr"] > 1.0 and actual_dscr > 1.0 else (300 if actual_dscr >= 1.0 else 0)
        dy_score = max(0, 200 * min(1.5, actual_dy / limits["min_dy"])) if limits["min_dy"] > 0 else 0
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


# ==========================================
# STATE MANAGEMENT
# ==========================================

def default_state() -> dict:
    return {
        "deal_id": generate_id("deal"),
        "deal_name": "Untitled Deal",
        "sponsor": "",
        "property_address": "",
        "property_type": "Multifamily",
        "transaction_type": "Acquisition",
        "lender_profile": "Bank / Credit Union",
        "purchase_price": 0.0,
        "appraisal": 0.0,
        "noi": 0.0,
        "target_ltv": 0.75,
        "target_ltc": 0.80,
        "target_dscr": 1.25,
        "target_dy": 0.085,
        "rate": 0.0525,
        "amort": 25,
        "term": 5,
        "is_io": False,
        "fees": 0.02,
        "closing_costs": 0.0,
        "reserves": 0.0,
        "mezz_debt": 0.0,
        "pref_equity": 0.0,
        "mezz_rate": 0.10,
        "pref_rate": 0.09,
        "rent_roll_dict": DEFAULT_RENT_ROLL.copy(),
        "last_saved_at": None,
        "unsaved_changes": False,
        "schema_version": SCHEMA_VERSION,
        "app_version": APP_VERSION,
    }


def normalize_loaded_state(state: dict) -> dict:
    normalized = default_state()
    if isinstance(state, dict):
        normalized.update(state)

    normalized["deal_id"] = str(normalized.get("deal_id") or generate_id("deal"))
    normalized["deal_name"] = str(normalized.get("deal_name") or "Untitled Deal").strip() or "Untitled Deal"
    normalized["sponsor"] = str(normalized.get("sponsor") or "")
    normalized["property_address"] = str(normalized.get("property_address") or "")

    if normalized.get("property_type") not in PROPERTY_TYPES:
        normalized["property_type"] = "Multifamily"
    if normalized.get("transaction_type") not in TRANSACTION_TYPES:
        normalized["transaction_type"] = "Acquisition"
    if normalized.get("lender_profile") not in LENDER_PROFILES_LIST:
        normalized["lender_profile"] = "Bank / Credit Union"

    normalized["purchase_price"] = max(0, safe_float(normalized.get("purchase_price")))
    normalized["appraisal"] = max(0, safe_float(normalized.get("appraisal")))
    normalized["noi"] = max(0, safe_float(normalized.get("noi")))
    normalized["target_ltv"] = normalize_percent(normalized.get("target_ltv"), 0.75, 1.25)
    normalized["target_ltc"] = normalize_percent(normalized.get("target_ltc"), 0.80, 1.25)
    normalized["target_dscr"] = max(0.01, safe_float(normalized.get("target_dscr"), 1.25))
    normalized["target_dy"] = normalize_percent(normalized.get("target_dy"), 0.085, 0.25)
    normalized["rate"] = normalize_percent(normalized.get("rate"), 0.0525, 0.30)
    normalized["amort"] = max(1, min(40, safe_int(normalized.get("amort"), 25)))
    normalized["term"] = max(1, min(40, safe_int(normalized.get("term"), 5)))
    normalized["is_io"] = bool(normalized.get("is_io", False))
    normalized["fees"] = normalize_percent(normalized.get("fees"), 0.02, 0.10)
    normalized["closing_costs"] = max(0, safe_float(normalized.get("closing_costs")))
    normalized["reserves"] = max(0, safe_float(normalized.get("reserves")))
    normalized["mezz_debt"] = max(0, safe_float(normalized.get("mezz_debt")))
    normalized["pref_equity"] = max(0, safe_float(normalized.get("pref_equity")))
    normalized["mezz_rate"] = normalize_percent(normalized.get("mezz_rate"), 0.10, 0.40)
    normalized["pref_rate"] = normalize_percent(normalized.get("pref_rate"), 0.09, 0.40)

    rr = normalized.get("rent_roll_dict", [])
    if isinstance(rr, dict):
        rr = [rr]
    normalized["rent_roll_dict"] = normalize_rent_roll_columns(pd.DataFrame(rr)).to_dict("records")
    normalized["schema_version"] = SCHEMA_VERSION
    normalized["app_version"] = APP_VERSION
    return normalized


def initialize_session_state():
    for key, value in default_state().items():
        if key not in st.session_state:
            st.session_state[key] = value


def extract_clean_state() -> dict:
    keys = list(default_state().keys())
    state = {}
    for key in keys:
        if key in st.session_state:
            val = st.session_state[key]
            if isinstance(val, np.integer):
                val = int(val)
            elif isinstance(val, np.floating):
                val = float(val)
            elif isinstance(val, np.ndarray):
                val = val.tolist()
            state[key] = val
    state = normalize_loaded_state(state)
    state["schema_version"] = SCHEMA_VERSION
    state["app_version"] = APP_VERSION
    return state


def stable_state_hash() -> str:
    state = extract_clean_state()
    for key in ["last_saved_at", "unsaved_changes", "app_version", "schema_version"]:
        state.pop(key, None)
    return hash_state(state)


class SensitivityEngine:
    @staticmethod
    def generate_matrix(state: dict) -> pd.DataFrame:
        scenarios = [("Base", 0, 0), ("Rate +1%", 0.01, 0), ("Rate -1%", -0.01, 0), ("NOI -10%", 0, -0.10), ("NOI +10%", 0, 0.10), ("Combined Stress", 0.01, -0.10)]
        results = []
        base_proceeds = None
        for scenario_name, rate_adj, noi_adj in scenarios:
            adjusted_rate = max(0.001, safe_float(state.get("rate", 0.0525)) + rate_adj)
            adjusted_noi = max(0, safe_float(state.get("noi", 0)) * (1 + noi_adj))
            loan, gate, _, _, _ = UnderwritingEngine.size_loan(
                adjusted_noi, safe_float(state.get("appraisal", 0)), safe_float(state.get("purchase_price", 0)),
                safe_float(state.get("closing_costs", 0)), safe_float(state.get("reserves", 0)), safe_float(state.get("fees", 0)), adjusted_rate,
                safe_int(state.get("amort", 25)), safe_int(state.get("term", 5)), bool(state.get("is_io", False)),
                safe_float(state.get("target_ltv", 0.75)), safe_float(state.get("target_ltc", 0.80)), safe_float(state.get("target_dscr", 1.25)), safe_float(state.get("target_dy", 0.085))
            )
            if scenario_name == "Base":
                base_proceeds = loan
            change_str = f"{((loan - base_proceeds) / base_proceeds * 100):+.1f}%" if base_proceeds and base_proceeds > 0 else "N/A"
            results.append({"Scenario": scenario_name, "Rate": f"{adjusted_rate * 100:.2f}%", "NOI": f"${adjusted_noi:,.0f}", "Max Proceeds": f"${loan:,.0f}", "Constraint": gate, "Change from Base": change_str})
        return pd.DataFrame(results)

    @staticmethod
    def proceeds_heatmap(state: dict) -> pd.DataFrame:
        rate_shocks = [-0.01, -0.005, 0, 0.005, 0.01]
        noi_shocks = [-0.10, -0.05, 0, 0.05, 0.10]
        rows = []
        for noi_adj in noi_shocks:
            row = {"NOI Shock": f"{noi_adj:+.0%}"}
            for rate_adj in rate_shocks:
                adjusted_rate = max(0.001, safe_float(state.get("rate", 0.0525)) + rate_adj)
                adjusted_noi = max(0, safe_float(state.get("noi", 0)) * (1 + noi_adj))
                loan, _, _, _, _ = UnderwritingEngine.size_loan(
                    adjusted_noi, safe_float(state.get("appraisal", 0)), safe_float(state.get("purchase_price", 0)), safe_float(state.get("closing_costs", 0)), safe_float(state.get("reserves", 0)), safe_float(state.get("fees", 0)), adjusted_rate,
                    safe_int(state.get("amort", 25)), safe_int(state.get("term", 5)), bool(state.get("is_io", False)), safe_float(state.get("target_ltv", 0.75)), safe_float(state.get("target_ltc", 0.80)), safe_float(state.get("target_dscr", 1.25)), safe_float(state.get("target_dy", 0.085))
                )
                row[f"Rate {rate_adj:+.1%}"] = loan
            rows.append(row)
        return pd.DataFrame(rows)


class ValidationEngine:
    @staticmethod
    def validate(state: dict) -> tuple:
        errors, warnings = [], []
        deal_name = str(state.get("deal_name", "")).strip()
        property_address = str(state.get("property_address", "")).strip()
        purchase_price = safe_float(state.get("purchase_price"))
        appraisal = safe_float(state.get("appraisal"))
        noi = safe_float(state.get("noi"))
        rate = safe_float(state.get("rate"))
        target_ltv = safe_float(state.get("target_ltv"))
        target_ltc = safe_float(state.get("target_ltc"))
        target_dscr = safe_float(state.get("target_dscr"))
        target_dy = safe_float(state.get("target_dy"))
        amort = safe_int(state.get("amort"))
        term = safe_int(state.get("term"))
        if not deal_name:
            errors.append("Deal name is required")
        if purchase_price > 0 and not property_address:
            warnings.append("Property address recommended")
        if purchase_price < 0:
            errors.append("Purchase price cannot be negative")
        if appraisal < 0:
            errors.append("Appraisal cannot be negative")
        if noi < 0:
            errors.append("NOI cannot be negative")
        if rate <= 0 and (purchase_price > 0 or noi > 0):
            errors.append("Interest rate must be positive")
        if rate > 0.30:
            errors.append("Interest rate cannot exceed 30%")
        if target_ltv <= 0:
            errors.append("Target LTV must be positive")
        if target_ltv > 1.25:
            errors.append("Target LTV cannot exceed 125%")
        if target_ltc <= 0:
            errors.append("Target LTC must be positive")
        if target_ltc > 1.25:
            warnings.append("Target LTC above 125% is unusually aggressive")
        if target_dscr <= 0:
            errors.append("Target DSCR must be positive")
        if target_dscr < 1.0:
            warnings.append("Target DSCR below 1.00x implies no cash-flow cushion")
        if target_dy <= 0:
            errors.append("Debt yield must be positive")
        if amort < 1:
            errors.append("Amortization must be at least 1 year")
        elif amort > 40:
            errors.append("Amortization cannot exceed 40 years")
        if term < 1:
            errors.append("Term must be at least 1 year")
        elif term > 40:
            errors.append("Term cannot exceed 40 years")
        if noi > 0 and appraisal > 0:
            cap_rate = noi / appraisal
            if 0 < cap_rate < 0.01:
                warnings.append(f"Implied cap rate {cap_rate:.2%} unusually low")
            elif cap_rate > 0.20:
                warnings.append(f"Implied cap rate {cap_rate:.2%} unusually high")
        if purchase_price > 0 and appraisal > 0:
            premium = (purchase_price - appraisal) / appraisal
            if premium > 0.30:
                warnings.append(f"Purchase price {premium:.1%} above appraisal")
            elif premium < -0.30:
                warnings.append(f"Purchase price is {abs(premium):.1%} below appraisal; verify values")
        rent_roll = state.get("rent_roll_dict", [])
        if isinstance(rent_roll, list) and rent_roll:
            tenant_names = []
            for idx, row in enumerate(rent_roll, start=1):
                tenant = str(row.get("Tenant", "")).strip()
                sf = safe_float(row.get("SF"))
                rent = safe_float(row.get("Monthly Rent"))
                term_remaining = safe_float(row.get("Remaining Term"))
                if tenant:
                    tenant_names.append(tenant.lower())
                if rent > 0 and sf <= 0:
                    warnings.append(f"Rent roll row {idx}: rent entered with zero SF")
                if term_remaining > 25:
                    warnings.append(f"Rent roll row {idx}: lease term above 25 years")
                if sf > 1_000_000:
                    warnings.append(f"Rent roll row {idx}: SF appears unusually high")
            duplicates = sorted({t for t in tenant_names if tenant_names.count(t) > 1 and t not in ["vacant", "empty", "available", "vacancy"]})
            if duplicates:
                warnings.append(f"Duplicate tenant names detected: {', '.join(duplicates[:5])}")
        return errors, warnings


def run_financial_self_tests() -> pd.DataFrame:
    rows = []
    def check(name: str, passed: bool, detail: str):
        rows.append({"Test": name, "Status": "✅ PASS" if passed else "❌ FAIL", "Detail": detail})
    loan, gate, gates, uses, equity = UnderwritingEngine.size_loan(1_000_000, 10_000_000, 9_000_000, 100_000, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
    check("Loan proceeds positive", loan > 0, f"${loan:,.0f}, Gate: {gate}")
    check("Binding gate equals minimum proceeds", abs(loan - min(gates.values())) < 1, f"Min: ${min(gates.values()):,.0f}")
    check("Total uses positive", uses > 0, f"${uses:,.0f}")
    check("Equity requirement calculated", isinstance(equity, float), f"${equity:,.0f}")
    amort, pmt, balloon = UnderwritingEngine.amort_schedule(1_000_000, 0.06, 25, 5, False)
    check("60 periods in 5-year term", len(amort) == 60, f"Periods: {len(amort)}")
    check("Monthly payment positive", pmt > 0, f"${pmt:,.2f}")
    check("Amortizing balloon below original", balloon < 1_000_000, f"Balloon: ${balloon:,.0f}")
    io_amort, io_pmt, io_balloon = UnderwritingEngine.amort_schedule(1_000_000, 0.06, 25, 5, True)
    check("IO payment equals interest-only amount", abs(io_pmt - 5_000) < 1, f"PMT: ${io_pmt:,.2f}")
    check("IO balloon equals original loan", abs(io_balloon - 1_000_000) < 1, f"Balloon: ${io_balloon:,.0f}")
    rr = pd.DataFrame([{"Tenant": "A", "SF": 10_000, "Remaining Term": 5, "Monthly Rent": 20_000}, {"Tenant": "Vacant", "SF": 2_000, "Remaining Term": 0, "Monthly Rent": 0}])
    sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr)
    check("Occupancy equals 83.3%", abs(occ - 0.8333) < 0.01, f"{occ:.2%}")
    check("Annual rent equals $240K", abs(ann_rent - 240_000) < 1, f"${ann_rent:,.0f}")
    check("Rent PSF equals $24", abs(psf - 24) < 0.1, f"${psf:.2f}")
    check("WALT equals 5.0 years", abs(walt - 5.0) < 0.01, f"{walt:.2f}")
    zero_loan, zero_gate, zero_gates, zero_uses, zero_equity = UnderwritingEngine.size_loan(0, 10_000_000, 9_000_000, 0, 0, 0.01, 0.06, 25, 5, False, 0.70, 0.80, 1.25, 0.09)
    check("Zero NOI produces zero income-based proceeds", zero_gates["DSCR"] == 0 and zero_gates["Debt Yield"] == 0 and zero_loan == 0, f"Loan: ${zero_loan:,.0f}")
    empty_sf, empty_occ, empty_rent, empty_psf, empty_walt, empty_exp = UnderwritingEngine.rent_roll_metrics(pd.DataFrame())
    check("Empty rent roll returns zero metrics", empty_sf == 0 and empty_occ == 0 and empty_rent == 0, "Empty rent roll handled safely")
    neg_loan, neg_gate, neg_gates, neg_uses, neg_equity = UnderwritingEngine.size_loan(-1_000_000, -10_000_000, -9_000_000, -100_000, -50_000, -0.01, -0.05, -25, -5, False, -0.70, -0.80, -1.25, -0.09)
    check("Negative inputs are sanitized", neg_loan >= 0 and neg_uses >= 0, f"Loan: ${neg_loan:,.0f}, Uses: ${neg_uses:,.0f}")
    be_occ = UnderwritingEngine.breakeven_occupancy(1_000_000, 0.90, 750_000)
    check("Breakeven occupancy calculation", abs(be_occ - 0.675) < 0.001, f"{be_occ:.2%}")
    stack = UnderwritingEngine.capital_stack(7_000_000, 1_000_000, 500_000, 1_500_000, 1_000_000, 0.06, 0.11, 0.09)
    check("Capital stack fixed charge coverage positive", stack["Fixed Charge Coverage"] > 0, f"{stack['Fixed Charge Coverage']:.2f}x")
    return pd.DataFrame(rows)


DatabaseManager.init_db()


# ==========================================
# MAIN APPLICATION - PART 2
# ==========================================

def main():
    """Main application entry point."""

    DatabaseManager.init_db()
    initialize_session_state()
    s = st.session_state
    if "_loading_deal" not in s:
        s["_loading_deal"] = False
    state_hash_before = stable_state_hash()

    with st.sidebar:
        st.title("🏛️ ALENZA OS")
        st.caption(f"v{APP_VERSION}")
        with st.expander("📁 DEAL MANAGER", expanded=True):
            new_name = st.text_input("New Deal Name", value="Untitled Deal", key="new_deal_name")
            if st.button("➕ New Deal", use_container_width=True, key="btn_new_deal"):
                fresh = default_state()
                fresh["deal_id"] = generate_id("deal")
                fresh["deal_name"] = new_name.strip() or "Untitled Deal"
                fresh["unsaved_changes"] = True
                fresh["last_saved_at"] = None
                for k, v in fresh.items():
                    s[k] = v
                st.rerun()

            st.markdown("---")
            deals_df = DatabaseManager.get_all_deals()
            if not deals_df.empty:
                deal_options = {}
                for _, row in deals_df.iterrows():
                    deal_id = str(row["id"])
                    name = str(row.get("name", "Untitled"))[:50]
                    updated = str(row.get("updated_at", ""))[:19]
                    label = f"{name} · {updated} · {deal_id[-8:]}"
                    deal_options[label] = deal_id
                selected_label = st.selectbox("Existing Deal", list(deal_options.keys()), key="load_deal_select")
                col_load, col_dup = st.columns(2)
                with col_load:
                    if st.button("📂 Load", use_container_width=True, key="btn_load"):
                        deal_id = deal_options[selected_label]
                        state = DatabaseManager.load_deal(deal_id)
                        if state:
                            s["_loading_deal"] = True
                            state = normalize_loaded_state(state)
                            for k, v in state.items():
                                s[k] = v
                            s.unsaved_changes = False
                            s["_loading_deal"] = False
                            st.rerun()
                        else:
                            st.error("Could not load selected deal.")
                with col_dup:
                    if st.button("📑 Duplicate", use_container_width=True, key="btn_duplicate"):
                        deal_id = deal_options[selected_label]
                        state = DatabaseManager.load_deal(deal_id)
                        if state:
                            copied = normalize_loaded_state(state)
                            copied["deal_id"] = generate_id("deal")
                            copied["deal_name"] = f"Copy of {copied.get('deal_name', 'Untitled Deal')}"
                            copied["last_saved_at"] = None
                            copied["unsaved_changes"] = True
                            for k, v in copied.items():
                                s[k] = v
                            st.success("Duplicated in memory. Click Save Deal to persist.")
                            st.rerun()
                        else:
                            st.error("Could not duplicate selected deal.")
                confirm_delete = st.checkbox("Confirm delete selected deal", key="confirm_delete_deal")
                if st.button("🗑️ Delete Selected Deal", use_container_width=True, key="btn_delete", disabled=not confirm_delete):
                    deal_id = deal_options[selected_label]
                    if DatabaseManager.delete_deal(deal_id):
                        st.success("✅ Deleted")
                        st.cache_data.clear()
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error("Delete failed.")
            else:
                st.info("No saved deals yet.")

            if st.button("💾 Save Deal", use_container_width=True, key="btn_save"):
                state = extract_clean_state()
                errors, warnings = ValidationEngine.validate(state)
                if errors:
                    for error in errors:
                        st.error(f"• {error}")
                else:
                    for warning in warnings:
                        st.warning(f"• {warning}")
                    deal_name = state.get("deal_name", "Untitled Deal")
                    if DatabaseManager.save_deal(state["deal_id"], deal_name, state):
                        s.unsaved_changes = False
                        s.last_saved_at = utc_now_iso()
                        st.success("✅ Deal saved!")
                    else:
                        st.error("Save failed.")

        if s.get("unsaved_changes"):
            st.warning("⚠️ Unsaved changes")
        if s.get("last_saved_at"):
            st.caption(f"Last saved: {str(s.last_saved_at)[:19]}")
        st.markdown("---")

        with st.expander("🏢 ASSET PROFILE", expanded=True):
            s.deal_name = st.text_input("Deal Name", value=s.get("deal_name", ""), key="asset_deal_name")
            s.sponsor = st.text_input("Sponsor", value=s.get("sponsor", ""), key="asset_sponsor")
            s.property_address = st.text_input("Property Address", value=s.get("property_address", ""), key="asset_property_address")
            current_pt = PROPERTY_TYPES.index(s.get("property_type", "Multifamily")) if s.get("property_type") in PROPERTY_TYPES else 0
            s.property_type = st.selectbox("Property Type", PROPERTY_TYPES, index=current_pt, key="asset_property_type")
            current_tt = TRANSACTION_TYPES.index(s.get("transaction_type", "Acquisition")) if s.get("transaction_type") in TRANSACTION_TYPES else 0
            s.transaction_type = st.selectbox("Transaction Type", TRANSACTION_TYPES, index=current_tt, key="asset_transaction_type")
            s.appraisal = st.number_input("Appraisal ($)", value=safe_float(s.get("appraisal", 0)), step=100000.0, min_value=0.0, format="%.0f", key="asset_appraisal")
            s.purchase_price = st.number_input("Cost Basis ($)", value=safe_float(s.get("purchase_price", 0)), step=100000.0, min_value=0.0, format="%.0f", key="asset_purchase_price")
            s.noi = st.number_input("Stabilized NOI ($)", value=safe_float(s.get("noi", 0)), step=10000.0, min_value=0.0, format="%.0f", key="asset_noi")

        with st.expander("📊 CREDIT POLICY", expanded=True):
            profiles = list(UnderwritingEngine.LENDER_LIMITS.keys())
            current_lp = profiles.index(s.get("lender_profile", "Bank / Credit Union")) if s.get("lender_profile") in profiles else 0
            s.lender_profile = st.selectbox("Lender Profile", profiles, index=current_lp, key="credit_lender_profile")
            limits = UnderwritingEngine.LENDER_LIMITS[s.lender_profile]
            s.target_ltv = st.slider("Max LTV %", 50.0, 95.0, float(normalize_percent(s.get("target_ltv", limits["max_ltv"]), limits["max_ltv"], 1.25) * 100), step=0.5, key="credit_target_ltv") / 100
            s.target_dscr = st.slider("Min DSCR", 1.0, 1.75, float(safe_float(s.get("target_dscr", limits["min_dscr"]))), step=0.05, key="credit_target_dscr")
            s.target_dy = st.slider("Min DY %", 5.0, 15.0, float(normalize_percent(s.get("target_dy", limits["min_dy"]), limits["min_dy"], 0.25) * 100), step=0.25, key="credit_target_dy") / 100
            s.target_ltc = st.slider("Max LTC %", 50.0, 100.0, float(normalize_percent(s.get("target_ltc", 0.80), 0.80, 1.25) * 100), step=0.5, key="credit_target_ltc") / 100

        with st.expander("💰 DEBT STRUCTURE", expanded=True):
            s.is_io = st.checkbox("Interest-Only Period", value=bool(s.get("is_io", False)), key="debt_is_io")
            s.rate = st.slider("Interest Rate %", 0.0, 15.0, float(normalize_percent(s.get("rate", 0.0525), 0.0525, 0.30) * 100), step=0.05, key="debt_rate") / 100
            s.amort = st.number_input("Amortization (Yrs)", value=max(1, safe_int(s.get("amort", 25))), step=1, min_value=1, max_value=40, key="debt_amort")
            s.term = st.number_input("Term (Yrs)", value=max(1, safe_int(s.get("term", 5))), step=1, min_value=1, max_value=40, key="debt_term")
            s.fees = st.slider("Financing Fees %", 0.0, 5.0, float(normalize_percent(s.get("fees", 0.02), 0.02, 0.10) * 100), step=0.05, key="debt_fees") / 100
            s.closing_costs = st.number_input("Closing Costs ($)", value=safe_float(s.get("closing_costs", 0)), step=1000.0, min_value=0.0, format="%.0f", key="debt_closing_costs")
            s.reserves = st.number_input("Required Reserves ($)", value=safe_float(s.get("reserves", 0)), step=1000.0, min_value=0.0, format="%.0f", key="debt_reserves")

    state_hash_after = stable_state_hash()
    if state_hash_after != state_hash_before and not s.get("_loading_deal", False) and not s.get("unsaved_changes", False):
        s.unsaved_changes = True

    rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))
    loan_amt, gate, gates, total_uses, req_equity = UnderwritingEngine.size_loan(s.noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, s.rate, s.amort, s.term, s.is_io, s.target_ltv, s.target_ltc, s.target_dscr, s.target_dy)
    amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(loan_amt, s.rate, s.amort, s.term, s.is_io)
    annual_ds = monthly_pmt * 12
    actual_ltv = loan_amt / s.appraisal if safe_float(s.appraisal) > 0 else 0
    actual_ltc = loan_amt / total_uses if total_uses > 0 else 0
    actual_dscr = s.noi / annual_ds if annual_ds > 0 else 0
    actual_dy = s.noi / loan_amt if loan_amt > 0 else 0
    tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr_df)
    score, tier = UnderwritingEngine.score_deal(actual_ltv, actual_ltc, actual_dscr, actual_dy, s.lender_profile)
    errors, warnings = ValidationEngine.validate(extract_clean_state())

    st.title(f"{s.sponsor or 'New Deal'} | {s.property_address or 'Property Profile'}")
    st.caption(f"ALENZA CAPITAL OS | CONSTRAINT: {gate} | {tier}")
    if errors:
        for error in errors[:3]:
            st.error(f"❌ {error}")
        if len(errors) > 3:
            st.error(f"... and {len(errors) - 3} more errors")
    if warnings:
        for warning in warnings[:2]:
            st.warning(f"⚠️ {warning}")

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
    col2.metric("ACTUAL LTV", f"{actual_ltv * 100:.1f}%")
    col3.metric("ACTUAL LTC", f"{actual_ltc * 100:.1f}%")
    col4.metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")
    col5.metric("BALLOON", f"${balloon:,.0f}")
    col6.metric("DEAL SCORE", f"{score}/1000", help=tier)
    st.markdown("---")

    tabs = st.tabs(["📊 Sizing & Risk", "🧪 Sensitivity", "📝 Rent Roll", "📅 Amortization", "🇨🇦 Canada Intel", "📈 Market Comps", "📎 Diligence Room", "💾 Save & Export", "✅ QA & Health"])

    with tabs[0]:
        col_left, col_right = st.columns([1.5, 1])
        with col_left:
            st.subheader("📐 Constraint Analysis")
            constraints_df = pd.DataFrame({"Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"], "Threshold": [f"{s.target_ltv * 100:.1f}%", f"{s.target_ltc * 100:.1f}%", f"{s.target_dscr:.2f}x", f"{s.target_dy * 100:.2f}%"], "Max Proceeds": [f"${gates.get('LTV', 0):,.0f}", f"${gates.get('LTC', 0):,.0f}", f"${gates.get('DSCR', 0):,.0f}", f"${gates.get('Debt Yield', 0):,.0f}"], "Binding": ["✅ ACTIVE" if gate == g else "" for g in ["LTV", "LTC", "DSCR", "Debt Yield"]]})
            st.dataframe(constraints_df, hide_index=True, use_container_width=True)
            st.subheader("💰 Sources & Uses")
            total_fees = loan_amt * s.fees
            su_df = pd.DataFrame({"Category": ["Cost Basis", "Closing Costs", "Reserves", "Financing Fees", "TOTAL USES"], "Uses": [f"${s.purchase_price:,.0f}", f"${s.closing_costs:,.0f}", f"${s.reserves:,.0f}", f"${total_fees:,.0f}", f"${total_uses:,.0f}"], "Sources": ["Senior Debt", "Sponsor Equity", "", "", "TOTAL SOURCES"], "Amount": [f"${loan_amt:,.0f}", f"${req_equity:,.0f}", "", "", f"${total_uses:,.0f}"]})
            st.dataframe(su_df, hide_index=True, use_container_width=True)
        with col_right:
            st.subheader("🔍 Risk Assessment")
            flags = []
            if actual_ltv > 0.75: flags.append(("high", f"⚠️ High Leverage: {actual_ltv * 100:.1f}% LTV"))
            elif actual_ltv < 0.55 and loan_amt > 0: flags.append(("low", f"✅ Conservative Leverage: {actual_ltv * 100:.1f}% LTV"))
            if actual_dscr < 1.20 and loan_amt > 0: flags.append(("high", f"⚠️ Tight Coverage: {actual_dscr:.2f}x DSCR"))
            elif actual_dscr > 1.50: flags.append(("low", f"✅ Strong Coverage: {actual_dscr:.2f}x DSCR"))
            if s.is_io: flags.append(("medium", "ℹ️ Interest-Only Structure"))
            if req_equity < 0: flags.append(("high", f"🚨 Negative Equity: ${abs(req_equity):,.0f}"))
            if walt > 0 and walt < 3: flags.append(("high", f"⚠️ Short WALT: {walt:.1f} years"))
            if exp1 > 0.30: flags.append(("high", f"🚨 High Rollover: {exp1 * 100:.1f}%"))
            if not flags: flags.append(("low", "✅ No Significant Risk Flags"))
            for severity, message in flags:
                if severity == "high": st.error(message)
                elif severity == "medium": st.warning(message)
                else: st.success(message)
            st.markdown("---")
            st.subheader("📊 Key Metrics")
            breakeven_occ = UnderwritingEngine.breakeven_occupancy(s.noi, occ, annual_ds)
            st.metric("Breakeven Occupancy", f"{breakeven_occ * 100:.1f}%", delta=f"Current: {occ * 100:.1f}%" if occ > 0 else None)
            st.metric("Required Equity", f"${req_equity:,.0f}")
            st.metric("Implied Cap Rate", f"{(s.noi / s.appraisal * 100):.2f}%" if s.appraisal > 0 else "N/A")
            st.markdown("---")
            st.subheader("🏗️ Capital Stack")
            s.mezz_debt = st.number_input("Mezzanine Debt ($)", value=safe_float(s.get("mezz_debt", 0)), min_value=0.0, step=100000.0, format="%.0f", key="capital_mezz_debt")
            s.pref_equity = st.number_input("Preferred Equity ($)", value=safe_float(s.get("pref_equity", 0)), min_value=0.0, step=100000.0, format="%.0f", key="capital_pref_equity")
            s.mezz_rate = st.slider("Mezz Rate %", 0.0, 25.0, float(normalize_percent(s.get("mezz_rate", 0.10), 0.10, 0.40) * 100), step=0.25, key="capital_mezz_rate") / 100
            s.pref_rate = st.slider("Pref Rate %", 0.0, 25.0, float(normalize_percent(s.get("pref_rate", 0.09), 0.09, 0.40) * 100), step=0.25, key="capital_pref_rate") / 100
            stack = UnderwritingEngine.capital_stack(loan_amt, s.mezz_debt, s.pref_equity, req_equity, s.noi, s.rate, s.mezz_rate, s.pref_rate)
            st.metric("Fixed Charge Coverage", f"{stack['Fixed Charge Coverage']:.2f}x")
            st.caption(f"Total Capital: ${stack['Total Capital']:,.0f} | Fixed Charges: ${stack['Fixed Charges']:,.0f}")

    with tabs[1]:
        st.subheader("🧪 Sensitivity Analysis")
        sensitivity_df = SensitivityEngine.generate_matrix(extract_clean_state())
        st.dataframe(sensitivity_df, hide_index=True, use_container_width=True)
        st.markdown("---")
        st.subheader("🔥 Proceeds Heatmap")
        heatmap_df = SensitivityEngine.proceeds_heatmap(extract_clean_state())
        currency_cols = [c for c in heatmap_df.columns if c != "NOI Shock"]
        st.dataframe(heatmap_df.style.format({col: "${:,.0f}" for col in currency_cols}), hide_index=True, use_container_width=True)
        if DEPENDENCIES.get("plotly"):
            try:
                import plotly.express as px
                plot_df = heatmap_df.set_index("NOI Shock")[currency_cols] / 1_000_000
                fig = px.imshow(plot_df, text_auto=".1f", aspect="auto", title="Max Proceeds Sensitivity ($MM)", labels={"x": "Rate Shock", "y": "NOI Shock", "color": "Max Proceeds ($MM)"})
                fig.update_traces(texttemplate="$%{z:.1f}MM", hovertemplate="NOI Shock: %{y}<br>Rate Shock: %{x}<br>Max Proceeds: $%{z:.1f}MM<extra></extra>")
                fig.update_layout(template="plotly_dark", paper_bgcolor="#0B0F19", plot_bgcolor="#0F172A", height=450, margin=dict(l=20, r=20, t=50, b=20))
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.info(f"Heatmap chart unavailable: {e}")
        st.markdown("---")
        st.subheader("🎯 Custom Stress Test")
        c1, c2, c3 = st.columns(3)
        with c1: rate_shock = st.slider("Rate Shock (bps)", -200, 200, 0, 25, key="custom_rate_shock") / 10000
        with c2: noi_shock = st.slider("NOI Shock (%)", -30, 30, 0, 5, key="custom_noi_shock") / 100
        with c3: ltv_shock = st.slider("LTV Adjustment (%)", -10, 10, 0, 1, key="custom_ltv_shock") / 100
        stressed_rate = max(0.001, s.rate + rate_shock)
        stressed_noi = max(0, s.noi * (1 + noi_shock))
        stressed_ltv = min(1.25, max(0.01, s.target_ltv + ltv_shock))
        stressed_loan, stressed_gate, _, _, _ = UnderwritingEngine.size_loan(stressed_noi, s.appraisal, s.purchase_price, s.closing_costs, s.reserves, s.fees, stressed_rate, s.amort, s.term, s.is_io, stressed_ltv, s.target_ltc, s.target_dscr, s.target_dy)
        _, stressed_monthly_pmt, _ = UnderwritingEngine.amort_schedule(stressed_loan, stressed_rate, s.amort, s.term, s.is_io)
        stressed_annual_ds = stressed_monthly_pmt * 12
        stressed_dscr = stressed_noi / stressed_annual_ds if stressed_annual_ds > 0 else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Stressed Proceeds", f"${stressed_loan:,.0f}", delta=f"${stressed_loan - loan_amt:,.0f}" if loan_amt > 0 else None)
        c2.metric("Constraint", stressed_gate)
        c3.metric("Stressed LTV", f"{(stressed_loan / s.appraisal * 100):.1f}%" if s.appraisal > 0 else "N/A")
        c4.metric("Stressed DSCR", f"{stressed_dscr:.2f}x")

    with tabs[2]:
        st.subheader("📝 Rent Roll Management")
        uploaded_file = st.file_uploader("Import Rent Roll (CSV or Excel)", type=["csv", "xlsx", "xls"], key="rent_roll_upload")
        if uploaded_file is not None:
            imported_df = None
            try:
                file_name = uploaded_file.name.lower()
                if file_name.endswith(".csv"):
                    imported_df = pd.read_csv(uploaded_file)
                elif file_name.endswith(".xlsx"):
                    if not DEPENDENCIES.get("excel_read"):
                        st.error("Install openpyxl to import .xlsx files.")
                    else:
                        imported_df = pd.read_excel(uploaded_file, engine="openpyxl")
                elif file_name.endswith(".xls"):
                    if not DEPENDENCIES.get("xls_read"):
                        st.error("Install xlrd to import .xls files.")
                    else:
                        imported_df = pd.read_excel(uploaded_file, engine="xlrd")
                if imported_df is not None:
                    imported_df = normalize_rent_roll_columns(imported_df)
                    st.write("Preview")
                    st.dataframe(imported_df, hide_index=True, use_container_width=True)
                    if st.button("✅ Apply Imported Rent Roll", key="apply_imported_rent_roll"):
                        s.rent_roll_dict = imported_df.to_dict("records")
                        s.unsaved_changes = True
                        st.success(f"✅ Imported {len(imported_df)} tenant records")
                        st.rerun()
            except Exception as e:
                st.error(f"Import failed: {str(e)[:200]}")
        edit_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))
        col_add, col_clear = st.columns(2)
        with col_add:
            if st.button("➕ Add Blank Row", use_container_width=True, key="add_rr_row"):
                rr_records = edit_df.to_dict("records")
                rr_records.append({"Tenant": "", "SF": 0, "Remaining Term": 0, "Monthly Rent": 0})
                s.rent_roll_dict = rr_records
                s.unsaved_changes = True
                st.rerun()
        with col_clear:
            if st.button("🧹 Clear Rent Roll", use_container_width=True, key="clear_rr"):
                s.rent_roll_dict = []
                s.unsaved_changes = True
                st.rerun()
        st.caption("After editing the table, click Update Rent Roll to apply changes to the model.")
        edited_df = st.data_editor(edit_df, num_rows="dynamic", use_container_width=True, hide_index=True, column_config={"Tenant": st.column_config.TextColumn("Tenant Name", width="large"), "SF": st.column_config.NumberColumn("Square Feet", min_value=0, step=100, format="%d"), "Remaining Term": st.column_config.NumberColumn("Lease Term (Yrs)", min_value=0.0, step=0.5, format="%.1f"), "Monthly Rent": st.column_config.NumberColumn("Monthly Rent ($)", min_value=0.0, step=100.0, format="$%.2f")}, key="rent_roll_editor")
        if st.button("💾 Update Rent Roll", use_container_width=True, key="update_rent_roll"):
            s.rent_roll_dict = normalize_rent_roll_columns(edited_df).to_dict("records")
            s.unsaved_changes = True
            st.rerun()
        if not edited_df.empty:
            st.markdown("---")
            st.subheader("📊 Rent Roll Analytics")
            met_total_sf, met_occ, met_ann_rent, met_psf, met_walt, met_exp1 = UnderwritingEngine.rent_roll_metrics(edited_df)
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total SF", f"{met_total_sf:,.0f}")
            c2.metric("Occupancy", f"{met_occ * 100:.1f}%")
            c3.metric("Annual Rent", f"${met_ann_rent:,.0f}")
            c4.metric("Rent PSF", f"${met_psf:.2f}")
            c5.metric("WALT (Yrs)", f"{met_walt:.2f}")
            c6.metric("12-Mo Rollover", f"{met_exp1 * 100:.1f}%")

    with tabs[3]:
        st.subheader(f"📅 Amortization Schedule - {s.term} Year Term")
        if amort_df is None or amort_df.empty:
            st.warning("Enter deal parameters to generate amortization schedule.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
            c2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
            c3.metric("Balloon Balance", f"${balloon:,.0f}")
            st.markdown("---")
            st.write("### Payment Structure")
            chart_data = amort_df.set_index("Period")[["Principal", "Interest"]]
            st.bar_chart(chart_data, use_container_width=True)
            st.write("### Outstanding Balance")
            balance_chart = amort_df.set_index("Period")[["Balance"]]
            st.line_chart(balance_chart, use_container_width=True)
            amort_view = amort_df.copy()
            amort_view["Year"] = ((amort_view["Period"] - 1) // 12) + 1
            annual_summary = amort_view.groupby("Year").agg({"Payment": "sum", "Principal": "sum", "Interest": "sum", "Balance": "last"}).reset_index()
            st.write("### Annual Summary")
            st.dataframe(annual_summary.style.format({"Payment": "${:,.2f}", "Principal": "${:,.2f}", "Interest": "${:,.2f}", "Balance": "${:,.2f}"}), use_container_width=True, hide_index=True)
            with st.expander("View Full Monthly Schedule"):
                st.dataframe(amort_view[["Period", "Payment", "Principal", "Interest", "Balance"]].style.format({"Payment": "${:,.2f}", "Principal": "${:,.2f}", "Interest": "${:,.2f}", "Balance": "${:,.2f}"}), use_container_width=True, height=400, hide_index=True)
    # ==========================================
    # TAB 4: CANADA INTEL
    # ==========================================

    with tabs[4]:
        st.subheader("🇨🇦 Canadian Sovereign Intelligence")
        st.caption("Bank of Canada, FX, labour-market signals, and rules-based market commentary.")

        @st.cache_data(ttl=3600, max_entries=12)
        def fetch_boc_rates_history(days: int = 365):
            series_map = {
                "FXUSDCAD": "USD/CAD",
                "FXEURCAD": "EUR/CAD",
                "BD.CDN.2YR.DQ.YLD": "2-Year Yield",
                "BD.CDN.5YR.DQ.YLD": "5-Year Yield",
                "BD.CDN.10YR.DQ.YLD": "10-Year Yield",
                "V122514": "Overnight Rate",
                "STATIC_ATABLE_V39079": "Overnight Target",
                "V39078": "Bank Rate",
            }

            try:
                series = ",".join(series_map.keys())
                url = f"https://www.bankofcanada.ca/valet/observations/{series}/json?recent={int(days)}"
                response = requests.get(url, timeout=15)
                response.raise_for_status()

                observations = response.json().get("observations", [])
                if not observations:
                    return {}, pd.DataFrame()

                rows = []

                for obs in observations:
                    row = {"Date": pd.to_datetime(obs.get("d"))}

                    for source_key, label in series_map.items():
                        try:
                            raw_value = obs.get(source_key, {}).get("v")
                            row[label] = float(raw_value) if raw_value not in [None, ""] else np.nan
                        except Exception:
                            row[label] = np.nan

                    rows.append(row)

                history_df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
                latest = {}

                for col in history_df.columns:
                    if col == "Date":
                        continue

                    valid = history_df[["Date", col]].dropna()

                    if not valid.empty:
                        latest[col] = {
                            "value": float(valid[col].iloc[-1]),
                            "date": valid["Date"].iloc[-1],
                        }

                # Derived fallback because the deposit rate is not always available in the same Valet pull.
                if "Deposit Rate" not in latest and "Overnight Target" in latest:
                    latest["Deposit Rate (derived)"] = {
                        "value": max(0.0, latest["Overnight Target"]["value"] - 0.05),
                        "date": latest["Overnight Target"]["date"],
                    }

                return latest, history_df

            except Exception as e:
                logger.error(f"BOC fetch failed: {e}")
                return {}, pd.DataFrame()

        @st.cache_data(ttl=86400, max_entries=4)
        def fetch_unemployment_history():
            try:
                url = "https://www150.statcan.gc.ca/n1/tbl/csv/14100287-eng.zip"
                response = requests.get(url, timeout=25)
                response.raise_for_status()

                with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                    csv_files = [f for f in z.namelist() if f.endswith(".csv")]

                    if not csv_files:
                        return pd.DataFrame()

                    with z.open(csv_files[0]) as f:
                        df = pd.read_csv(f)

                required_cols = {
                    "GEO",
                    "Labour force characteristics",
                    "Sex",
                    "Age group",
                    "REF_DATE",
                    "VALUE",
                }

                if not required_cols.issubset(set(df.columns)):
                    return pd.DataFrame()

                mask = (
                    (df["GEO"].astype(str).str.lower() == "canada")
                    & (df["Labour force characteristics"].astype(str).str.lower() == "unemployment rate")
                    & (df["Sex"].astype(str).str.lower() == "both sexes")
                    & (df["Age group"].astype(str).str.lower() == "15 years and over")
                )

                unemployment = df.loc[mask, ["REF_DATE", "VALUE"]].copy()
                unemployment.columns = ["Date", "Unemployment Rate"]
                unemployment["Date"] = pd.to_datetime(unemployment["Date"], errors="coerce")
                unemployment["Unemployment Rate"] = pd.to_numeric(
                    unemployment["Unemployment Rate"],
                    errors="coerce",
                )

                return unemployment.dropna().sort_values("Date").reset_index(drop=True)

            except Exception as e:
                logger.error(f"StatsCan fetch failed: {e}")
                return pd.DataFrame()

        @st.cache_data(ttl=86400, max_entries=4)
        def fetch_vacancy_rates():
            return pd.DataFrame(
                {
                    "Property Class": [
                        "Multifamily",
                        "Industrial",
                        "Retail",
                        "Office",
                        "Mixed-Use",
                        "Hospitality",
                        "Self-Storage",
                    ],
                    "National Vacancy": [2.1, 1.8, 5.2, 12.4, 4.5, 7.5, 3.8],
                    "Toronto": [1.5, 1.2, 4.8, 11.2, 3.8, 7.2, 3.2],
                    "Vancouver": [0.9, 1.1, 3.5, 9.8, 3.2, 6.4, 2.9],
                    "Montreal": [2.8, 2.1, 5.5, 13.1, 5.1, 8.0, 4.1],
                    "Calgary": [3.2, 3.5, 7.2, 18.5, 6.8, 9.1, 5.0],
                    "Trend": [
                        "Tightening",
                        "Tightening",
                        "Softening",
                        "High Vacancy",
                        "Stable",
                        "Mixed",
                        "Stable",
                    ],
                    "YoY Change": [-0.3, -0.2, 0.8, 1.5, 0.1, 0.4, 0.0],
                }
            )

        def latest_value(latest: dict, key: str):
            item = latest.get(key)
            return item.get("value") if isinstance(item, dict) else None

        def latest_date(latest: dict, key: str):
            item = latest.get(key)
            return item.get("date") if isinstance(item, dict) else None

        def build_market_commentary(boc_latest, boc_history, unemp_data, vacancy_data, deal_state):
            commentary = []

            y2 = latest_value(boc_latest, "2-Year Yield")
            y5 = latest_value(boc_latest, "5-Year Yield")
            y10 = latest_value(boc_latest, "10-Year Yield")
            target = latest_value(boc_latest, "Overnight Target")

            if y2 is not None and y10 is not None:
                spread = y10 - y2

                if spread < -0.50:
                    commentary.append(
                        (
                            "high",
                            "Yield Curve Deeply Inverted",
                            f"2s10s spread is {spread:.2f}%. Credit conditions should be underwritten conservatively.",
                        )
                    )
                elif spread < 0:
                    commentary.append(
                        (
                            "medium",
                            "Yield Curve Inverted",
                            f"2s10s spread is {spread:.2f}%. Stress rate and refinance assumptions.",
                        )
                    )
                elif spread < 0.50:
                    commentary.append(
                        (
                            "medium",
                            "Yield Curve Flat",
                            f"2s10s spread is {spread:.2f}%. Market is pricing uncertainty.",
                        )
                    )
                else:
                    commentary.append(
                        (
                            "low",
                            "Normal Yield Curve",
                            f"2s10s spread is {spread:.2f}%. Term structure is positively sloped.",
                        )
                    )

            if y5 is not None:
                if y5 < 3.0:
                    commentary.append(
                        (
                            "low",
                            "Lower 5-Year Benchmark",
                            f"5Y GoC is {y5:.2f}%. Current debt costs may be supportive.",
                        )
                    )
                elif y5 < 5.0:
                    commentary.append(
                        (
                            "medium",
                            "Moderate Rate Environment",
                            f"5Y GoC is {y5:.2f}%. Maintain DSCR cushion.",
                        )
                    )
                else:
                    commentary.append(
                        (
                            "high",
                            "High Rate Environment",
                            f"5Y GoC is {y5:.2f}%. Debt yield and equity requirements are likely binding.",
                        )
                    )

            if target is not None and y5 is not None:
                policy_spread = y5 - target

                if policy_spread < -0.50:
                    commentary.append(
                        (
                            "medium",
                            "Market Pricing Policy Easing",
                            f"5Y GoC trades {policy_spread:.2f}% below the overnight target.",
                        )
                    )
                elif policy_spread > 0.75:
                    commentary.append(
                        (
                            "medium",
                            "Long-End Premium",
                            f"5Y GoC trades {policy_spread:.2f}% above the overnight target.",
                        )
                    )

            if unemp_data is not None and not unemp_data.empty:
                latest_unemp = unemp_data["Unemployment Rate"].iloc[-1]
                six_month_ago = unemp_data["Unemployment Rate"].iloc[-6] if len(unemp_data) > 6 else latest_unemp
                six_month_change = latest_unemp - six_month_ago

                if latest_unemp < 5.5:
                    commentary.append(
                        (
                            "low",
                            "Strong Labour Market",
                            f"Unemployment is {latest_unemp:.1f}%. Demand fundamentals should be monitored but are supportive.",
                        )
                    )
                elif latest_unemp < 7.0:
                    commentary.append(
                        (
                            "medium",
                            "Moderate Labour Market",
                            f"Unemployment is {latest_unemp:.1f}%. Monitor tenant credit and local employment.",
                        )
                    )
                else:
                    commentary.append(
                        (
                            "high",
                            "Elevated Unemployment",
                            f"Unemployment is {latest_unemp:.1f}%. Tenant demand and collections risk may be elevated.",
                        )
                    )

                if six_month_change > 0.5:
                    commentary.append(
                        (
                            "high",
                            "Rising Unemployment Trend",
                            f"Unemployment rose {six_month_change:.1f}% over six months.",
                        )
                    )

            if vacancy_data is not None and not vacancy_data.empty:
                prop_type = deal_state.get("property_type", "Multifamily")
                row = vacancy_data[vacancy_data["Property Class"] == prop_type]

                if not row.empty:
                    vac = float(row["National Vacancy"].iloc[0])

                    if vac < 3:
                        commentary.append(
                            (
                                "low",
                                f"Tight {prop_type} Market",
                                f"Simulated national vacancy is {vac:.1f}%. Verify against broker data.",
                            )
                        )
                    elif vac < 8:
                        commentary.append(
                            (
                                "medium",
                                f"Balanced {prop_type} Market",
                                f"Simulated national vacancy is {vac:.1f}%. Use market vacancy in underwriting.",
                            )
                        )
                    else:
                        commentary.append(
                            (
                                "high",
                                f"High {prop_type} Vacancy",
                                f"Simulated national vacancy is {vac:.1f}%. Stress lease-up and exit cap assumptions.",
                            )
                        )

            return commentary

        st.write("### Data Range")

        history_days = st.select_slider(
            "Chart History",
            options=[30, 90, 180, 365, 730, 1825],
            value=365,
            format_func=lambda x: f"{x} Days ({x // 365}y)" if x >= 365 else f"{x} Days ({x // 30}m)",
        )

        boc_latest, boc_history = fetch_boc_rates_history(days=history_days)
        unemp_data = fetch_unemployment_history()
        vacancy_data = fetch_vacancy_rates()

        if unemp_data is not None and not unemp_data.empty:
            cutoff_date = pd.Timestamp.now() - pd.Timedelta(days=history_days)
            unemp_data = unemp_data[unemp_data["Date"] >= cutoff_date].copy()

        st.markdown("---")
        st.write("### Automated Market Commentary")
        st.caption("Rules-based analysis of current public data. Not AI-generated.")

        commentary = build_market_commentary(
            boc_latest,
            boc_history,
            unemp_data,
            vacancy_data,
            extract_clean_state(),
        )

        if commentary:
            for level, title, text in commentary:
                if level == "high":
                    st.error(f"**{title}** — {text}")
                elif level == "medium":
                    st.warning(f"**{title}** — {text}")
                else:
                    st.success(f"**{title}** — {text}")
        else:
            st.info("No market commentary available from current data.")

        st.markdown("---")
        st.write("### Bank of Canada Rates, FX, and Bond Yields")

        if boc_latest:
            metrics = [
                ("USD/CAD", "USD/CAD", "{:.4f}"),
                ("EUR/CAD", "EUR/CAD", "{:.4f}"),
                ("Overnight Target", "Overnight Target", "{:.2f}%"),
                ("Overnight Rate", "Overnight Rate", "{:.2f}%"),
                ("Bank Rate", "Bank Rate", "{:.2f}%"),
                ("Deposit Rate", "Deposit Rate (derived)", "{:.2f}%"),
                ("2Y Yield", "2-Year Yield", "{:.2f}%"),
                ("5Y Yield", "5-Year Yield", "{:.2f}%"),
                ("10Y Yield", "10-Year Yield", "{:.2f}%"),
            ]

            metric_cols = st.columns(3)

            for idx, (label, key, fmt) in enumerate(metrics):
                value = latest_value(boc_latest, key)
                metric_cols[idx % 3].metric(
                    label,
                    fmt.format(value) if value is not None else "N/A",
                )

            y2 = latest_value(boc_latest, "2-Year Yield")
            y10 = latest_value(boc_latest, "10-Year Yield")

            if y2 is not None and y10 is not None:
                st.metric("2s10s Spread", f"{(y10 - y2):.2f}%")

            rate_table = []

            for label, key, fmt in metrics:
                value = latest_value(boc_latest, key)
                d = latest_date(boc_latest, key)

                rate_table.append(
                    {
                        "Series": label,
                        "Value": fmt.format(value) if value is not None else "N/A",
                        "Observation Date": d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else "N/A",
                    }
                )

            st.dataframe(pd.DataFrame(rate_table), hide_index=True, use_container_width=True)

        else:
            st.warning("Bank of Canada data unavailable.")

        if boc_history is not None and not boc_history.empty:
            if DEPENDENCIES.get("plotly"):
                try:
                    import plotly.graph_objects as go

                    st.write("#### Bond Yield and Policy Rate History")

                    fig = go.Figure()

                    for col_name, color in {
                        "2-Year Yield": "#CFB87C",
                        "5-Year Yield": "#F59E0B",
                        "10-Year Yield": "#EF4444",
                        "Overnight Target": "#22C55E",
                        "Bank Rate": "#60A5FA",
                    }.items():
                        if col_name in boc_history.columns and boc_history[col_name].notna().any():
                            fig.add_trace(
                                go.Scatter(
                                    x=boc_history["Date"],
                                    y=boc_history[col_name],
                                    mode="lines",
                                    name=col_name,
                                    line=dict(color=color, width=2),
                                )
                            )

                    fig.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="#0B0F19",
                        plot_bgcolor="#0F172A",
                        height=500,
                        hovermode="x unified",
                        margin=dict(l=20, r=20, t=40, b=20),
                    )

                    st.plotly_chart(fig, use_container_width=True)

                    st.write("#### FX History")

                    fx_fig = go.Figure()

                    for col_name, color in {
                        "USD/CAD": "#3B82F6",
                        "EUR/CAD": "#A855F7",
                    }.items():
                        if col_name in boc_history.columns and boc_history[col_name].notna().any():
                            fx_fig.add_trace(
                                go.Scatter(
                                    x=boc_history["Date"],
                                    y=boc_history[col_name],
                                    mode="lines",
                                    name=col_name,
                                    line=dict(color=color, width=2),
                                )
                            )

                    fx_fig.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="#0B0F19",
                        plot_bgcolor="#0F172A",
                        height=400,
                        hovermode="x unified",
                        margin=dict(l=20, r=20, t=40, b=20),
                    )

                    st.plotly_chart(fx_fig, use_container_width=True)

                    y2 = latest_value(boc_latest, "2-Year Yield")
                    y5 = latest_value(boc_latest, "5-Year Yield")
                    y10 = latest_value(boc_latest, "10-Year Yield")

                    curve_df = pd.DataFrame(
                        {
                            "Tenor": ["2Y", "5Y", "10Y"],
                            "Yield": [y2, y5, y10],
                        }
                    ).dropna()

                    if not curve_df.empty:
                        st.write("#### Current Yield Curve")

                        curve_fig = go.Figure()

                        curve_fig.add_trace(
                            go.Scatter(
                                x=curve_df["Tenor"],
                                y=curve_df["Yield"],
                                mode="lines+markers+text",
                                text=[f"{v:.2f}%" for v in curve_df["Yield"]],
                                textposition="top center",
                                line=dict(color="#CFB87C", width=4),
                                marker=dict(size=12),
                            )
                        )

                        curve_fig.update_layout(
                            template="plotly_dark",
                            paper_bgcolor="#0B0F19",
                            plot_bgcolor="#0F172A",
                            height=350,
                            margin=dict(l=20, r=20, t=40, b=20),
                            yaxis=dict(ticksuffix="%"),
                        )

                        st.plotly_chart(curve_fig, use_container_width=True)

                except Exception as e:
                    st.info(f"Charts unavailable: {e}")

            else:
                chart_cols = [
                    c
                    for c in [
                        "2-Year Yield",
                        "5-Year Yield",
                        "10-Year Yield",
                        "Overnight Target",
                        "USD/CAD",
                        "EUR/CAD",
                    ]
                    if c in boc_history.columns
                ]

                if chart_cols:
                    st.line_chart(boc_history.set_index("Date")[chart_cols], use_container_width=True)

        st.markdown("---")
        st.write("### Canadian Labour Market")

        if unemp_data is not None and not unemp_data.empty:
            latest_unemp = unemp_data["Unemployment Rate"].iloc[-1]
            prev_month = unemp_data["Unemployment Rate"].iloc[-2] if len(unemp_data) > 1 else latest_unemp
            three_month_avg = unemp_data["Unemployment Rate"].tail(3).mean()
            year_ago = unemp_data["Unemployment Rate"].iloc[-12] if len(unemp_data) > 12 else unemp_data["Unemployment Rate"].iloc[0]

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Current Rate", f"{latest_unemp:.1f}%")
            c2.metric("Monthly Change", f"{latest_unemp - prev_month:+.1f}%")
            c3.metric("3-Month Avg", f"{three_month_avg:.1f}%")
            c4.metric("Year Ago", f"{year_ago:.1f}%", delta=f"{latest_unemp - year_ago:+.1f}%")

            if DEPENDENCIES.get("plotly"):
                try:
                    import plotly.graph_objects as go

                    unemp_view = unemp_data.copy()
                    unemp_view["3M MA"] = unemp_view["Unemployment Rate"].rolling(window=3).mean()

                    fig_ue = go.Figure()

                    fig_ue.add_trace(
                        go.Scatter(
                            x=unemp_view["Date"],
                            y=unemp_view["Unemployment Rate"],
                            mode="lines+markers",
                            name="Unemployment Rate",
                            line=dict(color="#EF4444", width=2),
                            marker=dict(size=4),
                        )
                    )

                    fig_ue.add_trace(
                        go.Scatter(
                            x=unemp_view["Date"],
                            y=unemp_view["3M MA"],
                            mode="lines",
                            name="3M MA",
                            line=dict(color="#F59E0B", width=2, dash="dash"),
                        )
                    )

                    fig_ue.add_hline(
                        y=unemp_view["Unemployment Rate"].mean(),
                        line_dash="dot",
                        line_color="#9CA3AF",
                        annotation_text="Period Avg",
                    )

                    fig_ue.update_layout(
                        template="plotly_dark",
                        paper_bgcolor="#0B0F19",
                        plot_bgcolor="#0F172A",
                        height=430,
                        hovermode="x unified",
                        margin=dict(l=20, r=20, t=40, b=20),
                    )

                    st.plotly_chart(fig_ue, use_container_width=True)

                except Exception as e:
                    st.info(f"Unemployment chart unavailable: {e}")

            else:
                st.line_chart(unemp_data.set_index("Date")["Unemployment Rate"], use_container_width=True)

            st.caption("Source: Statistics Canada Table 14-10-0287-01")

        else:
            st.warning("Unemployment data unavailable.")

        st.markdown("---")
        st.write("### Commercial Vacancy Rates")

        st.warning(
            "Vacancy data below is simulated placeholder intelligence. "
            "Use verified broker, CMHC, Altus, CBRE, or CoStar data for final underwriting."
        )

        if vacancy_data is not None and not vacancy_data.empty:
            cols = st.columns(min(len(vacancy_data), 4))

            for i, (_, row) in enumerate(vacancy_data.iterrows()):
                with cols[i % len(cols)]:
                    vac = row["National Vacancy"]
                    marker = "🟢" if vac < 3 else "🟡" if vac < 8 else "🔴"

                    st.metric(
                        f"{marker} {row['Property Class']}",
                        f"{vac:.1f}%",
                        delta=f"{row['YoY Change']:+.1f}% YoY",
                        delta_color="inverse" if row["YoY Change"] > 0 else "normal",
                    )

                    st.caption(str(row["Trend"]))

            user_vacancy = vacancy_data[vacancy_data["Property Class"] == s.property_type]

            if not user_vacancy.empty:
                st.info(
                    f"Your deal type ({s.property_type}) simulated national vacancy: "
                    f"{user_vacancy['National Vacancy'].iloc[0]:.1f}%."
                )

        st.markdown("---")
        st.write("### Deal Pricing vs Benchmark")

        y5 = latest_value(boc_latest, "5-Year Yield") if boc_latest else None

        if y5 is not None:
            current_rate = safe_float(s.rate) * 100
            spread_to_goc = current_rate - y5

            c1, c2, c3 = st.columns(3)
            c1.metric("Your Deal Rate", f"{current_rate:.2f}%")
            c2.metric("5Y GoC", f"{y5:.2f}%")
            c3.metric("Spread to GoC", f"{spread_to_goc:.2f}%")

            if spread_to_goc < 1.5:
                st.success("Tight pricing relative to the 5Y GoC benchmark.")
            elif spread_to_goc < 3.0:
                st.info("Standard CRE spread range relative to the 5Y GoC benchmark.")
            else:
                st.warning("Wide spread. Verify asset, sponsor, and structure risk assumptions.")

        st.markdown("---")
        st.write("### Federal Corporation Verification")

        corp_number = st.text_input(
            "Federal Corporation Number or BN9",
            placeholder="e.g., 123456-7",
        )

        if corp_number:
            @st.cache_data(ttl=86400, max_entries=50)
            def verify_corp(number):
                cleaned = re.sub(r"\D", "", str(number))

                if not cleaned:
                    return None

                try:
                    url = f"https://www.ic.gc.ca/app/scr/cc/CorporationsCanada/api/corporations/{cleaned}.json?lang=eng"
                    response = requests.get(url, timeout=10)

                    if response.status_code != 200:
                        return None

                    data = response.json()
                    corp = data[0] if isinstance(data, list) and data else data if isinstance(data, dict) else None

                    if not corp:
                        return None

                    return {"number": cleaned, "raw": corp}

                except Exception as e:
                    logger.warning(f"Corporation lookup failed: {e}")
                    return None

            result = verify_corp(corp_number)

            if result:
                st.success("Corporation record found.")
                st.json(result)
            else:
                st.warning("Corporation not found or registry unavailable.")

    # ==========================================
    # TAB 5: MARKET COMPS
    # ==========================================

    with tabs[5]:
        st.subheader("📈 Market Comparables")
        st.caption(
            "Simulated comparison stack by property type. "
            "Replace with verified broker, appraisal, CoStar, Altus, CMHC, or internal comp data before final underwriting."
        )

        @st.cache_data(ttl=3600, max_entries=100)
        def generate_market_comps(property_type: str, noi: float, appraisal: float, address_seed: str) -> pd.DataFrame:
            seed_base = int(abs(hash(f"{property_type}-{safe_float(noi)}-{safe_float(appraisal)}-{address_seed}")) % 10000)
            rng = np.random.default_rng(seed_base)
            today = datetime.now().date()

            base_caps = {
                "Multifamily": 0.045,
                "Industrial": 0.055,
                "Retail": 0.065,
                "Office": 0.075,
                "Mixed-Use": 0.060,
                "Hospitality": 0.080,
                "Self-Storage": 0.058,
            }

            comp_names = {
                "Multifamily": [
                    "Garden Apartment Portfolio",
                    "Urban Rental Mid-Rise",
                    "Transit-Oriented Rental",
                    "Stabilized Apartment Block",
                    "Suburban Multifamily Asset",
                ],
                "Industrial": [
                    "Last-Mile Logistics Facility",
                    "Small-Bay Industrial Park",
                    "Flex Industrial Centre",
                    "Distribution Warehouse",
                    "Light Industrial Complex",
                ],
                "Retail": [
                    "Neighbourhood Retail Plaza",
                    "Grocery-Anchored Centre",
                    "Service Retail Strip",
                    "Urban Main Street Retail",
                    "Convenience Retail Centre",
                ],
                "Office": [
                    "Suburban Office Centre",
                    "Medical Office Building",
                    "Downtown Office Asset",
                    "Professional Office Campus",
                    "Creative Office Property",
                ],
                "Mixed-Use": [
                    "Main Street Mixed-Use Block",
                    "Retail/Rental Urban Asset",
                    "Podium Retail Mixed-Use",
                    "Neighbourhood Mixed-Use Centre",
                    "Urban Residential/Retail Property",
                ],
                "Hospitality": [
                    "Limited-Service Hotel",
                    "Extended-Stay Hotel",
                    "Select-Service Hotel",
                    "Airport Hotel Asset",
                    "Boutique Hospitality Property",
                ],
                "Self-Storage": [
                    "Climate-Controlled Storage",
                    "Drive-Up Storage Facility",
                    "Urban Self-Storage Asset",
                    "Suburban Storage Centre",
                    "Multi-Level Storage Facility",
                ],
            }

            base_cap = base_caps.get(property_type, 0.060)
            names = comp_names.get(property_type, comp_names["Mixed-Use"])

            comps = []

            for i in range(5):
                cap_rate = max(0.030, base_cap + rng.uniform(-0.006, 0.006))

                if safe_float(noi) > 0:
                    comp_noi = safe_float(noi) * rng.uniform(0.70, 1.30)
                elif safe_float(appraisal) > 0:
                    comp_noi = safe_float(appraisal) * cap_rate * rng.uniform(0.75, 1.25)
                else:
                    comp_noi = rng.uniform(350_000, 1_500_000)

                estimated_value = comp_noi / cap_rate if cap_rate > 0 else 0
                sale_date = today - timedelta(days=int(rng.integers(20, 540)))

                if property_type == "Multifamily":
                    size_metric = "Units"
                    size_value = int(rng.integers(40, 260))
                elif property_type == "Hospitality":
                    size_metric = "Rooms"
                    size_value = int(rng.integers(60, 240))
                elif property_type == "Self-Storage":
                    size_metric = "NRSF"
                    size_value = int(rng.integers(25_000, 180_000))
                else:
                    size_metric = "NRA SF"
                    size_value = int(rng.integers(20_000, 250_000))

                comps.append(
                    {
                        "Comparable": names[i],
                        "Type": property_type,
                        "Distance": f"{rng.uniform(0.5, 8.0):.1f} km",
                        "Sale Date": sale_date.strftime("%Y-%m-%d"),
                        "Cap Rate": cap_rate,
                        "Estimated NOI": comp_noi,
                        "Estimated Value": estimated_value,
                        size_metric: size_value,
                        "Price/Unit or SF": estimated_value / max(size_value, 1),
                        "As Of": today.strftime("%Y-%m-%d"),
                        "Latitude": 43.6532 + rng.uniform(-0.08, 0.08),
                        "Longitude": -79.3832 + rng.uniform(-0.08, 0.08),
                    }
                )

            return pd.DataFrame(comps)

        comps_df = generate_market_comps(
            property_type=s.property_type,
            noi=s.noi,
            appraisal=s.appraisal,
            address_seed=s.property_address,
        )

        st.write("### Simulated Comparable Sales")

        display_cols = [c for c in comps_df.columns if c not in ["Latitude", "Longitude"]]

        format_map = {
            "Cap Rate": "{:.2%}",
            "Estimated NOI": "${:,.0f}",
            "Estimated Value": "${:,.0f}",
            "Price/Unit or SF": "${:,.0f}",
        }

        st.dataframe(
            comps_df[display_cols].style.format(format_map),
            hide_index=True,
            use_container_width=True,
        )

        st.warning(
            "These comps are simulated underwriting placeholders. "
            "Do not use them as final valuation evidence without verified market support."
        )

        st.markdown("---")

        st.write("### Subject vs Market Cap Rate")

        if safe_float(s.noi) > 0 and safe_float(s.appraisal) > 0:
            implied_cap = safe_float(s.noi) / safe_float(s.appraisal)
            avg_market_cap = comps_df["Cap Rate"].mean()
            median_market_cap = comps_df["Cap Rate"].median()
            implied_value_at_market = safe_float(s.noi) / avg_market_cap if avg_market_cap > 0 else 0
            value_delta = implied_value_at_market - safe_float(s.appraisal)

            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Your Implied Cap Rate", f"{implied_cap:.2%}")
            c2.metric("Avg Market Cap", f"{avg_market_cap:.2%}", delta=f"{implied_cap - avg_market_cap:+.2%}")
            c3.metric("Median Market Cap", f"{median_market_cap:.2%}")
            c4.metric("Value at Avg Cap", f"${implied_value_at_market:,.0f}", delta=f"${value_delta:,.0f}")

            if implied_cap < avg_market_cap - 0.005:
                st.warning(
                    "Subject implied cap rate is materially tighter than the simulated market set. "
                    "Verify rent growth, asset quality, and buyer demand assumptions."
                )
            elif implied_cap > avg_market_cap + 0.005:
                st.success(
                    "Subject implied cap rate is wider than the simulated market set. "
                    "This may indicate conservative valuation or higher perceived risk."
                )
            else:
                st.info("Subject implied cap rate is broadly in line with the simulated comp set.")

        else:
            st.info("Enter stabilized NOI and appraisal value to compare the subject cap rate against comps.")

        st.markdown("---")

        st.write("### Comp Location Map")

        if {"Latitude", "Longitude"}.issubset(comps_df.columns):
            map_df = comps_df.rename(columns={"Latitude": "lat", "Longitude": "lon"})[
                ["lat", "lon", "Comparable", "Cap Rate", "Estimated Value"]
            ].copy()

            st.map(
                map_df[["lat", "lon"]],
                zoom=10,
                use_container_width=True,
            )

            with st.expander("Map Data"):
                st.dataframe(
                    map_df.style.format(
                        {
                            "Cap Rate": "{:.2%}",
                            "Estimated Value": "${:,.0f}",
                        }
                    ),
                    hide_index=True,
                    use_container_width=True,
                )
        else:
            st.info("No coordinate data available for map display.")

    # ==========================================
    # TAB 6: DILIGENCE ROOM
    # ==========================================

    with tabs[6]:
        st.subheader("📎 Diligence Room & Document Vault")
        st.caption(
            "Upload and track underwriting documents. "
            "Files are stored in the local document vault and indexed by deal ID."
        )

        REQUIRED_DOCS = [
            "Appraisal",
            "Phase I ESA",
            "T12 Financials",
            "Rent Roll",
            "Sponsor Bio",
            "Purchase Agreement",
            "Environmental Report",
            "Structural Report",
        ]

        DOC_PRIORITY = {
            "Appraisal": "Required for Submission",
            "T12 Financials": "Required for Submission",
            "Rent Roll": "Required for Submission",
            "Sponsor Bio": "Required for Submission",
            "Purchase Agreement": "Required for Closing",
            "Phase I ESA": "Required for Closing",
            "Environmental Report": "Recommended",
            "Structural Report": "Recommended",
        }

        docs_df = DatabaseManager.get_deal_documents(s.deal_id)

        upload_col, checklist_col = st.columns([1, 1.15])

        with upload_col:
            st.write("### 📤 Upload Document")

            category = st.selectbox(
                "Document Category",
                REQUIRED_DOCS + ["Other"],
                key="doc_category",
            )

            uploaded_doc = st.file_uploader(
                "Select File",
                key="diligence_upload",
                type=[
                    "pdf",
                    "docx",
                    "xlsx",
                    "xls",
                    "csv",
                    "png",
                    "jpg",
                    "jpeg",
                ],
            )

            if uploaded_doc is not None:
                file_size_mb = len(uploaded_doc.getvalue()) / (1024 * 1024)

                st.caption(
                    f"Selected file size: {file_size_mb:.2f} MB "
                    f"/ {MAX_UPLOAD_SIZE_MB:.0f} MB limit"
                )

                if file_size_mb > MAX_UPLOAD_SIZE_MB:
                    st.error(
                        f"File exceeds the {MAX_UPLOAD_SIZE_MB:.0f} MB upload limit. "
                        "Choose a smaller file."
                    )

            if st.button(
                "📤 Upload to Vault",
                use_container_width=True,
                key="upload_to_vault",
                disabled=uploaded_doc is None,
            ):
                try:
                    if uploaded_doc is None:
                        st.warning("Select a file first.")
                    elif len(uploaded_doc.getvalue()) > MAX_UPLOAD_SIZE_MB * 1024 * 1024:
                        st.error(f"File exceeds {MAX_UPLOAD_SIZE_MB:.0f} MB limit.")
                    elif DatabaseManager.save_document(s.deal_id, uploaded_doc, category):
                        s.unsaved_changes = True
                        st.success(f"✅ Uploaded: {uploaded_doc.name}")
                        st.rerun()
                    else:
                        st.error("Upload failed.")
                except Exception as e:
                    st.error(f"Upload failed: {str(e)[:200]}")

            st.markdown("---")

            st.write("### 🧾 Vault Summary")

            if docs_df is not None and not docs_df.empty:
                total_files = len(docs_df)
                total_size_mb = docs_df["file_size"].fillna(0).sum() / (1024 * 1024)

                c1, c2 = st.columns(2)
                c1.metric("Files Uploaded", f"{total_files}")
                c2.metric("Vault Size", f"{total_size_mb:.2f} MB")
            else:
                st.info("No documents uploaded yet.")

        with checklist_col:
            st.write("### 📋 Required Document Checklist")

            uploaded_cats = (
                docs_df["category"].astype(str).tolist()
                if docs_df is not None and not docs_df.empty and "category" in docs_df.columns
                else []
            )

            checklist_rows = []

            for doc_name in REQUIRED_DOCS:
                is_uploaded = doc_name in uploaded_cats

                checklist_rows.append(
                    {
                        "Requirement": doc_name,
                        "Priority": DOC_PRIORITY.get(doc_name, "Recommended"),
                        "Status": "✅ Uploaded" if is_uploaded else "❌ Missing",
                    }
                )

            gap_df = pd.DataFrame(checklist_rows)

            st.dataframe(
                gap_df,
                hide_index=True,
                use_container_width=True,
            )

            missing_required = gap_df[
                (gap_df["Status"] == "❌ Missing")
                & (gap_df["Priority"].isin(["Required for Submission", "Required for Closing"]))
            ]

            if missing_required.empty:
                st.success("All required diligence categories have at least one uploaded document.")
            else:
                st.warning(
                    f"{len(missing_required)} required diligence item(s) are still missing."
                )

        st.markdown("---")

        st.write("### 📁 Vault Inventory")

        docs_df = DatabaseManager.get_deal_documents(s.deal_id)

        if docs_df is not None and not docs_df.empty:
            inventory_df = docs_df.copy()

            if "file_size" in inventory_df.columns:
                inventory_df["Size MB"] = inventory_df["file_size"].fillna(0) / (1024 * 1024)

            display_columns = [
                col
                for col in [
                    "id",
                    "filename",
                    "original_filename",
                    "category",
                    "Size MB",
                    "uploaded_at",
                ]
                if col in inventory_df.columns
            ]

            st.dataframe(
                inventory_df[display_columns].style.format(
                    {
                        "Size MB": "{:.2f}",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )

            st.markdown("---")

            st.write("### 🗑️ Delete Document")

            doc_options = {"-- Select Document --": None}

            for _, row in inventory_df.iterrows():
                doc_id = str(row["id"])
                file_label = str(row.get("filename", "Document"))
                category_label = str(row.get("category", "Other"))
                doc_options[f"{file_label} · {category_label} · {doc_id[-8:]}"] = doc_id

            selected_doc_label = st.selectbox(
                "Document",
                list(doc_options.keys()),
                key="doc_delete_select",
            )

            confirm_doc_delete = st.checkbox(
                "Confirm document delete",
                key="confirm_doc_delete",
            )

            if st.button(
                "🗑️ Delete Selected Document",
                use_container_width=True,
                key="delete_doc_button",
                disabled=(doc_options[selected_doc_label] is None or not confirm_doc_delete),
            ):
                selected_doc_id = doc_options[selected_doc_label]

                if DatabaseManager.delete_document(selected_doc_id):
                    s.unsaved_changes = True
                    st.success("✅ Document deleted.")
                    st.rerun()
                else:
                    st.error("Document delete failed.")

        else:
            st.info("No documents uploaded yet. Use the upload form to add diligence materials.")

        st.markdown("---")

        st.write("### 🔎 Diligence Notes")

        diligence_notes = st.text_area(
            "Internal diligence notes",
            value=s.get("diligence_notes", ""),
            height=140,
            placeholder="Add open diligence items, third-party report notes, waiver items, or closing conditions.",
            key="diligence_notes_input",
        )

        if diligence_notes != s.get("diligence_notes", ""):
            s.diligence_notes = diligence_notes
            s.unsaved_changes = True

    # ==========================================
    # TAB 7: SAVE & EXPORT
    # ==========================================

    with tabs[7]:
        st.subheader("💾 Save & Export Deal")
        st.caption("Save the current deal, export underwriting files, or download a full package.")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        safe_deal_name = (
            re.sub(r"[^a-zA-Z0-9_-]+", "_", str(s.get("deal_name", "alenza_deal")))
            .strip("_")[:60]
            or "alenza_deal"
        )

        state = extract_clean_state()
        sensitivity_df = SensitivityEngine.generate_matrix(state)
        heatmap_df = SensitivityEngine.proceeds_heatmap(state)
        versions_df = DatabaseManager.get_deal_versions(s.deal_id, limit=25)
        audit_df = DatabaseManager.get_audit_log(s.deal_id, limit=50)
        docs_df = DatabaseManager.get_deal_documents(s.deal_id)

        capital_stack = UnderwritingEngine.capital_stack(
            senior_debt=loan_amt,
            mezz_debt=safe_float(s.get("mezz_debt", 0)),
            pref_equity=safe_float(s.get("pref_equity", 0)),
            sponsor_equity=req_equity,
            noi=safe_float(s.noi),
            senior_rate=safe_float(s.rate),
            mezz_rate=safe_float(s.get("mezz_rate", 0.10)),
            pref_rate=safe_float(s.get("pref_rate", 0.09)),
        )

        save_col, excel_col, package_col = st.columns(3)

        with save_col:
            st.write("### 💾 Save to Database")

            if st.button("💾 Save Deal", use_container_width=True, key="export_save_deal"):
                errors_now, warnings_now = ValidationEngine.validate(state)

                if errors_now:
                    st.error("Deal has blocking validation errors.")
                    for error in errors_now:
                        st.error(f"• {error}")
                else:
                    if warnings_now:
                        for warning in warnings_now:
                            st.warning(f"• {warning}")

                    if DatabaseManager.save_deal(
                        state["deal_id"],
                        state.get("deal_name", "Untitled Deal"),
                        state,
                    ):
                        s.unsaved_changes = False
                        s.last_saved_at = utc_now_iso()
                        st.success(f"✅ Saved at {str(s.last_saved_at)[:19]}")
                    else:
                        st.error("Save failed.")

            if s.get("last_saved_at"):
                st.caption(f"Last saved: {str(s.last_saved_at)[:19]}")

            st.markdown("---")

            st.write("### 🔐 JSON Options")

            encrypt_json = st.checkbox(
                "Encrypt JSON export with password",
                value=False,
                key="encrypt_json_export",
            )

            export_password = ""

            if encrypt_json:
                export_password = st.text_input(
                    "Export Password",
                    type="password",
                    key="json_export_password",
                )

                if not DEPENDENCIES.get("crypto"):
                    st.warning("Install cryptography to enable encrypted JSON export.")

        with excel_col:
            st.write("### 📊 Excel Export")

            if DEPENDENCIES.get("excel_write"):
                try:
                    excel_output = io.BytesIO()

                    summary_df = pd.DataFrame(
                        {
                            "Metric": [
                                "Deal Name",
                                "Sponsor",
                                "Property Address",
                                "Property Type",
                                "Transaction Type",
                                "Lender Profile",
                                "Max Proceeds",
                                "Binding Constraint",
                                "Total Uses",
                                "Required Equity",
                                "Actual LTV",
                                "Actual LTC",
                                "Actual DSCR",
                                "Actual Debt Yield",
                                "Monthly Payment",
                                "Annual Debt Service",
                                "Balloon Balance",
                                "Score",
                                "Tier",
                                "As Of",
                            ],
                            "Value": [
                                s.deal_name,
                                s.sponsor,
                                s.property_address,
                                s.property_type,
                                s.transaction_type,
                                s.lender_profile,
                                loan_amt,
                                gate,
                                total_uses,
                                req_equity,
                                actual_ltv,
                                actual_ltc,
                                actual_dscr,
                                actual_dy,
                                monthly_pmt,
                                annual_ds,
                                balloon,
                                score,
                                tier,
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            ],
                        }
                    )

                    constraints_export = pd.DataFrame(
                        {
                            "Constraint": list(gates.keys()),
                            "Max Proceeds": list(gates.values()),
                            "Binding": [
                                "Yes" if key == gate else "No"
                                for key in gates.keys()
                            ],
                        }
                    )

                    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
                        summary_df.to_excel(writer, sheet_name="Summary", index=False)
                        constraints_export.to_excel(writer, sheet_name="Constraints", index=False)
                        pd.DataFrame([capital_stack]).to_excel(writer, sheet_name="Capital Stack", index=False)
                        rr_df.to_excel(writer, sheet_name="Rent Roll", index=False)
                        amort_df.to_excel(writer, sheet_name="Amortization", index=False)
                        sensitivity_df.to_excel(writer, sheet_name="Sensitivity", index=False)
                        heatmap_df.to_excel(writer, sheet_name="Heatmap Data", index=False)

                        if docs_df is not None and not docs_df.empty:
                            docs_df.to_excel(writer, sheet_name="Documents", index=False)

                        if versions_df is not None and not versions_df.empty:
                            versions_df.to_excel(writer, sheet_name="Versions", index=False)

                        workbook = writer.book

                        header_fmt = workbook.add_format(
                            {
                                "bold": True,
                                "bg_color": "#CFB87C",
                                "font_color": "#0B0F19",
                                "border": 1,
                            }
                        )

                        money_fmt = workbook.add_format({"num_format": "$#,##0"})
                        percent_fmt = workbook.add_format({"num_format": "0.00%"})
                        ratio_fmt = workbook.add_format({"num_format": "0.00x"})
                        default_fmt = workbook.add_format({"border": 0})

                        for sheet_name, worksheet in writer.sheets.items():
                            worksheet.freeze_panes(1, 0)
                            worksheet.set_row(0, None, header_fmt)
                            worksheet.set_column(0, 0, 24, default_fmt)
                            worksheet.set_column(1, 12, 18, default_fmt)

                        summary_ws = writer.sheets["Summary"]
                        summary_ws.set_column(0, 0, 28)
                        summary_ws.set_column(1, 1, 26)

                        # Apply basic value formatting by row in Summary.
                        for row_idx, metric in enumerate(summary_df["Metric"], start=1):
                            if metric in [
                                "Max Proceeds",
                                "Total Uses",
                                "Required Equity",
                                "Monthly Payment",
                                "Annual Debt Service",
                                "Balloon Balance",
                            ]:
                                summary_ws.write_number(row_idx, 1, safe_float(summary_df.iloc[row_idx - 1]["Value"]), money_fmt)
                            elif metric in [
                                "Actual LTV",
                                "Actual LTC",
                                "Actual Debt Yield",
                            ]:
                                summary_ws.write_number(row_idx, 1, safe_float(summary_df.iloc[row_idx - 1]["Value"]), percent_fmt)
                            elif metric == "Actual DSCR":
                                summary_ws.write_number(row_idx, 1, safe_float(summary_df.iloc[row_idx - 1]["Value"]), ratio_fmt)

                    st.download_button(
                        "📊 Download Excel Workbook",
                        data=excel_output.getvalue(),
                        file_name=f"{safe_deal_name}_{timestamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"Excel export failed: {str(e)[:300]}")
            else:
                st.warning("Install xlsxwriter for Excel export.")

        with package_col:
            st.write("### 📦 JSON / ZIP Package")

            try:
                if encrypt_json and export_password and DEPENDENCIES.get("crypto"):
                    json_payload = encrypt_deal_state(state, export_password)
                    json_filename = f"{safe_deal_name}_{timestamp}_encrypted.json"
                    json_mime = "application/octet-stream"
                else:
                    json_payload = json.dumps(state, indent=2, default=str)
                    json_filename = f"{safe_deal_name}_{timestamp}.json"
                    json_mime = "application/json"

                st.download_button(
                    "📄 Download JSON",
                    data=json_payload,
                    file_name=json_filename,
                    mime=json_mime,
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"JSON export failed: {str(e)[:200]}")

            try:
                zip_buffer = io.BytesIO()

                plain_state_json = json.dumps(state, indent=2, default=str)

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr("deal_state.json", plain_state_json)
                    zf.writestr("summary.csv", pd.DataFrame(
                        {
                            "Metric": [
                                "Deal Name",
                                "Sponsor",
                                "Property",
                                "Type",
                                "Transaction",
                                "Max Proceeds",
                                "Constraint",
                                "Actual LTV",
                                "Actual LTC",
                                "Actual DSCR",
                                "Debt Yield",
                                "Score",
                                "Tier",
                            ],
                            "Value": [
                                s.deal_name,
                                s.sponsor,
                                s.property_address,
                                s.property_type,
                                s.transaction_type,
                                loan_amt,
                                gate,
                                actual_ltv,
                                actual_ltc,
                                actual_dscr,
                                actual_dy,
                                score,
                                tier,
                            ],
                        }
                    ).to_csv(index=False))
                    zf.writestr("constraints.csv", pd.DataFrame(
                        {
                            "Constraint": list(gates.keys()),
                            "Max Proceeds": list(gates.values()),
                        }
                    ).to_csv(index=False))
                    zf.writestr("capital_stack.json", json.dumps(capital_stack, indent=2, default=str))
                    zf.writestr("rent_roll.csv", rr_df.to_csv(index=False))
                    zf.writestr("amortization.csv", amort_df.to_csv(index=False))
                    zf.writestr("sensitivity.csv", sensitivity_df.to_csv(index=False))
                    zf.writestr("heatmap_data.csv", heatmap_df.to_csv(index=False))

                    if docs_df is not None and not docs_df.empty:
                        zf.writestr("document_inventory.csv", docs_df.to_csv(index=False))

                    if versions_df is not None and not versions_df.empty:
                        zf.writestr("version_history.csv", versions_df.to_csv(index=False))

                    if audit_df is not None and not audit_df.empty:
                        zf.writestr("audit_log.csv", audit_df.to_csv(index=False))

                    zf.writestr(
                        "README.txt",
                        (
                            "ALENZA CAPITAL OS EXPORT PACKAGE\n"
                            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"Deal: {s.deal_name}\n"
                            f"App Version: {APP_VERSION}\n\n"
                            "This package contains indicative underwriting outputs only. "
                            "Final loan terms remain subject to credit approval, diligence, and definitive documentation.\n"
                        ),
                    )

                st.download_button(
                    "📦 Download ZIP Package",
                    data=zip_buffer.getvalue(),
                    file_name=f"{safe_deal_name}_{timestamp}.zip",
                    mime="application/zip",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"ZIP export failed: {str(e)[:300]}")

        st.markdown("---")

        st.write("### Export Contents")

        export_contents = pd.DataFrame(
            [
                {"File / Sheet": "Summary", "Included": "Yes"},
                {"File / Sheet": "Constraints", "Included": "Yes"},
                {"File / Sheet": "Capital Stack", "Included": "Yes"},
                {"File / Sheet": "Rent Roll", "Included": "Yes"},
                {"File / Sheet": "Amortization", "Included": "Yes"},
                {"File / Sheet": "Sensitivity", "Included": "Yes"},
                {"File / Sheet": "Heatmap Data", "Included": "Yes"},
                {"File / Sheet": "Document Inventory", "Included": "Yes" if docs_df is not None and not docs_df.empty else "No documents yet"},
                {"File / Sheet": "Version History", "Included": "Yes" if versions_df is not None and not versions_df.empty else "No versions yet"},
                {"File / Sheet": "Audit Log", "Included": "Yes" if audit_df is not None and not audit_df.empty else "No audit rows yet"},
            ]
        )

        st.dataframe(
            export_contents,
            hide_index=True,
            use_container_width=True,
        )

    # ==========================================
    # TAB 8: QA & HEALTH
    # ==========================================

    with tabs[8]:
        st.subheader("✅ Quality Assurance & System Health")
        st.caption("Financial regression checks, dependency status, deal validation, and version history.")

        qa_col, health_col = st.columns([1.25, 1])

        with qa_col:
            st.write("### 🔬 Financial Self-Tests")

            test_df = run_financial_self_tests()

            st.dataframe(
                test_df,
                hide_index=True,
                use_container_width=True,
            )

            total_tests = len(test_df)
            failed_tests = int((test_df["Status"] == "❌ FAIL").sum()) if not test_df.empty else 0
            passed_tests = total_tests - failed_tests

            c1, c2, c3 = st.columns(3)
            c1.metric("Total Tests", f"{total_tests}")
            c2.metric("Passed", f"{passed_tests}")
            c3.metric("Failed", f"{failed_tests}")

            if failed_tests > 0:
                st.error("Financial checks failed. Review outputs before relying on the model.")
            else:
                st.success("All financial regression checks passed.")

        with health_col:
            st.write("### 🖥️ System Health")

            health_checks = []

            try:
                deals_check = DatabaseManager.get_all_deals()
                health_checks.append(
                    {
                        "Component": "Database",
                        "Status": "✅ OK",
                        "Detail": f"{len(deals_check)} saved deal(s)",
                    }
                )
            except Exception as e:
                health_checks.append(
                    {
                        "Component": "Database",
                        "Status": "❌ FAIL",
                        "Detail": str(e)[:150],
                    }
                )

            try:
                DATA_DIR.mkdir(parents=True, exist_ok=True)
                DOC_DIR.mkdir(parents=True, exist_ok=True)
                health_checks.append(
                    {
                        "Component": "Storage",
                        "Status": "✅ OK",
                        "Detail": str(DATA_DIR),
                    }
                )
            except Exception as e:
                health_checks.append(
                    {
                        "Component": "Storage",
                        "Status": "❌ FAIL",
                        "Detail": str(e)[:150],
                    }
                )

            for dep, available in DEPENDENCIES.items():
                if dep in ["ocr", "pdf", "crypto", "xls_read"]:
                    importance = "Optional"
                elif dep in ["excel_write", "excel_read"]:
                    importance = "Recommended"
                else:
                    importance = "Recommended"

                health_checks.append(
                    {
                        "Component": f"Module: {dep}",
                        "Status": "✅ Available" if available else "⚠️ Missing",
                        "Detail": importance,
                    }
                )

            encryption_key_present = bool(get_db_encryption_key())

            health_checks.append(
                {
                    "Component": "DB Encryption Key",
                    "Status": "✅ Present" if encryption_key_present else "⚠️ Not Set",
                    "Detail": "ALENZA_DB_ENCRYPTION_KEY",
                }
            )

            st.dataframe(
                pd.DataFrame(health_checks),
                hide_index=True,
                use_container_width=True,
            )

        st.markdown("---")

        validation_col, version_col = st.columns(2)

        with validation_col:
            st.write("### 🔍 Current Deal Validation")

            current_state = extract_clean_state()
            validation_errors, validation_warnings = ValidationEngine.validate(current_state)

            if validation_errors:
                st.error(f"{len(validation_errors)} blocking validation error(s).")
                for error in validation_errors:
                    st.error(f"❌ {error}")
            else:
                st.success("No blocking validation errors.")

            if validation_warnings:
                st.warning(f"{len(validation_warnings)} review warning(s).")
                for warning in validation_warnings:
                    st.warning(f"⚠️ {warning}")
            else:
                st.info("No review warnings.")

            st.markdown("---")

            st.write("### 📌 Current Model Snapshot")

            snapshot_df = pd.DataFrame(
                {
                    "Metric": [
                        "Deal ID",
                        "Schema Version",
                        "App Version",
                        "Unsaved Changes",
                        "Loan Amount",
                        "Binding Constraint",
                        "Actual DSCR",
                        "Actual LTV",
                        "Actual LTC",
                        "Actual Debt Yield",
                    ],
                    "Value": [
                        s.deal_id,
                        SCHEMA_VERSION,
                        APP_VERSION,
                        bool(s.get("unsaved_changes", False)),
                        f"${loan_amt:,.0f}",
                        gate,
                        f"{actual_dscr:.2f}x",
                        f"{actual_ltv:.2%}",
                        f"{actual_ltc:.2%}",
                        f"{actual_dy:.2%}",
                    ],
                }
            )

            st.dataframe(
                snapshot_df,
                hide_index=True,
                use_container_width=True,
            )

        with version_col:
            st.write("### 📜 Deal Version History")

            versions_df = DatabaseManager.get_deal_versions(s.deal_id, limit=15)

            if versions_df is not None and not versions_df.empty:
                st.dataframe(
                    versions_df,
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.info("No version history yet. Save the deal to create version records.")

            st.markdown("---")

            st.write("### 🧾 Audit Log")

            audit_df = DatabaseManager.get_audit_log(s.deal_id, limit=25)

            if audit_df is not None and not audit_df.empty:
                st.dataframe(
                    audit_df,
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.info("No audit records yet.")

        st.markdown("---")

        st.write("### 🧪 Manual QA Checklist")

        checklist_df = pd.DataFrame(
            [
                {
                    "Check": "Deal has name and property profile",
                    "Status": "✅ Pass" if s.get("deal_name") and s.get("property_type") else "⚠️ Review",
                },
                {
                    "Check": "NOI and appraisal populated",
                    "Status": "✅ Pass" if safe_float(s.noi) > 0 and safe_float(s.appraisal) > 0 else "⚠️ Review",
                },
                {
                    "Check": "Rent roll has usable records",
                    "Status": "✅ Pass" if rr_df is not None and not rr_df.empty else "⚠️ Review",
                },
                {
                    "Check": "Loan proceeds calculated",
                    "Status": "✅ Pass" if loan_amt > 0 else "⚠️ Review",
                },
                {
                    "Check": "DSCR is above 1.00x",
                    "Status": "✅ Pass" if actual_dscr >= 1.0 else "⚠️ Review",
                },
                {
                    "Check": "Document vault initialized",
                    "Status": "✅ Pass" if DOC_DIR.exists() else "❌ Fail",
                },
            ]
        )

        st.dataframe(
            checklist_df,
            hide_index=True,
            use_container_width=True,
        )

    # ==========================================
    # FOOTER
    # ==========================================

    st.markdown("---")

    st.caption(
        "⚠️ **DISCLAIMER:** ALENZA CAPITAL OS is an indicative modeling tool for commercial real estate underwriting. "
        "Outputs do not constitute a loan commitment, appraisal, valuation opinion, investment recommendation, legal advice, "
        "tax advice, or credit approval. Final terms remain subject to formal credit approval, third-party diligence, "
        "verified market data, and definitive documentation."
    )

    if s.get("unsaved_changes"):
        st.sidebar.warning(
            f"⏰ Unsaved changes — last check: {datetime.now().strftime('%H:%M:%S')}"
        )


# ==========================================
# APP ENTRY POINT
# ==========================================

if __name__ == "__main__":
    main()
