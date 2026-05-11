"""
Alenza Capital OS v3.0.3
Enterprise underwriting workspace
Midnight Slate and CU Gold theme
"""

import streamlit as st
import streamlit.components.v1 as components
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

APP_VERSION = "3.0.3"
SCHEMA_VERSION = 4
MAX_UPLOAD_SIZE_MB = 50
AUTOSAVE_INTERVAL_SECONDS = 300

DATA_DIR = Path("alenza_data")
DOC_DIR = DATA_DIR / "documents"
BACKUP_DIR = DATA_DIR / "backups"
EXPORT_DIR = DATA_DIR / "exports"
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
    {
        "Tenant": "Main Anchor",
        "SF": 25000,
        "Remaining Term": 5.5,
        "Monthly Rent": 45000,
    },
    {
        "Tenant": "In-Line A",
        "SF": 3500,
        "Remaining Term": 1.2,
        "Monthly Rent": 8000,
    },
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
        "numpy_financial": False,
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
        import plotly.graph_objects as go
        deps["plotly"] = True
    except ImportError:
        pass

    try:
        from cryptography.fernet import Fernet
        deps["crypto"] = True
    except ImportError:
        pass

    try:
        import numpy_financial as npf
        deps["numpy_financial"] = True
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


def safe_ratio(numerator: Any, denominator: Any) -> float:
    numerator = safe_float(numerator)
    denominator = safe_float(denominator)

    if denominator <= 0:
        return 0.0

    return numerator / denominator


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
        "pf_revenue_growth",
        "pf_expense_growth",
        "pf_expense_ratio",
        "pf_terminal_growth",
        "pf_exit_cap",
        "pf_selling_costs",
        "pf_projection_years",
        "diligence_notes",
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


def inject_responsive_css():
    st.markdown(
        """
        <style>
        @media (max-width: 900px) {
            section[data-testid="stSidebar"] {
                width: 18rem !important;
            }

            div[data-testid="column"] {
                width: 100% !important;
                flex: 1 1 100% !important;
            }

            .stMetric {
                padding: 0.35rem 0;
            }

            div[data-testid="stTabs"] button {
                font-size: 0.78rem;
                padding: 0.35rem 0.45rem;
            }
        }

        div[data-testid="stMetric"] {
            background: rgba(15, 23, 42, 0.35);
            border: 1px solid rgba(207, 184, 124, 0.18);
            border-radius: 14px;
            padding: 0.75rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
# DATABASE
# ==========================================

def get_db_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    DOC_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)

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
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)

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
        "Bank / Credit Union": {
            "max_ltv": 0.75,
            "min_dscr": 1.25,
            "min_dy": 0.08,
        },
        "LifeCo / Core": {
            "max_ltv": 0.65,
            "min_dscr": 1.35,
            "min_dy": 0.09,
        },
        "Bridge / Private": {
            "max_ltv": 0.85,
            "min_dscr": 1.00,
            "min_dy": 0.07,
        },
        "CMHC Multifamily": {
            "max_ltv": 0.95,
            "min_dscr": 1.10,
            "min_dy": 0.05,
        },
    }

    @staticmethod
    def size_loan(
        noi,
        appraisal,
        purchase_price,
        closing_costs,
        reserves,
        fees_pct,
        rate,
        amort,
        term,
        is_io,
        target_ltv,
        target_ltc,
        target_dscr,
        target_dy,
    ) -> tuple:
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

        gates = {
            "LTV": 0.0,
            "LTC": 0.0,
            "DSCR": 0.0,
            "Debt Yield": 0.0,
        }

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
        required_equity = total_uses - loan

        return loan, gate, gates, total_uses, required_equity

    @staticmethod
    def amort_schedule(loan_amt, rate, amort_yrs, term_yrs, is_io) -> tuple:
        loan_amt = max(0, safe_float(loan_amt))
        rate = normalize_percent(rate, 0.0525, 0.30)
        amort_yrs = max(1, min(40, safe_int(amort_yrs)))
        term_yrs = max(1, min(40, safe_int(term_yrs)))

        if loan_amt <= 0:
            return (
                pd.DataFrame(columns=["Period", "Payment", "Principal", "Interest", "Balance"]),
                0.0,
                0.0,
            )

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

            schedule.append(
                {
                    "Period": period,
                    "Payment": monthly_pmt if balance > 0 else 0.0,
                    "Principal": principal_paid,
                    "Interest": interest,
                    "Balance": balance,
                }
            )

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

        occupied = df[
            (~df["Tenant"].str.lower().isin(vacant_keywords))
            & (df["SF"] > 0)
        ]

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

        breakeven = annual_debt_service / current_noi * current_occupancy
        return max(0.0, min(1.5, breakeven))

    @staticmethod
    def capital_stack(
        senior_debt: float,
        mezz_debt: float,
        pref_equity: float,
        sponsor_equity: float,
        noi: float,
        senior_rate: float,
        mezz_rate: float,
        pref_rate: float,
    ) -> dict:
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
        limits = UnderwritingEngine.LENDER_LIMITS.get(
            profile,
            UnderwritingEngine.LENDER_LIMITS["Bank / Credit Union"],
        )

        actual_ltv = max(0, safe_float(actual_ltv))
        actual_ltc = max(0, safe_float(actual_ltc))
        actual_dscr = max(0, safe_float(actual_dscr))
        actual_dy = max(0, safe_float(actual_dy))

        ltv_score = max(0, 300 * (1 - actual_ltv / limits["max_ltv"])) if limits["max_ltv"] > 0 else 0

        if limits["min_dscr"] > 1.0 and actual_dscr > 1.0:
            dscr_score = max(0, 300 * (actual_dscr - 1.0) / (limits["min_dscr"] - 1.0))
        else:
            dscr_score = 300 if actual_dscr >= 1.0 else 0

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
        "pf_revenue_growth": 0.03,
        "pf_expense_growth": 0.035,
        "pf_expense_ratio": 0.40,
        "pf_terminal_growth": 0.02,
        "pf_exit_cap": 0.065,
        "pf_selling_costs": 0.015,
        "pf_projection_years": 10,
        "rent_roll_dict": DEFAULT_RENT_ROLL.copy(),
        "diligence_notes": "",
        "last_saved_at": None,
        "last_autosaved_at": None,
        "autosave_enabled": True,
        "autosave_status": "",
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
    normalized["diligence_notes"] = str(normalized.get("diligence_notes") or "")

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
    normalized["pf_revenue_growth"] = normalize_percent(normalized.get("pf_revenue_growth"), 0.03, 0.20)
    normalized["pf_expense_growth"] = normalize_percent(normalized.get("pf_expense_growth"), 0.035, 0.20)
    normalized["pf_expense_ratio"] = normalize_percent(normalized.get("pf_expense_ratio"), 0.40, 0.95)
    normalized["pf_terminal_growth"] = normalize_percent(normalized.get("pf_terminal_growth"), 0.02, 0.10)
    normalized["pf_exit_cap"] = normalize_percent(normalized.get("pf_exit_cap"), 0.065, 0.20)
    normalized["pf_selling_costs"] = normalize_percent(normalized.get("pf_selling_costs"), 0.015, 0.10)
    normalized["pf_projection_years"] = max(5, min(30, safe_int(normalized.get("pf_projection_years"), 10)))

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

    for key in [
        "last_saved_at",
        "last_autosaved_at",
        "autosave_status",
        "unsaved_changes",
        "app_version",
        "schema_version",
    ]:
        state.pop(key, None)

    return hash_state(state)


# ==========================================
# SENSITIVITY ENGINE
# ==========================================

class SensitivityEngine:
    @staticmethod
    def generate_matrix(state: dict) -> pd.DataFrame:
        scenarios = [
            ("Base", 0, 0),
            ("Rate +1%", 0.01, 0),
            ("Rate -1%", -0.01, 0),
            ("NOI -10%", 0, -0.10),
            ("NOI +10%", 0, 0.10),
            ("Combined Stress", 0.01, -0.10),
        ]

        results = []
        base_proceeds = None

        for scenario_name, rate_adj, noi_adj in scenarios:
            adjusted_rate = max(0.001, safe_float(state.get("rate", 0.0525)) + rate_adj)
            adjusted_noi = max(0, safe_float(state.get("noi", 0)) * (1 + noi_adj))

            loan, gate, _, _, _ = UnderwritingEngine.size_loan(
                noi=adjusted_noi,
                appraisal=safe_float(state.get("appraisal", 0)),
                purchase_price=safe_float(state.get("purchase_price", 0)),
                closing_costs=safe_float(state.get("closing_costs", 0)),
                reserves=safe_float(state.get("reserves", 0)),
                fees_pct=safe_float(state.get("fees", 0)),
                rate=adjusted_rate,
                amort=safe_int(state.get("amort", 25)),
                term=safe_int(state.get("term", 5)),
                is_io=bool(state.get("is_io", False)),
                target_ltv=safe_float(state.get("target_ltv", 0.75)),
                target_ltc=safe_float(state.get("target_ltc", 0.80)),
                target_dscr=safe_float(state.get("target_dscr", 1.25)),
                target_dy=safe_float(state.get("target_dy", 0.085)),
            )

            if scenario_name == "Base":
                base_proceeds = loan

            if base_proceeds and base_proceeds > 0:
                change_str = f"{((loan - base_proceeds) / base_proceeds * 100):+.1f}%"
            else:
                change_str = "N/A"

            results.append(
                {
                    "Scenario": scenario_name,
                    "Rate": f"{adjusted_rate * 100:.2f}%",
                    "NOI": f"${adjusted_noi:,.0f}",
                    "Max Proceeds": f"${loan:,.0f}",
                    "Constraint": gate,
                    "Change from Base": change_str,
                }
            )

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
                    noi=adjusted_noi,
                    appraisal=safe_float(state.get("appraisal", 0)),
                    purchase_price=safe_float(state.get("purchase_price", 0)),
                    closing_costs=safe_float(state.get("closing_costs", 0)),
                    reserves=safe_float(state.get("reserves", 0)),
                    fees_pct=safe_float(state.get("fees", 0)),
                    rate=adjusted_rate,
                    amort=safe_int(state.get("amort", 25)),
                    term=safe_int(state.get("term", 5)),
                    is_io=bool(state.get("is_io", False)),
                    target_ltv=safe_float(state.get("target_ltv", 0.75)),
                    target_ltc=safe_float(state.get("target_ltc", 0.80)),
                    target_dscr=safe_float(state.get("target_dscr", 1.25)),
                    target_dy=safe_float(state.get("target_dy", 0.085)),
                )

                row[f"Rate {rate_adj:+.1%}"] = loan

            rows.append(row)

        return pd.DataFrame(rows)


# ==========================================
# INVESTMENT & PRO FORMA ENGINE
# ==========================================

class InvestmentEngine:
    @staticmethod
    def calculate_pro_forma(
        stabilized_noi: float,
        revenue_growth: float,
        expense_growth: float,
        expense_ratio: float = 0.40,
        years: int = 10,
    ) -> pd.DataFrame:
        rows = []

        stabilized_noi = max(0, safe_float(stabilized_noi))
        revenue_growth = normalize_percent(revenue_growth, 0.03, 0.20)
        expense_growth = normalize_percent(expense_growth, 0.035, 0.20)
        expense_ratio = normalize_percent(expense_ratio, 0.40, 0.95)
        years = max(1, min(30, safe_int(years, 10)))

        if stabilized_noi <= 0:
            revenue = 0.0
            expenses = 0.0
        else:
            revenue = stabilized_noi / max(1 - expense_ratio, 0.01)
            expenses = revenue * expense_ratio

        for year in range(1, years + 1):
            noi = revenue - expenses
            margin = safe_ratio(noi, revenue)

            rows.append(
                {
                    "Year": year,
                    "Revenue": revenue,
                    "Expenses": expenses,
                    "Projected NOI": noi,
                    "NOI Margin": margin,
                    "Revenue Growth": 0.0 if year == 1 else revenue_growth,
                    "Expense Growth": 0.0 if year == 1 else expense_growth,
                }
            )

            revenue *= 1 + revenue_growth
            expenses *= 1 + expense_growth

        return pd.DataFrame(rows)

    @staticmethod
    def solve_returns(
        purchase_price: float,
        loan_amt: float,
        pro_forma_df: pd.DataFrame,
        exit_cap: float,
        selling_costs: float,
        annual_ds: float,
        balloon_balance: float,
        terminal_growth: float = 0.02,
    ) -> Dict[str, Any]:
        purchase_price = max(0, safe_float(purchase_price))
        loan_amt = max(0, safe_float(loan_amt))
        exit_cap = normalize_percent(exit_cap, 0.06, 0.20)
        selling_costs = normalize_percent(selling_costs, 0.015, 0.10)
        annual_ds = max(0, safe_float(annual_ds))
        balloon_balance = max(0, safe_float(balloon_balance))
        terminal_growth = normalize_percent(terminal_growth, 0.02, 0.10)

        equity_in = max(0, purchase_price - loan_amt)

        if equity_in <= 0 or pro_forma_df is None or pro_forma_df.empty:
            return {
                "IRR": 0.0,
                "Equity Multiple": 0.0,
                "Net Exit Proceeds": 0.0,
                "Gross Exit Value": 0.0,
                "Total Cash Flow": 0.0,
                "Exit NOI": 0.0,
            }

        cash_flows = [-equity_in]

        for _, row in pro_forma_df.iterrows():
            cash_flows.append(max(0, safe_float(row.get("Projected NOI"))) - annual_ds)

        exit_noi = safe_float(pro_forma_df.iloc[-1]["Projected NOI"]) * (1 + terminal_growth)
        gross_exit_value = exit_noi / exit_cap if exit_cap > 0 else 0
        net_exit_value = gross_exit_value * (1 - selling_costs)
        net_exit_proceeds = net_exit_value - balloon_balance

        cash_flows[-1] += net_exit_proceeds

        try:
            if DEPENDENCIES.get("numpy_financial"):
                import numpy_financial as npf
                irr = npf.irr(cash_flows)
            else:
                irr = 0.0

            if irr is None or np.isnan(irr):
                irr = 0.0

        except Exception:
            irr = 0.0

        total_distributions = sum(cash_flows[1:])
        equity_multiple = total_distributions / equity_in if equity_in > 0 else 0.0

        return {
            "IRR": irr,
            "Equity Multiple": equity_multiple,
            "Net Exit Proceeds": net_exit_proceeds,
            "Gross Exit Value": gross_exit_value,
            "Total Cash Flow": total_distributions,
            "Exit NOI": exit_noi,
        }


# ==========================================
# VALIDATION ENGINE
# ==========================================

class ValidationEngine:
    @staticmethod
    def validate(state: dict) -> tuple:
        errors = []
        warnings = []

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

            duplicates = sorted(
                {
                    t
                    for t in tenant_names
                    if tenant_names.count(t) > 1
                    and t not in ["vacant", "empty", "available", "vacancy"]
                }
            )

            if duplicates:
                warnings.append(f"Duplicate tenant names detected: {', '.join(duplicates[:5])}")

        return errors, warnings


# ==========================================
# FINANCIAL SELF TESTS
# ==========================================

def run_financial_self_tests() -> pd.DataFrame:
    rows = []

    def check(name: str, passed: bool, detail: str):
        rows.append(
            {
                "Test": name,
                "Status": "✅ PASS" if passed else "❌ FAIL",
                "Detail": detail,
            }
        )

    loan, gate, gates, uses, equity = UnderwritingEngine.size_loan(
        1_000_000,
        10_000_000,
        9_000_000,
        100_000,
        0,
        0.01,
        0.06,
        25,
        5,
        False,
        0.70,
        0.80,
        1.25,
        0.09,
    )

    check("Loan proceeds positive", loan > 0, f"${loan:,.0f}, Gate: {gate}")
    check("Binding gate equals minimum proceeds", abs(loan - min(gates.values())) < 1, f"Min: ${min(gates.values()):,.0f}")
    check("Total uses positive", uses > 0, f"${uses:,.0f}")
    check("Equity requirement calculated", isinstance(equity, float), f"${equity:,.0f}")

    amort, pmt, balloon_test = UnderwritingEngine.amort_schedule(
        1_000_000,
        0.06,
        25,
        5,
        False,
    )

    check("60 periods in 5-year term", len(amort) == 60, f"Periods: {len(amort)}")
    check("Monthly payment positive", pmt > 0, f"${pmt:,.2f}")
    check("Amortizing balloon below original", balloon_test < 1_000_000, f"Balloon: ${balloon_test:,.0f}")

    io_amort, io_pmt, io_balloon = UnderwritingEngine.amort_schedule(
        1_000_000,
        0.06,
        25,
        5,
        True,
    )

    check("IO payment equals interest-only amount", abs(io_pmt - 5_000) < 1, f"PMT: ${io_pmt:,.2f}")
    check("IO balloon equals original loan", abs(io_balloon - 1_000_000) < 1, f"Balloon: ${io_balloon:,.0f}")

    rr = pd.DataFrame(
        [
            {"Tenant": "A", "SF": 10_000, "Remaining Term": 5, "Monthly Rent": 20_000},
            {"Tenant": "Vacant", "SF": 2_000, "Remaining Term": 0, "Monthly Rent": 0},
        ]
    )

    sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr)

    check("Occupancy equals 83.3%", abs(occ - 0.8333) < 0.01, f"{occ:.2%}")
    check("Annual rent equals $240K", abs(ann_rent - 240_000) < 1, f"${ann_rent:,.0f}")
    check("Rent PSF equals $24", abs(psf - 24) < 0.1, f"${psf:.2f}")
    check("WALT equals 5.0 years", abs(walt - 5.0) < 0.01, f"{walt:.2f}")

    zero_loan, zero_gate, zero_gates, zero_uses, zero_equity = UnderwritingEngine.size_loan(
        0,
        10_000_000,
        9_000_000,
        0,
        0,
        0.01,
        0.06,
        25,
        5,
        False,
        0.70,
        0.80,
        1.25,
        0.09,
    )

    check(
        "Zero NOI produces zero income-based proceeds",
        zero_gates["DSCR"] == 0 and zero_gates["Debt Yield"] == 0 and zero_loan == 0,
        f"Loan: ${zero_loan:,.0f}",
    )

    empty_sf, empty_occ, empty_rent, empty_psf, empty_walt, empty_exp = UnderwritingEngine.rent_roll_metrics(
        pd.DataFrame()
    )

    check(
        "Empty rent roll returns zero metrics",
        empty_sf == 0 and empty_occ == 0 and empty_rent == 0,
        "Empty rent roll handled safely",
    )

    neg_loan, neg_gate, neg_gates, neg_uses, neg_equity = UnderwritingEngine.size_loan(
        -1_000_000,
        -10_000_000,
        -9_000_000,
        -100_000,
        -50_000,
        -0.01,
        -0.05,
        -25,
        -5,
        False,
        -0.70,
        -0.80,
        -1.25,
        -0.09,
    )

    check(
        "Negative inputs are sanitized",
        neg_loan >= 0 and neg_uses >= 0,
        f"Loan: ${neg_loan:,.0f}, Uses: ${neg_uses:,.0f}",
    )

    be_occ = UnderwritingEngine.breakeven_occupancy(
        current_noi=1_000_000,
        current_occupancy=0.90,
        annual_debt_service=750_000,
    )

    check("Breakeven occupancy calculation", abs(be_occ - 0.675) < 0.001, f"{be_occ:.2%}")

    stack = UnderwritingEngine.capital_stack(
        senior_debt=7_000_000,
        mezz_debt=1_000_000,
        pref_equity=500_000,
        sponsor_equity=1_500_000,
        noi=1_000_000,
        senior_rate=0.06,
        mezz_rate=0.11,
        pref_rate=0.09,
    )

    check(
        "Capital stack fixed charge coverage positive",
        stack["Fixed Charge Coverage"] > 0,
        f"{stack['Fixed Charge Coverage']:.2f}x",
    )

    pro_forma_df = InvestmentEngine.calculate_pro_forma(
        stabilized_noi=1_000_000,
        revenue_growth=0.03,
        expense_growth=0.035,
        expense_ratio=0.40,
        years=10,
    )

    check(
        "Pro forma produces 10 years",
        len(pro_forma_df) == 10,
        f"Rows: {len(pro_forma_df)}",
    )

    returns = InvestmentEngine.solve_returns(
        purchase_price=10_000_000,
        loan_amt=6_000_000,
        pro_forma_df=pro_forma_df,
        exit_cap=0.06,
        selling_costs=0.015,
        annual_ds=400_000,
        balloon_balance=5_500_000,
        terminal_growth=0.02,
    )

    check(
        "Returns engine produces equity multiple",
        returns["Equity Multiple"] >= 0,
        f"{returns['Equity Multiple']:.2f}x",
    )

    return pd.DataFrame(rows)


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

            try:
                st.toast(f"Autosaved at {now[:19]}", icon="✅")
            except Exception:
                pass
        else:
            st.session_state.autosave_status = f"Saved at {now[:19]}"

            try:
                st.toast("Deal saved", icon="✅")
            except Exception:
                pass

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


def try_autorefresh(interval_ms: int, key: str):
    try:
        from streamlit_autorefresh import st_autorefresh
        return st_autorefresh(interval=interval_ms, key=key)
    except Exception:
        return None


DatabaseManager.init_db()


# ==========================================
# MAIN APPLICATION
# ==========================================

def main():
    """Main application entry point."""

    DatabaseManager.init_db()
    initialize_session_state()

    s = st.session_state

    inject_responsive_css()

    components.html(
        """
        <script>
        document.addEventListener("keydown", function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
                e.preventDefault();
            }
        });
        </script>
        """,
        height=0,
    )

    if "_loading_deal" not in s:
        s["_loading_deal"] = False

    if "autosave_enabled" not in s:
        s["autosave_enabled"] = True

    if s.get("autosave_enabled", True):
        try_autorefresh(
            interval_ms=AUTOSAVE_INTERVAL_SECONDS * 1000,
            key="autosave_timer",
        )

    state_hash_before = stable_state_hash()

    pre_state = extract_clean_state()
    pre_errors, pre_warnings = ValidationEngine.validate(pre_state)

    pre_loan, pre_gate, pre_gates, pre_total_uses, pre_req_equity = UnderwritingEngine.size_loan(
        noi=pre_state.get("noi", 0),
        appraisal=pre_state.get("appraisal", 0),
        purchase_price=pre_state.get("purchase_price", 0),
        closing_costs=pre_state.get("closing_costs", 0),
        reserves=pre_state.get("reserves", 0),
        fees_pct=pre_state.get("fees", 0),
        rate=pre_state.get("rate", 0.0525),
        amort=pre_state.get("amort", 25),
        term=pre_state.get("term", 5),
        is_io=pre_state.get("is_io", False),
        target_ltv=pre_state.get("target_ltv", 0.75),
        target_ltc=pre_state.get("target_ltc", 0.80),
        target_dscr=pre_state.get("target_dscr", 1.25),
        target_dy=pre_state.get("target_dy", 0.085),
    )

    pre_mezz = safe_float(pre_state.get("mezz_debt", 0))
    pre_pref = safe_float(pre_state.get("pref_equity", 0))
    pre_sponsor_equity = pre_total_uses - pre_loan - pre_mezz - pre_pref
    pre_total_sources = pre_loan + pre_mezz + pre_pref + pre_sponsor_equity
    pre_balance_diff = pre_total_sources - pre_total_uses

    # ==========================================
    # SIDEBAR
    # ==========================================

    with st.sidebar:
        st.title("🏛️ ALENZA OS")
        st.caption(f"v{APP_VERSION}")

        if pre_errors:
            st.error(f"🚨 {len(pre_errors)} validation issue(s)")
        elif abs(pre_balance_diff) > 1:
            st.error(f"🚨 Unbalanced S&U: ${pre_balance_diff:,.0f}")
        elif pre_sponsor_equity < 0:
            st.warning(f"💸 Cash-Out: ${abs(pre_sponsor_equity):,.0f}")
        else:
            st.success("✅ Model balanced")

        sidebar_mode = st.radio(
            "Input Mode",
            ["📂 Deal Selection", "🏢 Asset Setup", "⚖️ Underwriting"],
            horizontal=False,
            key="sidebar_input_mode",
        )

        st.markdown("---")

        if sidebar_mode == "📂 Deal Selection":
            with st.expander("📁 DEAL MANAGER", expanded=True):
                new_name = st.text_input(
                    "New Deal Name",
                    value="Untitled Deal",
                    key="new_deal_name",
                )

                if st.button("➕ New Deal", use_container_width=True, key="btn_new_deal"):
                    fresh = default_state()
                    fresh["deal_id"] = generate_id("deal")
                    fresh["deal_name"] = new_name.strip() or "Untitled Deal"
                    fresh["unsaved_changes"] = True
                    fresh["last_saved_at"] = None
                    fresh["last_autosaved_at"] = None

                    for k, v in fresh.items():
                        s[k] = v

                    st.toast("New deal created", icon="🆕")
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

                    selected_label = st.selectbox(
                        "Existing Deal",
                        list(deal_options.keys()),
                        key="load_deal_select",
                    )

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
                                st.toast("Deal loaded", icon="📂")
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
                                copied["last_autosaved_at"] = None
                                copied["unsaved_changes"] = True

                                for k, v in copied.items():
                                    s[k] = v

                                st.toast("Deal duplicated", icon="📑")
                                st.rerun()
                            else:
                                st.error("Could not duplicate selected deal.")

                    confirm_delete = st.checkbox(
                        "Confirm delete selected deal",
                        key="confirm_delete_deal",
                    )

                    if st.button(
                        "🗑️ Delete Selected Deal",
                        use_container_width=True,
                        key="btn_delete",
                        disabled=not confirm_delete,
                    ):
                        deal_id = deal_options[selected_label]

                        if DatabaseManager.delete_deal(deal_id):
                            st.success("✅ Deleted")
                            st.toast("Deal deleted", icon="🗑️")
                            st.cache_data.clear()
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.error("Delete failed.")
                else:
                    st.info("No saved deals yet.")

                if st.button("💾 Save Deal", use_container_width=True, key="btn_save"):
                    if save_current_deal(reason="MANUAL_SAVE"):
                        st.success("✅ Deal saved!")

        elif sidebar_mode == "🏢 Asset Setup":
            with st.expander("🏢 ASSET PROFILE", expanded=True):
                s.deal_name = st.text_input(
                    "Deal Name",
                    value=s.get("deal_name", ""),
                    key="asset_deal_name",
                )

                s.sponsor = st.text_input(
                    "Sponsor",
                    value=s.get("sponsor", ""),
                    key="asset_sponsor",
                )

                s.property_address = st.text_input(
                    "Property Address",
                    value=s.get("property_address", ""),
                    key="asset_property_address",
                )

                current_pt = (
                    PROPERTY_TYPES.index(s.get("property_type", "Multifamily"))
                    if s.get("property_type") in PROPERTY_TYPES
                    else 0
                )

                s.property_type = st.selectbox(
                    "Property Type",
                    PROPERTY_TYPES,
                    index=current_pt,
                    key="asset_property_type",
                )

                current_tt = (
                    TRANSACTION_TYPES.index(s.get("transaction_type", "Acquisition"))
                    if s.get("transaction_type") in TRANSACTION_TYPES
                    else 0
                )

                s.transaction_type = st.selectbox(
                    "Transaction Type",
                    TRANSACTION_TYPES,
                    index=current_tt,
                    key="asset_transaction_type",
                )

                s.appraisal = st.number_input(
                    "Appraisal ($)",
                    value=safe_float(s.get("appraisal", 0)),
                    step=100000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_appraisal",
                )

                s.purchase_price = st.number_input(
                    "Cost Basis ($)",
                    value=safe_float(s.get("purchase_price", 0)),
                    step=100000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_purchase_price",
                )

                s.noi = st.number_input(
                    "Stabilized NOI ($)",
                    value=safe_float(s.get("noi", 0)),
                    step=10000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_noi",
                )

        else:
            with st.expander("📊 CREDIT POLICY", expanded=True):
                profiles = list(UnderwritingEngine.LENDER_LIMITS.keys())

                current_lp = (
                    profiles.index(s.get("lender_profile", "Bank / Credit Union"))
                    if s.get("lender_profile") in profiles
                    else 0
                )

                s.lender_profile = st.selectbox(
                    "Lender Profile",
                    profiles,
                    index=current_lp,
                    key="credit_lender_profile",
                )

                limits = UnderwritingEngine.LENDER_LIMITS[s.lender_profile]

                s.target_ltv = (
                    st.slider(
                        "Max LTV %",
                        50.0,
                        95.0,
                        float(normalize_percent(s.get("target_ltv", limits["max_ltv"]), limits["max_ltv"], 1.25) * 100),
                        step=0.5,
                        key="credit_target_ltv",
                    )
                    / 100
                )

                s.target_dscr = st.slider(
                    "Min DSCR",
                    1.0,
                    1.75,
                    float(safe_float(s.get("target_dscr", limits["min_dscr"]))),
                    step=0.05,
                    key="credit_target_dscr",
                )

                s.target_dy = (
                    st.slider(
                        "Min DY %",
                        5.0,
                        15.0,
                        float(normalize_percent(s.get("target_dy", limits["min_dy"]), limits["min_dy"], 0.25) * 100),
                        step=0.25,
                        key="credit_target_dy",
                    )
                    / 100
                )

                s.target_ltc = (
                    st.slider(
                        "Max LTC %",
                        50.0,
                        100.0,
                        float(normalize_percent(s.get("target_ltc", 0.80), 0.80, 1.25) * 100),
                        step=0.5,
                        key="credit_target_ltc",
                    )
                    / 100
                )

            with st.expander("💰 DEBT STRUCTURE", expanded=True):
                s.is_io = st.checkbox(
                    "Interest-Only Period",
                    value=bool(s.get("is_io", False)),
                    key="debt_is_io",
                )

                s.rate = (
                    st.slider(
                        "Interest Rate %",
                        0.0,
                        15.0,
                        float(normalize_percent(s.get("rate", 0.0525), 0.0525, 0.30) * 100),
                        step=0.05,
                        key="debt_rate",
                    )
                    / 100
                )

                s.amort = st.number_input(
                    "Amortization (Yrs)",
                    value=max(1, safe_int(s.get("amort", 25))),
                    step=1,
                    min_value=1,
                    max_value=40,
                    key="debt_amort",
                )

                s.term = st.number_input(
                    "Term (Yrs)",
                    value=max(1, safe_int(s.get("term", 5))),
                    step=1,
                    min_value=1,
                    max_value=40,
                    key="debt_term",
                )

                s.fees = (
                    st.slider(
                        "Financing Fees %",
                        0.0,
                        5.0,
                        float(normalize_percent(s.get("fees", 0.02), 0.02, 0.10) * 100),
                        step=0.05,
                        key="debt_fees",
                    )
                    / 100
                )

                s.closing_costs = st.number_input(
                    "Closing Costs ($)",
                    value=safe_float(s.get("closing_costs", 0)),
                    step=1000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="debt_closing_costs",
                )

                s.reserves = st.number_input(
                    "Required Reserves ($)",
                    value=safe_float(s.get("reserves", 0)),
                    step=1000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="debt_reserves",
                )

        st.markdown("---")

        s.autosave_enabled = st.checkbox(
            "Auto-save every 5 minutes",
            value=s.get("autosave_enabled", True),
            key="autosave_toggle_sidebar",
        )

        if s.get("autosave_status"):
            st.caption(f"🕐 {s.autosave_status}")

        if s.get("unsaved_changes"):
            st.warning("⚠️ Unsaved changes")

        if s.get("last_saved_at"):
            st.caption(f"Last saved: {str(s.last_saved_at)[:19]}")

    state_hash_after = stable_state_hash()

    if (
        state_hash_after != state_hash_before
        and not s.get("_loading_deal", False)
        and not s.get("unsaved_changes", False)
    ):
        s.unsaved_changes = True

    rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

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
        target_dy=s.target_dy,
    )

    amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(
        loan_amt,
        s.rate,
        s.amort,
        s.term,
        s.is_io,
    )

    annual_ds = monthly_pmt * 12

    actual_ltv = safe_ratio(loan_amt, s.appraisal)
    actual_ltc = safe_ratio(loan_amt, total_uses)
    actual_dscr = safe_ratio(s.noi, annual_ds)
    actual_dy = safe_ratio(s.noi, loan_amt)

    tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr_df)

    score, tier = UnderwritingEngine.score_deal(
        actual_ltv,
        actual_ltc,
        actual_dscr,
        actual_dy,
        s.lender_profile,
    )

    quick_pro_forma_df = InvestmentEngine.calculate_pro_forma(
        stabilized_noi=s.noi,
        revenue_growth=s.get("pf_revenue_growth", 0.03),
        expense_growth=s.get("pf_expense_growth", 0.035),
        expense_ratio=s.get("pf_expense_ratio", 0.40),
        years=s.get("pf_projection_years", 10),
    )

    quick_returns = InvestmentEngine.solve_returns(
        purchase_price=s.purchase_price,
        loan_amt=loan_amt,
        pro_forma_df=quick_pro_forma_df,
        exit_cap=s.get("pf_exit_cap", 0.065),
        selling_costs=s.get("pf_selling_costs", 0.015),
        annual_ds=annual_ds,
        balloon_balance=balloon,
        terminal_growth=s.get("pf_terminal_growth", 0.02),
    )

    errors, warnings = ValidationEngine.validate(extract_clean_state())

    autosaved_now = maybe_autosave_current_deal()

    if autosaved_now:
        try:
            st.toast("Autosaved", icon="✅")
        except Exception:
            pass

    mezz_source = safe_float(s.get("mezz_debt", 0))
    pref_source = safe_float(s.get("pref_equity", 0))
    sponsor_equity_source = total_uses - loan_amt - mezz_source - pref_source
    total_sources = loan_amt + mezz_source + pref_source + sponsor_equity_source
    balance_diff = total_sources - total_uses

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

    st.markdown("### Debt Capacity")

    debt_kpi_cols = st.columns(4)
    debt_kpi_cols[0].metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
    debt_kpi_cols[1].metric("CONSTRAINT", gate)
    debt_kpi_cols[2].metric("ACTUAL LTV", f"{actual_ltv * 100:.1f}%")
    debt_kpi_cols[3].metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")

    st.markdown("### Investment / Returns")

    equity_label = "CASH OUT" if sponsor_equity_source < 0 else "REQ. EQUITY"
    equity_value = abs(sponsor_equity_source)
    equity_delta = "Surplus" if sponsor_equity_source < 0 else f"{safe_ratio(sponsor_equity_source, total_uses):.1%} of uses" if total_uses > 0 else None

    investment_kpi_cols = st.columns(4)
    investment_kpi_cols[0].metric(equity_label, f"${equity_value:,.0f}", delta=equity_delta)
    investment_kpi_cols[1].metric("DEAL SCORE", f"{score}/1000", help=tier)
    investment_kpi_cols[2].metric("PROJECTED IRR", f"{quick_returns['IRR']:.2%}")
    investment_kpi_cols[3].metric("BALLOON", f"${balloon:,.0f}")

    if abs(balance_diff) > 1:
        st.error(f"🚨 UNBALANCED SOURCES & USES: Difference ${balance_diff:,.0f}")
    elif sponsor_equity_source < 0:
        st.warning(f"💸 Cash-out / surplus proceeds implied: ${abs(sponsor_equity_source):,.0f}")

    st.markdown("---")

    tabs = st.tabs(
        [
            "📊 Sizing & Risk",
            "🧪 Sensitivity",
            "📝 Rent Roll",
            "📅 Amortization",
            "📈 Pro Forma",
            "🇨🇦 Canada Intel",
            "📈 Market Comps",
            "📎 Diligence Room",
            "💾 Save & Export",
            "✅ QA & Health",
        ]
    )

    fragment = getattr(st, "fragment", None)

    def highlight_base_row(row):
        if row.get("Scenario") == "Base":
            return ["background-color: rgba(207, 184, 124, 0.35); font-weight: 700;"] * len(row)
        return [""] * len(row)

    def highlight_heatmap_base(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        if "NOI Shock" in data.columns:
            base_rows = data.index[data["NOI Shock"] == "+0%"].tolist()
            if base_rows and "Rate +0.0%" in data.columns:
                styles.loc[base_rows[0], "Rate +0.0%"] = "background-color: rgba(207, 184, 124, 0.45); font-weight: 700;"
        return styles

    # ==========================================
    # TAB 0: SIZING & RISK
    # ==========================================

    with tabs[0]:
        col_left, col_right = st.columns([1.5, 1])

        with col_left:
            st.subheader("📐 Constraint Analysis")

            constraints_df = pd.DataFrame(
                {
                    "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
                    "Threshold": [
                        f"{s.target_ltv * 100:.1f}%",
                        f"{s.target_ltc * 100:.1f}%",
                        f"{s.target_dscr:.2f}x",
                        f"{s.target_dy * 100:.2f}%",
                    ],
                    "Max Proceeds": [
                        f"${gates.get('LTV', 0):,.0f}",
                        f"${gates.get('LTC', 0):,.0f}",
                        f"${gates.get('DSCR', 0):,.0f}",
                        f"${gates.get('Debt Yield', 0):,.0f}",
                    ],
                    "Binding": [
                        "✅ ACTIVE" if gate == g else ""
                        for g in ["LTV", "LTC", "DSCR", "Debt Yield"]
                    ],
                }
            )

            st.dataframe(constraints_df, hide_index=True, use_container_width=True)

            st.subheader("💰 Sources & Uses")

            total_fees = loan_amt * s.fees

            su_df = pd.DataFrame(
                {
                    "Uses": [
                        "Cost Basis",
                        "Closing Costs",
                        "Reserves",
                        "Financing Fees",
                        "TOTAL USES",
                    ],
                    "Use Amount": [
                        s.purchase_price,
                        s.closing_costs,
                        s.reserves,
                        total_fees,
                        total_uses,
                    ],
                    "Sources": [
                        "Senior Debt",
                        "Mezzanine Debt",
                        "Preferred Equity",
                        "Sponsor Equity / Cash-Out",
                        "TOTAL SOURCES",
                    ],
                    "Source Amount": [
                        loan_amt,
                        mezz_source,
                        pref_source,
                        sponsor_equity_source,
                        total_sources,
                    ],
                }
            )

            st.dataframe(
                su_df.style.format(
                    {
                        "Use Amount": "${:,.0f}",
                        "Source Amount": "${:,.0f}",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )

            if abs(balance_diff) > 1:
                st.error(f"🚨 UNBALANCED SOURCES & USES: Difference ${balance_diff:,.0f}")
            else:
                st.success("✅ Sources & Uses balanced")

            if sponsor_equity_source < 0:
                st.warning(f"Cash-out / surplus proceeds implied: ${abs(sponsor_equity_source):,.0f}")

        with col_right:
            st.subheader("🔍 Risk Assessment")

            flags = []

            if actual_ltv > 0.75:
                flags.append(("high", f"⚠️ High Leverage: {actual_ltv * 100:.1f}% LTV"))
            elif actual_ltv < 0.55 and loan_amt > 0:
                flags.append(("low", f"✅ Conservative Leverage: {actual_ltv * 100:.1f}% LTV"))

            if actual_dscr < 1.20 and loan_amt > 0:
                flags.append(("high", f"⚠️ Tight Coverage: {actual_dscr:.2f}x DSCR"))
            elif actual_dscr > 1.50:
                flags.append(("low", f"✅ Strong Coverage: {actual_dscr:.2f}x DSCR"))

            if s.is_io:
                flags.append(("medium", "ℹ️ Interest-Only Structure"))

            if sponsor_equity_source < 0:
                flags.append(("medium", f"ℹ️ Cash-Out Structure: ${abs(sponsor_equity_source):,.0f} surplus proceeds"))

            if walt > 0 and walt < 3:
                flags.append(("high", f"⚠️ Short WALT: {walt:.1f} years"))

            if exp1 > 0.30:
                flags.append(("high", f"🚨 High Rollover: {exp1 * 100:.1f}%"))

            if not flags:
                flags.append(("low", "✅ No Significant Risk Flags"))

            for severity, message in flags:
                if severity == "high":
                    st.error(message)
                elif severity == "medium":
                    st.warning(message)
                else:
                    st.success(message)

            st.markdown("---")
            st.subheader("📊 Key Metrics")

            breakeven_occ = UnderwritingEngine.breakeven_occupancy(
                current_noi=s.noi,
                current_occupancy=occ,
                annual_debt_service=annual_ds,
            )

            st.metric(
                "Breakeven Occupancy",
                f"{breakeven_occ * 100:.1f}%",
                delta=f"Current: {occ * 100:.1f}%" if occ > 0 else None,
            )

            if sponsor_equity_source < 0:
                st.metric("Surplus Proceeds", f"${abs(sponsor_equity_source):,.0f}")
            else:
                st.metric("Required Equity", f"${sponsor_equity_source:,.0f}")

            st.metric(
                "Implied Cap Rate",
                f"{safe_ratio(s.noi, s.appraisal) * 100:.2f}%" if s.appraisal > 0 else "N/A",
            )

            st.markdown("---")
            st.subheader("🏗️ Capital Stack")

            s.mezz_debt = st.number_input(
                "Mezzanine Debt ($)",
                value=safe_float(s.get("mezz_debt", 0)),
                min_value=0.0,
                step=100000.0,
                format="%.0f",
                key="capital_mezz_debt",
            )

            s.pref_equity = st.number_input(
                "Preferred Equity ($)",
                value=safe_float(s.get("pref_equity", 0)),
                min_value=0.0,
                step=100000.0,
                format="%.0f",
                key="capital_pref_equity",
            )

            s.mezz_rate = (
                st.slider(
                    "Mezz Rate %",
                    0.0,
                    25.0,
                    float(normalize_percent(s.get("mezz_rate", 0.10), 0.10, 0.40) * 100),
                    step=0.25,
                    key="capital_mezz_rate",
                )
                / 100
            )

            s.pref_rate = (
                st.slider(
                    "Pref Rate %",
                    0.0,
                    25.0,
                    float(normalize_percent(s.get("pref_rate", 0.09), 0.09, 0.40) * 100),
                    step=0.25,
                    key="capital_pref_rate",
                )
                / 100
            )

            stack = UnderwritingEngine.capital_stack(
                senior_debt=loan_amt,
                mezz_debt=s.mezz_debt,
                pref_equity=s.pref_equity,
                sponsor_equity=max(0, sponsor_equity_source),
                noi=s.noi,
                senior_rate=s.rate,
                mezz_rate=s.mezz_rate,
                pref_rate=s.pref_rate,
            )

            st.metric("Fixed Charge Coverage", f"{stack['Fixed Charge Coverage']:.2f}x")

            st.caption(
                f"Total Capital: ${stack['Total Capital']:,.0f} | "
                f"Fixed Charges: ${stack['Fixed Charges']:,.0f}"
            )

    # ==========================================
    # TAB 1: SENSITIVITY
    # ==========================================

    with tabs[1]:
        st.subheader("🧪 Sensitivity Analysis")

        sensitivity_df = SensitivityEngine.generate_matrix(extract_clean_state())

        st.dataframe(
            sensitivity_df.style.apply(highlight_base_row, axis=1),
            hide_index=True,
            use_container_width=True,
        )

        st.markdown("---")
        st.subheader("🔥 Proceeds Heatmap")

        heatmap_df = SensitivityEngine.proceeds_heatmap(extract_clean_state())
        currency_cols = [c for c in heatmap_df.columns if c != "NOI Shock"]

        st.dataframe(
            heatmap_df.style.format({col: "${:,.0f}" for col in currency_cols}).apply(
                highlight_heatmap_base,
                axis=None,
            ),
            hide_index=True,
            use_container_width=True,
        )

        if DEPENDENCIES.get("plotly"):
            try:
                import plotly.express as px

                plot_df = heatmap_df.set_index("NOI Shock")[currency_cols] / 1_000_000

                fig = px.imshow(
                    plot_df,
                    text_auto=".1f",
                    aspect="auto",
                    title="Max Proceeds Sensitivity ($MM)",
                    labels={
                        "x": "Rate Shock",
                        "y": "NOI Shock",
                        "color": "Max Proceeds ($MM)",
                    },
                )

                fig.update_traces(
                    texttemplate="$%{z:.1f}MM",
                    hovertemplate="NOI Shock: %{y}<br>Rate Shock: %{x}<br>Max Proceeds: $%{z:.1f}MM<extra></extra>",
                )

                fig.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="#0B0F19",
                    plot_bgcolor="#0F172A",
                    height=450,
                    margin=dict(l=20, r=20, t=50, b=20),
                )

                st.plotly_chart(fig, use_container_width=True)

            except Exception as e:
                st.info(f"Heatmap chart unavailable: {e}")

        st.markdown("---")
        st.subheader("🎯 Custom Stress Test")

        c1, c2, c3 = st.columns(3)

        with c1:
            rate_shock = (
                st.slider(
                    "Rate Shock (bps)",
                    -200,
                    200,
                    0,
                    25,
                    key="custom_rate_shock",
                )
                / 10000
            )

        with c2:
            noi_shock = (
                st.slider(
                    "NOI Shock (%)",
                    -30,
                    30,
                    0,
                    5,
                    key="custom_noi_shock",
                )
                / 100
            )

        with c3:
            ltv_shock = (
                st.slider(
                    "LTV Adjustment (%)",
                    -10,
                    10,
                    0,
                    1,
                    key="custom_ltv_shock",
                )
                / 100
            )

        stressed_rate = max(0.001, s.rate + rate_shock)
        stressed_noi = max(0, s.noi * (1 + noi_shock))
        stressed_ltv = min(1.25, max(0.01, s.target_ltv + ltv_shock))

        stressed_loan, stressed_gate, _, _, _ = UnderwritingEngine.size_loan(
            noi=stressed_noi,
            appraisal=s.appraisal,
            purchase_price=s.purchase_price,
            closing_costs=s.closing_costs,
            reserves=s.reserves,
            fees_pct=s.fees,
            rate=stressed_rate,
            amort=s.amort,
            term=s.term,
            is_io=s.is_io,
            target_ltv=stressed_ltv,
            target_ltc=s.target_ltc,
            target_dscr=s.target_dscr,
            target_dy=s.target_dy,
        )

        _, stressed_monthly_pmt, _ = UnderwritingEngine.amort_schedule(
            stressed_loan,
            stressed_rate,
            s.amort,
            s.term,
            s.is_io,
        )

        stressed_annual_ds = stressed_monthly_pmt * 12
        stressed_dscr = safe_ratio(stressed_noi, stressed_annual_ds)

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Stressed Proceeds",
            f"${stressed_loan:,.0f}",
            delta=f"${stressed_loan - loan_amt:,.0f}" if loan_amt > 0 else None,
        )

        c2.metric("Constraint", stressed_gate)

        c3.metric(
            "Stressed LTV",
            f"{safe_ratio(stressed_loan, s.appraisal) * 100:.1f}%" if s.appraisal > 0 else "N/A",
        )

        c4.metric("Stressed DSCR", f"{stressed_dscr:.2f}x")

    # ==========================================
    # TAB 2: RENT ROLL
    # ==========================================

    with tabs[2]:
        def render_rent_roll_tab():
            st.subheader("📝 Rent Roll Management")

            uploaded_file = st.file_uploader(
                "Import Rent Roll (CSV or Excel)",
                type=["csv", "xlsx", "xls"],
                key="rent_roll_upload",
            )

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
                            st.toast(f"Imported {len(imported_df)} tenant records", icon="✅")
                            st.rerun()

                except Exception as e:
                    st.error(f"Import failed: {str(e)[:200]}")

            edit_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

            col_add, col_clear, col_save = st.columns([1, 1, 1])

            with col_add:
                if st.button("➕ Add Blank Row", use_container_width=True, key="add_rr_row"):
                    rr_records = edit_df.to_dict("records")
                    rr_records.append(
                        {
                            "Tenant": "",
                            "SF": 0,
                            "Remaining Term": 0,
                            "Monthly Rent": 0,
                        }
                    )
                    s.rent_roll_dict = rr_records
                    s.unsaved_changes = True
                    st.toast("Blank rent roll row added", icon="➕")
                    st.rerun()

            with col_clear:
                if st.button("🧹 Clear Rent Roll", use_container_width=True, key="clear_rr"):
                    s.rent_roll_dict = []
                    s.unsaved_changes = True
                    st.toast("Rent roll cleared", icon="🧹")
                    st.rerun()

            with col_save:
                if st.button("💾 Save Deal", use_container_width=True, key="rent_roll_fragment_save"):
                    if save_current_deal(reason="MANUAL_SAVE"):
                        st.success("✅ Deal saved")
                        st.toast("Deal saved", icon="✅")

            st.caption(
                "Edits sync automatically. Use Recalculate Now to refresh the full model immediately, "
                "or Save Deal to commit directly from this tab."
            )

            edited_df = st.data_editor(
                edit_df,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Tenant": st.column_config.TextColumn("Tenant Name", width="large"),
                    "SF": st.column_config.NumberColumn(
                        "Square Feet",
                        min_value=0,
                        step=100,
                        format="%d",
                    ),
                    "Remaining Term": st.column_config.NumberColumn(
                        "Lease Term (Yrs)",
                        min_value=0.0,
                        step=0.5,
                        format="%.1f",
                    ),
                    "Monthly Rent": st.column_config.NumberColumn(
                        "Monthly Rent ($)",
                        min_value=0.0,
                        step=100.0,
                        format="$%.2f",
                    ),
                },
                key="rent_roll_editor",
            )

            normalized_editor_df = normalize_rent_roll_columns(edited_df)
            existing_rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

            rent_roll_changed = not normalized_editor_df.equals(existing_rr_df)

            if rent_roll_changed:
                s.rent_roll_dict = normalized_editor_df.to_dict("records")
                s.unsaved_changes = True
                s.autosave_status = f"Rent roll edited at {datetime.now().strftime('%H:%M:%S')}"

                st.warning(
                    "Rent roll changes are staged. Click Recalculate Now to refresh the full model header immediately."
                )

            col_recalc, col_status = st.columns([1, 2])

            with col_recalc:
                if st.button("🟨 Recalculate Now", use_container_width=True, key="recalc_rent_roll_now"):
                    s.rent_roll_dict = normalized_editor_df.to_dict("records")
                    s.unsaved_changes = True
                    s.autosave_status = f"Rent roll synced at {datetime.now().strftime('%H:%M:%S')}"
                    st.toast("Rent roll synced", icon="✅")
                    st.rerun()

            with col_status:
                if rent_roll_changed:
                    st.caption("Status: staged changes detected. Full-page metrics refresh after Recalculate Now.")
                elif s.get("unsaved_changes"):
                    st.caption("Status: rent roll synced; deal has unsaved changes.")
                else:
                    st.caption("Status: rent roll synced and no unsaved changes detected.")

            if not normalized_editor_df.empty:
                st.markdown("---")
                st.subheader("📊 Rent Roll Analytics")

                met_total_sf, met_occ, met_ann_rent, met_psf, met_walt, met_exp1 = (
                    UnderwritingEngine.rent_roll_metrics(normalized_editor_df)
                )

                c1, c2, c3, c4, c5, c6 = st.columns(6)

                c1.metric("Total SF", f"{met_total_sf:,.0f}")
                c2.metric("Occupancy", f"{met_occ * 100:.1f}%")
                c3.metric("Annual Rent", f"${met_ann_rent:,.0f}")
                c4.metric("Rent PSF", f"${met_psf:.2f}")
                c5.metric("WALT (Yrs)", f"{met_walt:.2f}")
                c6.metric("12-Mo Rollover", f"{met_exp1 * 100:.1f}")

            else:
                st.info("No rent roll rows currently entered.")

        if fragment:
            fragment(render_rent_roll_tab)()
        else:
            render_rent_roll_tab()

    # ==========================================
    # TAB 3: AMORTIZATION
    # ==========================================

    with tabs[3]:
        st.subheader(f"📅 Amortization Schedule - {s.term} Year Term")

        if amort_df is None or amort_df.empty:
            st.warning("Enter deal parameters to generate amortization schedule.")
        else:
            c1, c2, c3 = st.columns(3)

            c1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
            c2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
            c3.metric(
                "Balloon Balance",
                f"${balloon:,.0f}",
                delta=f"{safe_ratio(balloon, loan_amt):.1%} of Original" if loan_amt > 0 else None,
                delta_color="normal",
            )

            st.markdown("---")

            st.write("### 📉 Term vs. Balloon Analysis")
            st.caption("Visualizing principal paydown velocity over various exit dates.")

            term_scenarios = [3, 5, 7, 10, 15]

            if safe_int(s.term) not in term_scenarios:
                term_scenarios.append(safe_int(s.term))

            term_scenarios = sorted(
                [
                    t
                    for t in term_scenarios
                    if t > 0 and t <= 40
                ]
            )

            scenario_data = []

            for t in term_scenarios:
                _, _, b_bal = UnderwritingEngine.amort_schedule(
                    loan_amt,
                    s.rate,
                    s.amort,
                    t,
                    s.is_io,
                )

                scenario_data.append(
                    {
                        "Term (Yrs)": t,
                        "Balloon Balance": b_bal,
                        "Paydown %": 1 - safe_ratio(b_bal, loan_amt) if loan_amt > 0 else 0,
                    }
                )

            sens_df = pd.DataFrame(scenario_data)

            if DEPENDENCIES.get("plotly") and not sens_df.empty:
                try:
                    import plotly.graph_objects as go
                    from plotly.subplots import make_subplots

                    fig_sens = make_subplots(specs=[[{"secondary_y": True}]])

                    fig_sens.add_trace(
                        go.Bar(
                            x=sens_df["Term (Yrs)"],
                            y=sens_df["Balloon Balance"],
                            name="Balloon Amount",
                            marker_color="#1E293B",
                            hovertemplate="Term: %{x}yr<br>Balloon: $%{y:,.0f}<extra></extra>",
                        ),
                        secondary_y=False,
                    )

                    fig_sens.add_trace(
                        go.Scatter(
                            x=sens_df["Term (Yrs)"],
                            y=sens_df["Paydown %"],
                            name="Principal Paydown %",
                            line=dict(color="#CFB87C", width=4),
                            mode="lines+markers",
                            hovertemplate="Term: %{x}yr<br>Paydown: %{y:.1%}<extra></extra>",
                        ),
                        secondary_y=True,
                    )

                    fig_sens.update_layout(
                        template="plotly_dark",
                        title="Balloon Exposure by Loan Term",
                        paper_bgcolor="#0B0F19",
                        plot_bgcolor="rgba(0,0,0,0)",
                        xaxis=dict(title="Loan Term (Years)", tickmode="linear"),
                        yaxis=dict(
                            title="Balloon Balance ($)",
                            gridcolor="rgba(255,255,255,0.05)",
                        ),
                        yaxis2=dict(
                            title="Paydown %",
                            tickformat=".0%",
                            showgrid=False,
                        ),
                        height=400,
                        margin=dict(l=20, r=20, t=50, b=20),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1,
                        ),
                    )

                    st.plotly_chart(fig_sens, use_container_width=True)

                except Exception as e:
                    st.info(f"Term vs. balloon chart unavailable: {e}")
                    st.dataframe(
                        sens_df.style.format(
                            {
                                "Balloon Balance": "${:,.0f}",
                                "Paydown %": "{:.1%}",
                            }
                        ),
                        hide_index=True,
                        use_container_width=True,
                    )
            else:
                st.dataframe(
                    sens_df.style.format(
                        {
                            "Balloon Balance": "${:,.0f}",
                            "Paydown %": "{:.1%}",
                        }
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

            amort_view = amort_df.copy()
            amort_view["Year"] = ((amort_view["Period"] - 1) // 12) + 1

            annual_summary = (
                amort_view.groupby("Year")
                .agg(
                    {
                        "Payment": "sum",
                        "Principal": "sum",
                        "Interest": "sum",
                        "Balance": "last",
                    }
                )
                .reset_index()
            )

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.write("#### Payment Structure")

                if DEPENDENCIES.get("plotly"):
                    try:
                        import plotly.graph_objects as go

                        pay_df = amort_df.copy()

                        fig_pay = go.Figure()

                        fig_pay.add_trace(
                            go.Bar(
                                x=pay_df["Period"],
                                y=pay_df["Principal"],
                                name="Principal",
                                marker_color="#CFB87C",
                            )
                        )

                        fig_pay.add_trace(
                            go.Bar(
                                x=pay_df["Period"],
                                y=pay_df["Interest"],
                                name="Interest",
                                marker_color="#1E293B",
                            )
                        )

                        fig_pay.update_layout(
                            barmode="stack",
                            template="plotly_dark",
                            paper_bgcolor="#0B0F19",
                            plot_bgcolor="rgba(0,0,0,0)",
                            height=360,
                            margin=dict(l=20, r=20, t=20, b=20),
                            xaxis_title="Period",
                            yaxis_title="Payment",
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1,
                            ),
                        )

                        st.plotly_chart(fig_pay, use_container_width=True)

                    except Exception:
                        st.bar_chart(
                            amort_df.set_index("Period")[["Principal", "Interest"]],
                            use_container_width=True,
                        )
                else:
                    st.bar_chart(
                        amort_df.set_index("Period")[["Principal", "Interest"]],
                        use_container_width=True,
                    )

            with col_chart2:
                st.write("#### Paydown Curve")

                if DEPENDENCIES.get("plotly"):
                    try:
                        import plotly.graph_objects as go

                        fig_bal = go.Figure()

                        fig_bal.add_trace(
                            go.Scatter(
                                x=amort_df["Period"],
                                y=amort_df["Balance"],
                                mode="lines",
                                name="Balance",
                                line=dict(color="#CFB87C", width=3),
                            )
                        )

                        fig_bal.update_layout(
                            template="plotly_dark",
                            paper_bgcolor="#0B0F19",
                            plot_bgcolor="rgba(0,0,0,0)",
                            height=360,
                            margin=dict(l=20, r=20, t=20, b=20),
                            xaxis_title="Period",
                            yaxis_title="Balance",
                        )

                        st.plotly_chart(fig_bal, use_container_width=True)

                    except Exception:
                        st.line_chart(
                            amort_df.set_index("Period")[["Balance"]],
                            use_container_width=True,
                        )
                else:
                    st.line_chart(
                        amort_df.set_index("Period")[["Balance"]],
                        use_container_width=True,
                    )

            with st.expander("📊 View Annual Summary & Full Schedule"):
                st.write("### Annual Summary")

                st.dataframe(
                    annual_summary.style.format(
                        {
                            "Payment": "${:,.2f}",
                            "Principal": "${:,.2f}",
                            "Interest": "${:,.2f}",
                            "Balance": "${:,.2f}",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.write("### Full Monthly Schedule")

                st.dataframe(
                    amort_view[["Period", "Payment", "Principal", "Interest", "Balance"]].style.format(
                        {
                            "Payment": "${:,.2f}",
                            "Principal": "${:,.2f}",
                            "Interest": "${:,.2f}",
                            "Balance": "${:,.2f}",
                        }
                    ),
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                )

    # ==========================================
    # TAB 4: PRO FORMA & RETURNS
    # ==========================================

    with tabs[4]:
        st.subheader("📈 10-Year Operating Pro Forma")
        st.caption("Revenue / expense growth, margin compression, exit valuation, and levered return profile.")

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            s.pf_revenue_growth = (
                st.slider(
                    "Revenue Growth %",
                    0.0,
                    8.0,
                    float(normalize_percent(s.get("pf_revenue_growth", 0.03), 0.03, 0.20) * 100),
                    0.25,
                    key="pf_revenue_growth_input",
                )
                / 100
            )

        with c2:
            s.pf_expense_growth = (
                st.slider(
                    "Expense Growth %",
                    0.0,
                    10.0,
                    float(normalize_percent(s.get("pf_expense_growth", 0.035), 0.035, 0.20) * 100),
                    0.25,
                    key="pf_expense_growth_input",
                )
                / 100
            )

        with c3:
            s.pf_expense_ratio = (
                st.slider(
                    "Initial Expense Ratio %",
                    10.0,
                    80.0,
                    float(normalize_percent(s.get("pf_expense_ratio", 0.40), 0.40, 0.95) * 100),
                    1.0,
                    key="pf_expense_ratio_input",
                )
                / 100
            )

        with c4:
            s.pf_projection_years = st.number_input(
                "Projection Years",
                min_value=5,
                max_value=30,
                value=max(5, min(30, safe_int(s.get("pf_projection_years", 10)))),
                step=1,
                key="pf_projection_years_input",
            )

        c5, c6, c7 = st.columns(3)

        with c5:
            implied_cap = (
                safe_ratio(s.noi, s.appraisal) * 100
                if safe_float(s.appraisal) > 0
                else 6.0
            )

            s.pf_exit_cap = (
                st.slider(
                    "Exit Cap Rate %",
                    3.0,
                    12.0,
                    float(min(max(safe_float(s.get("pf_exit_cap", (implied_cap + 0.50) / 100)) * 100, 3.0), 12.0)),
                    0.25,
                    key="pf_exit_cap_input",
                )
                / 100
            )

        with c6:
            s.pf_terminal_growth = (
                st.slider(
                    "Terminal NOI Growth %",
                    0.0,
                    5.0,
                    float(normalize_percent(s.get("pf_terminal_growth", 0.02), 0.02, 0.10) * 100),
                    0.25,
                    key="pf_terminal_growth_input",
                )
                / 100
            )

        with c7:
            s.pf_selling_costs = (
                st.slider(
                    "Selling Costs %",
                    0.0,
                    5.0,
                    float(normalize_percent(s.get("pf_selling_costs", 0.015), 0.015, 0.10) * 100),
                    0.25,
                    key="pf_selling_costs_input",
                )
                / 100
            )

        pro_forma_df = InvestmentEngine.calculate_pro_forma(
            stabilized_noi=s.noi,
            revenue_growth=s.pf_revenue_growth,
            expense_growth=s.pf_expense_growth,
            expense_ratio=s.pf_expense_ratio,
            years=s.pf_projection_years,
        )

        st.dataframe(
            pro_forma_df.style.format(
                {
                    "Revenue": "${:,.0f}",
                    "Expenses": "${:,.0f}",
                    "Projected NOI": "${:,.0f}",
                    "NOI Margin": "{:.2%}",
                    "Revenue Growth": "{:.2%}",
                    "Expense Growth": "{:.2%}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        if not pro_forma_df.empty:
            first_margin = safe_float(pro_forma_df["NOI Margin"].iloc[0])
            final_margin = safe_float(pro_forma_df["NOI Margin"].iloc[-1])
            margin_delta = final_margin - first_margin

            if margin_delta < -0.03:
                st.warning(
                    f"Margin compression detected: NOI margin declines {abs(margin_delta):.2%} over the projection."
                )
            elif margin_delta > 0.03:
                st.success(
                    f"Margin expansion detected: NOI margin increases {margin_delta:.2%} over the projection."
                )
            else:
                st.info("NOI margin remains broadly stable across the projection period.")

        st.markdown("---")
        st.subheader("💰 Levered Investment Performance")

        returns = InvestmentEngine.solve_returns(
            purchase_price=s.purchase_price,
            loan_amt=loan_amt,
            pro_forma_df=pro_forma_df,
            exit_cap=s.pf_exit_cap,
            selling_costs=s.pf_selling_costs,
            annual_ds=annual_ds,
            balloon_balance=balloon,
            terminal_growth=s.pf_terminal_growth,
        )

        r1, r2, r3, r4, r5 = st.columns(5)

        r1.metric("Levered IRR", f"{returns['IRR']:.2%}")
        r2.metric("Equity Multiple", f"{returns['Equity Multiple']:.2f}x")
        r3.metric("Exit NOI", f"${returns['Exit NOI']:,.0f}")
        r4.metric("Gross Exit Value", f"${returns['Gross Exit Value']:,.0f}")
        r5.metric("Net Exit Proceeds", f"${returns['Net Exit Proceeds']:,.0f}")

        st.metric(
            "Total Cash Flow",
            f"${returns['Total Cash Flow']:,.0f}",
        )

        st.markdown("---")
        st.write("### Annual Cash Flow After Debt Service")

        cf_data = pro_forma_df.copy()
        cf_data["Cash Flow After Debt Service"] = cf_data["Projected NOI"] - annual_ds

        if DEPENDENCIES.get("plotly"):
            try:
                import plotly.express as px

                fig_returns = px.bar(
                    cf_data,
                    x="Year",
                    y="Cash Flow After Debt Service",
                    title="Annual Cash Flow After Debt Service",
                    color_discrete_sequence=["#CFB87C"],
                )

                fig_returns.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="#0B0F19",
                    plot_bgcolor="#0F172A",
                    height=420,
                    margin=dict(l=20, r=20, t=50, b=20),
                )

                st.plotly_chart(fig_returns, use_container_width=True)

                sensitivity_rows = []

                for exit_cap_shock in [-0.005, 0, 0.005]:
                    for terminal_growth_shock in [-0.005, 0, 0.005]:
                        scenario_exit_cap = max(0.001, s.pf_exit_cap + exit_cap_shock)
                        scenario_terminal_growth = max(0, s.pf_terminal_growth + terminal_growth_shock)

                        scenario_returns = InvestmentEngine.solve_returns(
                            purchase_price=s.purchase_price,
                            loan_amt=loan_amt,
                            pro_forma_df=pro_forma_df,
                            exit_cap=scenario_exit_cap,
                            selling_costs=s.pf_selling_costs,
                            annual_ds=annual_ds,
                            balloon_balance=balloon,
                            terminal_growth=scenario_terminal_growth,
                        )

                        sensitivity_rows.append(
                            {
                                "Exit Cap": f"{scenario_exit_cap:.2%}",
                                "Terminal Growth": f"{scenario_terminal_growth:.2%}",
                                "Levered IRR": scenario_returns["IRR"],
                                "Equity Multiple": scenario_returns["Equity Multiple"],
                            }
                        )

                st.write("### Exit Cap / Terminal Growth Sensitivity")

                returns_sens_df = pd.DataFrame(sensitivity_rows)

                st.dataframe(
                    returns_sens_df.style.format(
                        {
                            "Levered IRR": "{:.2%}",
                            "Equity Multiple": "{:.2f}x",
                        }
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

            except Exception as e:
                st.info(f"Returns chart unavailable: {e}")

        else:
            st.bar_chart(
                cf_data.set_index("Year")[["Cash Flow After Debt Service"]],
                use_container_width=True,
            )
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
        "pf_revenue_growth": 0.03,
        "pf_expense_growth": 0.035,
        "pf_expense_ratio": 0.40,
        "pf_terminal_growth": 0.02,
        "pf_exit_cap": 0.065,
        "pf_selling_costs": 0.015,
        "pf_projection_years": 10,
        "rent_roll_dict": DEFAULT_RENT_ROLL.copy(),
        "diligence_notes": "",
        "last_saved_at": None,
        "last_autosaved_at": None,
        "autosave_enabled": True,
        "autosave_status": "",
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
    normalized["diligence_notes"] = str(normalized.get("diligence_notes") or "")

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
    normalized["pf_revenue_growth"] = normalize_percent(normalized.get("pf_revenue_growth"), 0.03, 0.20)
    normalized["pf_expense_growth"] = normalize_percent(normalized.get("pf_expense_growth"), 0.035, 0.20)
    normalized["pf_expense_ratio"] = normalize_percent(normalized.get("pf_expense_ratio"), 0.40, 0.95)
    normalized["pf_terminal_growth"] = normalize_percent(normalized.get("pf_terminal_growth"), 0.02, 0.10)
    normalized["pf_exit_cap"] = normalize_percent(normalized.get("pf_exit_cap"), 0.065, 0.20)
    normalized["pf_selling_costs"] = normalize_percent(normalized.get("pf_selling_costs"), 0.015, 0.10)
    normalized["pf_projection_years"] = max(5, min(30, safe_int(normalized.get("pf_projection_years"), 10)))

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

    for key in [
        "last_saved_at",
        "last_autosaved_at",
        "autosave_status",
        "unsaved_changes",
        "app_version",
        "schema_version",
    ]:
        state.pop(key, None)

    return hash_state(state)


# ==========================================
# SENSITIVITY ENGINE
# ==========================================

class SensitivityEngine:
    @staticmethod
    def generate_matrix(state: dict) -> pd.DataFrame:
        scenarios = [
            ("Base", 0, 0),
            ("Rate +1%", 0.01, 0),
            ("Rate -1%", -0.01, 0),
            ("NOI -10%", 0, -0.10),
            ("NOI +10%", 0, 0.10),
            ("Combined Stress", 0.01, -0.10),
        ]

        results = []
        base_proceeds = None

        for scenario_name, rate_adj, noi_adj in scenarios:
            adjusted_rate = max(0.001, safe_float(state.get("rate", 0.0525)) + rate_adj)
            adjusted_noi = max(0, safe_float(state.get("noi", 0)) * (1 + noi_adj))

            loan, gate, _, _, _ = UnderwritingEngine.size_loan(
                noi=adjusted_noi,
                appraisal=safe_float(state.get("appraisal", 0)),
                purchase_price=safe_float(state.get("purchase_price", 0)),
                closing_costs=safe_float(state.get("closing_costs", 0)),
                reserves=safe_float(state.get("reserves", 0)),
                fees_pct=safe_float(state.get("fees", 0)),
                rate=adjusted_rate,
                amort=safe_int(state.get("amort", 25)),
                term=safe_int(state.get("term", 5)),
                is_io=bool(state.get("is_io", False)),
                target_ltv=safe_float(state.get("target_ltv", 0.75)),
                target_ltc=safe_float(state.get("target_ltc", 0.80)),
                target_dscr=safe_float(state.get("target_dscr", 1.25)),
                target_dy=safe_float(state.get("target_dy", 0.085)),
            )

            if scenario_name == "Base":
                base_proceeds = loan

            if base_proceeds and base_proceeds > 0:
                change_str = f"{((loan - base_proceeds) / base_proceeds * 100):+.1f}%"
            else:
                change_str = "N/A"

            results.append(
                {
                    "Scenario": scenario_name,
                    "Rate": f"{adjusted_rate * 100:.2f}%",
                    "NOI": f"${adjusted_noi:,.0f}",
                    "Max Proceeds": f"${loan:,.0f}",
                    "Constraint": gate,
                    "Change from Base": change_str,
                }
            )

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
                    noi=adjusted_noi,
                    appraisal=safe_float(state.get("appraisal", 0)),
                    purchase_price=safe_float(state.get("purchase_price", 0)),
                    closing_costs=safe_float(state.get("closing_costs", 0)),
                    reserves=safe_float(state.get("reserves", 0)),
                    fees_pct=safe_float(state.get("fees", 0)),
                    rate=adjusted_rate,
                    amort=safe_int(state.get("amort", 25)),
                    term=safe_int(state.get("term", 5)),
                    is_io=bool(state.get("is_io", False)),
                    target_ltv=safe_float(state.get("target_ltv", 0.75)),
                    target_ltc=safe_float(state.get("target_ltc", 0.80)),
                    target_dscr=safe_float(state.get("target_dscr", 1.25)),
                    target_dy=safe_float(state.get("target_dy", 0.085)),
                )

                row[f"Rate {rate_adj:+.1%}"] = loan

            rows.append(row)

        return pd.DataFrame(rows)


# ==========================================
# INVESTMENT & PRO FORMA ENGINE
# ==========================================

class InvestmentEngine:
    @staticmethod
    def calculate_pro_forma(
        stabilized_noi: float,
        revenue_growth: float,
        expense_growth: float,
        expense_ratio: float = 0.40,
        years: int = 10,
    ) -> pd.DataFrame:
        rows = []

        stabilized_noi = max(0, safe_float(stabilized_noi))
        revenue_growth = normalize_percent(revenue_growth, 0.03, 0.20)
        expense_growth = normalize_percent(expense_growth, 0.035, 0.20)
        expense_ratio = normalize_percent(expense_ratio, 0.40, 0.95)
        years = max(1, min(30, safe_int(years, 10)))

        if stabilized_noi <= 0:
            revenue = 0.0
            expenses = 0.0
        else:
            revenue = stabilized_noi / max(1 - expense_ratio, 0.01)
            expenses = revenue * expense_ratio

        for year in range(1, years + 1):
            noi = revenue - expenses
            margin = safe_ratio(noi, revenue)

            rows.append(
                {
                    "Year": year,
                    "Revenue": revenue,
                    "Expenses": expenses,
                    "Projected NOI": noi,
                    "NOI Margin": margin,
                    "Revenue Growth": 0.0 if year == 1 else revenue_growth,
                    "Expense Growth": 0.0 if year == 1 else expense_growth,
                }
            )

            revenue *= 1 + revenue_growth
            expenses *= 1 + expense_growth

        return pd.DataFrame(rows)

    @staticmethod
    def solve_returns(
        purchase_price: float,
        loan_amt: float,
        pro_forma_df: pd.DataFrame,
        exit_cap: float,
        selling_costs: float,
        annual_ds: float,
        balloon_balance: float,
        terminal_growth: float = 0.02,
    ) -> Dict[str, Any]:
        purchase_price = max(0, safe_float(purchase_price))
        loan_amt = max(0, safe_float(loan_amt))
        exit_cap = normalize_percent(exit_cap, 0.06, 0.20)
        selling_costs = normalize_percent(selling_costs, 0.015, 0.10)
        annual_ds = max(0, safe_float(annual_ds))
        balloon_balance = max(0, safe_float(balloon_balance))
        terminal_growth = normalize_percent(terminal_growth, 0.02, 0.10)

        equity_in = max(0, purchase_price - loan_amt)

        if equity_in <= 0 or pro_forma_df is None or pro_forma_df.empty:
            return {
                "IRR": 0.0,
                "Equity Multiple": 0.0,
                "Net Exit Proceeds": 0.0,
                "Gross Exit Value": 0.0,
                "Total Cash Flow": 0.0,
                "Exit NOI": 0.0,
            }

        cash_flows = [-equity_in]

        for _, row in pro_forma_df.iterrows():
            cash_flows.append(max(0, safe_float(row.get("Projected NOI"))) - annual_ds)

        exit_noi = safe_float(pro_forma_df.iloc[-1]["Projected NOI"]) * (1 + terminal_growth)
        gross_exit_value = exit_noi / exit_cap if exit_cap > 0 else 0
        net_exit_value = gross_exit_value * (1 - selling_costs)
        net_exit_proceeds = net_exit_value - balloon_balance

        cash_flows[-1] += net_exit_proceeds

        try:
            if DEPENDENCIES.get("numpy_financial"):
                import numpy_financial as npf
                irr = npf.irr(cash_flows)
            else:
                irr = 0.0

            if irr is None or np.isnan(irr):
                irr = 0.0

        except Exception:
            irr = 0.0

        total_distributions = sum(cash_flows[1:])
        equity_multiple = total_distributions / equity_in if equity_in > 0 else 0.0

        return {
            "IRR": irr,
            "Equity Multiple": equity_multiple,
            "Net Exit Proceeds": net_exit_proceeds,
            "Gross Exit Value": gross_exit_value,
            "Total Cash Flow": total_distributions,
            "Exit NOI": exit_noi,
        }


# ==========================================
# VALIDATION ENGINE
# ==========================================

class ValidationEngine:
    @staticmethod
    def validate(state: dict) -> tuple:
        errors = []
        warnings = []

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

            duplicates = sorted(
                {
                    t
                    for t in tenant_names
                    if tenant_names.count(t) > 1
                    and t not in ["vacant", "empty", "available", "vacancy"]
                }
            )

            if duplicates:
                warnings.append(f"Duplicate tenant names detected: {', '.join(duplicates[:5])}")

        return errors, warnings


# ==========================================
# FINANCIAL SELF TESTS
# ==========================================

def run_financial_self_tests() -> pd.DataFrame:
    rows = []

    def check(name: str, passed: bool, detail: str):
        rows.append(
            {
                "Test": name,
                "Status": "✅ PASS" if passed else "❌ FAIL",
                "Detail": detail,
            }
        )

    loan, gate, gates, uses, equity = UnderwritingEngine.size_loan(
        1_000_000,
        10_000_000,
        9_000_000,
        100_000,
        0,
        0.01,
        0.06,
        25,
        5,
        False,
        0.70,
        0.80,
        1.25,
        0.09,
    )

    check("Loan proceeds positive", loan > 0, f"${loan:,.0f}, Gate: {gate}")
    check("Binding gate equals minimum proceeds", abs(loan - min(gates.values())) < 1, f"Min: ${min(gates.values()):,.0f}")
    check("Total uses positive", uses > 0, f"${uses:,.0f}")
    check("Equity requirement calculated", isinstance(equity, float), f"${equity:,.0f}")

    amort, pmt, balloon_test = UnderwritingEngine.amort_schedule(
        1_000_000,
        0.06,
        25,
        5,
        False,
    )

    check("60 periods in 5-year term", len(amort) == 60, f"Periods: {len(amort)}")
    check("Monthly payment positive", pmt > 0, f"${pmt:,.2f}")
    check("Amortizing balloon below original", balloon_test < 1_000_000, f"Balloon: ${balloon_test:,.0f}")

    io_amort, io_pmt, io_balloon = UnderwritingEngine.amort_schedule(
        1_000_000,
        0.06,
        25,
        5,
        True,
    )

    check("IO payment equals interest-only amount", abs(io_pmt - 5_000) < 1, f"PMT: ${io_pmt:,.2f}")
    check("IO balloon equals original loan", abs(io_balloon - 1_000_000) < 1, f"Balloon: ${io_balloon:,.0f}")

    rr = pd.DataFrame(
        [
            {"Tenant": "A", "SF": 10_000, "Remaining Term": 5, "Monthly Rent": 20_000},
            {"Tenant": "Vacant", "SF": 2_000, "Remaining Term": 0, "Monthly Rent": 0},
        ]
    )

    sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr)

    check("Occupancy equals 83.3%", abs(occ - 0.8333) < 0.01, f"{occ:.2%}")
    check("Annual rent equals $240K", abs(ann_rent - 240_000) < 1, f"${ann_rent:,.0f}")
    check("Rent PSF equals $24", abs(psf - 24) < 0.1, f"${psf:.2f}")
    check("WALT equals 5.0 years", abs(walt - 5.0) < 0.01, f"{walt:.2f}")

    zero_loan, zero_gate, zero_gates, zero_uses, zero_equity = UnderwritingEngine.size_loan(
        0,
        10_000_000,
        9_000_000,
        0,
        0,
        0.01,
        0.06,
        25,
        5,
        False,
        0.70,
        0.80,
        1.25,
        0.09,
    )

    check(
        "Zero NOI produces zero income-based proceeds",
        zero_gates["DSCR"] == 0 and zero_gates["Debt Yield"] == 0 and zero_loan == 0,
        f"Loan: ${zero_loan:,.0f}",
    )

    empty_sf, empty_occ, empty_rent, empty_psf, empty_walt, empty_exp = UnderwritingEngine.rent_roll_metrics(
        pd.DataFrame()
    )

    check(
        "Empty rent roll returns zero metrics",
        empty_sf == 0 and empty_occ == 0 and empty_rent == 0,
        "Empty rent roll handled safely",
    )

    neg_loan, neg_gate, neg_gates, neg_uses, neg_equity = UnderwritingEngine.size_loan(
        -1_000_000,
        -10_000_000,
        -9_000_000,
        -100_000,
        -50_000,
        -0.01,
        -0.05,
        -25,
        -5,
        False,
        -0.70,
        -0.80,
        -1.25,
        -0.09,
    )

    check(
        "Negative inputs are sanitized",
        neg_loan >= 0 and neg_uses >= 0,
        f"Loan: ${neg_loan:,.0f}, Uses: ${neg_uses:,.0f}",
    )

    be_occ = UnderwritingEngine.breakeven_occupancy(
        current_noi=1_000_000,
        current_occupancy=0.90,
        annual_debt_service=750_000,
    )

    check("Breakeven occupancy calculation", abs(be_occ - 0.675) < 0.001, f"{be_occ:.2%}")

    stack = UnderwritingEngine.capital_stack(
        senior_debt=7_000_000,
        mezz_debt=1_000_000,
        pref_equity=500_000,
        sponsor_equity=1_500_000,
        noi=1_000_000,
        senior_rate=0.06,
        mezz_rate=0.11,
        pref_rate=0.09,
    )

    check(
        "Capital stack fixed charge coverage positive",
        stack["Fixed Charge Coverage"] > 0,
        f"{stack['Fixed Charge Coverage']:.2f}x",
    )

    pro_forma_df = InvestmentEngine.calculate_pro_forma(
        stabilized_noi=1_000_000,
        revenue_growth=0.03,
        expense_growth=0.035,
        expense_ratio=0.40,
        years=10,
    )

    check(
        "Pro forma produces 10 years",
        len(pro_forma_df) == 10,
        f"Rows: {len(pro_forma_df)}",
    )

    returns = InvestmentEngine.solve_returns(
        purchase_price=10_000_000,
        loan_amt=6_000_000,
        pro_forma_df=pro_forma_df,
        exit_cap=0.06,
        selling_costs=0.015,
        annual_ds=400_000,
        balloon_balance=5_500_000,
        terminal_growth=0.02,
    )

    check(
        "Returns engine produces equity multiple",
        returns["Equity Multiple"] >= 0,
        f"{returns['Equity Multiple']:.2f}x",
    )

    return pd.DataFrame(rows)


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

            try:
                st.toast(f"Autosaved at {now[:19]}", icon="✅")
            except Exception:
                pass
        else:
            st.session_state.autosave_status = f"Saved at {now[:19]}"

            try:
                st.toast("Deal saved", icon="✅")
            except Exception:
                pass

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


def try_autorefresh(interval_ms: int, key: str):
    try:
        from streamlit_autorefresh import st_autorefresh
        return st_autorefresh(interval=interval_ms, key=key)
    except Exception:
        return None


DatabaseManager.init_db()


# ==========================================
# MAIN APPLICATION
# ==========================================

def main():
    """Main application entry point."""

    DatabaseManager.init_db()
    initialize_session_state()

    s = st.session_state

    inject_responsive_css()

    components.html(
        """
        <script>
        document.addEventListener("keydown", function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
                e.preventDefault();
            }
        });
        </script>
        """,
        height=0,
    )

    if "_loading_deal" not in s:
        s["_loading_deal"] = False

    if "autosave_enabled" not in s:
        s["autosave_enabled"] = True

    if s.get("autosave_enabled", True):
        try_autorefresh(
            interval_ms=AUTOSAVE_INTERVAL_SECONDS * 1000,
            key="autosave_timer",
        )

    state_hash_before = stable_state_hash()

    pre_state = extract_clean_state()
    pre_errors, pre_warnings = ValidationEngine.validate(pre_state)

    pre_loan, pre_gate, pre_gates, pre_total_uses, pre_req_equity = UnderwritingEngine.size_loan(
        noi=pre_state.get("noi", 0),
        appraisal=pre_state.get("appraisal", 0),
        purchase_price=pre_state.get("purchase_price", 0),
        closing_costs=pre_state.get("closing_costs", 0),
        reserves=pre_state.get("reserves", 0),
        fees_pct=pre_state.get("fees", 0),
        rate=pre_state.get("rate", 0.0525),
        amort=pre_state.get("amort", 25),
        term=pre_state.get("term", 5),
        is_io=pre_state.get("is_io", False),
        target_ltv=pre_state.get("target_ltv", 0.75),
        target_ltc=pre_state.get("target_ltc", 0.80),
        target_dscr=pre_state.get("target_dscr", 1.25),
        target_dy=pre_state.get("target_dy", 0.085),
    )

    pre_mezz = safe_float(pre_state.get("mezz_debt", 0))
    pre_pref = safe_float(pre_state.get("pref_equity", 0))
    pre_sponsor_equity = pre_total_uses - pre_loan - pre_mezz - pre_pref
    pre_total_sources = pre_loan + pre_mezz + pre_pref + pre_sponsor_equity
    pre_balance_diff = pre_total_sources - pre_total_uses

    # ==========================================
    # SIDEBAR
    # ==========================================

    with st.sidebar:
        st.title("🏛️ ALENZA OS")
        st.caption(f"v{APP_VERSION}")

        if pre_errors:
            st.error(f"🚨 {len(pre_errors)} validation issue(s)")
        elif abs(pre_balance_diff) > 1:
            st.error(f"🚨 Unbalanced S&U: ${pre_balance_diff:,.0f}")
        elif pre_sponsor_equity < 0:
            st.warning(f"💸 Cash-Out: ${abs(pre_sponsor_equity):,.0f}")
        else:
            st.success("✅ Model balanced")

        sidebar_mode = st.radio(
            "Input Mode",
            ["📂 Deal Selection", "🏢 Asset Setup", "⚖️ Underwriting"],
            horizontal=False,
            key="sidebar_input_mode",
        )

        st.markdown("---")

        if sidebar_mode == "📂 Deal Selection":
            with st.expander("📁 DEAL MANAGER", expanded=True):
                new_name = st.text_input(
                    "New Deal Name",
                    value="Untitled Deal",
                    key="new_deal_name",
                )

                if st.button("➕ New Deal", use_container_width=True, key="btn_new_deal"):
                    fresh = default_state()
                    fresh["deal_id"] = generate_id("deal")
                    fresh["deal_name"] = new_name.strip() or "Untitled Deal"
                    fresh["unsaved_changes"] = True
                    fresh["last_saved_at"] = None
                    fresh["last_autosaved_at"] = None

                    for k, v in fresh.items():
                        s[k] = v

                    st.toast("New deal created", icon="🆕")
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

                    selected_label = st.selectbox(
                        "Existing Deal",
                        list(deal_options.keys()),
                        key="load_deal_select",
                    )

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
                                st.toast("Deal loaded", icon="📂")
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
                                copied["last_autosaved_at"] = None
                                copied["unsaved_changes"] = True

                                for k, v in copied.items():
                                    s[k] = v

                                st.toast("Deal duplicated", icon="📑")
                                st.rerun()
                            else:
                                st.error("Could not duplicate selected deal.")

                    confirm_delete = st.checkbox(
                        "Confirm delete selected deal",
                        key="confirm_delete_deal",
                    )

                    if st.button(
                        "🗑️ Delete Selected Deal",
                        use_container_width=True,
                        key="btn_delete",
                        disabled=not confirm_delete,
                    ):
                        deal_id = deal_options[selected_label]

                        if DatabaseManager.delete_deal(deal_id):
                            st.success("✅ Deleted")
                            st.toast("Deal deleted", icon="🗑️")
                            st.cache_data.clear()
                            time.sleep(0.5)
                            st.rerun()
                        else:
                            st.error("Delete failed.")
                else:
                    st.info("No saved deals yet.")

                if st.button("💾 Save Deal", use_container_width=True, key="btn_save"):
                    if save_current_deal(reason="MANUAL_SAVE"):
                        st.success("✅ Deal saved!")

        elif sidebar_mode == "🏢 Asset Setup":
            with st.expander("🏢 ASSET PROFILE", expanded=True):
                s.deal_name = st.text_input(
                    "Deal Name",
                    value=s.get("deal_name", ""),
                    key="asset_deal_name",
                )

                s.sponsor = st.text_input(
                    "Sponsor",
                    value=s.get("sponsor", ""),
                    key="asset_sponsor",
                )

                s.property_address = st.text_input(
                    "Property Address",
                    value=s.get("property_address", ""),
                    key="asset_property_address",
                )

                current_pt = (
                    PROPERTY_TYPES.index(s.get("property_type", "Multifamily"))
                    if s.get("property_type") in PROPERTY_TYPES
                    else 0
                )

                s.property_type = st.selectbox(
                    "Property Type",
                    PROPERTY_TYPES,
                    index=current_pt,
                    key="asset_property_type",
                )

                current_tt = (
                    TRANSACTION_TYPES.index(s.get("transaction_type", "Acquisition"))
                    if s.get("transaction_type") in TRANSACTION_TYPES
                    else 0
                )

                s.transaction_type = st.selectbox(
                    "Transaction Type",
                    TRANSACTION_TYPES,
                    index=current_tt,
                    key="asset_transaction_type",
                )

                s.appraisal = st.number_input(
                    "Appraisal ($)",
                    value=safe_float(s.get("appraisal", 0)),
                    step=100000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_appraisal",
                )

                s.purchase_price = st.number_input(
                    "Cost Basis ($)",
                    value=safe_float(s.get("purchase_price", 0)),
                    step=100000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_purchase_price",
                )

                s.noi = st.number_input(
                    "Stabilized NOI ($)",
                    value=safe_float(s.get("noi", 0)),
                    step=10000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="asset_noi",
                )

        else:
            with st.expander("📊 CREDIT POLICY", expanded=True):
                profiles = list(UnderwritingEngine.LENDER_LIMITS.keys())

                current_lp = (
                    profiles.index(s.get("lender_profile", "Bank / Credit Union"))
                    if s.get("lender_profile") in profiles
                    else 0
                )

                s.lender_profile = st.selectbox(
                    "Lender Profile",
                    profiles,
                    index=current_lp,
                    key="credit_lender_profile",
                )

                limits = UnderwritingEngine.LENDER_LIMITS[s.lender_profile]

                s.target_ltv = (
                    st.slider(
                        "Max LTV %",
                        50.0,
                        95.0,
                        float(normalize_percent(s.get("target_ltv", limits["max_ltv"]), limits["max_ltv"], 1.25) * 100),
                        step=0.5,
                        key="credit_target_ltv",
                    )
                    / 100
                )

                s.target_dscr = st.slider(
                    "Min DSCR",
                    1.0,
                    1.75,
                    float(safe_float(s.get("target_dscr", limits["min_dscr"]))),
                    step=0.05,
                    key="credit_target_dscr",
                )

                s.target_dy = (
                    st.slider(
                        "Min DY %",
                        5.0,
                        15.0,
                        float(normalize_percent(s.get("target_dy", limits["min_dy"]), limits["min_dy"], 0.25) * 100),
                        step=0.25,
                        key="credit_target_dy",
                    )
                    / 100
                )

                s.target_ltc = (
                    st.slider(
                        "Max LTC %",
                        50.0,
                        100.0,
                        float(normalize_percent(s.get("target_ltc", 0.80), 0.80, 1.25) * 100),
                        step=0.5,
                        key="credit_target_ltc",
                    )
                    / 100
                )

            with st.expander("💰 DEBT STRUCTURE", expanded=True):
                s.is_io = st.checkbox(
                    "Interest-Only Period",
                    value=bool(s.get("is_io", False)),
                    key="debt_is_io",
                )

                s.rate = (
                    st.slider(
                        "Interest Rate %",
                        0.0,
                        15.0,
                        float(normalize_percent(s.get("rate", 0.0525), 0.0525, 0.30) * 100),
                        step=0.05,
                        key="debt_rate",
                    )
                    / 100
                )

                s.amort = st.number_input(
                    "Amortization (Yrs)",
                    value=max(1, safe_int(s.get("amort", 25))),
                    step=1,
                    min_value=1,
                    max_value=40,
                    key="debt_amort",
                )

                s.term = st.number_input(
                    "Term (Yrs)",
                    value=max(1, safe_int(s.get("term", 5))),
                    step=1,
                    min_value=1,
                    max_value=40,
                    key="debt_term",
                )

                s.fees = (
                    st.slider(
                        "Financing Fees %",
                        0.0,
                        5.0,
                        float(normalize_percent(s.get("fees", 0.02), 0.02, 0.10) * 100),
                        step=0.05,
                        key="debt_fees",
                    )
                    / 100
                )

                s.closing_costs = st.number_input(
                    "Closing Costs ($)",
                    value=safe_float(s.get("closing_costs", 0)),
                    step=1000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="debt_closing_costs",
                )

                s.reserves = st.number_input(
                    "Required Reserves ($)",
                    value=safe_float(s.get("reserves", 0)),
                    step=1000.0,
                    min_value=0.0,
                    format="%.0f",
                    key="debt_reserves",
                )

        st.markdown("---")

        s.autosave_enabled = st.checkbox(
            "Auto-save every 5 minutes",
            value=s.get("autosave_enabled", True),
            key="autosave_toggle_sidebar",
        )

        if s.get("autosave_status"):
            st.caption(f"🕐 {s.autosave_status}")

        if s.get("unsaved_changes"):
            st.warning("⚠️ Unsaved changes")

        if s.get("last_saved_at"):
            st.caption(f"Last saved: {str(s.last_saved_at)[:19]}")

    state_hash_after = stable_state_hash()

    if (
        state_hash_after != state_hash_before
        and not s.get("_loading_deal", False)
        and not s.get("unsaved_changes", False)
    ):
        s.unsaved_changes = True

    rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

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
        target_dy=s.target_dy,
    )

    amort_df, monthly_pmt, balloon = UnderwritingEngine.amort_schedule(
        loan_amt,
        s.rate,
        s.amort,
        s.term,
        s.is_io,
    )

    annual_ds = monthly_pmt * 12

    actual_ltv = safe_ratio(loan_amt, s.appraisal)
    actual_ltc = safe_ratio(loan_amt, total_uses)
    actual_dscr = safe_ratio(s.noi, annual_ds)
    actual_dy = safe_ratio(s.noi, loan_amt)

    tot_sf, occ, ann_rent, psf, walt, exp1 = UnderwritingEngine.rent_roll_metrics(rr_df)

    score, tier = UnderwritingEngine.score_deal(
        actual_ltv,
        actual_ltc,
        actual_dscr,
        actual_dy,
        s.lender_profile,
    )

    quick_pro_forma_df = InvestmentEngine.calculate_pro_forma(
        stabilized_noi=s.noi,
        revenue_growth=s.get("pf_revenue_growth", 0.03),
        expense_growth=s.get("pf_expense_growth", 0.035),
        expense_ratio=s.get("pf_expense_ratio", 0.40),
        years=s.get("pf_projection_years", 10),
    )

    quick_returns = InvestmentEngine.solve_returns(
        purchase_price=s.purchase_price,
        loan_amt=loan_amt,
        pro_forma_df=quick_pro_forma_df,
        exit_cap=s.get("pf_exit_cap", 0.065),
        selling_costs=s.get("pf_selling_costs", 0.015),
        annual_ds=annual_ds,
        balloon_balance=balloon,
        terminal_growth=s.get("pf_terminal_growth", 0.02),
    )

    errors, warnings = ValidationEngine.validate(extract_clean_state())

    autosaved_now = maybe_autosave_current_deal()

    if autosaved_now:
        try:
            st.toast("Autosaved", icon="✅")
        except Exception:
            pass

    mezz_source = safe_float(s.get("mezz_debt", 0))
    pref_source = safe_float(s.get("pref_equity", 0))
    sponsor_equity_source = total_uses - loan_amt - mezz_source - pref_source
    total_sources = loan_amt + mezz_source + pref_source + sponsor_equity_source
    balance_diff = total_sources - total_uses

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

    st.markdown("### Debt Capacity")

    debt_kpi_cols = st.columns(4)
    debt_kpi_cols[0].metric("MAX PROCEEDS", f"${loan_amt:,.0f}")
    debt_kpi_cols[1].metric("CONSTRAINT", gate)
    debt_kpi_cols[2].metric("ACTUAL LTV", f"{actual_ltv * 100:.1f}%")
    debt_kpi_cols[3].metric("ACTUAL DSCR", f"{actual_dscr:.2f}x")

    st.markdown("### Investment / Returns")

    equity_label = "CASH OUT" if sponsor_equity_source < 0 else "REQ. EQUITY"
    equity_value = abs(sponsor_equity_source)
    equity_delta = "Surplus" if sponsor_equity_source < 0 else f"{safe_ratio(sponsor_equity_source, total_uses):.1%} of uses" if total_uses > 0 else None

    investment_kpi_cols = st.columns(4)
    investment_kpi_cols[0].metric(equity_label, f"${equity_value:,.0f}", delta=equity_delta)
    investment_kpi_cols[1].metric("DEAL SCORE", f"{score}/1000", help=tier)
    investment_kpi_cols[2].metric("PROJECTED IRR", f"{quick_returns['IRR']:.2%}")
    investment_kpi_cols[3].metric("BALLOON", f"${balloon:,.0f}")

    if abs(balance_diff) > 1:
        st.error(f"🚨 UNBALANCED SOURCES & USES: Difference ${balance_diff:,.0f}")
    elif sponsor_equity_source < 0:
        st.warning(f"💸 Cash-out / surplus proceeds implied: ${abs(sponsor_equity_source):,.0f}")

    st.markdown("---")

    tabs = st.tabs(
        [
            "📊 Sizing & Risk",
            "🧪 Sensitivity",
            "📝 Rent Roll",
            "📅 Amortization",
            "📈 Pro Forma",
            "🇨🇦 Canada Intel",
            "📈 Market Comps",
            "📎 Diligence Room",
            "💾 Save & Export",
            "✅ QA & Health",
        ]
    )

    fragment = getattr(st, "fragment", None)

    def highlight_base_row(row):
        if row.get("Scenario") == "Base":
            return ["background-color: rgba(207, 184, 124, 0.35); font-weight: 700;"] * len(row)
        return [""] * len(row)

    def highlight_heatmap_base(data):
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        if "NOI Shock" in data.columns:
            base_rows = data.index[data["NOI Shock"] == "+0%"].tolist()
            if base_rows and "Rate +0.0%" in data.columns:
                styles.loc[base_rows[0], "Rate +0.0%"] = "background-color: rgba(207, 184, 124, 0.45); font-weight: 700;"
        return styles

    # ==========================================
    # TAB 0: SIZING & RISK
    # ==========================================

    with tabs[0]:
        col_left, col_right = st.columns([1.5, 1])

        with col_left:
            st.subheader("📐 Constraint Analysis")

            constraints_df = pd.DataFrame(
                {
                    "Constraint": ["LTV", "LTC", "DSCR", "Debt Yield"],
                    "Threshold": [
                        f"{s.target_ltv * 100:.1f}%",
                        f"{s.target_ltc * 100:.1f}%",
                        f"{s.target_dscr:.2f}x",
                        f"{s.target_dy * 100:.2f}%",
                    ],
                    "Max Proceeds": [
                        f"${gates.get('LTV', 0):,.0f}",
                        f"${gates.get('LTC', 0):,.0f}",
                        f"${gates.get('DSCR', 0):,.0f}",
                        f"${gates.get('Debt Yield', 0):,.0f}",
                    ],
                    "Binding": [
                        "✅ ACTIVE" if gate == g else ""
                        for g in ["LTV", "LTC", "DSCR", "Debt Yield"]
                    ],
                }
            )

            st.dataframe(constraints_df, hide_index=True, use_container_width=True)

            st.subheader("💰 Sources & Uses")

            total_fees = loan_amt * s.fees

            su_df = pd.DataFrame(
                {
                    "Uses": [
                        "Cost Basis",
                        "Closing Costs",
                        "Reserves",
                        "Financing Fees",
                        "TOTAL USES",
                    ],
                    "Use Amount": [
                        s.purchase_price,
                        s.closing_costs,
                        s.reserves,
                        total_fees,
                        total_uses,
                    ],
                    "Sources": [
                        "Senior Debt",
                        "Mezzanine Debt",
                        "Preferred Equity",
                        "Sponsor Equity / Cash-Out",
                        "TOTAL SOURCES",
                    ],
                    "Source Amount": [
                        loan_amt,
                        mezz_source,
                        pref_source,
                        sponsor_equity_source,
                        total_sources,
                    ],
                }
            )

            st.dataframe(
                su_df.style.format(
                    {
                        "Use Amount": "${:,.0f}",
                        "Source Amount": "${:,.0f}",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )

            if abs(balance_diff) > 1:
                st.error(f"🚨 UNBALANCED SOURCES & USES: Difference ${balance_diff:,.0f}")
            else:
                st.success("✅ Sources & Uses balanced")

            if sponsor_equity_source < 0:
                st.warning(f"Cash-out / surplus proceeds implied: ${abs(sponsor_equity_source):,.0f}")

        with col_right:
            st.subheader("🔍 Risk Assessment")

            flags = []

            if actual_ltv > 0.75:
                flags.append(("high", f"⚠️ High Leverage: {actual_ltv * 100:.1f}% LTV"))
            elif actual_ltv < 0.55 and loan_amt > 0:
                flags.append(("low", f"✅ Conservative Leverage: {actual_ltv * 100:.1f}% LTV"))

            if actual_dscr < 1.20 and loan_amt > 0:
                flags.append(("high", f"⚠️ Tight Coverage: {actual_dscr:.2f}x DSCR"))
            elif actual_dscr > 1.50:
                flags.append(("low", f"✅ Strong Coverage: {actual_dscr:.2f}x DSCR"))

            if s.is_io:
                flags.append(("medium", "ℹ️ Interest-Only Structure"))

            if sponsor_equity_source < 0:
                flags.append(("medium", f"ℹ️ Cash-Out Structure: ${abs(sponsor_equity_source):,.0f} surplus proceeds"))

            if walt > 0 and walt < 3:
                flags.append(("high", f"⚠️ Short WALT: {walt:.1f} years"))

            if exp1 > 0.30:
                flags.append(("high", f"🚨 High Rollover: {exp1 * 100:.1f}%"))

            if not flags:
                flags.append(("low", "✅ No Significant Risk Flags"))

            for severity, message in flags:
                if severity == "high":
                    st.error(message)
                elif severity == "medium":
                    st.warning(message)
                else:
                    st.success(message)

            st.markdown("---")
            st.subheader("📊 Key Metrics")

            breakeven_occ = UnderwritingEngine.breakeven_occupancy(
                current_noi=s.noi,
                current_occupancy=occ,
                annual_debt_service=annual_ds,
            )

            st.metric(
                "Breakeven Occupancy",
                f"{breakeven_occ * 100:.1f}%",
                delta=f"Current: {occ * 100:.1f}%" if occ > 0 else None,
            )

            if sponsor_equity_source < 0:
                st.metric("Surplus Proceeds", f"${abs(sponsor_equity_source):,.0f}")
            else:
                st.metric("Required Equity", f"${sponsor_equity_source:,.0f}")

            st.metric(
                "Implied Cap Rate",
                f"{safe_ratio(s.noi, s.appraisal) * 100:.2f}%" if s.appraisal > 0 else "N/A",
            )

            st.markdown("---")
            st.subheader("🏗️ Capital Stack")

            s.mezz_debt = st.number_input(
                "Mezzanine Debt ($)",
                value=safe_float(s.get("mezz_debt", 0)),
                min_value=0.0,
                step=100000.0,
                format="%.0f",
                key="capital_mezz_debt",
            )

            s.pref_equity = st.number_input(
                "Preferred Equity ($)",
                value=safe_float(s.get("pref_equity", 0)),
                min_value=0.0,
                step=100000.0,
                format="%.0f",
                key="capital_pref_equity",
            )

            s.mezz_rate = (
                st.slider(
                    "Mezz Rate %",
                    0.0,
                    25.0,
                    float(normalize_percent(s.get("mezz_rate", 0.10), 0.10, 0.40) * 100),
                    step=0.25,
                    key="capital_mezz_rate",
                )
                / 100
            )

            s.pref_rate = (
                st.slider(
                    "Pref Rate %",
                    0.0,
                    25.0,
                    float(normalize_percent(s.get("pref_rate", 0.09), 0.09, 0.40) * 100),
                    step=0.25,
                    key="capital_pref_rate",
                )
                / 100
            )

            stack = UnderwritingEngine.capital_stack(
                senior_debt=loan_amt,
                mezz_debt=s.mezz_debt,
                pref_equity=s.pref_equity,
                sponsor_equity=max(0, sponsor_equity_source),
                noi=s.noi,
                senior_rate=s.rate,
                mezz_rate=s.mezz_rate,
                pref_rate=s.pref_rate,
            )

            st.metric("Fixed Charge Coverage", f"{stack['Fixed Charge Coverage']:.2f}x")

            st.caption(
                f"Total Capital: ${stack['Total Capital']:,.0f} | "
                f"Fixed Charges: ${stack['Fixed Charges']:,.0f}"
            )

    # ==========================================
    # TAB 1: SENSITIVITY
    # ==========================================

    with tabs[1]:
        st.subheader("🧪 Sensitivity Analysis")

        sensitivity_df = SensitivityEngine.generate_matrix(extract_clean_state())

        st.dataframe(
            sensitivity_df.style.apply(highlight_base_row, axis=1),
            hide_index=True,
            use_container_width=True,
        )

        st.markdown("---")
        st.subheader("🔥 Proceeds Heatmap")

        heatmap_df = SensitivityEngine.proceeds_heatmap(extract_clean_state())
        currency_cols = [c for c in heatmap_df.columns if c != "NOI Shock"]

        st.dataframe(
            heatmap_df.style.format({col: "${:,.0f}" for col in currency_cols}).apply(
                highlight_heatmap_base,
                axis=None,
            ),
            hide_index=True,
            use_container_width=True,
        )

        if DEPENDENCIES.get("plotly"):
            try:
                import plotly.express as px

                plot_df = heatmap_df.set_index("NOI Shock")[currency_cols] / 1_000_000

                fig = px.imshow(
                    plot_df,
                    text_auto=".1f",
                    aspect="auto",
                    title="Max Proceeds Sensitivity ($MM)",
                    labels={
                        "x": "Rate Shock",
                        "y": "NOI Shock",
                        "color": "Max Proceeds ($MM)",
                    },
                )

                fig.update_traces(
                    texttemplate="$%{z:.1f}MM",
                    hovertemplate="NOI Shock: %{y}<br>Rate Shock: %{x}<br>Max Proceeds: $%{z:.1f}MM<extra></extra>",
                )

                fig.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="#0B0F19",
                    plot_bgcolor="#0F172A",
                    height=450,
                    margin=dict(l=20, r=20, t=50, b=20),
                )

                st.plotly_chart(fig, use_container_width=True)

            except Exception as e:
                st.info(f"Heatmap chart unavailable: {e}")

        st.markdown("---")
        st.subheader("🎯 Custom Stress Test")

        c1, c2, c3 = st.columns(3)

        with c1:
            rate_shock = (
                st.slider(
                    "Rate Shock (bps)",
                    -200,
                    200,
                    0,
                    25,
                    key="custom_rate_shock",
                )
                / 10000
            )

        with c2:
            noi_shock = (
                st.slider(
                    "NOI Shock (%)",
                    -30,
                    30,
                    0,
                    5,
                    key="custom_noi_shock",
                )
                / 100
            )

        with c3:
            ltv_shock = (
                st.slider(
                    "LTV Adjustment (%)",
                    -10,
                    10,
                    0,
                    1,
                    key="custom_ltv_shock",
                )
                / 100
            )

        stressed_rate = max(0.001, s.rate + rate_shock)
        stressed_noi = max(0, s.noi * (1 + noi_shock))
        stressed_ltv = min(1.25, max(0.01, s.target_ltv + ltv_shock))

        stressed_loan, stressed_gate, _, _, _ = UnderwritingEngine.size_loan(
            noi=stressed_noi,
            appraisal=s.appraisal,
            purchase_price=s.purchase_price,
            closing_costs=s.closing_costs,
            reserves=s.reserves,
            fees_pct=s.fees,
            rate=stressed_rate,
            amort=s.amort,
            term=s.term,
            is_io=s.is_io,
            target_ltv=stressed_ltv,
            target_ltc=s.target_ltc,
            target_dscr=s.target_dscr,
            target_dy=s.target_dy,
        )

        _, stressed_monthly_pmt, _ = UnderwritingEngine.amort_schedule(
            stressed_loan,
            stressed_rate,
            s.amort,
            s.term,
            s.is_io,
        )

        stressed_annual_ds = stressed_monthly_pmt * 12
        stressed_dscr = safe_ratio(stressed_noi, stressed_annual_ds)

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Stressed Proceeds",
            f"${stressed_loan:,.0f}",
            delta=f"${stressed_loan - loan_amt:,.0f}" if loan_amt > 0 else None,
        )

        c2.metric("Constraint", stressed_gate)

        c3.metric(
            "Stressed LTV",
            f"{safe_ratio(stressed_loan, s.appraisal) * 100:.1f}%" if s.appraisal > 0 else "N/A",
        )

        c4.metric("Stressed DSCR", f"{stressed_dscr:.2f}x")

    # ==========================================
    # TAB 2: RENT ROLL
    # ==========================================

    with tabs[2]:
        def render_rent_roll_tab():
            st.subheader("📝 Rent Roll Management")

            uploaded_file = st.file_uploader(
                "Import Rent Roll (CSV or Excel)",
                type=["csv", "xlsx", "xls"],
                key="rent_roll_upload",
            )

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
                            st.toast(f"Imported {len(imported_df)} tenant records", icon="✅")
                            st.rerun()

                except Exception as e:
                    st.error(f"Import failed: {str(e)[:200]}")

            edit_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

            col_add, col_clear, col_save = st.columns([1, 1, 1])

            with col_add:
                if st.button("➕ Add Blank Row", use_container_width=True, key="add_rr_row"):
                    rr_records = edit_df.to_dict("records")
                    rr_records.append(
                        {
                            "Tenant": "",
                            "SF": 0,
                            "Remaining Term": 0,
                            "Monthly Rent": 0,
                        }
                    )
                    s.rent_roll_dict = rr_records
                    s.unsaved_changes = True
                    st.toast("Blank rent roll row added", icon="➕")
                    st.rerun()

            with col_clear:
                if st.button("🧹 Clear Rent Roll", use_container_width=True, key="clear_rr"):
                    s.rent_roll_dict = []
                    s.unsaved_changes = True
                    st.toast("Rent roll cleared", icon="🧹")
                    st.rerun()

            with col_save:
                if st.button("💾 Save Deal", use_container_width=True, key="rent_roll_fragment_save"):
                    if save_current_deal(reason="MANUAL_SAVE"):
                        st.success("✅ Deal saved")
                        st.toast("Deal saved", icon="✅")

            st.caption(
                "Edits sync automatically. Use Recalculate Now to refresh the full model immediately, "
                "or Save Deal to commit directly from this tab."
            )

            edited_df = st.data_editor(
                edit_df,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Tenant": st.column_config.TextColumn("Tenant Name", width="large"),
                    "SF": st.column_config.NumberColumn(
                        "Square Feet",
                        min_value=0,
                        step=100,
                        format="%d",
                    ),
                    "Remaining Term": st.column_config.NumberColumn(
                        "Lease Term (Yrs)",
                        min_value=0.0,
                        step=0.5,
                        format="%.1f",
                    ),
                    "Monthly Rent": st.column_config.NumberColumn(
                        "Monthly Rent ($)",
                        min_value=0.0,
                        step=100.0,
                        format="$%.2f",
                    ),
                },
                key="rent_roll_editor",
            )

            normalized_editor_df = normalize_rent_roll_columns(edited_df)
            existing_rr_df = normalize_rent_roll_columns(pd.DataFrame(s.get("rent_roll_dict", [])))

            rent_roll_changed = not normalized_editor_df.equals(existing_rr_df)

            if rent_roll_changed:
                s.rent_roll_dict = normalized_editor_df.to_dict("records")
                s.unsaved_changes = True
                s.autosave_status = f"Rent roll edited at {datetime.now().strftime('%H:%M:%S')}"

                st.warning(
                    "Rent roll changes are staged. Click Recalculate Now to refresh the full model header immediately."
                )

            col_recalc, col_status = st.columns([1, 2])

            with col_recalc:
                if st.button("🟨 Recalculate Now", use_container_width=True, key="recalc_rent_roll_now"):
                    s.rent_roll_dict = normalized_editor_df.to_dict("records")
                    s.unsaved_changes = True
                    s.autosave_status = f"Rent roll synced at {datetime.now().strftime('%H:%M:%S')}"
                    st.toast("Rent roll synced", icon="✅")
                    st.rerun()

            with col_status:
                if rent_roll_changed:
                    st.caption("Status: staged changes detected. Full-page metrics refresh after Recalculate Now.")
                elif s.get("unsaved_changes"):
                    st.caption("Status: rent roll synced; deal has unsaved changes.")
                else:
                    st.caption("Status: rent roll synced and no unsaved changes detected.")

            if not normalized_editor_df.empty:
                st.markdown("---")
                st.subheader("📊 Rent Roll Analytics")

                met_total_sf, met_occ, met_ann_rent, met_psf, met_walt, met_exp1 = (
                    UnderwritingEngine.rent_roll_metrics(normalized_editor_df)
                )

                c1, c2, c3, c4, c5, c6 = st.columns(6)

                c1.metric("Total SF", f"{met_total_sf:,.0f}")
                c2.metric("Occupancy", f"{met_occ * 100:.1f}%")
                c3.metric("Annual Rent", f"${met_ann_rent:,.0f}")
                c4.metric("Rent PSF", f"${met_psf:.2f}")
                c5.metric("WALT (Yrs)", f"{met_walt:.2f}")
                c6.metric("12-Mo Rollover", f"{met_exp1 * 100:.1f}")

            else:
                st.info("No rent roll rows currently entered.")

        if fragment:
            fragment(render_rent_roll_tab)()
        else:
            render_rent_roll_tab()

    # ==========================================
    # TAB 3: AMORTIZATION
    # ==========================================

    with tabs[3]:
        st.subheader(f"📅 Amortization Schedule - {s.term} Year Term")

        if amort_df is None or amort_df.empty:
            st.warning("Enter deal parameters to generate amortization schedule.")
        else:
            c1, c2, c3 = st.columns(3)

            c1.metric("Monthly Payment", f"${monthly_pmt:,.2f}")
            c2.metric("Annual Debt Service", f"${annual_ds:,.0f}")
            c3.metric(
                "Balloon Balance",
                f"${balloon:,.0f}",
                delta=f"{safe_ratio(balloon, loan_amt):.1%} of Original" if loan_amt > 0 else None,
                delta_color="normal",
            )

            st.markdown("---")

            st.write("### 📉 Term vs. Balloon Analysis")
            st.caption("Visualizing principal paydown velocity over various exit dates.")

            term_scenarios = [3, 5, 7, 10, 15]

            if safe_int(s.term) not in term_scenarios:
                term_scenarios.append(safe_int(s.term))

            term_scenarios = sorted(
                [
                    t
                    for t in term_scenarios
                    if t > 0 and t <= 40
                ]
            )

            scenario_data = []

            for t in term_scenarios:
                _, _, b_bal = UnderwritingEngine.amort_schedule(
                    loan_amt,
                    s.rate,
                    s.amort,
                    t,
                    s.is_io,
                )

                scenario_data.append(
                    {
                        "Term (Yrs)": t,
                        "Balloon Balance": b_bal,
                        "Paydown %": 1 - safe_ratio(b_bal, loan_amt) if loan_amt > 0 else 0,
                    }
                )

            sens_df = pd.DataFrame(scenario_data)

            if DEPENDENCIES.get("plotly") and not sens_df.empty:
                try:
                    import plotly.graph_objects as go
                    from plotly.subplots import make_subplots

                    fig_sens = make_subplots(specs=[[{"secondary_y": True}]])

                    fig_sens.add_trace(
                        go.Bar(
                            x=sens_df["Term (Yrs)"],
                            y=sens_df["Balloon Balance"],
                            name="Balloon Amount",
                            marker_color="#1E293B",
                            hovertemplate="Term: %{x}yr<br>Balloon: $%{y:,.0f}<extra></extra>",
                        ),
                        secondary_y=False,
                    )

                    fig_sens.add_trace(
                        go.Scatter(
                            x=sens_df["Term (Yrs)"],
                            y=sens_df["Paydown %"],
                            name="Principal Paydown %",
                            line=dict(color="#CFB87C", width=4),
                            mode="lines+markers",
                            hovertemplate="Term: %{x}yr<br>Paydown: %{y:.1%}<extra></extra>",
                        ),
                        secondary_y=True,
                    )

                    fig_sens.update_layout(
                        template="plotly_dark",
                        title="Balloon Exposure by Loan Term",
                        paper_bgcolor="#0B0F19",
                        plot_bgcolor="rgba(0,0,0,0)",
                        xaxis=dict(title="Loan Term (Years)", tickmode="linear"),
                        yaxis=dict(
                            title="Balloon Balance ($)",
                            gridcolor="rgba(255,255,255,0.05)",
                        ),
                        yaxis2=dict(
                            title="Paydown %",
                            tickformat=".0%",
                            showgrid=False,
                        ),
                        height=400,
                        margin=dict(l=20, r=20, t=50, b=20),
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1,
                        ),
                    )

                    st.plotly_chart(fig_sens, use_container_width=True)

                except Exception as e:
                    st.info(f"Term vs. balloon chart unavailable: {e}")
                    st.dataframe(
                        sens_df.style.format(
                            {
                                "Balloon Balance": "${:,.0f}",
                                "Paydown %": "{:.1%}",
                            }
                        ),
                        hide_index=True,
                        use_container_width=True,
                    )
            else:
                st.dataframe(
                    sens_df.style.format(
                        {
                            "Balloon Balance": "${:,.0f}",
                            "Paydown %": "{:.1%}",
                        }
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

            amort_view = amort_df.copy()
            amort_view["Year"] = ((amort_view["Period"] - 1) // 12) + 1

            annual_summary = (
                amort_view.groupby("Year")
                .agg(
                    {
                        "Payment": "sum",
                        "Principal": "sum",
                        "Interest": "sum",
                        "Balance": "last",
                    }
                )
                .reset_index()
            )

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                st.write("#### Payment Structure")

                if DEPENDENCIES.get("plotly"):
                    try:
                        import plotly.graph_objects as go

                        pay_df = amort_df.copy()

                        fig_pay = go.Figure()

                        fig_pay.add_trace(
                            go.Bar(
                                x=pay_df["Period"],
                                y=pay_df["Principal"],
                                name="Principal",
                                marker_color="#CFB87C",
                            )
                        )

                        fig_pay.add_trace(
                            go.Bar(
                                x=pay_df["Period"],
                                y=pay_df["Interest"],
                                name="Interest",
                                marker_color="#1E293B",
                            )
                        )

                        fig_pay.update_layout(
                            barmode="stack",
                            template="plotly_dark",
                            paper_bgcolor="#0B0F19",
                            plot_bgcolor="rgba(0,0,0,0)",
                            height=360,
                            margin=dict(l=20, r=20, t=20, b=20),
                            xaxis_title="Period",
                            yaxis_title="Payment",
                            legend=dict(
                                orientation="h",
                                yanchor="bottom",
                                y=1.02,
                                xanchor="right",
                                x=1,
                            ),
                        )

                        st.plotly_chart(fig_pay, use_container_width=True)

                    except Exception:
                        st.bar_chart(
                            amort_df.set_index("Period")[["Principal", "Interest"]],
                            use_container_width=True,
                        )
                else:
                    st.bar_chart(
                        amort_df.set_index("Period")[["Principal", "Interest"]],
                        use_container_width=True,
                    )

            with col_chart2:
                st.write("#### Paydown Curve")

                if DEPENDENCIES.get("plotly"):
                    try:
                        import plotly.graph_objects as go

                        fig_bal = go.Figure()

                        fig_bal.add_trace(
                            go.Scatter(
                                x=amort_df["Period"],
                                y=amort_df["Balance"],
                                mode="lines",
                                name="Balance",
                                line=dict(color="#CFB87C", width=3),
                            )
                        )

                        fig_bal.update_layout(
                            template="plotly_dark",
                            paper_bgcolor="#0B0F19",
                            plot_bgcolor="rgba(0,0,0,0)",
                            height=360,
                            margin=dict(l=20, r=20, t=20, b=20),
                            xaxis_title="Period",
                            yaxis_title="Balance",
                        )

                        st.plotly_chart(fig_bal, use_container_width=True)

                    except Exception:
                        st.line_chart(
                            amort_df.set_index("Period")[["Balance"]],
                            use_container_width=True,
                        )
                else:
                    st.line_chart(
                        amort_df.set_index("Period")[["Balance"]],
                        use_container_width=True,
                    )

            with st.expander("📊 View Annual Summary & Full Schedule"):
                st.write("### Annual Summary")

                st.dataframe(
                    annual_summary.style.format(
                        {
                            "Payment": "${:,.2f}",
                            "Principal": "${:,.2f}",
                            "Interest": "${:,.2f}",
                            "Balance": "${:,.2f}",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

                st.write("### Full Monthly Schedule")

                st.dataframe(
                    amort_view[["Period", "Payment", "Principal", "Interest", "Balance"]].style.format(
                        {
                            "Payment": "${:,.2f}",
                            "Principal": "${:,.2f}",
                            "Interest": "${:,.2f}",
                            "Balance": "${:,.2f}",
                        }
                    ),
                    use_container_width=True,
                    height=400,
                    hide_index=True,
                )

    # ==========================================
    # TAB 4: PRO FORMA & RETURNS
    # ==========================================

    with tabs[4]:
        st.subheader("📈 10-Year Operating Pro Forma")
        st.caption("Revenue / expense growth, margin compression, exit valuation, and levered return profile.")

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            s.pf_revenue_growth = (
                st.slider(
                    "Revenue Growth %",
                    0.0,
                    8.0,
                    float(normalize_percent(s.get("pf_revenue_growth", 0.03), 0.03, 0.20) * 100),
                    0.25,
                    key="pf_revenue_growth_input",
                )
                / 100
            )

        with c2:
            s.pf_expense_growth = (
                st.slider(
                    "Expense Growth %",
                    0.0,
                    10.0,
                    float(normalize_percent(s.get("pf_expense_growth", 0.035), 0.035, 0.20) * 100),
                    0.25,
                    key="pf_expense_growth_input",
                )
                / 100
            )

        with c3:
            s.pf_expense_ratio = (
                st.slider(
                    "Initial Expense Ratio %",
                    10.0,
                    80.0,
                    float(normalize_percent(s.get("pf_expense_ratio", 0.40), 0.40, 0.95) * 100),
                    1.0,
                    key="pf_expense_ratio_input",
                )
                / 100
            )

        with c4:
            s.pf_projection_years = st.number_input(
                "Projection Years",
                min_value=5,
                max_value=30,
                value=max(5, min(30, safe_int(s.get("pf_projection_years", 10)))),
                step=1,
                key="pf_projection_years_input",
            )

        c5, c6, c7 = st.columns(3)

        with c5:
            implied_cap = (
                safe_ratio(s.noi, s.appraisal) * 100
                if safe_float(s.appraisal) > 0
                else 6.0
            )

            s.pf_exit_cap = (
                st.slider(
                    "Exit Cap Rate %",
                    3.0,
                    12.0,
                    float(min(max(safe_float(s.get("pf_exit_cap", (implied_cap + 0.50) / 100)) * 100, 3.0), 12.0)),
                    0.25,
                    key="pf_exit_cap_input",
                )
                / 100
            )

        with c6:
            s.pf_terminal_growth = (
                st.slider(
                    "Terminal NOI Growth %",
                    0.0,
                    5.0,
                    float(normalize_percent(s.get("pf_terminal_growth", 0.02), 0.02, 0.10) * 100),
                    0.25,
                    key="pf_terminal_growth_input",
                )
                / 100
            )

        with c7:
            s.pf_selling_costs = (
                st.slider(
                    "Selling Costs %",
                    0.0,
                    5.0,
                    float(normalize_percent(s.get("pf_selling_costs", 0.015), 0.015, 0.10) * 100),
                    0.25,
                    key="pf_selling_costs_input",
                )
                / 100
            )

        pro_forma_df = InvestmentEngine.calculate_pro_forma(
            stabilized_noi=s.noi,
            revenue_growth=s.pf_revenue_growth,
            expense_growth=s.pf_expense_growth,
            expense_ratio=s.pf_expense_ratio,
            years=s.pf_projection_years,
        )

        st.dataframe(
            pro_forma_df.style.format(
                {
                    "Revenue": "${:,.0f}",
                    "Expenses": "${:,.0f}",
                    "Projected NOI": "${:,.0f}",
                    "NOI Margin": "{:.2%}",
                    "Revenue Growth": "{:.2%}",
                    "Expense Growth": "{:.2%}",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        if not pro_forma_df.empty:
            first_margin = safe_float(pro_forma_df["NOI Margin"].iloc[0])
            final_margin = safe_float(pro_forma_df["NOI Margin"].iloc[-1])
            margin_delta = final_margin - first_margin

            if margin_delta < -0.03:
                st.warning(
                    f"Margin compression detected: NOI margin declines {abs(margin_delta):.2%} over the projection."
                )
            elif margin_delta > 0.03:
                st.success(
                    f"Margin expansion detected: NOI margin increases {margin_delta:.2%} over the projection."
                )
            else:
                st.info("NOI margin remains broadly stable across the projection period.")

        st.markdown("---")
        st.subheader("💰 Levered Investment Performance")

        returns = InvestmentEngine.solve_returns(
            purchase_price=s.purchase_price,
            loan_amt=loan_amt,
            pro_forma_df=pro_forma_df,
            exit_cap=s.pf_exit_cap,
            selling_costs=s.pf_selling_costs,
            annual_ds=annual_ds,
            balloon_balance=balloon,
            terminal_growth=s.pf_terminal_growth,
        )

        r1, r2, r3, r4, r5 = st.columns(5)

        r1.metric("Levered IRR", f"{returns['IRR']:.2%}")
        r2.metric("Equity Multiple", f"{returns['Equity Multiple']:.2f}x")
        r3.metric("Exit NOI", f"${returns['Exit NOI']:,.0f}")
        r4.metric("Gross Exit Value", f"${returns['Gross Exit Value']:,.0f}")
        r5.metric("Net Exit Proceeds", f"${returns['Net Exit Proceeds']:,.0f}")

        st.metric(
            "Total Cash Flow",
            f"${returns['Total Cash Flow']:,.0f}",
        )

        st.markdown("---")
        st.write("### Annual Cash Flow After Debt Service")

        cf_data = pro_forma_df.copy()
        cf_data["Cash Flow After Debt Service"] = cf_data["Projected NOI"] - annual_ds

        if DEPENDENCIES.get("plotly"):
            try:
                import plotly.express as px

                fig_returns = px.bar(
                    cf_data,
                    x="Year",
                    y="Cash Flow After Debt Service",
                    title="Annual Cash Flow After Debt Service",
                    color_discrete_sequence=["#CFB87C"],
                )

                fig_returns.update_layout(
                    template="plotly_dark",
                    paper_bgcolor="#0B0F19",
                    plot_bgcolor="#0F172A",
                    height=420,
                    margin=dict(l=20, r=20, t=50, b=20),
                )

                st.plotly_chart(fig_returns, use_container_width=True)

                sensitivity_rows = []

                for exit_cap_shock in [-0.005, 0, 0.005]:
                    for terminal_growth_shock in [-0.005, 0, 0.005]:
                        scenario_exit_cap = max(0.001, s.pf_exit_cap + exit_cap_shock)
                        scenario_terminal_growth = max(0, s.pf_terminal_growth + terminal_growth_shock)

                        scenario_returns = InvestmentEngine.solve_returns(
                            purchase_price=s.purchase_price,
                            loan_amt=loan_amt,
                            pro_forma_df=pro_forma_df,
                            exit_cap=scenario_exit_cap,
                            selling_costs=s.pf_selling_costs,
                            annual_ds=annual_ds,
                            balloon_balance=balloon,
                            terminal_growth=scenario_terminal_growth,
                        )

                        sensitivity_rows.append(
                            {
                                "Exit Cap": f"{scenario_exit_cap:.2%}",
                                "Terminal Growth": f"{scenario_terminal_growth:.2%}",
                                "Levered IRR": scenario_returns["IRR"],
                                "Equity Multiple": scenario_returns["Equity Multiple"],
                            }
                        )

                st.write("### Exit Cap / Terminal Growth Sensitivity")

                returns_sens_df = pd.DataFrame(sensitivity_rows)

                st.dataframe(
                    returns_sens_df.style.format(
                        {
                            "Levered IRR": "{:.2%}",
                            "Equity Multiple": "{:.2f}x",
                        }
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

            except Exception as e:
                st.info(f"Returns chart unavailable: {e}")

        else:
            st.bar_chart(
                cf_data.set_index("Year")[["Cash Flow After Debt Service"]],
                use_container_width=True,
            )
