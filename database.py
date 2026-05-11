"""
Alenza Capital OS - Database Module
Enhanced database management with connection pooling and migration support
"""

import sqlite3
import json
import shutil
import pandas as pd
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple
from contextlib import contextmanager
from config import config
from security import SecurityManager
import logging

logger = logging.getLogger(__name__)

class DatabasePool:
    """Simple connection pool for SQLite"""
    
    def __init__(self, max_connections: int = 5):
        self.max_connections = max_connections
        self._connections = []
        self._in_use = set()
    
    @contextmanager
    def get_connection(self):
        """Get a connection from the pool"""
        conn = None
        try:
            # Try to reuse existing connection
            if self._connections:
                conn = self._connections.pop()
            else:
                conn = self._create_connection()
            
            self._in_use.add(id(conn))
            yield conn
            
        finally:
            if conn:
                self._in_use.discard(id(conn))
                if len(self._connections) < self.max_connections:
                    self._connections.append(conn)
                else:
                    conn.close()
    
    def _create_connection(self) -> sqlite3.Connection:
        """Create a new database connection"""
        conn = sqlite3.connect(
            config.DB_PATH, 
            timeout=config.DB_TIMEOUT,
            detect_types=sqlite3.PARSE_DECLTYPES
        )
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA busy_timeout=30000")
        conn.row_factory = sqlite3.Row
        return conn

# Global connection pool
db_pool = DatabasePool(max_connections=config.DB_POOL_SIZE)

class DatabaseManager:
    """Enhanced database manager with migration support"""
    
    @staticmethod
    def initialize_database():
        """Initialize database with current schema"""
        config.initialize_directories()
        
        with db_pool.get_connection() as conn:
            # Create schema version table
            conn.execute("""
                CREATE TABLE IF NOT EXISTS schema_version (
                    version INTEGER PRIMARY KEY,
                    applied_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    description TEXT
                )
            """)
            
            # Get current version
            cursor = conn.execute(
                "SELECT MAX(version) as version FROM schema_version"
            )
            row = cursor.fetchone()
            current_version = row["version"] if row["version"] else 0
            
            # Apply migrations
            DatabaseManager._apply_migrations(conn, current_version)
    
    @staticmethod
    def _apply_migrations(conn: sqlite3.Connection, current_version: int):
        """Apply pending migrations"""
        migrations = DatabaseManager._get_migrations()
        
        for version, description, sql in migrations:
            if version > current_version:
                logger.info(f"Applying migration {version}: {description}")
                try:
                    conn.executescript(sql)
                    conn.execute(
                        "INSERT INTO schema_version (version, description) VALUES (?, ?)",
                        (version, description)
                    )
                    conn.commit()
                    logger.info(f"Migration {version} applied successfully")
                except Exception as e:
                    logger.error(f"Migration {version} failed: {e}")
                    conn.rollback()
                    raise
    
    @staticmethod
    def _get_migrations() -> List[Tuple[int, str, str]]:
        """Define database migrations"""
        return [
            (1, "Initial schema", """
                CREATE TABLE IF NOT EXISTS deals (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    state_json TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS audit_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user TEXT NOT NULL,
                    action TEXT NOT NULL,
                    details TEXT,
                    deal_id TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE TABLE IF NOT EXISTS documents (
                    id TEXT PRIMARY KEY,
                    deal_id TEXT NOT NULL,
                    filename TEXT NOT NULL,
                    original_filename TEXT,
                    category TEXT NOT NULL,
                    path TEXT NOT NULL,
                    file_size INTEGER,
                    file_hash TEXT,
                    uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
                
                CREATE INDEX IF NOT EXISTS idx_deals_updated 
                ON deals(updated_at DESC);
                
                CREATE INDEX IF NOT EXISTS idx_audit_timestamp 
                ON audit_log(timestamp DESC);
            """),
            
            (2, "Add deal versioning", """
                CREATE TABLE IF NOT EXISTS deal_versions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    deal_id TEXT NOT NULL,
                    state_json TEXT,
                    changed_by TEXT,
                    change_summary TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (deal_id) REFERENCES deals(id) ON DELETE CASCADE
                );
                
                CREATE INDEX IF NOT EXISTS idx_versions_deal 
                ON deal_versions(deal_id, created_at DESC);
            """),
            
            (3, "Add encryption support", """
                ALTER TABLE deals ADD COLUMN encrypted INTEGER DEFAULT 0;
                
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                );
            """),
        ]
    
    @staticmethod
    def save_deal(deal_id: str, name: str, state: dict) -> bool:
        """Save deal with encryption and versioning"""
        try:
            # Encrypt sensitive data
            if config.ENCRYPTION_ENABLED:
                state = SecurityManager.encrypt_sensitive_data(state)
            
            state_json = json.dumps(state, default=str)
            safe_name = str(name or "Untitled Deal").strip()
            
            with db_pool.get_connection() as conn:
                # Save deal
                conn.execute("""
                    INSERT OR REPLACE INTO deals 
                    (id, name, state_json, updated_at, encrypted) 
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    deal_id, safe_name, state_json, 
                    datetime.now(timezone.utc), 
                    1 if config.ENCRYPTION_ENABLED else 0
                ))
                
                # Save version
                old_state = {}
                row = conn.execute(
                    "SELECT state_json FROM deals WHERE id = ?", 
                    (deal_id,)
                ).fetchone()
                
                if row:
                    try:
                        old_state = json.loads(row["state_json"])
                    except:
                        pass
                
                change_summary = DatabaseManager._summarize_changes(old_state, state)
                
                conn.execute("""
                    INSERT INTO deal_versions 
                    (deal_id, state_json, changed_by, change_summary) 
                    VALUES (?, ?, ?, ?)
                """, (deal_id, state_json, DatabaseManager.get_current_user(), change_summary))
                
                conn.commit()
                logger.info(f"Deal saved: {safe_name}")
                return True
                
        except Exception as e:
            logger.error(f"Failed to save deal: {e}")
            return False
    
    @staticmethod
    def load_deal(deal_id: str) -> Optional[dict]:
        """Load and decrypt deal"""
        try:
            with db_pool.get_connection() as conn:
                row = conn.execute(
                    "SELECT state_json, encrypted FROM deals WHERE id = ?",
                    (deal_id,)
                ).fetchone()
                
                if not row:
                    return None
                
                state = json.loads(row["state_json"])
                
                # Decrypt if needed
                if row["encrypted"] and config.ENCRYPTION_ENABLED:
                    state = SecurityManager.decrypt_sensitive_data(state)
                
                return state
                
        except Exception as e:
            logger.error(f"Failed to load deal: {e}")
            return None
    
    @staticmethod
    def get_current_user() -> str:
        """Get current authenticated user"""
        # This would integrate with your auth system
        import os
        return os.getenv("ALENZA_USER", "Local User")
    
    @staticmethod
    def _summarize_changes(old_state: dict, new_state: dict) -> str:
        """Generate human-readable change summary"""
        tracked_fields = [
            "deal_name", "purchase_price", "appraisal", "noi",
            "rate", "amort", "term", "target_ltv", "target_dscr"
        ]
        
        changes = []
        for field in tracked_fields:
            old_val = old_state.get(field) if old_state else None
            new_val = new_state.get(field) if new_state else None
            
            if str(old_val) != str(new_val):
                changes.append(f"{field}: {old_val} → {new_val}")
        
        return "; ".join(changes) if changes else "No material changes"
    
    @staticmethod
    def backup_database(backup_path: Optional[Path] = None) -> Path:
        """Create database backup"""
        if backup_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = config.DATA_DIR / f"backup_{timestamp}.db"
        
        with db_pool.get_connection() as conn:
            # Flush WAL to main database
            conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        
        # Create backup
        shutil.copy2(config.DB_PATH, backup_path)
        logger.info(f"Database backed up to {backup_path}")
        
        return backup_path
    
    @staticmethod
    def cleanup_old_data(retention_days: int = 365):
        """Clean up old audit logs and versions"""
        cutoff = datetime.now(timezone.utc) - timedelta(days=retention_days)
        
        with db_pool.get_connection() as conn:
            # Archive old audit logs
            conn.execute(
                "DELETE FROM audit_log WHERE timestamp < ?",
                (cutoff,)
            )
            
            # Archive old versions
            conn.execute("""
                DELETE FROM deal_versions 
                WHERE created_at < ? 
                AND id NOT IN (
                    SELECT id FROM deal_versions 
                    GROUP BY deal_id 
                    ORDER BY created_at DESC 
                    LIMIT 10
                )
            """, (cutoff,))
            
            conn.commit()
            logger.info(f"Cleaned up data older than {retention_days} days")
    
    @staticmethod
    def get_deal_versions(deal_id: str, limit: int = 20) -> pd.DataFrame:
        """Get version history for a deal"""
        try:
            with db_pool.get_connection() as conn:
                return pd.read_sql_query("""
                    SELECT changed_by, change_summary, created_at 
                    FROM deal_versions 
                    WHERE deal_id = ? 
                    ORDER BY created_at DESC 
                    LIMIT ?
                """, conn, params=(deal_id, limit))
        except Exception as e:
            logger.error(f"Failed to get versions: {e}")
            return pd.DataFrame()
    
    @staticmethod
    def get_all_deals() -> pd.DataFrame:
        """Get all deals"""
        try:
            with db_pool.get_connection() as conn:
                return pd.read_sql_query("""
                    SELECT id, name, created_at, updated_at 
                    FROM deals 
                    ORDER BY updated_at DESC
                """, conn)
        except Exception as e:
            logger.error(f"Failed to get deals: {e}")
            return pd.DataFrame()
