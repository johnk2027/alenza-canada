"""
Alenza Capital OS - Configuration Module
Centralized configuration management with environment support
"""

import os
from pathlib import Path
from typing import Dict, Any
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class Config:
    """Application configuration with environment-specific settings"""
    
    # Application
    APP_NAME = "Alenza Capital OS"
    APP_VERSION = "3.0"
    SCHEMA_VERSION = 3
    DEBUG = os.getenv("ALENZA_DEBUG", "false").lower() == "true"
    
    # Paths
    BASE_DIR = Path(__file__).resolve().parent
    DATA_DIR = BASE_DIR / "alenza_data"
    DB_PATH = DATA_DIR / "alenza_platform.db"
    DOC_DIR = DATA_DIR / "documents"
    LOG_DIR = DATA_DIR / "logs"
    
    # Security
    SECRET_KEY = os.getenv("ALENZA_SECRET_KEY", "change-this-in-production-please!")
    ENCRYPTION_ENABLED = os.getenv("ALENZA_ENCRYPTION", "true").lower() == "true"
    
    # Database
    DB_TIMEOUT = 30
    DB_POOL_SIZE = 5
    DB_WAL_MODE = True
    
    # File Upload Limits
    MAX_UPLOAD_SIZE_MB = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50"))
    ALLOWED_EXTENSIONS = {".csv", ".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg"}
    
    # API Settings
    API_RETRY_COUNT = 3
    API_RETRY_BACKOFF = 0.6
    API_TIMEOUT = 12
    
    # Cache Settings
    CACHE_TTL_MARKET_DATA = 3600  # 1 hour
    CACHE_TTL_CORP_DATA = 86400   # 24 hours
    CACHE_TTL_GEOLOCATION = 86400  # 24 hours
    
    # Financial Limits
    MAX_AMORT_YEARS = 40
    MIN_AMORT_YEARS = 1
    MAX_RATE_PERCENT = 30.0
    MAX_LTV_PERCENT = 125.0
    MAX_RENT_ROLL_ROWS = 1000
    
    # Property Types
    PROPERTY_TYPES = [
        "Multifamily", "Industrial", "Retail", "Office", 
        "Mixed-Use", "Hospitality", "Self-Storage"
    ]
    
    # Lender Profiles
    LENDER_PROFILES = {
        "Bank / Credit Union": {"max_ltv": 0.75, "min_dscr": 1.25, "min_dy": 0.08},
        "LifeCo / Core": {"max_ltv": 0.65, "min_dscr": 1.35, "min_dy": 0.09},
        "Bridge / Private": {"max_ltv": 0.85, "min_dscr": 1.00, "min_dy": 0.07},
        "CMHC Multifamily": {"max_ltv": 0.95, "min_dscr": 1.10, "min_dy": 0.05},
    }
    
    # Required Documents
    REQUIRED_DOCS = [
        "Appraisal", "Phase I ESA", "T12 Financials", "Rent Roll",
        "Sponsor Bio", "Purchase Agreement", "Environmental Report"
    ]
    
    @classmethod
    def initialize_directories(cls):
        """Create all necessary directories"""
        for directory in [cls.DATA_DIR, cls.DOC_DIR, cls.LOG_DIR]:
            directory.mkdir(parents=True, exist_ok=True)
    
    @classmethod
    def get_database_url(cls) -> str:
        """Get database connection string"""
        return f"sqlite:///{cls.DB_PATH}"

# Production overrides
class ProductionConfig(Config):
    DEBUG = False
    ENCRYPTION_ENABLED = True

class DevelopmentConfig(Config):
    DEBUG = True
    ENCRYPTION_ENABLED = False

# Select configuration based on environment
ENV = os.getenv("ALENZA_ENV", "development")
config = ProductionConfig() if ENV == "production" else DevelopmentConfig()
