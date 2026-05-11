"""
Alenza Capital OS - Security Module
Handles encryption, hashing, and secure data handling
"""

import json
import base64
import hashlib
from typing import Dict, Any, Optional
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from config import config
import logging

logger = logging.getLogger(__name__)

class SecurityManager:
    """Manages encryption and security operations"""
    
    _fernet_instance = None
    
    @classmethod
    def _get_fernet(cls) -> Fernet:
        """Get or create Fernet instance for symmetric encryption"""
        if cls._fernet_instance is None:
            # Derive a 256-bit key from the secret
            kdf = PBKDF2HMAC(
                algorithm=hashes.SHA256(),
                length=32,
                salt=b"alenza_os_salt_2024",  # In production, use random salt stored in env
                iterations=100000,
            )
            key = base64.urlsafe_b64encode(
                kdf.derive(config.SECRET_KEY.encode())
            )
            cls._fernet_instance = Fernet(key)
        return cls._fernet_instance
    
    @classmethod
    def encrypt_sensitive_data(cls, data: Dict[str, Any]) -> Dict[str, Any]:
        """Encrypt sensitive fields in deal data"""
        if not config.ENCRYPTION_ENABLED:
            return data
        
        try:
            sensitive_fields = [
                "purchase_price", "appraisal", "noi", 
                "closing_costs", "reserves", "rent_roll_dict"
            ]
            
            sensitive_data = {}
            clean_data = data.copy()
            
            # Extract sensitive fields
            for field in sensitive_fields:
                if field in clean_data:
                    sensitive_data[field] = clean_data.pop(field)
                    clean_data[field] = "[ENCRYPTED]"
            
            # Encrypt
            fernet = cls._get_fernet()
            encrypted_bytes = fernet.encrypt(
                json.dumps(sensitive_data, default=str).encode()
            )
            clean_data["_encrypted"] = base64.b64encode(encrypted_bytes).decode()
            clean_data["_encryption_version"] = "1.0"
            
            return clean_data
        except Exception as e:
            logger.error(f"Encryption failed: {e}")
            return data  # Fallback to unencrypted
    
    @classmethod
    def decrypt_sensitive_data(cls, data: Dict[str, Any]) -> Dict[str, Any]:
        """Decrypt sensitive fields in deal data"""
        if not config.ENCRYPTION_ENABLED or "_encrypted" not in data:
            return data
        
        try:
            fernet = cls._get_fernet()
            encrypted_bytes = base64.b64decode(data["_encrypted"])
            decrypted_json = fernet.decrypt(encrypted_bytes)
            sensitive_data = json.loads(decrypted_json)
            
            # Merge decrypted data
            decrypted = data.copy()
            decrypted.update(sensitive_data)
            
            # Clean up encryption metadata
            decrypted.pop("_encrypted", None)
            decrypted.pop("_encryption_version", None)
            
            return decrypted
        except Exception as e:
            logger.error(f"Decryption failed: {e}")
            return data
    
    @classmethod
    def hash_file(cls, content: bytes) -> str:
        """Calculate SHA-256 hash of file content"""
        return hashlib.sha256(content).hexdigest()
    
    @classmethod
    def sanitize_filename(cls, filename: str) -> str:
        """Sanitize filename to prevent path traversal"""
        import re
        from pathlib import Path
        
        # Get name and extension
        path = Path(filename)
        name = path.stem or "file"
        ext = path.suffix[:10]  # Limit extension length
        
        # Remove potentially dangerous characters
        name = re.sub(r'[^a-zA-Z0-9_.-]', '_', name)
        name = name[:100]  # Limit length
        
        # Add random suffix to prevent collisions
        import uuid
        name = f"{name}_{uuid.uuid4().hex[:8]}"
        
        return f"{name}{ext}"
