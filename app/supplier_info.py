"""Supplier contact information database.

This module reads all supplier information from a single YAML file:
    config/suppliers.yaml

This is the SINGLE SOURCE OF TRUTH for:
- Supplier name mappings (ORGA name -> Voucher name)
- Contact details (address, phone, GPS)

Your team can edit suppliers.yaml directly - no coding required!
"""
import logging
import os
from pathlib import Path
from typing import Dict, Optional
import yaml

logger = logging.getLogger(__name__)

# Cache for loaded suppliers
_suppliers_cache: Dict[str, dict] = {}
_last_load_time: float = 0


def _get_config_path() -> Path:
    """Get path to suppliers.yaml config file."""
    possible_paths = [
        Path(__file__).parent.parent / "config" / "suppliers.yaml",
        Path("config/suppliers.yaml"),
        Path("./config/suppliers.yaml"),
    ]
    
    for path in possible_paths:
        if path.exists():
            return path
    
    return possible_paths[0]


def _load_suppliers() -> None:
    """Load all suppliers from YAML config file."""
    global _suppliers_cache, _last_load_time
    
    config_path = _get_config_path()
    
    # Check if file was modified since last load
    if config_path.exists():
        mtime = config_path.stat().st_mtime
        if mtime <= _last_load_time and _suppliers_cache:
            return  # Already loaded and up-to-date
        _last_load_time = mtime
    
    _suppliers_cache = {}
    
    if not config_path.exists():
        logger.warning(f"Suppliers config not found: {config_path}")
        return
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        if not config:
            return
        
        # Load all categories into a flat dictionary
        # Keys are normalized to uppercase for case-insensitive lookup
        for category_name, suppliers in config.items():
            if not isinstance(suppliers, dict):
                continue
            
            for orga_name, info in suppliers.items():
                if not isinstance(info, dict):
                    continue
                
                # Normalize key to uppercase
                key = orga_name.upper().strip()
                
                _suppliers_cache[key] = {
                    "display_name": info.get("name", orga_name),
                    "address": info.get("address", "") or "",
                    "phone": info.get("phone", "") or "",
                    "gps": info.get("gps", "") or "",
                    "category": category_name
                }
        
        logger.info(f"Loaded {len(_suppliers_cache)} suppliers from {config_path}")
        
    except Exception as e:
        logger.error(f"Error loading suppliers: {e}")


def get_supplier_info(supplier_name: str, category: str = None) -> dict:
    """Look up supplier information by name (case-insensitive).
    
    Args:
        supplier_name: The supplier name from ORGA
        category: Optional category hint (not used, kept for compatibility)
    
    Returns:
        dict with 'display_name', 'address', 'phone', 'gps'
    """
    if not supplier_name:
        return {}
    
    # Ensure suppliers are loaded
    _load_suppliers()
    
    # Normalize for lookup
    name_upper = supplier_name.upper().strip()
    
    # Remove (TR) suffix if present
    import re
    name_upper = re.sub(r'\s*\(TR\)\s*$', '', name_upper, flags=re.IGNORECASE)
    name_upper = re.sub(r'\s+TR\s*$', '', name_upper, flags=re.IGNORECASE)
    name_upper = name_upper.strip()
    
    # Try exact match first
    if name_upper in _suppliers_cache:
        return _suppliers_cache[name_upper].copy()
    
    # Try partial match - look for key in name or name in key
    for key, info in _suppliers_cache.items():
        if key in name_upper or name_upper in key:
            return info.copy()
    
    # Try matching first significant word
    words = name_upper.split()
    if words:
        first_word = words[0]
        for key, info in _suppliers_cache.items():
            key_words = key.split()
            if key_words and first_word == key_words[0]:
                return info.copy()
    
    # Return default with formatted name
    return {
        "display_name": supplier_name.strip().upper() if supplier_name.isupper() else supplier_name.strip().title(),
        "address": "",
        "phone": "",
        "gps": ""
    }


def get_canonical_name(supplier_name: str, category: str = None) -> str:
    """Get the canonical (voucher template) name for a supplier.
    
    Args:
        supplier_name: The supplier name from ORGA
        category: Optional category hint (not used)
    
    Returns:
        The correct full name as it should appear on vouchers.
    """
    info = get_supplier_info(supplier_name, category)
    return info.get("display_name", supplier_name)


# Pre-load suppliers on module import
_load_suppliers()
