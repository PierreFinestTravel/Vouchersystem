"""Canonical Supplier Name Mapping System.

This module provides backwards compatibility with the old name_mapper interface.
All functionality is now in supplier_info.py which reads from config/suppliers.yaml.

The suppliers.yaml file is the SINGLE SOURCE OF TRUTH for:
- Supplier name mappings (ORGA name -> Voucher name)  
- Contact details (address, phone, GPS)
"""
import logging
from typing import List, Tuple

from .supplier_info import get_canonical_name, get_supplier_info

logger = logging.getLogger(__name__)

# Track suspicious names during a run (for debug report)
_suspicious_names_log: List[Tuple[str, str]] = []


def get_hotel_name(orga_name: str) -> str:
    """Get canonical hotel name."""
    return get_canonical_name(orga_name, 'hotels')


def get_golf_name(orga_name: str) -> str:
    """Get canonical golf course/club name."""
    return get_canonical_name(orga_name, 'golf')


def get_activity_name(orga_name: str) -> str:
    """Get canonical activity name."""
    return get_canonical_name(orga_name, 'activities')


def get_restaurant_name(orga_name: str) -> str:
    """Get canonical restaurant name."""
    return get_canonical_name(orga_name, 'restaurants')


def get_transfer_name(orga_name: str) -> str:
    """Get canonical transfer company name."""
    return get_canonical_name(orga_name, 'transfers')


def get_car_rental_name(orga_name: str) -> str:
    """Get canonical car rental company name."""
    return get_canonical_name(orga_name, 'car_rental')


def get_rental_clubs_name(orga_name: str) -> str:
    """Get canonical rental clubs company name."""
    return get_canonical_name(orga_name, 'rental_clubs')


def get_suspicious_names_log() -> List[Tuple[str, str]]:
    """Get list of suspicious names encountered during this run."""
    return list(_suspicious_names_log)


def clear_suspicious_names_log() -> None:
    """Clear the suspicious names log (call at start of new run)."""
    global _suspicious_names_log
    _suspicious_names_log = []
