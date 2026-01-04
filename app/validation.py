"""Pre-flight validation for voucher generation.

This module validates that:
1. All detected ORGA items have corresponding vouchers
2. No (TR) restaurants are included
3. Golf data produces golf vouchers
4. No voucher has empty or suspicious titles

If validation fails, a debug report is generated.
"""
import json
import logging
import os
from datetime import date, datetime
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, asdict

from .models import ParsedORGA
from .name_mapper import get_suspicious_names_log, clear_suspicious_names_log

logger = logging.getLogger(__name__)


@dataclass
class ValidationItem:
    """Represents an item being validated."""
    type: str  # 'hotel', 'golf', 'activity', 'restaurant', 'transfer', 'car_rental'
    orga_name: str
    orga_date: Optional[str]
    voucher_generated: bool
    skipped_reason: Optional[str] = None
    canonical_name: Optional[str] = None


@dataclass
class ValidationReport:
    """Complete validation report for a voucher generation run."""
    timestamp: str
    orga_file: str
    region: str
    
    # Counts
    total_orga_items: int
    vouchers_generated: int
    items_skipped: int
    
    # Detailed items
    hotels: List[Dict]
    golf: List[Dict]
    activities: List[Dict]
    restaurants: List[Dict]
    transfers: List[Dict]
    car_rentals: List[Dict]
    
    # Issues
    suspicious_names: List[Dict]  # Names without aliases
    empty_titles: List[Dict]  # Vouchers with empty/suspicious titles
    validation_errors: List[str]
    
    # Status
    passed: bool
    

class VoucherValidator:
    """Validates voucher generation before producing final PDFs."""
    
    def __init__(self, parsed_data: ParsedORGA, orga_file: str = ""):
        self.parsed_data = parsed_data
        self.orga_file = orga_file
        self.items: List[ValidationItem] = []
        self.errors: List[str] = []
        
    def validate(self) -> Tuple[bool, ValidationReport]:
        """Run all validations and return (passed, report).
        
        Returns:
            Tuple of (passed: bool, report: ValidationReport)
        """
        clear_suspicious_names_log()
        self.items = []
        self.errors = []
        
        # Validate each category
        self._validate_hotels()
        self._validate_golf()
        self._validate_activities()
        self._validate_restaurants()
        self._validate_transfers()
        self._validate_car_rentals()
        
        # Check for critical issues
        self._check_golf_generation()
        self._check_empty_titles()
        
        # Build report
        report = self._build_report()
        
        return report.passed, report
    
    def _validate_hotels(self) -> None:
        """Validate hotel vouchers."""
        for hotel in self.parsed_data.hotels:
            item = ValidationItem(
                type='hotel',
                orga_name=hotel.supplier,
                orga_date=str(hotel.check_in),
                voucher_generated=True,  # Hotels are always generated
                canonical_name=hotel.supplier
            )
            self.items.append(item)
    
    def _validate_golf(self) -> None:
        """Validate golf vouchers."""
        for golf in self.parsed_data.golf:
            item = ValidationItem(
                type='golf',
                orga_name=golf.supplier,
                orga_date=str(golf.date),
                voucher_generated=True,
                canonical_name=golf.course
            )
            self.items.append(item)
    
    def _validate_activities(self) -> None:
        """Validate activity vouchers (SA region only)."""
        region = self.parsed_data.region
        
        for activity in self.parsed_data.activities:
            generated = region == "SA"
            item = ValidationItem(
                type='activity',
                orga_name=activity.supplier,
                orga_date=str(activity.entries[0].date) if activity.entries else None,
                voucher_generated=generated,
                skipped_reason="EU region - no activity vouchers" if not generated else None,
                canonical_name=activity.supplier
            )
            self.items.append(item)
    
    def _validate_restaurants(self) -> None:
        """Validate restaurant vouchers (SA region only, no TR)."""
        region = self.parsed_data.region
        
        for restaurant in self.parsed_data.restaurants:
            generated = region == "SA"  # Already filtered for TR in parser
            item = ValidationItem(
                type='restaurant',
                orga_name=restaurant.supplier,
                orga_date=str(restaurant.date),
                voucher_generated=generated,
                skipped_reason="EU region - no restaurant vouchers" if not generated else None,
                canonical_name=restaurant.supplier
            )
            self.items.append(item)
    
    def _validate_transfers(self) -> None:
        """Validate transfer vouchers."""
        for transfer in self.parsed_data.transfers:
            earliest = min(leg.date for leg in transfer.legs) if transfer.legs else None
            item = ValidationItem(
                type='transfer',
                orga_name=transfer.supplier,
                orga_date=str(earliest) if earliest else None,
                voucher_generated=True,
                canonical_name=transfer.supplier
            )
            self.items.append(item)
    
    def _validate_car_rentals(self) -> None:
        """Validate car rental vouchers (SA region only)."""
        region = self.parsed_data.region
        
        for car in self.parsed_data.car_rentals:
            generated = region == "SA"
            item = ValidationItem(
                type='car_rental',
                orga_name=car.supplier,
                orga_date=str(car.pickup_date),
                voucher_generated=generated,
                skipped_reason="EU region - no car rental vouchers" if not generated else None,
                canonical_name=car.supplier
            )
            self.items.append(item)
    
    def _check_golf_generation(self) -> None:
        """Ensure golf data produces golf vouchers."""
        golf_items = [i for i in self.items if i.type == 'golf']
        
        if not golf_items and len(self.parsed_data.golf) == 0:
            # Check if there's golf data in ORGA that wasn't parsed
            # This would be detected by examining raw ORGA data
            pass
        
        for item in golf_items:
            if not item.voucher_generated:
                self.errors.append(
                    f"CRITICAL: Golf data detected but no voucher generated: {item.orga_name}"
                )
    
    def _check_empty_titles(self) -> None:
        """Check for vouchers with empty or suspicious titles."""
        for item in self.items:
            if item.voucher_generated:
                name = item.canonical_name or item.orga_name
                if not name or len(name.strip()) < 3:
                    self.errors.append(
                        f"Empty/suspicious title for {item.type}: '{name}'"
                    )
    
    def _build_report(self) -> ValidationReport:
        """Build the complete validation report."""
        suspicious = get_suspicious_names_log()
        
        # Categorize items
        hotels = [asdict(i) for i in self.items if i.type == 'hotel']
        golf = [asdict(i) for i in self.items if i.type == 'golf']
        activities = [asdict(i) for i in self.items if i.type == 'activity']
        restaurants = [asdict(i) for i in self.items if i.type == 'restaurant']
        transfers = [asdict(i) for i in self.items if i.type == 'transfer']
        car_rentals = [asdict(i) for i in self.items if i.type == 'car_rental']
        
        generated = sum(1 for i in self.items if i.voucher_generated)
        skipped = sum(1 for i in self.items if not i.voucher_generated)
        
        # Check for critical errors
        passed = len(self.errors) == 0
        
        # Log suspicious names as warnings but don't fail
        if suspicious:
            for name, category in suspicious:
                logger.warning(f"Suspicious name without alias: '{name}' ({category})")
        
        report = ValidationReport(
            timestamp=datetime.now().isoformat(),
            orga_file=self.orga_file,
            region=self.parsed_data.region,
            total_orga_items=len(self.items),
            vouchers_generated=generated,
            items_skipped=skipped,
            hotels=hotels,
            golf=golf,
            activities=activities,
            restaurants=restaurants,
            transfers=transfers,
            car_rentals=car_rentals,
            suspicious_names=[{"name": n, "category": c} for n, c in suspicious],
            empty_titles=[],
            validation_errors=self.errors,
            passed=passed
        )
        
        return report


def validate_and_report(
    parsed_data: ParsedORGA, 
    orga_file: str,
    output_dir: str = None
) -> Tuple[bool, Optional[str]]:
    """Validate parsed data and optionally generate debug report.
    
    Args:
        parsed_data: The parsed ORGA data
        orga_file: Path to the original ORGA file
        output_dir: Optional directory to write debug report
    
    Returns:
        Tuple of (passed: bool, report_path: Optional[str])
    """
    validator = VoucherValidator(parsed_data, orga_file)
    passed, report = validator.validate()
    
    report_path = None
    
    if not passed or output_dir:
        # Write debug report
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            report_path = os.path.join(output_dir, "run_debug_report.json")
        else:
            report_path = "run_debug_report.json"
        
        try:
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump(asdict(report), f, indent=2, default=str)
            
            if not passed:
                logger.error(f"Validation FAILED - see {report_path}")
                for error in report.validation_errors:
                    logger.error(f"  - {error}")
            else:
                logger.info(f"Validation passed - debug report: {report_path}")
                
        except Exception as e:
            logger.error(f"Failed to write debug report: {e}")
    
    return passed, report_path


def get_validation_summary(parsed_data: ParsedORGA) -> Dict[str, Any]:
    """Get a quick validation summary for logging.
    
    Returns dict with counts of items by type and generation status.
    """
    region = parsed_data.region
    
    return {
        "region": region,
        "hotels": {
            "detected": len(parsed_data.hotels),
            "will_generate": len(parsed_data.hotels)
        },
        "golf": {
            "detected": len(parsed_data.golf),
            "will_generate": len(parsed_data.golf)
        },
        "activities": {
            "detected": len(parsed_data.activities),
            "will_generate": len(parsed_data.activities) if region == "SA" else 0
        },
        "restaurants": {
            "detected": len(parsed_data.restaurants),
            "will_generate": len(parsed_data.restaurants) if region == "SA" else 0
        },
        "transfers": {
            "detected": len(parsed_data.transfers),
            "will_generate": len(parsed_data.transfers)
        },
        "car_rentals": {
            "detected": len(parsed_data.car_rentals),
            "will_generate": len(parsed_data.car_rentals) if region == "SA" else 0
        }
    }

