"""Data models for the voucher generator."""
from dataclasses import dataclass, field
from datetime import datetime, date
from typing import Optional, List
from enum import Enum


class VoucherType(Enum):
    HOTEL = "hotel"
    TRANSFER = "transfer"
    CAR_RENTAL = "car_rental"
    ACTIVITY = "activity"
    RESTAURANT = "restaurant"
    GOLF = "golf"


class BoardBasis(Enum):
    RO = "Room Only"
    BB = "Bed & Breakfast"
    HB = "Half Board"
    FB = "Full Board"
    FB_PLUS = "Full Board Plus"
    AI = "All Inclusive"


@dataclass
class HotelStay:
    """Represents a hotel stay with check-in/out dates."""
    supplier: str
    region_city: str
    room_type: str
    board: str
    check_in: date
    check_out: date
    nights: int
    notes: str = ""
    status: str = ""
    
    # Contact info (to be looked up from supplier database)
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass
class TransferLeg:
    """Represents one leg of a transfer."""
    date: date
    pickup_location: str
    dropoff_location: str
    pickup_time: str = ""
    dropoff_time: str = ""
    flight_number: str = ""
    flight_time: str = ""
    notes: str = ""


@dataclass
class TransferVoucher:
    """Represents a transfer voucher with multiple legs."""
    supplier: str
    legs: List[TransferLeg] = field(default_factory=list)
    notes: str = ""
    
    # Contact info
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass
class ActivityEntry:
    """Represents an activity on a specific date."""
    date: date
    activity_name: str
    time: str = ""
    notes: str = ""


@dataclass
class ActivityVoucher:
    """Represents an activity/tour voucher."""
    supplier: str
    entries: List[ActivityEntry] = field(default_factory=list)
    notes: str = ""
    
    # Contact info
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass
class RestaurantVoucher:
    """Represents a restaurant voucher."""
    supplier: str
    date: date
    time: str = ""
    notes: str = ""
    
    # Contact info
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass
class CarRentalVoucher:
    """Represents a car rental voucher."""
    supplier: str
    car_group: str
    pickup_date: date
    pickup_location: str
    dropoff_date: date
    dropoff_location: str
    notes: str = ""
    
    # Contact info
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass 
class GolfVoucher:
    """Represents a golf voucher."""
    supplier: str
    course: str
    date: date
    tee_time: str = ""
    cart: str = ""
    rental_set: str = ""
    notes: str = ""
    
    # Contact info
    address: str = ""
    phone: str = ""
    gps: str = ""


@dataclass
class ParsedORGA:
    """Container for all parsed ORGA data.
    
    NOTE: This contains SERVICE data from ORGA only (hotels, transfers, etc.)
    TRAVELLER NAMES must come from separate client files, NOT from ORGA.
    """
    hotels: List[HotelStay] = field(default_factory=list)
    transfers: List[TransferVoucher] = field(default_factory=list)
    activities: List[ActivityVoucher] = field(default_factory=list)
    restaurants: List[RestaurantVoucher] = field(default_factory=list)
    car_rentals: List[CarRentalVoucher] = field(default_factory=list)
    golf: List[GolfVoucher] = field(default_factory=list)
    
    # Region detection (SA = South Africa, EU = Europe)
    # This affects which voucher types are generated
    region: str = "SA"  # Default to SA
    
    # Metadata from ORGA header (for reference only - NOT for voucher names)
    # IMPORTANT: These fields are NOT used for traveller names on vouchers
    # Traveller names MUST come from uploaded client files (SINGLE or GROUP mode)
    client_name: str = ""  # DO NOT use for voucher traveller names
    trip_number: str = ""
    pax: int = 0
    dates: str = ""

