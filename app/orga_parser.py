"""ORGA Excel parser module.

This module parses the ORGA Excel file and extracts all service information
(hotels, transfers, activities, restaurants, car rentals, golf) into structured data.
"""
import logging
from datetime import datetime, date
from typing import Optional, List, Dict, Any, Tuple
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .models import (
    ParsedORGA, HotelStay, TransferVoucher, TransferLeg,
    ActivityVoucher, ActivityEntry, RestaurantVoucher,
    CarRentalVoucher, GolfVoucher
)
from .supplier_info import get_supplier_info

logger = logging.getLogger(__name__)

# Column indices (1-based as in Excel)
# Hotel columns
COL_DAYS = 1
COL_DAY = 2
COL_DATE = 3
COL_REGION_CITY = 4
COL_HOTEL_SUPPLIER = 5
COL_ROOM = 6
COL_BOARD = 7
COL_HOTEL_STATUS = 8
COL_HOTEL_NOTES = 20
COL_HOTEL_STATUS2 = 21
COL_HOTEL_INVOICE = 22

# Golf columns
COL_GOLF_SUPPLIER = 23
COL_GOLF_COURSE = 24
COL_TEE_TIME = 25
COL_DRIVING_RANGE = 26
COL_GOLF_CART = 27
COL_RENTAL_SET = 28
COL_GOLF_NOTES = 29
COL_GOLF_STATUS = 30
COL_GOLF_INVOICE = 31

# Activity columns
COL_ACTIVITY_SUPPLIER = 32
COL_ACTIVITY_NAME = 33
COL_ACTIVITY_TIME = 34
COL_ACTIVITY_NOTES = 35
COL_ACTIVITY_STATUS = 36
COL_ACTIVITY_INVOICE = 37

# Transfer columns
COL_TRANSFER_SUPPLIER = 38
COL_TRANSFER_ROUTE = 39
COL_SERVICE_TYPE = 40
COL_PICKUP_TIME = 41
COL_DROPOFF_TIME = 42
COL_FLIGHT_NUM = 43
COL_FLIGHT_TIME = 44
COL_TRAVEL_TIME = 45
COL_TRANSFER_NOTES = 46
COL_TRANSFER_STATUS = 47
COL_TRANSFER_INVOICE = 48


def get_cell_value(ws: Worksheet, row: int, col: int) -> Optional[str]:
    """Get cell value as string, handling None and whitespace."""
    val = ws.cell(row, col).value
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    return str(val).strip() if str(val).strip() else None


def parse_date(val: Any) -> Optional[date]:
    """Parse a date value from various formats."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    # Try parsing string formats
    try:
        return datetime.strptime(str(val), "%Y-%m-%d").date()
    except ValueError:
        pass
    try:
        return datetime.strptime(str(val), "%d.%m.%Y").date()
    except ValueError:
        pass
    return None


def find_header_row(ws: Worksheet) -> int:
    """Find the row containing column headers."""
    for row in range(1, 20):
        # Look for the "Days" header in column 1
        val = get_cell_value(ws, row, COL_DAYS)
        if val and str(val).lower() == "days":
            return row
    return 10  # Default based on analysis


def find_data_start_row(ws: Worksheet, header_row: int) -> int:
    """Find the first row with actual data after headers."""
    for row in range(header_row + 1, header_row + 10):
        date_val = get_cell_value(ws, row, COL_DATE)
        if date_val and parse_date(date_val):
            # Skip example rows
            days_val = get_cell_value(ws, row, COL_DAYS)
            if days_val and str(days_val).lower() == "e.g":
                continue
            return row
    return header_row + 2


def is_car_rental_row(route_val: str) -> bool:
    """Check if a transfer row is actually a car rental."""
    if not route_val:
        return False
    route_lower = route_val.lower()
    return "rental car" in route_lower or "group o" in route_lower or "group " in route_lower


def parse_orga(file_path: str) -> ParsedORGA:
    """Parse an ORGA Excel file and extract all service data."""
    logger.info(f"Parsing ORGA file: {file_path}")
    
    wb = load_workbook(file_path, data_only=True)
    
    # Find the correct sheet (prefer "Orga" sheets)
    ws = None
    for sheet_name in wb.sheetnames:
        if "orga" in sheet_name.lower():
            ws = wb[sheet_name]
            logger.info(f"Using sheet: {sheet_name}")
            break
    
    if ws is None:
        ws = wb.active
        logger.info(f"Using active sheet: {ws.title}")
    
    result = ParsedORGA()
    
    # Extract metadata from header rows
    for row in range(1, 10):
        label = get_cell_value(ws, row, 1)
        value = get_cell_value(ws, row, 4)
        if label and value:
            label_lower = label.lower()
            if "lead name" in label_lower:
                result.client_name = str(value)
            elif "pax" in label_lower:
                try:
                    result.pax = int(value)
                except ValueError:
                    pass
            elif "dates" in label_lower:
                result.dates = str(value)
            elif "trip number" in label_lower:
                result.trip_number = str(value)
    
    # Find header and data rows
    header_row = find_header_row(ws)
    data_start = find_data_start_row(ws, header_row)
    logger.info(f"Header row: {header_row}, Data starts: {data_start}")
    
    # Collect all data rows
    data_rows = []
    for row in range(data_start, ws.max_row + 1):
        date_val = get_cell_value(ws, row, COL_DATE)
        current_date = parse_date(date_val)
        
        if current_date is None:
            # Check if this is an "action" or notes row
            col1_val = get_cell_value(ws, row, 1)
            if col1_val and ("action" in str(col1_val).lower() or "book" in str(col1_val).lower()):
                break  # End of data rows
            continue
        
        data_rows.append({
            "row": row,
            "date": current_date,
            "days": get_cell_value(ws, row, COL_DAYS),
            # Hotel
            "region_city": get_cell_value(ws, row, COL_REGION_CITY),
            "hotel_supplier": get_cell_value(ws, row, COL_HOTEL_SUPPLIER),
            "room": get_cell_value(ws, row, COL_ROOM),
            "board": get_cell_value(ws, row, COL_BOARD),
            "hotel_status": get_cell_value(ws, row, COL_HOTEL_STATUS),
            "hotel_notes": get_cell_value(ws, row, COL_HOTEL_NOTES),
            # Golf
            "golf_supplier": get_cell_value(ws, row, COL_GOLF_SUPPLIER),
            "golf_course": get_cell_value(ws, row, COL_GOLF_COURSE),
            "tee_time": get_cell_value(ws, row, COL_TEE_TIME),
            "golf_cart": get_cell_value(ws, row, COL_GOLF_CART),
            "rental_set": get_cell_value(ws, row, COL_RENTAL_SET),
            "golf_notes": get_cell_value(ws, row, COL_GOLF_NOTES),
            # Activity
            "activity_supplier": get_cell_value(ws, row, COL_ACTIVITY_SUPPLIER),
            "activity_name": get_cell_value(ws, row, COL_ACTIVITY_NAME),
            "activity_time": get_cell_value(ws, row, COL_ACTIVITY_TIME),
            "activity_notes": get_cell_value(ws, row, COL_ACTIVITY_NOTES),
            # Transfer
            "transfer_supplier": get_cell_value(ws, row, COL_TRANSFER_SUPPLIER),
            "transfer_route": get_cell_value(ws, row, COL_TRANSFER_ROUTE),
            "service_type": get_cell_value(ws, row, COL_SERVICE_TYPE),
            "pickup_time": get_cell_value(ws, row, COL_PICKUP_TIME),
            "dropoff_time": get_cell_value(ws, row, COL_DROPOFF_TIME),
            "flight_num": get_cell_value(ws, row, COL_FLIGHT_NUM),
            "flight_time": get_cell_value(ws, row, COL_FLIGHT_TIME),
            "transfer_notes": get_cell_value(ws, row, COL_TRANSFER_NOTES),
            "transfer_status": get_cell_value(ws, row, COL_TRANSFER_STATUS),
        })
    
    logger.info(f"Found {len(data_rows)} data rows")
    
    # Parse hotels - group consecutive stays at the same hotel
    result.hotels = parse_hotels(data_rows)
    logger.info(f"Parsed {len(result.hotels)} hotel stays")
    
    # Parse transfers - group by supplier
    result.transfers = parse_transfers(data_rows)
    logger.info(f"Parsed {len(result.transfers)} transfer vouchers")
    
    # Parse car rentals
    result.car_rentals = parse_car_rentals(data_rows)
    logger.info(f"Parsed {len(result.car_rentals)} car rental vouchers")
    
    # Parse activities - group by supplier
    result.activities = parse_activities(data_rows)
    logger.info(f"Parsed {len(result.activities)} activity vouchers")
    
    # Parse restaurants
    result.restaurants = parse_restaurants(data_rows)
    logger.info(f"Parsed {len(result.restaurants)} restaurant vouchers")
    
    # Parse golf
    result.golf = parse_golf(data_rows)
    logger.info(f"Parsed {len(result.golf)} golf vouchers")
    
    return result


def parse_hotels(data_rows: List[Dict]) -> List[HotelStay]:
    """Parse hotel stays from data rows, grouping consecutive nights."""
    hotels = []
    current_hotel = None
    current_start = None
    
    for i, row in enumerate(data_rows):
        hotel_supplier = row.get("hotel_supplier")
        
        if hotel_supplier:
            # Clean up supplier name (remove trailing whitespace/newlines)
            hotel_supplier = hotel_supplier.strip().split('\n')[0].strip()
            
            if current_hotel is None or hotel_supplier.lower() != current_hotel.lower():
                # New hotel stay - save previous if exists
                if current_hotel is not None:
                    # Find the checkout date (current row's date)
                    checkout = row["date"]
                    nights = (checkout - current_start).days
                    
                    hotels.append(HotelStay(
                        supplier=current_hotel,
                        region_city=current_region,
                        room_type=current_room or "",
                        board=current_board or "",
                        check_in=current_start,
                        check_out=checkout,
                        nights=nights,
                        notes=current_notes or "",
                        status=current_status or ""
                    ))
                
                # Start new hotel
                current_hotel = hotel_supplier
                current_start = row["date"]
                current_region = row.get("region_city", "")
                current_room = row.get("room", "")
                current_board = row.get("board", "")
                current_notes = row.get("hotel_notes", "")
                current_status = row.get("hotel_status", "")
            else:
                # Same hotel - update room/notes if provided
                if row.get("room"):
                    current_room = row.get("room")
                if row.get("hotel_notes"):
                    if current_notes:
                        current_notes += "\n" + row.get("hotel_notes")
                    else:
                        current_notes = row.get("hotel_notes")
    
    # Don't forget the last hotel
    if current_hotel is not None:
        # For the last hotel, checkout is day after the last row
        last_date = data_rows[-1]["date"] if data_rows else current_start
        # Find next day for checkout
        from datetime import timedelta
        checkout = last_date + timedelta(days=1)
        nights = (checkout - current_start).days
        
        hotels.append(HotelStay(
            supplier=current_hotel,
            region_city=current_region,
            room_type=current_room or "",
            board=current_board or "",
            check_in=current_start,
            check_out=checkout,
            nights=nights,
            notes=current_notes or "",
            status=current_status or ""
        ))
    
    # Add supplier info
    for hotel in hotels:
        info = get_supplier_info(hotel.supplier)
        hotel.address = info.get("address", "")
        hotel.phone = info.get("phone", "")
        hotel.gps = info.get("gps", "")
    
    return hotels


def parse_transfers(data_rows: List[Dict]) -> List[TransferVoucher]:
    """Parse transfers from data rows, grouping by supplier."""
    transfers_by_supplier: Dict[str, TransferVoucher] = {}
    
    for row in data_rows:
        supplier = row.get("transfer_supplier")
        route = row.get("transfer_route")
        
        if not supplier:
            continue
        
        # Handle multi-line suppliers/routes
        suppliers = [s.strip() for s in supplier.split('\n') if s.strip()]
        routes = [r.strip() for r in (route or "").split('\n') if r.strip()]
        pickup_times = (row.get("pickup_time") or "").split('\n')
        dropoff_times = (row.get("dropoff_time") or "").split('\n')
        flight_nums = (row.get("flight_num") or "").split('\n')
        
        for idx, sup in enumerate(suppliers):
            rt = routes[idx] if idx < len(routes) else ""
            
            # Skip car rental entries
            if is_car_rental_row(rt):
                continue
            
            # Skip flight-only entries (Airlink, etc.) - they might be handled separately
            if "flight" in rt.lower() and "airport" not in rt.lower():
                continue
                
            sup_key = sup.lower().strip()
            
            if sup_key not in transfers_by_supplier:
                info = get_supplier_info(sup)
                transfers_by_supplier[sup_key] = TransferVoucher(
                    supplier=sup,
                    address=info.get("address", ""),
                    phone=info.get("phone", ""),
                    gps=info.get("gps", "")
                )
            
            # Parse pickup and dropoff from route
            pickup_loc = ""
            dropoff_loc = ""
            route_notes = ""
            
            if "-" in rt:
                parts = rt.split("-", 1)
                if len(parts) == 2:
                    first_part = parts[0].strip()
                    second_part = parts[1].strip()
                    
                    # Check if it's "Trf - Location" format
                    if first_part.lower() in ["trf", "transfer"]:
                        dropoff_loc = second_part
                    # Check if it has "incl." with additional info
                    elif "incl." in second_part.lower():
                        incl_parts = second_part.split("incl.", 1)
                        dropoff_loc = incl_parts[0].strip()
                        route_notes = "Includes: " + incl_parts[1].strip() if len(incl_parts) > 1 else ""
                        pickup_loc = first_part
                    else:
                        pickup_loc = first_part
                        dropoff_loc = second_part
            else:
                dropoff_loc = rt
            
            # Combine notes
            all_notes = []
            if route_notes:
                all_notes.append(route_notes)
            if row.get("transfer_notes"):
                all_notes.append(row.get("transfer_notes"))
            
            leg = TransferLeg(
                date=row["date"],
                pickup_location=pickup_loc,
                dropoff_location=dropoff_loc,
                pickup_time=pickup_times[idx] if idx < len(pickup_times) else "",
                dropoff_time=dropoff_times[idx] if idx < len(dropoff_times) else "",
                flight_number=flight_nums[idx] if idx < len(flight_nums) else "",
                notes="\n".join(all_notes)
            )
            transfers_by_supplier[sup_key].legs.append(leg)
    
    return list(transfers_by_supplier.values())


def parse_car_rentals(data_rows: List[Dict]) -> List[CarRentalVoucher]:
    """Parse car rental information from data rows."""
    car_rentals = []
    car_rental_data = {}
    
    for row in data_rows:
        route = row.get("transfer_route")
        if not route:
            continue
        
        routes = [r.strip() for r in route.split('\n') if r.strip()]
        
        for rt in routes:
            if is_car_rental_row(rt):
                # Extract car group info
                if "group" in rt.lower():
                    if "car_group" not in car_rental_data:
                        car_rental_data["car_group"] = rt
                        car_rental_data["pickup_date"] = row["date"]
                    car_rental_data["dropoff_date"] = row["date"]
                    
                if "collect" in rt.lower() or "pickup" in rt.lower():
                    car_rental_data["pickup_location"] = rt
                if "drop" in rt.lower() or "return" in rt.lower():
                    car_rental_data["dropoff_location"] = rt
    
    if car_rental_data.get("car_group"):
        info = get_supplier_info("pace car rental")  # Default car rental company
        
        # Parse car group for supplier info
        car_group = car_rental_data.get("car_group", "")
        
        car_rental = CarRentalVoucher(
            supplier="Pace Car Rental",  # Default, can be detected from ORGA
            car_group=car_group,
            pickup_date=car_rental_data.get("pickup_date"),
            pickup_location=car_rental_data.get("pickup_location", "Cape Town International Airport"),
            dropoff_date=car_rental_data.get("dropoff_date"),
            dropoff_location=car_rental_data.get("dropoff_location", "Cape Town International Airport"),
            address=info.get("address", ""),
            phone=info.get("phone", ""),
            gps=info.get("gps", "")
        )
        car_rentals.append(car_rental)
    
    return car_rentals


def parse_activities(data_rows: List[Dict]) -> List[ActivityVoucher]:
    """Parse activities from data rows, grouping by supplier."""
    activities_by_supplier: Dict[str, ActivityVoucher] = {}
    
    # Keywords that indicate restaurant (NOT activity)
    restaurant_only_keywords = ["dinner", "lunch"]
    
    # Keywords that indicate activity (even if at a restaurant-sounding venue)
    activity_keywords = ["tasting", "tour", "tickets", "watching", "drive", "panorama", "route", "safari"]
    
    for row in data_rows:
        supplier = row.get("activity_supplier")
        activity = row.get("activity_name")
        
        if not supplier:
            continue
        
        # Handle multi-line entries
        suppliers = [s.strip() for s in supplier.split('\n') if s.strip()]
        activities = [a.strip() for a in (activity or "").split('\n') if a.strip()]
        times = [t.strip() for t in (row.get("activity_time") or "").split('\n') if t.strip()]
        notes_list = [n.strip() for n in (row.get("activity_notes") or "").split('\n') if n.strip()]
        
        for idx, sup in enumerate(suppliers):
            act = activities[idx] if idx < len(activities) else ""
            time = times[idx] if idx < len(times) else ""
            notes = notes_list[idx] if idx < len(notes_list) else ""
            
            combined_text = (sup + " " + act + " " + notes).lower()
            
            # Check if this is explicitly an activity (e.g., wine tasting)
            is_explicit_activity = any(kw in combined_text for kw in activity_keywords)
            
            # Check if this is a restaurant meal
            is_restaurant_meal = any(kw in combined_text for kw in restaurant_only_keywords)
            
            # If it's a restaurant meal and NOT an activity, skip (handle in parse_restaurants)
            if is_restaurant_meal and not is_explicit_activity:
                continue
            
            # Check if this is a game drive at a safari lodge (should be part of hotel)
            if "game drive" in act.lower():
                continue  # Will be included in hotel voucher
            
            sup_key = sup.lower().strip()
            
            if sup_key not in activities_by_supplier:
                info = get_supplier_info(sup)
                activities_by_supplier[sup_key] = ActivityVoucher(
                    supplier=sup,
                    address=info.get("address", ""),
                    phone=info.get("phone", ""),
                    gps=info.get("gps", "")
                )
            
            entry = ActivityEntry(
                date=row["date"],
                activity_name=act or sup,  # Use supplier name if no activity specified
                time=time,
                notes=notes
            )
            activities_by_supplier[sup_key].entries.append(entry)
    
    return list(activities_by_supplier.values())


def parse_restaurants(data_rows: List[Dict]) -> List[RestaurantVoucher]:
    """Parse restaurant vouchers from data rows."""
    restaurants = []
    
    # Keywords for restaurant meals only
    restaurant_keywords = ["dinner", "lunch"]
    
    # Keywords that indicate activity (NOT restaurant) even at dining venues
    activity_keywords = ["tasting", "tour", "tickets", "watching"]
    
    for row in data_rows:
        supplier = row.get("activity_supplier")
        activity = row.get("activity_name")
        
        if not supplier:
            continue
        
        # Handle multi-line entries
        suppliers = [s.strip() for s in supplier.split('\n') if s.strip()]
        activities = [a.strip() for a in (activity or "").split('\n') if a.strip()]
        times = [t.strip() for t in (row.get("activity_time") or "").split('\n') if t.strip()]
        notes_list = [n.strip() for n in (row.get("activity_notes") or "").split('\n') if n.strip()]
        
        for idx, sup in enumerate(suppliers):
            act = activities[idx] if idx < len(activities) else ""
            time = times[idx] if idx < len(times) else ""
            notes = notes_list[idx] if idx < len(notes_list) else ""
            
            combined_text = (sup + " " + act + " " + notes).lower()
            
            # Check if this is a restaurant meal
            is_restaurant_meal = any(kw in combined_text for kw in restaurant_keywords)
            
            # Check if this is actually an activity (like wine tasting)
            is_activity = any(kw in combined_text for kw in activity_keywords)
            
            # Only include if it's a restaurant meal AND NOT an activity
            if not is_restaurant_meal or is_activity:
                continue
            
            info = get_supplier_info(sup)
            
            restaurant = RestaurantVoucher(
                supplier=sup,
                date=row["date"],
                time=time,
                notes=notes or act,  # Include activity description as notes
                address=info.get("address", ""),
                phone=info.get("phone", ""),
                gps=info.get("gps", "")
            )
            restaurants.append(restaurant)
    
    return restaurants


def parse_golf(data_rows: List[Dict]) -> List[GolfVoucher]:
    """Parse golf vouchers from data rows."""
    golf_vouchers = []
    
    for row in data_rows:
        supplier = row.get("golf_supplier")
        course = row.get("golf_course")
        
        if not supplier or not course:
            continue
        
        info = get_supplier_info(supplier)
        
        golf = GolfVoucher(
            supplier=supplier,
            course=course,
            date=row["date"],
            tee_time=row.get("tee_time", ""),
            cart=row.get("golf_cart", ""),
            rental_set=row.get("rental_set", ""),
            notes=row.get("golf_notes", ""),
            address=info.get("address", ""),
            phone=info.get("phone", ""),
            gps=info.get("gps", "")
        )
        golf_vouchers.append(golf)
    
    return golf_vouchers

