"""Supplier contact information database.

This module contains known supplier details (address, phone, GPS) that are used
to fill in voucher headers. Information is derived from the example vouchers.
"""

SUPPLIER_INFO = {
    # Hotels
    "home suite station house": {
        "display_name": "HOME SUITE HOTEL STATION HOUSE",
        "address": "19 Kloof Rd Cape Town",
        "phone": "+27 (0)82 362 2603",
        "gps": "S 33° 55' 9.119\", E 18° 23' 13.740\""
    },
    "whale rock lodge": {
        "display_name": "WHALE ROCK LUXURY LODGE",
        "address": "37 Springfield Avenue, Westcliff, Hermanus",
        "phone": "+27 (0)28 313 0014",
        "gps": "S 34° 24' 50.4\", E 19° 15' 21.6\""
    },
    "whale rock": {
        "display_name": "WHALE ROCK LUXURY LODGE",
        "address": "37 Springfield Avenue, Westcliff, Hermanus",
        "phone": "+27 (0)28 313 0014",
        "gps": "S 34° 24' 50.4\", E 19° 15' 21.6\""
    },
    "mgm wilderness": {
        "display_name": "MGM WILDERNESS",
        "address": "29 Roland Krynauw St, Wilderness, 6560",
        "phone": "+27 (0) 83 292 0753",
        "gps": "S 33.9959°, E 22.5875°"
    },
    "wedgeview": {
        "display_name": "WEDGEVIEW COUNTRY HOUSE & SPA",
        "address": "Bonniemile, Stellenbosch, 7604",
        "phone": "+27 (0)21 881 3525",
        "gps": "S 33° 58' 5.464\", E 18° 51' 37.828\""
    },
    "umlani": {
        "display_name": "UMLANI BUSH CAMP",
        "address": "Timbavati Game Reserve Hoedspruit",
        "phone": "+27 (0)21 785 5547",
        "gps": "S 24° 19' 54.588\", E 31° 18' 29.808\""
    },
    "ukuthula bush lodge": {
        "display_name": "UKUTHULA BUSH LODGE",
        "address": "Hoedspruit, Limpopo",
        "phone": "+27 (0)15 793 0267",
        "gps": "S 24° 21' 36.0\", E 30° 58' 12.0\""
    },
    "ukuthula": {
        "display_name": "UKUTHULA BUSH LODGE",
        "address": "Hoedspruit, Limpopo",
        "phone": "+27 (0)15 793 0267",
        "gps": "S 24° 21' 36.0\", E 30° 58' 12.0\""
    },
    
    # Transfer Companies
    "osprey tours": {
        "display_name": "OSPREY TOURS",
        "address": "",
        "phone": "+27 (0)81 032 7936",
        "gps": ""
    },
    "percy tours": {
        "display_name": "PERCY TOURS",
        "address": "46 Main Road, Hermanus, South Africa 7200",
        "phone": "+27(0)72 062 8500",
        "gps": "S 34.43370819091797, E 19.224746704101562"
    },
    
    # Car Rental
    "pace car rental": {
        "display_name": "PACE CAR RENTAL",
        "address": "Unit 6 Airport Business Park, Michigan Road Airport Industria Cape Town WP 7525 South Africa",
        "phone": "+27 (0)21 386 2411",
        "gps": "S 33.9761, E 18.5650"
    },
    "cabs car": {
        "display_name": "CABS CAR RENTAL",
        "address": "Cape Town International Airport",
        "phone": "+27 (0)21 380 5500",
        "gps": ""
    },
    
    # Activities
    "table mountain": {
        "display_name": "TABLE MOUNTAIN AERIAL CABLEWAY",
        "address": "Tafelberg Rd, Gardens, Cape Town, 8001",
        "phone": "+27 (0)21 424 8181",
        "gps": "S 33.9483°, E18.4029°"
    },
    "table mountain tickets": {
        "display_name": "TABLE MOUNTAIN AERIAL CABLEWAY",
        "address": "Tafelberg Rd, Gardens, Cape Town, 8001",
        "phone": "+27 (0)21 424 8181",
        "gps": "S 33.9483°, E18.4029°"
    },
    "ernie els": {
        "display_name": "ERNIE ELS WINES",
        "address": "Annandale Road, Stellenbosch",
        "phone": "+27 (0)21 881 3588",
        "gps": "S 33° 56' 24.0\", E 18° 52' 12.0\""
    },
    "guardian peak": {
        "display_name": "GUARDIAN PEAK GRILL & WINERY",
        "address": "Annandale Road, Stellenbosch",
        "phone": "+27 (0)21 881 3899",
        "gps": "S 33° 56' 30.0\", E 18° 52' 18.0\""
    },
    
    # Restaurants
    "the bungalow": {
        "display_name": "THE BUNGALOW",
        "address": "3 Victoria Road, Clifton, Cape Town",
        "phone": "+27 (0)21 438 2018",
        "gps": "S 33° 56' 16.8\", E 18° 22' 33.6\""
    },
    "char'd grill": {
        "display_name": "CHAR'D GRILL",
        "address": "Hermanus Waterfront, Market Square St",
        "phone": "+27 (0)28 312 1986",
        "gps": ""
    },
    "perlemoen restaurant": {
        "display_name": "PERLEMOEN RESTAURANT",
        "address": "Hermanus",
        "phone": "",
        "gps": ""
    },
    
    # Tours/Whale Watching
    "whale watching tour": {
        "display_name": "WHALE WATCHING BOAT TRIP",
        "address": "Hermanus New Harbour",
        "phone": "",
        "gps": ""
    },
    "whale watching": {
        "display_name": "WHALE WATCHING BOAT TRIP",
        "address": "Hermanus New Harbour, Hermanus",
        "phone": "+27 (0)28 312 2222",
        "gps": "S 34° 25' 12.0\", E 19° 15' 0.0\""
    },
    
    # Wine Estates / Activities
    "ernie els wines": {
        "display_name": "ERNIE ELS WINES",
        "address": "Annandale Road, Stellenbosch 7600",
        "phone": "+27 (0) 21 881 3588",
        "gps": "-34.01401835114784, 18.848032645432273"
    },
    "guardian peak grill": {
        "display_name": "GUARDIAN PEAK GRILL & WINERY",
        "address": "Annandale Road, Stellenbosch",
        "phone": "+27 (0)21 881 3899",
        "gps": "S 33° 56' 30.0\", E 18° 52' 18.0\""
    },
    
    # More restaurants
    "char'd grill": {
        "display_name": "CHAR'D GRILL & WINE BAR",
        "address": "Hermanus Waterfront, Market Square St",
        "phone": "+27 (0)28 312 1986",
        "gps": ""
    },
    "char'd grill & wine bar": {
        "display_name": "CHAR'D GRILL & WINE BAR",
        "address": "Hermanus Waterfront, Market Square St",
        "phone": "+27 (0)28 312 1986",
        "gps": ""
    },
}


def get_supplier_info(supplier_name: str) -> dict:
    """Look up supplier information by name (case-insensitive partial match)."""
    if not supplier_name:
        return {}
    
    name_lower = supplier_name.lower().strip()
    
    # Try exact match first
    if name_lower in SUPPLIER_INFO:
        return SUPPLIER_INFO[name_lower]
    
    # Try partial match
    for key, info in SUPPLIER_INFO.items():
        if key in name_lower or name_lower in key:
            return info
    
    # Return default with capitalized name
    return {
        "display_name": supplier_name.upper(),
        "address": "",
        "phone": "",
        "gps": ""
    }

