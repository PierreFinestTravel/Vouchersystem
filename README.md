# Finest Travel Africa - Automatic Travel Voucher Generator

A FastAPI-based web application that automatically generates travel vouchers from ORGA Excel files.

## Features

- **Automatic Parsing**: Reads ORGA Excel files and automatically detects:
  - Hotel stays (with check-in/check-out dates, room type, board basis)
  - Transfers (grouped by supplier with multiple pickup/dropoff points)
  - Car rentals (with pickup/dropoff locations and dates)
  - Activities and tours
  - Restaurant reservations
  - Golf bookings

- **Smart Voucher Generation**: Creates professionally formatted vouchers using the standard template
- **PDF Merging**: Combines all vouchers into a single PDF, sorted by:
  1. Hotels
  2. Transfers
  3. Car Rental
  4. Activities & Tours
  5. Restaurants
  6. Golf

## Prerequisites

1. **Python 3.8+**
2. **LibreOffice** (required for PDF conversion)
   - Download from: https://www.libreoffice.org/download/

## Installation

1. Clone this repository:
   ```bash
   git clone <repository-url>
   cd Finest-Travel-Africa---Automating-travel-Documentation
   ```

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure the voucher template is in place:
   - The `_Voucher blank.docx` template should be in the `templates/` folder

4. Install LibreOffice (if not already installed):
   - Windows: Download from https://www.libreoffice.org/download/
   - macOS: `brew install --cask libreoffice`
   - Linux: `sudo apt install libreoffice`

## Usage

1. Start the server:
   ```bash
   python run.py
   ```

2. Open http://localhost:8000 in your browser

3. Fill in the form:
   - **Traveller Names** (required): Enter names as they should appear on vouchers
     - Example: `Mr John Smith & Mrs Jane Smith`
   - **Reference Number** (optional): Booking reference
   - **Group / Pax Info** (optional): Additional group information
   - **ORGA Excel File** (required): Upload the ORGA `.xlsx` file

4. Click "Generate Vouchers"

5. Download the merged PDF containing all vouchers

## Project Structure

```
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI application
│   ├── models.py            # Data models
│   ├── orga_parser.py       # ORGA Excel parsing logic
│   ├── voucher_generator.py # Voucher document generation
│   ├── pdf_merger.py        # PDF conversion and merging
│   └── supplier_info.py     # Supplier contact database
├── templates/
│   └── _Voucher blank.docx  # Voucher template
├── requirements.txt
├── run.py                   # Application entry point
└── README.md
```

## ORGA Excel Format

The parser expects the standard ORGA format with:
- Row 10: Column headers
- Data rows starting from row 11/12

### Detected Columns:
- **Columns 1-8, 20-22**: Hotel information (Region, Supplier, Room, Board, Status, Notes)
- **Columns 23-31**: Golf information (Supplier, Course, Tee Time, Cart, Rental Set)
- **Columns 32-37**: Activity information (Supplier, Activity, Time, Notes)
- **Columns 38-48**: Transfer information (Supplier, Route, Service Type, Times, Flight info)

## Adding New Suppliers

To add new supplier contact information, edit `app/supplier_info.py`:

```python
SUPPLIER_INFO = {
    "supplier name lowercase": {
        "display_name": "SUPPLIER NAME",
        "address": "Full Address",
        "phone": "+27 (0)XX XXX XXXX",
        "gps": "GPS Coordinates"
    },
    # ...
}
```

## Troubleshooting

### "LibreOffice not found"
Install LibreOffice from https://www.libreoffice.org/download/

### "Template not found"
Ensure `_Voucher blank.docx` is in the `templates/` folder

### "No services found"
Check that:
- The ORGA file has the correct sheet name (containing "Orga")
- Data rows have dates in column C
- The file structure matches the expected format

## API Endpoints

- `GET /` - Main HTML interface
- `POST /generate` - Generate vouchers
  - Form fields: `traveller_names`, `ref_no`, `group_text`, `orga_file`
  - Returns: PDF file
- `GET /health` - Health check endpoint

## License

Internal tool for Finest Travel Africa.
