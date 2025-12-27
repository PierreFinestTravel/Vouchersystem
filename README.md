# Finest Travel Africa - Automatic Travel Voucher Generator v2

A FastAPI-based web application that automatically generates travel vouchers from ORGA Excel files.

## New in Version 2.0 - SINGLE/GROUP Mode

The application now supports two modes:

### SINGLE Mode
- For individual clients (FIT - Free Independent Traveler)
- Upload ORGA Excel + Client confirmation document (.docx)
- Generates **ONE PDF** with all services
- Client names are automatically extracted from the confirmation file

### GROUP Mode
- For group trips with multiple clients
- Upload ORGA Excel + Group booking sheet (.xlsx)
- Generates **ONE PDF PER ROOM**
- Clients sharing a room get the same PDF
- Returns a ZIP file with all PDFs

## Features

- **Automatic Parsing**: Reads ORGA Excel files and detects:
  - Hotel stays (with check-in/check-out dates, room type, board basis)
  - Transfers (grouped by supplier with pickup/dropoff points)
  - Car rentals
  - Activities and tours
  - Restaurant reservations
  - Golf bookings

- **Trip ID Validation**: Ensures ORGA and client files belong to the same trip
- **Smart Name Extraction**: Automatically extracts traveller names from client files
- **Professional Formatting**: Generates vouchers matching the company template
- **PDF Merging**: Combines all vouchers into organized PDFs

## Prerequisites

1. **Python 3.8+**
2. **Microsoft Word** (for PDF conversion on Windows) or **LibreOffice**

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

## Usage

1. Start the server:
   ```bash
   python run.py
   ```

2. Open http://localhost:8000 in your browser

3. Select mode:
   - **SINGLE**: For individual client trips
   - **GROUP**: For group trips with room assignments

4. Enter **Trip ID** (4-digit code like 1008, 1115, 1222)

5. Upload files:
   - **ORGA Excel** (required)
   - **Client file** (varies by mode)

6. Click "Generate Vouchers"

7. Download:
   - **SINGLE mode**: One PDF file
   - **GROUP mode**: ZIP file with per-room PDFs

## File Formats

### ORGA Excel
Standard ORGA format with services in columns:
- Columns 1-8, 20-22: Hotel information
- Columns 23-31: Golf information
- Columns 32-37: Activity information
- Columns 38-48: Transfer information

### SINGLE Client File (.docx)
Confirmation document containing "Kundennamen:" or "Traveller names:" with client names

### GROUP Client File (.xlsx)
Booking sheet with columns:
- Room (number)
- Last Name
- First Name

Rows without a room number are considered sharing with the previous room.

## Project Structure

```
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI application (v2 with SINGLE/GROUP mode)
│   ├── models.py            # Data models
│   ├── orga_parser.py       # ORGA Excel parsing logic
│   ├── voucher_generator.py # Voucher document generation
│   ├── pdf_merger.py        # PDF conversion and merging
│   ├── supplier_info.py     # Supplier contact database
│   └── client_parser.py     # Client file parsing (NEW)
├── templates/
│   └── _Voucher blank.docx
├── requirements.txt
├── run.py
└── README.md
```

## Trip ID Examples

| Trip ID | Meaning |
|---------|---------|
| 1008 | October 8 departure |
| 1115 | November 15 departure |
| 1222 | December 22 departure |

## API Endpoints

- `GET /` - Main HTML interface (v2)
- `POST /generate` - Generate vouchers
  - Form fields: `mode`, `trip_id`, `ref_no`, `orga_file`, `single_client_file` or `group_client_file`
  - Returns: PDF (SINGLE) or ZIP (GROUP)
- `GET /health` - Health check endpoint

## Error Handling

- **Trip ID mismatch**: Blocks processing if ORGA and client file don't match
- **Missing files**: Clear error messages for missing required files
- **Invalid format**: Validation for file types and Trip ID format

## License

Internal tool for Finest Travel Africa.
