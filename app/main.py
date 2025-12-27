"""Main FastAPI application for the Travel Voucher Generator."""
import logging
import os
import tempfile
import shutil
import zipfile
from pathlib import Path
from typing import Optional, List

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from .orga_parser import parse_orga
from .voucher_generator import VoucherGenerator
from .pdf_merger import process_vouchers_to_pdf
from .client_parser import (
    parse_single_client_file, 
    parse_group_client_file,
    validate_trip_ids,
    extract_trip_id,
    RoomGroup
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Create FastAPI app
app = FastAPI(
    title="Travel Voucher Generator",
    description="Automatic travel voucher generation from ORGA Excel files",
    version="2.0.0"
)

# Get the base directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Template path
TEMPLATE_PATH = os.environ.get(
    "VOUCHER_TEMPLATE_PATH",
    str(BASE_DIR / "templates" / "_Voucher blank.docx")
)


def get_html_page() -> str:
    """Generate the HTML page for the frontend."""
    return '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Finest Travel Africa - Voucher Generator</title>
    <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500;600;700&family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
    <style>
        :root {
            --primary: #1a3a2f;
            --primary-light: #2d5a48;
            --gold: #c9a962;
            --gold-light: #dfc788;
            --cream: #f5f0e6;
            --cream-dark: #e8dfc9;
            --text-dark: #1a1a1a;
            --text-light: #f5f0e6;
            --shadow: 0 4px 24px rgba(0,0,0,0.12);
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Montserrat', sans-serif;
            background: linear-gradient(135deg, var(--primary) 0%, #0d1f18 100%);
            min-height: 100vh;
            color: var(--text-light);
            position: relative;
            overflow-x: hidden;
        }

        body::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M30 0L60 30L30 60L0 30z' fill='%23c9a962' fill-opacity='0.03'/%3E%3C/svg%3E");
            pointer-events: none;
        }

        .container {
            max-width: 720px;
            margin: 0 auto;
            padding: 40px 24px;
            position: relative;
            z-index: 1;
        }

        header {
            text-align: center;
            margin-bottom: 36px;
        }

        .logo {
            width: 100px;
            height: 100px;
            margin: 0 auto 20px;
            background: linear-gradient(135deg, var(--gold) 0%, var(--gold-light) 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: var(--shadow), 0 0 40px rgba(201, 169, 98, 0.2);
            animation: float 6s ease-in-out infinite;
        }

        @keyframes float {
            0%, 100% { transform: translateY(0px); }
            50% { transform: translateY(-8px); }
        }

        .logo svg {
            width: 50px;
            height: 50px;
            fill: var(--primary);
        }

        h1 {
            font-family: 'Cormorant Garamond', serif;
            font-size: 2.2rem;
            font-weight: 600;
            color: var(--gold);
            letter-spacing: 2px;
            margin-bottom: 8px;
        }

        .subtitle {
            font-size: 0.85rem;
            color: var(--gold-light);
            opacity: 0.8;
            letter-spacing: 3px;
            text-transform: uppercase;
        }

        .card {
            background: linear-gradient(180deg, rgba(255,255,255,0.08) 0%, rgba(255,255,255,0.04) 100%);
            backdrop-filter: blur(20px);
            border: 1px solid rgba(201, 169, 98, 0.2);
            border-radius: 20px;
            padding: 32px;
            box-shadow: var(--shadow);
        }

        .form-group {
            margin-bottom: 24px;
        }

        label {
            display: block;
            font-size: 0.7rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: var(--gold);
            margin-bottom: 8px;
        }

        label .required {
            color: #ff6b6b;
            margin-left: 4px;
        }

        input[type="text"],
        input[type="file"] {
            width: 100%;
            padding: 14px 18px;
            background: rgba(0,0,0,0.3);
            border: 1px solid rgba(201, 169, 98, 0.3);
            border-radius: 10px;
            font-family: 'Montserrat', sans-serif;
            font-size: 0.95rem;
            color: var(--text-light);
            transition: all 0.3s ease;
        }

        input[type="text"]:focus,
        input[type="file"]:focus {
            outline: none;
            border-color: var(--gold);
            box-shadow: 0 0 20px rgba(201, 169, 98, 0.2);
        }

        input[type="text"]::placeholder {
            color: rgba(245, 240, 230, 0.4);
        }

        input[type="file"] {
            cursor: pointer;
        }

        input[type="file"]::file-selector-button {
            padding: 8px 16px;
            margin-right: 12px;
            background: linear-gradient(135deg, var(--gold) 0%, var(--gold-light) 100%);
            border: none;
            border-radius: 6px;
            font-family: 'Montserrat', sans-serif;
            font-size: 0.8rem;
            font-weight: 500;
            color: var(--primary);
            cursor: pointer;
            transition: all 0.3s ease;
        }

        /* Mode Toggle */
        .mode-toggle {
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
        }

        .mode-option {
            flex: 1;
            position: relative;
        }

        .mode-option input[type="radio"] {
            position: absolute;
            opacity: 0;
            pointer-events: none;
        }

        .mode-option label {
            display: flex;
            flex-direction: column;
            align-items: center;
            padding: 20px;
            background: rgba(0,0,0,0.2);
            border: 2px solid rgba(201, 169, 98, 0.2);
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.3s ease;
            text-transform: none;
            letter-spacing: normal;
        }

        .mode-option label .mode-icon {
            font-size: 24px;
            margin-bottom: 8px;
        }

        .mode-option label .mode-title {
            font-weight: 600;
            font-size: 1rem;
            color: var(--gold);
            margin-bottom: 4px;
        }

        .mode-option label .mode-desc {
            font-size: 0.7rem;
            color: rgba(245, 240, 230, 0.6);
            text-align: center;
        }

        .mode-option input[type="radio"]:checked + label {
            border-color: var(--gold);
            background: rgba(201, 169, 98, 0.15);
            box-shadow: 0 0 20px rgba(201, 169, 98, 0.2);
        }

        .btn-submit {
            width: 100%;
            padding: 16px 28px;
            background: linear-gradient(135deg, var(--gold) 0%, var(--gold-light) 100%);
            border: none;
            border-radius: 10px;
            font-family: 'Montserrat', sans-serif;
            font-size: 0.95rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 2px;
            color: var(--primary);
            cursor: pointer;
            transition: all 0.3s ease;
            position: relative;
            overflow: hidden;
        }

        .btn-submit:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 32px rgba(201, 169, 98, 0.4);
        }

        .btn-submit:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
        }

        .btn-submit.loading {
            color: transparent;
        }

        .btn-submit.loading::after {
            content: '';
            position: absolute;
            width: 24px;
            height: 24px;
            top: 50%;
            left: 50%;
            margin: -12px 0 0 -12px;
            border: 3px solid var(--primary);
            border-top-color: transparent;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .message {
            margin-top: 20px;
            padding: 14px 18px;
            border-radius: 10px;
            font-size: 0.85rem;
            text-align: center;
            display: none;
        }

        .message.success {
            display: block;
            background: rgba(76, 175, 80, 0.2);
            border: 1px solid rgba(76, 175, 80, 0.4);
            color: #81c784;
        }

        .message.error {
            display: block;
            background: rgba(244, 67, 54, 0.2);
            border: 1px solid rgba(244, 67, 54, 0.4);
            color: #ef9a9a;
        }

        .divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(201, 169, 98, 0.3), transparent);
            margin: 24px 0;
        }

        .info-text {
            font-size: 0.75rem;
            color: rgba(245, 240, 230, 0.5);
            margin-top: 6px;
        }

        /* Client file section - changes based on mode */
        #single-file-section, #group-file-section {
            display: none;
        }

        #single-file-section.active, #group-file-section.active {
            display: block;
        }

        footer {
            text-align: center;
            margin-top: 32px;
            font-size: 0.75rem;
            color: rgba(245, 240, 230, 0.4);
        }

        @media (max-width: 600px) {
            .container { padding: 20px 16px; }
            h1 { font-size: 1.6rem; }
            .card { padding: 24px 18px; }
            .mode-toggle { flex-direction: column; }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="logo">
                <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <path d="M12 2L2 7l10 5 10-5-10-5zM2 17l10 5 10-5M2 12l10 5 10-5"/>
                </svg>
            </div>
            <h1>Finest Travel Africa</h1>
            <p class="subtitle">Voucher Generator v2</p>
        </header>

        <div class="card">
            <form id="voucherForm" enctype="multipart/form-data">
                
                <!-- Mode Toggle -->
                <label style="margin-bottom: 12px;">Mode<span class="required">*</span></label>
                <div class="mode-toggle">
                    <div class="mode-option">
                        <input type="radio" name="mode" id="mode-single" value="single" checked>
                        <label for="mode-single">
                            <span class="mode-icon">👤</span>
                            <span class="mode-title">SINGLE</span>
                            <span class="mode-desc">One client / One PDF</span>
                        </label>
                    </div>
                    <div class="mode-option">
                        <input type="radio" name="mode" id="mode-group" value="group">
                        <label for="mode-group">
                            <span class="mode-icon">👥</span>
                            <span class="mode-title">GROUP</span>
                            <span class="mode-desc">Multiple clients / Per-room PDFs</span>
                        </label>
                    </div>
                </div>

                <div class="divider"></div>

                <!-- Trip ID -->
                <div class="form-group">
                    <label>Trip ID<span class="required">*</span></label>
                    <input type="text" name="trip_id" id="trip_id" 
                           placeholder="e.g., 1008, 1115, 1222" maxlength="4" required>
                    <p class="info-text">4-digit departure code (must match ORGA and client file)</p>
                </div>

                <!-- Reference Number (optional) -->
                <div class="form-group">
                    <label>Reference Number</label>
                    <input type="text" name="ref_no" id="ref_no" 
                           placeholder="e.g., 15123">
                </div>

                <div class="divider"></div>

                <!-- ORGA File -->
                <div class="form-group">
                    <label>ORGA Excel File<span class="required">*</span></label>
                    <input type="file" name="orga_file" id="orga_file" 
                           accept=".xlsx,.xls" required>
                    <p class="info-text">Upload the ORGA Excel file (.xlsx)</p>
                </div>

                <!-- SINGLE mode: Client confirmation file -->
                <div id="single-file-section" class="active">
                    <div class="form-group">
                        <label>Client Confirmation File<span class="required">*</span></label>
                        <input type="file" name="single_client_file" id="single_client_file" 
                               accept=".docx">
                        <p class="info-text">Upload the client confirmation document (.docx)</p>
                    </div>
                </div>

                <!-- GROUP mode: Client booking sheet -->
                <div id="group-file-section">
                    <div class="form-group">
                        <label>Group Client Excel<span class="required">*</span></label>
                        <input type="file" name="group_client_file" id="group_client_file" 
                               accept=".xlsx,.xls">
                        <p class="info-text">Upload the group booking sheet with room assignments (.xlsx)</p>
                    </div>
                </div>

                <button type="submit" class="btn-submit" id="submitBtn">
                    Generate Vouchers
                </button>

                <div class="message" id="message"></div>
            </form>
        </div>

        <footer>
            <p>&copy; 2025 Finest Travel Africa. Internal Tool.</p>
        </footer>
    </div>

    <script>
        const form = document.getElementById('voucherForm');
        const submitBtn = document.getElementById('submitBtn');
        const messageDiv = document.getElementById('message');
        const singleSection = document.getElementById('single-file-section');
        const groupSection = document.getElementById('group-file-section');
        const modeSingle = document.getElementById('mode-single');
        const modeGroup = document.getElementById('mode-group');
        const singleFileInput = document.getElementById('single_client_file');
        const groupFileInput = document.getElementById('group_client_file');

        // Toggle mode sections
        function updateModeSection() {
            if (modeSingle.checked) {
                singleSection.classList.add('active');
                groupSection.classList.remove('active');
                singleFileInput.required = true;
                groupFileInput.required = false;
            } else {
                singleSection.classList.remove('active');
                groupSection.classList.add('active');
                singleFileInput.required = false;
                groupFileInput.required = true;
            }
        }

        modeSingle.addEventListener('change', updateModeSection);
        modeGroup.addEventListener('change', updateModeSection);

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            const mode = document.querySelector('input[name="mode"]:checked').value;
            const tripId = document.getElementById('trip_id').value.trim();
            const orgaFile = document.getElementById('orga_file').files[0];
            
            // Validate
            if (!tripId || tripId.length !== 4) {
                showMessage('Please enter a valid 4-digit Trip ID', 'error');
                return;
            }
            
            if (!orgaFile) {
                showMessage('Please select an ORGA Excel file', 'error');
                return;
            }
            
            if (mode === 'single' && !singleFileInput.files[0]) {
                showMessage('Please select a client confirmation file', 'error');
                return;
            }
            
            if (mode === 'group' && !groupFileInput.files[0]) {
                showMessage('Please select a group client Excel file', 'error');
                return;
            }
            
            // Show loading state
            submitBtn.classList.add('loading');
            submitBtn.disabled = true;
            messageDiv.style.display = 'none';
            
            try {
                const formData = new FormData(form);
                
                const response = await fetch('/generate', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const contentType = response.headers.get('Content-Type');
                    const blob = await response.blob();
                    const contentDisposition = response.headers.get('Content-Disposition');
                    
                    let filename = mode === 'group' ? 'Vouchers.zip' : 'Travel_Vouchers.pdf';
                    if (contentDisposition) {
                        const match = contentDisposition.match(/filename="?([^"]+)"?/);
                        if (match) filename = match[1];
                    }
                    
                    // Download
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    
                    const msg = mode === 'group' 
                        ? 'Vouchers generated! ZIP file with per-room PDFs is downloading.'
                        : 'Voucher generated! PDF is downloading.';
                    showMessage(msg, 'success');
                } else {
                    const error = await response.json();
                    showMessage(error.detail || 'An error occurred', 'error');
                }
            } catch (error) {
                showMessage('An error occurred: ' + error.message, 'error');
            } finally {
                submitBtn.classList.remove('loading');
                submitBtn.disabled = false;
            }
        });

        function showMessage(text, type) {
            messageDiv.textContent = text;
            messageDiv.className = 'message ' + type;
        }
    </script>
</body>
</html>'''


@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main HTML page."""
    return get_html_page()


@app.post("/generate")
async def generate_vouchers(
    mode: str = Form(..., description="Mode: 'single' or 'group'"),
    trip_id: str = Form(..., description="4-digit Trip ID"),
    ref_no: str = Form("", description="Reference number (optional)"),
    orga_file: UploadFile = File(..., description="ORGA Excel file"),
    single_client_file: Optional[UploadFile] = File(None, description="Single client .docx file"),
    group_client_file: Optional[UploadFile] = File(None, description="Group client .xlsx file")
):
    """Generate vouchers based on mode (SINGLE or GROUP)."""
    logger.info(f"Request: mode={mode}, trip_id={trip_id}, ref={ref_no}")
    
    # Validate mode
    if mode not in ['single', 'group']:
        raise HTTPException(status_code=400, detail="Invalid mode. Use 'single' or 'group'")
    
    # Validate Trip ID
    if not trip_id or len(trip_id) != 4 or not trip_id.isdigit():
        raise HTTPException(status_code=400, detail="Trip ID must be a 4-digit number")
    
    # Validate file presence based on mode
    if mode == 'single' and (not single_client_file or not single_client_file.filename):
        raise HTTPException(status_code=400, detail="Single mode requires a client confirmation file (.docx)")
    
    if mode == 'group' and (not group_client_file or not group_client_file.filename):
        raise HTTPException(status_code=400, detail="Group mode requires a client booking sheet (.xlsx)")
    
    # Create temp directory
    temp_dir = tempfile.mkdtemp(prefix="voucher_gen_")
    
    try:
        # Save ORGA file
        orga_path = os.path.join(temp_dir, orga_file.filename)
        with open(orga_path, "wb") as f:
            content = await orga_file.read()
            f.write(content)
        
        # Save and validate client file
        if mode == 'single':
            client_path = os.path.join(temp_dir, single_client_file.filename)
            with open(client_path, "wb") as f:
                content = await single_client_file.read()
                f.write(content)
            client_filename = single_client_file.filename
        else:
            client_path = os.path.join(temp_dir, group_client_file.filename)
            with open(client_path, "wb") as f:
                content = await group_client_file.read()
                f.write(content)
            client_filename = group_client_file.filename
        
        # Validate Trip IDs match
        is_valid, orga_id, client_id = validate_trip_ids(orga_file.filename, client_filename)
        
        # Also check against user-provided Trip ID
        if orga_id != trip_id and orga_id != "?":
            raise HTTPException(
                status_code=400,
                detail=f"Trip ID mismatch: You entered '{trip_id}' but ORGA file has '{orga_id}'"
            )
        
        if not is_valid and client_id != "?":
            raise HTTPException(
                status_code=400,
                detail=f"Trip ID mismatch. ORGA file: {orga_id}, Client file: {client_id}. Please make sure both files belong to the same trip."
            )
        
        # Check template
        if not os.path.exists(TEMPLATE_PATH):
            raise HTTPException(status_code=500, detail=f"Voucher template not found: {TEMPLATE_PATH}")
        
        # Parse ORGA
        try:
            parsed_data = parse_orga(orga_path)
        except Exception as e:
            logger.error(f"Failed to parse ORGA: {e}")
            raise HTTPException(status_code=400, detail=f"Failed to parse ORGA file: {str(e)}")
        
        # Generate vouchers based on mode
        if mode == 'single':
            return await generate_single_mode(
                parsed_data, client_path, temp_dir, trip_id, ref_no
            )
        else:
            return await generate_group_mode(
                parsed_data, client_path, temp_dir, trip_id, ref_no
            )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Unexpected error")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred: {str(e)}")


async def generate_single_mode(parsed_data, client_path: str, temp_dir: str, trip_id: str, ref_no: str):
    """Generate vouchers for SINGLE mode - one PDF for all services.
    
    CRITICAL: Names are ONLY extracted from the uploaded client file.
    - Never from ORGA file
    - Never from ORGA filename
    - Never guessed or generated as fallback
    - If parsing fails, return error
    """
    # Extract traveller names from CLIENT FILE ONLY - NO FALLBACK
    logger.info(f"SINGLE mode: Parsing names from CLIENT FILE: {os.path.basename(client_path)}")
    logger.info("NOTE: Names are extracted from client file ONLY, never from ORGA")
    
    try:
        names = parse_single_client_file(client_path)
        
        # STRICT: If no names found, FAIL - do not use any fallback
        if not names:
            logger.error("SINGLE mode FAILED: No traveller names found in client file")
            raise HTTPException(
                status_code=400, 
                detail="FAILED: Could not extract traveller names from client file. "
                       "The file must contain 'Kundennamen:', 'Traveller names:', or similar pattern. "
                       "Names are NOT taken from ORGA - they MUST be in the client confirmation file."
            )
        
        traveller_names = " & ".join(names)
        logger.info(f"SINGLE mode: Successfully extracted names from CLIENT FILE: {names}")
        
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"SINGLE mode FAILED: Error parsing client file: {e}")
        raise HTTPException(status_code=400, detail=f"Failed to parse client file: {str(e)}")
    
    logger.info(f"SINGLE mode - Using travellers: {traveller_names} (from client file)")
    
    # Generate vouchers
    generator = VoucherGenerator(TEMPLATE_PATH)
    vouchers_dir = os.path.join(temp_dir, "vouchers")
    
    vouchers = generator.generate_all(
        parsed_data=parsed_data,
        traveller_names=traveller_names,
        ref_no=ref_no,
        output_dir=vouchers_dir
    )
    
    if not vouchers:
        raise HTTPException(status_code=400, detail="No services found in ORGA file")
    
    logger.info(f"Generated {len(vouchers)} vouchers")
    
    # Convert to PDF and merge
    output_dir = os.path.join(temp_dir, "output")
    safe_name = "".join(c for c in traveller_names if c.isalnum() or c in (' ', '-', '_', '&'))
    safe_name = safe_name.replace(' ', '_').replace('&', '_')[:30]
    final_pdf_name = f"{trip_id}_{safe_name}_Vouchers.pdf"
    
    try:
        final_pdf_path = process_vouchers_to_pdf(vouchers, output_dir, final_pdf_name)
    except Exception as e:
        logger.error(f"PDF conversion failed: {e}")
        raise HTTPException(status_code=500, detail=f"PDF conversion failed: {str(e)}")
    
    return FileResponse(
        path=final_pdf_path,
        filename=final_pdf_name,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{final_pdf_name}"'}
    )


async def generate_group_mode(parsed_data, client_path: str, temp_dir: str, trip_id: str, ref_no: str):
    """Generate vouchers for GROUP mode - one PDF per room.
    
    CRITICAL: Names are ONLY extracted from the uploaded GROUP client file.
    - Never from ORGA file
    - Never from ORGA filename
    - Never guessed or generated as fallback
    - If parsing fails, return error
    """
    # Parse room assignments from GROUP CLIENT FILE ONLY - NO FALLBACK
    logger.info(f"GROUP mode: Parsing names from GROUP CLIENT FILE: {os.path.basename(client_path)}")
    logger.info("NOTE: Names are extracted from group client Excel ONLY, never from ORGA")
    
    try:
        rooms = parse_group_client_file(client_path)
        
        # STRICT: If no rooms found, FAIL - do not use any fallback
        if not rooms:
            logger.error("GROUP mode FAILED: No room assignments found in group client file")
            raise HTTPException(
                status_code=400, 
                detail="FAILED: Could not extract room assignments from GROUP client file. "
                       "The file must have columns 'Room', 'Last Name', 'First Name' with valid client data. "
                       "Names are NOT taken from ORGA - they MUST be in the group booking Excel file."
            )
        
        logger.info(f"GROUP mode: Successfully extracted {len(rooms)} rooms from GROUP CLIENT FILE")
        for room in rooms:
            logger.info(f"  Room {room.room_number}: {room.occupants} (from group client file)")
            
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"GROUP mode FAILED: Error parsing group client file: {e}")
        raise HTTPException(status_code=400, detail=f"Failed to parse group client file: {str(e)}")
    
    logger.info(f"GROUP mode - {len(rooms)} rooms found")
    
    # Generate PDFs for each room
    pdf_files = []
    
    for room in rooms:
        traveller_names = room.get_names_display()
        
        # Validate names are not empty
        if not traveller_names or len(traveller_names.strip()) < 2:
            logger.error(f"Room {room.room_number} has empty or invalid names - skipping")
            continue
        
        logger.info(f"Processing Room {room.room_number}: {traveller_names}")
        
        # Generate vouchers for this room
        generator = VoucherGenerator(TEMPLATE_PATH)
        room_vouchers_dir = os.path.join(temp_dir, f"room_{room.room_number}")
        
        vouchers = generator.generate_all(
            parsed_data=parsed_data,
            traveller_names=traveller_names,
            ref_no=ref_no,
            output_dir=room_vouchers_dir
        )
        
        if not vouchers:
            logger.warning(f"No vouchers generated for Room {room.room_number}")
            continue
        
        # Convert to PDF
        room_output_dir = os.path.join(temp_dir, f"output_room_{room.room_number}")
        safe_names = room.get_filename_safe()
        pdf_name = f"{trip_id}_{safe_names}.pdf"
        
        try:
            pdf_path = process_vouchers_to_pdf(vouchers, room_output_dir, pdf_name)
            pdf_files.append((pdf_path, pdf_name))
        except Exception as e:
            logger.error(f"PDF conversion failed for room {room.room_number}: {e}")
            continue
    
    if not pdf_files:
        raise HTTPException(status_code=500, detail="Failed to generate any PDFs")
    
    # If only one room, return single PDF
    if len(pdf_files) == 1:
        path, name = pdf_files[0]
        return FileResponse(
            path=path,
            filename=name,
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{name}"'}
        )
    
    # Multiple rooms - create ZIP
    zip_name = f"{trip_id}_Group_Vouchers.zip"
    zip_path = os.path.join(temp_dir, zip_name)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for pdf_path, pdf_name in pdf_files:
            zf.write(pdf_path, pdf_name)
    
    logger.info(f"Created ZIP with {len(pdf_files)} PDFs: {zip_path}")
    
    return FileResponse(
        path=zip_path,
        filename=zip_name,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'}
    )


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {
        "status": "healthy",
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "version": "2.0.0"
    }
