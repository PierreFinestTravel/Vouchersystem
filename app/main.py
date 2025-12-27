"""Main FastAPI application for the Travel Voucher Generator."""
import logging
import os
import tempfile
import shutil
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles

from .orga_parser import parse_orga
from .voucher_generator import VoucherGenerator
from .pdf_merger import process_vouchers_to_pdf

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
    version="1.0.0"
)

# Get the base directory
BASE_DIR = Path(__file__).resolve().parent.parent

# Template path - can be configured via environment variable
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
            max-width: 680px;
            margin: 0 auto;
            padding: 40px 24px;
            position: relative;
            z-index: 1;
        }

        header {
            text-align: center;
            margin-bottom: 48px;
        }

        .logo {
            width: 120px;
            height: 120px;
            margin: 0 auto 24px;
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
            width: 60px;
            height: 60px;
            fill: var(--primary);
        }

        h1 {
            font-family: 'Cormorant Garamond', serif;
            font-size: 2.5rem;
            font-weight: 600;
            color: var(--gold);
            letter-spacing: 2px;
            margin-bottom: 8px;
        }

        .subtitle {
            font-size: 0.9rem;
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
            padding: 40px;
            box-shadow: var(--shadow);
        }

        .form-group {
            margin-bottom: 28px;
        }

        label {
            display: block;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 1.5px;
            color: var(--gold);
            margin-bottom: 10px;
        }

        label .required {
            color: #ff6b6b;
            margin-left: 4px;
        }

        input[type="text"],
        input[type="file"] {
            width: 100%;
            padding: 16px 20px;
            background: rgba(0,0,0,0.3);
            border: 1px solid rgba(201, 169, 98, 0.3);
            border-radius: 10px;
            font-family: 'Montserrat', sans-serif;
            font-size: 1rem;
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
            padding: 10px 20px;
            margin-right: 16px;
            background: linear-gradient(135deg, var(--gold) 0%, var(--gold-light) 100%);
            border: none;
            border-radius: 6px;
            font-family: 'Montserrat', sans-serif;
            font-size: 0.85rem;
            font-weight: 500;
            color: var(--primary);
            cursor: pointer;
            transition: all 0.3s ease;
        }

        input[type="file"]::file-selector-button:hover {
            transform: scale(1.02);
        }

        .btn-submit {
            width: 100%;
            padding: 18px 32px;
            background: linear-gradient(135deg, var(--gold) 0%, var(--gold-light) 100%);
            border: none;
            border-radius: 10px;
            font-family: 'Montserrat', sans-serif;
            font-size: 1rem;
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

        .btn-submit:active {
            transform: translateY(0);
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
            margin-top: 24px;
            padding: 16px 20px;
            border-radius: 10px;
            font-size: 0.9rem;
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

        .download-link {
            display: inline-block;
            margin-top: 12px;
            padding: 12px 24px;
            background: var(--gold);
            color: var(--primary);
            text-decoration: none;
            font-weight: 600;
            border-radius: 8px;
            transition: all 0.3s ease;
        }

        .download-link:hover {
            background: var(--gold-light);
            transform: scale(1.02);
        }

        .info-text {
            font-size: 0.8rem;
            color: rgba(245, 240, 230, 0.5);
            margin-top: 8px;
        }

        .divider {
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(201, 169, 98, 0.3), transparent);
            margin: 32px 0;
        }

        footer {
            text-align: center;
            margin-top: 40px;
            font-size: 0.8rem;
            color: rgba(245, 240, 230, 0.4);
        }

        @media (max-width: 600px) {
            .container {
                padding: 24px 16px;
            }
            
            h1 {
                font-size: 1.8rem;
            }
            
            .card {
                padding: 28px 20px;
            }
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
            <p class="subtitle">Voucher Generator</p>
        </header>

        <div class="card">
            <form id="voucherForm" enctype="multipart/form-data">
                <div class="form-group">
                    <label>Traveller Names<span class="required">*</span></label>
                    <input type="text" name="traveller_names" id="traveller_names" 
                           placeholder="e.g., Mr John Smith & Mrs Jane Smith" required>
                    <p class="info-text">Enter names exactly as they should appear on vouchers</p>
                </div>

                <div class="form-group">
                    <label>Reference Number</label>
                    <input type="text" name="ref_no" id="ref_no" 
                           placeholder="e.g., 15123">
                </div>

                <div class="form-group">
                    <label>Group / Pax Info</label>
                    <input type="text" name="group_text" id="group_text" 
                           placeholder="e.g., 2 Adults">
                </div>

                <div class="divider"></div>

                <div class="form-group">
                    <label>ORGA Excel File<span class="required">*</span></label>
                    <input type="file" name="orga_file" id="orga_file" 
                           accept=".xlsx,.xls" required>
                    <p class="info-text">Upload the ORGA Excel file (.xlsx)</p>
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

        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            // Validate required fields
            const travellerNames = document.getElementById('traveller_names').value.trim();
            const orgaFile = document.getElementById('orga_file').files[0];
            
            if (!travellerNames) {
                showMessage('Please enter traveller names', 'error');
                return;
            }
            
            if (!orgaFile) {
                showMessage('Please select an ORGA Excel file', 'error');
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
                    // Download the PDF
                    const blob = await response.blob();
                    const contentDisposition = response.headers.get('Content-Disposition');
                    let filename = 'Travel_Vouchers.pdf';
                    
                    if (contentDisposition) {
                        const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
                        if (filenameMatch) {
                            filename = filenameMatch[1];
                        }
                    }
                    
                    // Create download link
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    
                    showMessage('Vouchers generated successfully! Your download should start automatically.', 'success');
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
    traveller_names: str = Form(..., description="Traveller names (required)"),
    ref_no: str = Form("", description="Reference number (optional)"),
    group_text: str = Form("", description="Group/Pax info (optional)"),
    orga_file: UploadFile = File(..., description="ORGA Excel file")
):
    """Generate vouchers from ORGA Excel file.
    
    Accepts the ORGA Excel file and traveller details, generates all vouchers,
    and returns a merged PDF file.
    """
    logger.info(f"Received request: travellers='{traveller_names}', ref='{ref_no}'")
    
    # Validate traveller names
    if not traveller_names or not traveller_names.strip():
        raise HTTPException(status_code=400, detail="Traveller names are required")
    
    # Validate file type
    if not orga_file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Please upload an Excel file (.xlsx or .xls)")
    
    # Create temp directory for processing
    temp_dir = tempfile.mkdtemp(prefix="voucher_gen_")
    
    try:
        # Save uploaded file
        orga_path = os.path.join(temp_dir, orga_file.filename)
        with open(orga_path, "wb") as f:
            content = await orga_file.read()
            f.write(content)
        
        logger.info(f"Saved ORGA file to {orga_path}")
        
        # Check if template exists
        if not os.path.exists(TEMPLATE_PATH):
            raise HTTPException(
                status_code=500,
                detail=f"Voucher template not found. Please ensure the template is at: {TEMPLATE_PATH}"
            )
        
        # Parse ORGA file
        try:
            parsed_data = parse_orga(orga_path)
        except Exception as e:
            logger.error(f"Failed to parse ORGA: {e}")
            raise HTTPException(status_code=400, detail=f"Failed to parse ORGA file: {str(e)}")
        
        # Count total vouchers
        total_vouchers = (
            len(parsed_data.hotels) +
            len(parsed_data.transfers) +
            len(parsed_data.car_rentals) +
            len(parsed_data.activities) +
            len(parsed_data.restaurants) +
            len(parsed_data.golf)
        )
        
        if total_vouchers == 0:
            raise HTTPException(
                status_code=400,
                detail="No services found in the ORGA file. Please check the file format."
            )
        
        logger.info(f"Found {total_vouchers} services to generate vouchers for")
        logger.info(f"  Hotels: {len(parsed_data.hotels)}")
        logger.info(f"  Transfers: {len(parsed_data.transfers)}")
        logger.info(f"  Car Rentals: {len(parsed_data.car_rentals)}")
        logger.info(f"  Activities: {len(parsed_data.activities)}")
        logger.info(f"  Restaurants: {len(parsed_data.restaurants)}")
        logger.info(f"  Golf: {len(parsed_data.golf)}")
        
        # Generate vouchers
        generator = VoucherGenerator(TEMPLATE_PATH)
        vouchers_dir = os.path.join(temp_dir, "vouchers")
        
        try:
            vouchers = generator.generate_all(
                parsed_data=parsed_data,
                traveller_names=traveller_names.strip(),
                ref_no=ref_no.strip(),
                group_text=group_text.strip(),
                output_dir=vouchers_dir
            )
        except Exception as e:
            logger.error(f"Failed to generate vouchers: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to generate vouchers: {str(e)}")
        
        logger.info(f"Generated {len(vouchers)} voucher documents")
        
        # Convert to PDF and merge
        output_dir = os.path.join(temp_dir, "output")
        
        # Generate filename based on client name or traveller names
        safe_name = "".join(c for c in traveller_names if c.isalnum() or c in (' ', '-', '_'))
        safe_name = safe_name.strip().replace(' ', '_')[:30]
        final_pdf_name = f"{safe_name}_Travel_Vouchers.pdf"
        
        try:
            final_pdf_path = process_vouchers_to_pdf(
                vouchers=vouchers,
                output_dir=output_dir,
                final_pdf_name=final_pdf_name
            )
        except Exception as e:
            logger.error(f"Failed to create PDF: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to create PDF: {str(e)}")
        
        logger.info(f"Created final PDF: {final_pdf_path}")
        
        # Return the PDF file
        return FileResponse(
            path=final_pdf_path,
            filename=final_pdf_name,
            media_type="application/pdf",
            headers={
                "Content-Disposition": f'attachment; filename="{final_pdf_name}"'
            }
        )
        
    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Unexpected error during voucher generation")
        raise HTTPException(status_code=500, detail=f"An unexpected error occurred: {str(e)}")
    
    finally:
        # Note: We don't clean up temp_dir here because FileResponse needs the file
        # In production, you'd want to set up a cleanup job or use a streaming response
        pass


@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "template_exists": os.path.exists(TEMPLATE_PATH)}

