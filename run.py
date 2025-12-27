"""Run the Travel Voucher Generator application."""
import uvicorn
import os
import sys
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))


def main():
    """Run the FastAPI application."""
    # Check if template exists
    template_path = project_root / "templates" / "_Voucher blank.docx"
    if not template_path.exists():
        print("=" * 60)
        print("WARNING: Voucher template not found!")
        print(f"Expected location: {template_path}")
        print()
        print("Please copy '_Voucher blank.docx' to the 'templates' folder.")
        print("=" * 60)
        print()
    
    # Check for PDF conversion method
    from app.pdf_merger import get_conversion_method, find_libreoffice
    method = get_conversion_method()
    
    if method == "docx2pdf":
        print("PDF conversion: Using Microsoft Word (docx2pdf)")
    elif method == "libreoffice":
        print(f"PDF conversion: Using LibreOffice ({find_libreoffice()})")
    else:
        print("=" * 60)
        print("WARNING: No PDF conversion method available!")
        print()
        print("Please install one of the following:")
        print("1. Microsoft Word (recommended for Windows)")
        print("2. LibreOffice: https://www.libreoffice.org/download/")
        print("=" * 60)
        print()
    
    print()
    print("Starting Travel Voucher Generator...")
    print("Open http://localhost:8000 in your browser")
    print("Press Ctrl+C to stop the server")
    print()
    
    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )


if __name__ == "__main__":
    main()

