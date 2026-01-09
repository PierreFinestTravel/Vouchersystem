"""PDF conversion and merging module.

This module handles converting Word documents to PDF and merging multiple PDFs
into a single final document.

Supports multiple conversion methods:
1. docx2pdf (uses Microsoft Word on Windows) - preferred
2. LibreOffice headless mode - fallback
"""
import logging
import os
import subprocess
import platform
import tempfile
from typing import List, Tuple
from datetime import date
from pathlib import Path

from PyPDF2 import PdfMerger, PdfReader

logger = logging.getLogger(__name__)

# Track which conversion method is available
_conversion_method = None


def check_docx2pdf_available() -> bool:
    """Check if docx2pdf (Microsoft Word) is available."""
    try:
        import docx2pdf
        # On Windows, check if Word is installed by trying to import win32com
        if platform.system() == "Windows":
            import win32com.client
            return True
    except ImportError:
        pass
    return False


def find_libreoffice() -> str:
    """Find LibreOffice installation path."""
    system = platform.system()
    
    if system == "Windows":
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            r"C:\Program Files\LibreOffice\program\soffice.com",
        ]
    elif system == "Darwin":  # macOS
        possible_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ]
    else:  # Linux
        possible_paths = [
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/usr/local/bin/libreoffice",
        ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # Try to find in PATH
    try:
        result = subprocess.run(
            ["which", "libreoffice"] if system != "Windows" else ["where", "libreoffice"],
            capture_output=True,
            text=True
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip().split('\n')[0]
    except Exception:
        pass
    
    return None


def get_conversion_method() -> str:
    """Determine which PDF conversion method to use."""
    global _conversion_method
    
    if _conversion_method is not None:
        return _conversion_method
    
    # Try docx2pdf first (uses Microsoft Word on Windows)
    if check_docx2pdf_available():
        _conversion_method = "docx2pdf"
        logger.info("Using docx2pdf (Microsoft Word) for PDF conversion")
        return _conversion_method
    
    # Try LibreOffice
    if find_libreoffice():
        _conversion_method = "libreoffice"
        logger.info("Using LibreOffice for PDF conversion")
        return _conversion_method
    
    _conversion_method = "none"
    return _conversion_method


def convert_docx_to_pdf_with_word(docx_path: str, output_dir: str) -> str:
    """Convert Word document to PDF using Microsoft Word via docx2pdf."""
    from docx2pdf import convert
    
    os.makedirs(output_dir, exist_ok=True)
    
    docx_filename = os.path.basename(docx_path)
    pdf_filename = os.path.splitext(docx_filename)[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    logger.info(f"Converting to PDF (Word): {docx_path}")
    
    try:
        convert(docx_path, pdf_path)
    except Exception as e:
        raise RuntimeError(f"PDF conversion with Word failed: {str(e)}")
    
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF file was not created: {pdf_path}")
    
    return pdf_path


def convert_docx_to_pdf_with_libreoffice(docx_path: str, output_dir: str) -> str:
    """Convert Word document to PDF using LibreOffice."""
    libreoffice_path = find_libreoffice()
    
    if not libreoffice_path:
        raise RuntimeError("LibreOffice not found")
    
    os.makedirs(output_dir, exist_ok=True)
    
    cmd = [
        libreoffice_path,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ]
    
    logger.info(f"Converting to PDF (LibreOffice): {docx_path}")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60
        )
        
        if result.returncode != 0:
            logger.error(f"LibreOffice error: {result.stderr}")
            raise RuntimeError(f"PDF conversion failed: {result.stderr}")
        
    except subprocess.TimeoutExpired:
        raise RuntimeError("PDF conversion timed out")
    
    docx_filename = os.path.basename(docx_path)
    pdf_filename = os.path.splitext(docx_filename)[0] + ".pdf"
    pdf_path = os.path.join(output_dir, pdf_filename)
    
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF file was not created: {pdf_path}")
    
    return pdf_path


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """Convert a Word document to PDF.
    
    Automatically selects the best available method:
    1. Microsoft Word (via docx2pdf) - preferred on Windows
    2. LibreOffice headless mode - fallback
    
    Returns the path to the generated PDF file.
    """
    method = get_conversion_method()
    
    if method == "docx2pdf":
        return convert_docx_to_pdf_with_word(docx_path, output_dir)
    elif method == "libreoffice":
        return convert_docx_to_pdf_with_libreoffice(docx_path, output_dir)
    else:
        raise RuntimeError(
            "No PDF conversion method available. Please install either:\n"
            "1. Microsoft Word (recommended for Windows)\n"
            "2. LibreOffice (https://www.libreoffice.org/download/)"
        )


def convert_all_to_pdf(
    vouchers: List[Tuple[str, str, date]],
    output_dir: str
) -> List[Tuple[str, str, date]]:
    """Convert all voucher documents to PDF.
    
    Args:
        vouchers: List of (docx_path, voucher_type, date) tuples
        output_dir: Directory to store PDF files
        
    Returns:
        List of (pdf_path, voucher_type, date) tuples
    """
    pdf_vouchers = []
    
    for docx_path, voucher_type, voucher_date in vouchers:
        try:
            pdf_path = convert_docx_to_pdf(docx_path, output_dir)
            pdf_vouchers.append((pdf_path, voucher_type, voucher_date))
        except Exception as e:
            logger.error(f"Failed to convert {docx_path}: {e}")
            raise
    
    return pdf_vouchers


def sort_vouchers(vouchers: List[Tuple[str, str, date]]) -> List[Tuple[str, str, date]]:
    """Sort vouchers by type priority, then by date.
    
    Order: Hotels, Transfers, Car Rental, Activities, Restaurants, Golf
    """
    type_priority = {
        "hotel": 1,
        "transfer": 2,
        "car_rental": 3,
        "activity": 4,
        "restaurant": 5,
        "golf": 6,
    }
    
    return sorted(vouchers, key=lambda v: (type_priority.get(v[1], 99), v[2]))


def merge_pdfs(
    pdf_files: List[str],
    output_path: str
) -> str:
    """Merge multiple PDF files into a single PDF.
    
    Args:
        pdf_files: List of paths to PDF files
        output_path: Path for the merged output PDF
        
    Returns:
        Path to the merged PDF file
    """
    logger.info(f"Merging {len(pdf_files)} PDFs into {output_path}")
    
    merger = PdfMerger()
    
    for pdf_file in pdf_files:
        if os.path.exists(pdf_file):
            merger.append(pdf_file)
        else:
            logger.warning(f"PDF file not found, skipping: {pdf_file}")
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    merger.write(output_path)
    merger.close()
    
    if not os.path.exists(output_path):
        raise RuntimeError(f"Failed to create merged PDF: {output_path}")
    
    logger.info(f"Successfully created merged PDF: {output_path}")
    return output_path


def process_vouchers_to_pdf(
    vouchers: List[Tuple[str, str, date]],
    output_dir: str,
    final_pdf_name: str = "Travel_Vouchers.pdf"
) -> str:
    """Process all vouchers: convert to PDF, sort, and merge.
    
    Args:
        vouchers: List of (docx_path, voucher_type, date) tuples
        output_dir: Working directory for intermediate files
        final_pdf_name: Name for the final merged PDF
        
    Returns:
        Path to the final merged PDF
    """
    if not vouchers:
        raise ValueError("No vouchers to process")
    
    # Create subdirectory for PDFs
    pdf_dir = os.path.join(output_dir, "pdfs")
    os.makedirs(pdf_dir, exist_ok=True)
    
    # Convert all to PDF
    logger.info(f"Converting {len(vouchers)} vouchers to PDF...")
    pdf_vouchers = convert_all_to_pdf(vouchers, pdf_dir)
    
    # Sort by type and date
    sorted_vouchers = sort_vouchers(pdf_vouchers)
    
    # Get just the PDF paths in sorted order
    sorted_pdf_paths = [v[0] for v in sorted_vouchers]
    
    # Merge into final PDF
    final_path = os.path.join(output_dir, final_pdf_name)
    merge_pdfs(sorted_pdf_paths, final_path)
    
    return final_path


def process_vouchers_to_zip(
    vouchers: List[Tuple[str, str, date]],
    output_dir: str,
    final_zip_name: str = "Travel_Vouchers.zip"
) -> str:
    """Process all vouchers: sort and package as DOCX files in a ZIP.
    
    NO PDF conversion - returns original .docx files.
    This is much faster as it skips LibreOffice conversion.
    
    Args:
        vouchers: List of (docx_path, voucher_type, date) tuples
        output_dir: Working directory for output
        final_zip_name: Name for the final ZIP file
        
    Returns:
        Path to the final ZIP file
    """
    import zipfile
    import shutil
    
    if not vouchers:
        raise ValueError("No vouchers to process")
    
    os.makedirs(output_dir, exist_ok=True)
    
    # Sort by type and date
    sorted_vouchers = sort_vouchers(vouchers)
    
    # Create ZIP file with sorted DOCX files
    final_path = os.path.join(output_dir, final_zip_name)
    
    logger.info(f"Packaging {len(vouchers)} vouchers into ZIP (no PDF conversion)...")
    
    with zipfile.ZipFile(final_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for idx, (docx_path, voucher_type, voucher_date) in enumerate(sorted_vouchers, 1):
            if os.path.exists(docx_path):
                # Name files with order number for easy sorting
                original_name = os.path.basename(docx_path)
                # Add index prefix to maintain sort order
                new_name = f"{idx:02d}_{original_name}"
                zf.write(docx_path, new_name)
                logger.info(f"Added to ZIP: {new_name}")
            else:
                logger.warning(f"DOCX file not found, skipping: {docx_path}")
    
    if not os.path.exists(final_path):
        raise RuntimeError(f"Failed to create ZIP: {final_path}")
    
    logger.info(f"Successfully created ZIP with {len(vouchers)} vouchers: {final_path}")
    return final_path


def merge_docx_files(
    vouchers: List[Tuple[str, str, date]],
    output_path: str
) -> str:
    """Merge multiple DOCX files into a single document.
    
    Args:
        vouchers: List of (docx_path, voucher_type, date) tuples
        output_path: Path for the merged output DOCX
        
    Returns:
        Path to the merged DOCX file
    """
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    if not vouchers:
        raise ValueError("No vouchers to merge")
    
    # Sort by type and date
    sorted_vouchers = sort_vouchers(vouchers)
    
    logger.info(f"Merging {len(sorted_vouchers)} vouchers into single DOCX...")
    
    # Start with the first document as base
    first_docx_path = sorted_vouchers[0][0]
    merged_doc = Document(first_docx_path)
    
    # Append remaining documents
    for docx_path, voucher_type, voucher_date in sorted_vouchers[1:]:
        if not os.path.exists(docx_path):
            logger.warning(f"DOCX file not found, skipping: {docx_path}")
            continue
        
        # Add page break before each new voucher
        merged_doc.add_page_break()
        
        # Open the document to append
        sub_doc = Document(docx_path)
        
        # Copy all elements from sub_doc to merged_doc
        for element in sub_doc.element.body:
            merged_doc.element.body.append(element)
    
    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Save merged document
    merged_doc.save(output_path)
    
    if not os.path.exists(output_path):
        raise RuntimeError(f"Failed to create merged DOCX: {output_path}")
    
    logger.info(f"Successfully created merged DOCX with {len(sorted_vouchers)} vouchers: {output_path}")
    return output_path


def process_vouchers_to_single_docx(
    vouchers: List[Tuple[str, str, date]],
    output_dir: str,
    final_docx_name: str = "Travel_Vouchers.docx"
) -> str:
    """Process all vouchers: sort and merge into a single DOCX file.
    
    Args:
        vouchers: List of (docx_path, voucher_type, date) tuples
        output_dir: Working directory for output
        final_docx_name: Name for the final merged DOCX
        
    Returns:
        Path to the final merged DOCX file
    """
    if not vouchers:
        raise ValueError("No vouchers to process")
    
    os.makedirs(output_dir, exist_ok=True)
    
    final_path = os.path.join(output_dir, final_docx_name)
    
    return merge_docx_files(vouchers, final_path)