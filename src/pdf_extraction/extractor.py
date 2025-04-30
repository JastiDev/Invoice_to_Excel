"""
PDF extraction module for extracting text from PDF files
"""
import pdfplumber  # For text extraction from PDF
import pytesseract  # For OCR if PDF is scanned
from PIL import Image  # For handling image data
import io


def extract_text_from_pdf(pdf_path):
    """
    Extract text from PDF (works for both text-based and scanned PDFs)
    
    Args:
        pdf_path (str): Path to the PDF file
        
    Returns:
        str: Extracted text from PDF
    """
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Try to extract text directly (works for text-based PDFs)
            page_text = page.extract_text()
            if page_text:
                text += f"\n=== Page {page_num} ===\n" + page_text + "\n"
            else:
                # If no text found, it's likely a scanned PDF - use OCR
                img = page.to_image(resolution=300).original
                # Convert to grayscale and enhance contrast
                img = img.convert('L')
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr = img_byte_arr.getvalue()
                # Configure OCR for better recognition
                custom_config = '--oem 3 --psm 6'
                page_text = pytesseract.image_to_string(Image.open(io.BytesIO(img_byte_arr)), config=custom_config)
                text += f"\n=== Page {page_num} (OCR) ===\n" + page_text + "\n"
    return text 