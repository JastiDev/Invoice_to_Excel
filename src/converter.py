"""
Main converter module that ties all components together
"""
from src.pdf_extraction.extractor import extract_text_from_pdf
from src.text_processing.processor import parse_invoice_text
from src.excel_output.export import export_to_excel
from src.gui.app import create_gui


def invoice_pdf_to_excel(pdf_path, output_excel_path, log_callback=None):
    """
    Main function to process PDF and export to Excel
    
    Args:
        pdf_path (str): Path to the PDF file
        output_excel_path (str): Path to save Excel file
        log_callback (function, optional): Callback for logging
        
    Returns:
        bool: True if successful, False otherwise
    """
    # Step 1: Extract text from PDF
    if log_callback:
        log_callback("Extracting text from PDF...")
    text = extract_text_from_pdf(pdf_path)
    
    # Step 2: Parse the text to extract structured data
    if log_callback:
        log_callback("Parsing extracted text...")
    invoice_items = parse_invoice_text(text)
    
    # Step 3: Create DataFrame and export to Excel
    if invoice_items:
        if log_callback:
            log_callback(f"Found {len(invoice_items)} items in the invoice")
        
        if log_callback:
            log_callback("Applying Excel formatting...")
        
        success = export_to_excel(invoice_items, output_excel_path)
        
        if success and log_callback:
            log_callback(f"Data successfully exported to {output_excel_path}")
            log_callback(f"Number of records processed: {len(invoice_items)}")
        return success
    else:
        if log_callback:
            log_callback("No invoice data could be extracted from the PDF.")
        return False


def run_application():
    """
    Run the GUI application
    """
    root = create_gui(invoice_pdf_to_excel)
    root.mainloop()


if __name__ == "__main__":
    run_application() 