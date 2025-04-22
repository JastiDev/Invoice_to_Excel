# Invoice PDF to Excel Converter

A Python application that converts PDF invoices into structured Excel spreadsheets using OCR technology. The application features a user-friendly GUI interface and supports both text-based and scanned PDF invoices.

## Features

- **PDF Text Extraction**: Supports both native text-based PDFs and scanned documents
- **OCR Capabilities**: Uses Tesseract OCR for processing scanned invoices
- **Smart Data Parsing**: Extracts key invoice information including:
  - Purchase quantities
  - Product codes
  - Brand names
  - Product descriptions
  - Cost per packet
  - Total cost
  - Unit cost calculations
- **Excel Output**:
  - Automatically formatted Excel spreadsheets
  - Currency formatting for cost columns
  - Auto-adjusted column widths
  - Timestamp-based file naming
- **User-Friendly GUI**:
  - Simple file selection interface
  - Progress feedback
  - Error handling and notifications

## Requirements

- Python 3.x
- Required Python packages:
  - pandas
  - pdfplumber
  - pytesseract
  - Pillow (PIL)
  - openpyxl
- Tesseract OCR engine installed on your system
  - Windows: Install from [Tesseract GitHub Releases](https://github.com/UB-Mannheim/tesseract/wiki)
  - Default path: `C:\Program Files\Tesseract-OCR\tesseract.exe`

## Installation

1. Clone this repository or download the source code
2. Install required Python packages:

```bash
pip install pandas pdfplumber pytesseract Pillow openpyxl
```

3. Install Tesseract OCR engine for your operating system
4. Ensure Tesseract is added to your system PATH

## Usage

1. Run the application:

```bash
python main.py
```

2. Using the GUI:
   - Click "Browse..." to select your input PDF invoice
   - Choose a save location for the Excel output file
   - Click "Process PDF" to start the conversion
   - Wait for the success message

## Output Format

The generated Excel file will contain the following columns:

- Purchased
- Received
- Code1 (CAS/PK/BAG)
- Code2 (Product Code)
- Brand
- Description
- CostPerPacket
- TotalCost
- BarInParanthesis
- UnitCost

## Error Handling

The application includes robust error handling for:

- Invalid PDF files
- OCR processing issues
- Data parsing errors
- File access permissions
- Missing Tesseract installation

## Notes

- The application is optimized for specific invoice formats
- OCR accuracy may vary depending on the quality of scanned documents
- Make sure you have appropriate permissions for input/output file locations

## License

This project is licensed under the MIT License - see the LICENSE file for details.
