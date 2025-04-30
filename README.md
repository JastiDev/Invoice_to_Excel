# Invoice to Excel Converter

A Python application that converts PDF invoices to Excel format, with support for both text-based and scanned PDFs. The application includes OCR capabilities for processing scanned documents and features a modern GUI interface with real-time processing logs.

## Features

- Converts PDF invoices to Excel format
- Supports both text-based and scanned PDFs using OCR
- Smart data extraction with OCR error correction:
  - Handles common OCR mistakes in numbers and text
  - Corrects product codes automatically
  - Normalizes unit measurements (oz, lb, pc)
- Extracts key information including:
  - Purchase and received quantities
  - Product codes (CAS/PK/BAG)
  - Brand information
  - Product descriptions
  - Cost per packet
  - Total cost
  - Unit cost calculations
- Modern GUI interface with:
  - Real-time processing logs
  - File selection dialogs
  - Progress tracking
  - Error handling
- Automatic Excel formatting with currency formatting
- Modular architecture for maintainability and extensibility

## Project Structure

The codebase is organized into modules, each with specific responsibilities:

```
Invoice_to_Excel/
├── main.py                  # Entry point for the application
├── src/                     # Source code directory
│   ├── converter.py         # Main converter logic that ties modules together
│   ├── pdf_extraction/      # PDF text extraction module
│   │   └── extractor.py     # Functions for extracting text from PDFs
│   ├── text_processing/     # Text processing module
│   │   └── processor.py     # Functions for cleaning and parsing invoice text
│   ├── excel_output/        # Excel export module
│   │   └── export.py        # Functions for formatting and exporting to Excel
│   └── gui/                 # GUI module
│       └── app.py           # User interface implementation
├── requirements.txt         # Project dependencies
└── README.md                # This file
```

## Prerequisites

### For Users (Running the EXE)

1. Windows Operating System
2. [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) installed (minimum version 5.0.0)
   - Install to the default location: `C:\Program Files\Tesseract-OCR`
   - Add Tesseract to your system PATH
   - Verify installation by running `tesseract --version` in command prompt

### For Developers (Running from Source)

1. Python 3.x
2. Required Python packages (install using `pip install -r requirements.txt`):
   - pandas (>=2.2.3): Data manipulation and Excel export
   - pdfplumber (>=0.11.6): PDF text extraction
   - pytesseract (>=0.3.13): OCR processing
   - Pillow (>=11.2.1): Image processing
   - openpyxl (>=3.1.5): Excel file creation
   - python-dateutil (>=2.9.0): Date handling
   - pyinstaller (>=6.13.0): For creating executable

## Installation

### Option 1: Running the Executable

1. Download `Invoice_to_Excel.exe` from the `dist` folder
2. Install Tesseract OCR:
   - Download from [Tesseract GitHub Releases](https://github.com/UB-Mannheim/tesseract/wiki)
   - Run installer and choose default location
   - Add to system PATH during installation
3. Double-click the executable to run

### Option 2: Running from Source

1. Clone this repository
2. Install Python 3.x
3. Install Tesseract OCR as described above
4. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```
5. Run the application:
   ```bash
   python main.py
   ```

## Usage

1. Launch the application
2. Click "Browse..." to select your PDF invoice
3. Choose where to save the Excel output file
4. Click "Process PDF" to start the conversion
5. Monitor progress in the log window
6. Excel file will be created with formatted data

## Excel Output Format

The generated Excel file will contain the following columns:

- Purchased: Quantity purchased
- Received: Quantity received
- Code1: Primary product code (CAS/PK/BAG)
- Code2: Secondary product code
- Brand: Product brand
- Description: Product type/category
- Product: Full product description
- CostPerPacket: Cost per packet (currency formatted)
- TotalCost: Total cost (currency formatted)
- BarInParanthesis: Units per packet
- UnitCost: Cost per unit (currency formatted)
- Tentative: Calculated tentative price (currency formatted)

## Development

### Module Descriptions

1. **PDF Extraction Module** (`src/pdf_extraction/extractor.py`):

   - Extracts text from PDF files
   - Uses pdfplumber for text-based PDFs
   - Uses pytesseract OCR for scanned PDFs

2. **Text Processing Module** (`src/text_processing/processor.py`):

   - Cleans and parses OCR text
   - Handles detection of invoice items, quantities, codes, prices
   - Contains OCR error correction logic

3. **Excel Output Module** (`src/excel_output/export.py`):

   - Formats and exports data to Excel
   - Applies proper column formatting and width adjustments

4. **GUI Module** (`src/gui/app.py`):

   - Implements the user interface
   - Provides file selection, processing status, and log display

5. **Converter Module** (`src/converter.py`):
   - Ties all components together
   - Orchestrates the conversion process

### Extending the Application

To add new features:

1. Identify which module should contain the functionality
2. Implement your feature in the appropriate module
3. Update the converter module if needed to integrate your changes
4. Test thoroughly

## Building the Executable

To create the executable yourself:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=NONE --name="Invoice_to_Excel" main.py
```

The executable will be created in the `dist` directory.

## Troubleshooting

1. **Tesseract Error**:

   - Verify Tesseract is installed in `C:\Program Files\Tesseract-OCR`
   - Check if Tesseract is in system PATH
   - Run `tesseract --version` to verify installation

2. **PDF Not Reading**:

   - Ensure the PDF is not password protected
   - Check if the PDF is readable (try opening in a PDF viewer)
   - For scanned PDFs, ensure good image quality

3. **Excel File Issues**:

   - Check if the output Excel file is not already open
   - Verify you have write permissions in the output directory
   - Ensure enough disk space is available

4. **OCR Quality Issues**:
   - Ensure PDF scan quality is good
   - Check if the PDF is properly oriented
   - Verify Tesseract installation is complete with all language packs

## Support

For issues and feature requests, please create an issue in the repository.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
