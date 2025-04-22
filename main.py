import pandas as pd
import pdfplumber  # For text extraction from PDF
import pytesseract  # For OCR if PDF is scanned
from PIL import Image  # For handling image data
import io
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os

# Function to extract text from PDF (works for both text-based and scanned PDFs)
def extract_text_from_pdf(pdf_path):
    text = ""
    print("\n" + "="*80)
    print("                               PDF TEXT EXTRACTION                             ")
    print("="*80)
    print(f"Processing PDF file: {pdf_path}")
    
    with pdfplumber.open(pdf_path) as pdf:
        print(f"\nTotal pages in PDF: {len(pdf.pages)}")
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"\nProcessing page {page_num}/{len(pdf.pages)}:")
            print("-"*40)
            
            # Try to extract text directly (works for text-based PDFs)
            page_text = page.extract_text()
            if page_text:
                print(f"✓ Successfully extracted text directly from page {page_num}")
                text += f"\n=== Page {page_num} ===\n" + page_text + "\n"
            else:
                print(f"! No direct text found on page {page_num}, attempting OCR...")
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
                print(f"✓ OCR processing completed for page {page_num}")
                text += f"\n=== Page {page_num} (OCR) ===\n" + page_text + "\n"
    
    print("\n" + "="*80)
    print("                               EXTRACTED TEXT                                  ")
    print("="*80)
    
    # Split text into lines and print with line numbers
    lines = text.split('\n')
    non_empty_lines = [line for line in lines if line.strip()]
    print(f"Total non-empty lines extracted: {len(non_empty_lines)}")
    print("-"*80)
    
    for i, line in enumerate(lines, 1):
        if line.strip():  # Only print non-empty lines
            if line.startswith('=== Page'):
                print(f"\n{line}")
            else:
                print(f"Line {i:3d}: {line}")
    
    print("\n" + "-"*80)
    print("                              END OF EXTRACTED TEXT                           ")
    print("="*80 + "\n")
    return text

# Function to clean OCR text
def clean_ocr_text(text):
    # Convert common OCR errors in numbers, but only for specific patterns
    # Only replace when these characters appear as standalone numbers
    replacements = {
        '.loz': 'oz',
        '.Lb': 'lb',
        '-': ' ',
        '*': '',
        '(': ' ',
        ')': ' ',
        '  ': ' '  # Replace double spaces with single space
    }
    
    for old, new in replacements.items():
        text = text.replace(old, new)
    
    # Remove multiple spaces
    text = ' '.join(text.split())
    return text

# Function to clean and normalize a line of text
def clean_line(line):
    # Special case for "ES) )" -> "55"
    line = line.replace('ES) )', '55')
    
    # Remove special characters that aren't needed
    line = line.replace('*', ' ')
    line = line.replace('-', ' ')
    line = line.replace(':', '.')
    line = line.replace('<', ' ')
    line = line.replace('>', ' ')
    line = line.replace('\\', ' ')
    line = line.replace('%', ' ')
    line = line.replace('$', ' ')
    line = line.replace('=', ' ')
    
    # Normalize spaces
    line = ' '.join(line.split())
    
    # Normalize units
    line = line.replace('.loz', 'oz').replace('.Lb', 'lb')
    line = line.replace('oz.', 'oz').replace('lb.', 'lb')
    
    return line

# Function to convert OCR text to number
def convert_to_number(text):
    # First handle the complete pattern to avoid double conversion
    if 'ES) )' in text:
        # text = text.replace('ES) )', '55')
        text = text.replace('ES)', '5')
    else:
        # Only if the complete pattern isn't found, handle individual cases
        text = text.replace('ES)', '5')
        # text = text.replace(')', '5')
    
    # Other special case conversions for known OCR errors
    text = text.replace('QO', '0')
    text = text.replace('al', '1')
    text = text.replace('aI', '1')
    text = text.replace('il', '1')
    text = text.replace('iI', '1')
    text = text.replace('ul', '1')
    text = text.replace('él', '1')
    text = text.replace('993', 'Q93')  # Special case for product code
    
    # Only convert when the text looks like it should be a number
    if any(c.isdigit() for c in text):
        # Convert common OCR errors in numbers
        number_text = text.replace('O', '0')\
                         .replace('o', '0')\
                         .replace('l', '1')\
                         .replace('i', '1')\
                         .replace('I', '1')\
                         .replace('Z', '2')\
                         .replace('z', '2')\
                         .replace('S', '5')\
                         .replace('E', '5')  # E -> 5 conversion
        # Keep only digits and decimal points
        number_text = ''.join(c for c in number_text if c.isdigit() or c == '.')
        try:
            return float(number_text)
        except ValueError:
            return None
    return None

# Function to parse extracted text and find the required data
def parse_invoice_text(text):
    items = []
    
    for line in text.split('\n'):
        # Skip empty lines and headers/footers
        if not line.strip() or any(x in line for x in ['CONTINUED', 'COPY', 'Free!', 'Suggested']):
            continue
        
        # Clean the line
        original_line = line
        line = clean_line(line)
        
        # Only process lines that look like item entries
        if not any(x in line for x in ['CAS', 'PK', 'BAG']):
            continue
            
        try:
            # Split the line into parts
            parts = line.split()
            
            # Find the index of CAS/PK/BAG
            code_index = -1
            for i, part in enumerate(parts):
                if part in ['CAS', 'PK', 'BAG']:
                    code_index = i
                    break
            
            if code_index == -1 or code_index + 1 >= len(parts):
                continue
                
            # Extract Purchased and Received quantities
            purchased = None
            received = 0  # Default to 0
            purchased_idx = -1
            
            # Handle special case for "ES) )" -> "5"
            if 'ES) )' in line:
                purchased = 5
                received = 5
            else:
                # First number before CAS/PK/BAG is purchased
                for i in range(code_index):
                    # Check for known OCR patterns first
                    if parts[i] in ['il', 'iL', 'al', 'aI']:
                        purchased = 1
                        purchased_idx = i
                        break
                    num = convert_to_number(parts[i])
                    if num is not None and num <= 100:  # Reasonable quantity check
                        purchased = int(num)
                        purchased_idx = i
                        break
                
                # Look for received quantity (O, 0, or a number)
                if purchased_idx != -1 and purchased_idx + 1 < len(parts):
                    next_part = parts[purchased_idx + 1]
                    if next_part in ['O', '0', 'QO', 'iL', 'il', 'al', 'aI']:  # Zero indicators
                        received = 0
                    else:
                        received_num = convert_to_number(next_part)
                        if received_num is not None:
                            received = int(received_num)
            
            if purchased is None:
                print(f"Could not find valid purchased quantity in line: {original_line}")
                continue
                
            code1 = parts[code_index]  # CAS/PK/BAG
            code2 = parts[code_index+1]  # Product code
            
            # Special case for product code 993 -> Q93
            if code2 == '993':
                code2 = 'Q93'
            
            # Find the cost per packet and total (should be the last two numbers)
            # But make sure they're reasonable (not extremely large numbers)
            numbers = []
            for part in reversed(parts):
                num = convert_to_number(part)
                if num is not None:
                    # Skip unreasonably large numbers (likely OCR errors)
                    if num < 1000:  # Assuming no item costs more than 1000
                        numbers.insert(0, num)
                        if len(numbers) == 2:
                            break
            
            if len(numbers) < 2:
                print(f"Could not find valid cost and total in line: {original_line}")
                continue
            
            cost_per_packet = numbers[0]  # First number is cost per packet
            # Calculate total cost based on received quantity
            total_cost = round(received * cost_per_packet, 2)
            
            # Extract the quantity in parentheses and full description
            bar = 0
            description_parts = []
            
            # Start collecting description after code2
            for part in parts[code_index+2:]:
                # Stop if we hit a cost number
                if convert_to_number(part) in numbers:
                    break
                    
                # Check for quantity in parentheses
                if '(' in part and ')' in part:
                    num_str = part[part.index('(')+1:part.index(')')]
                    num = convert_to_number(num_str)
                    if num is not None:
                        bar = int(num)
                
                description_parts.append(part)
            
            # Join all parts for full description
            full_description = ' '.join(description_parts)
            
            # Extract brand and description parts
            brand_pattern = r'(Deep|Bre|Mirch|Bansi|Britanni|Sujata|Chandan|Hem|MDH)'
            brand_match = re.search(brand_pattern, full_description)
            brand = brand_match.group(1) if brand_match else "Unknown"
            
            # Get description and product parts
            if brand_match:
                rest = full_description[brand_match.end():].strip()
                # Split on first period or space
                split_chars = ['.', ' ']
                split_index = len(rest)  # Default to end if no split char found
                for char in split_chars:
                    pos = rest.find(char)
                    if pos != -1 and pos < split_index:
                        split_index = pos
                
                description = rest[:split_index].strip()
                if split_index < len(rest):
                    product = rest[split_index + 1:].strip()
                else:
                    product = ""
                
                # If description contains a period, split there
                if '.' in description:
                    description = description.split('.')[0]
                
                # Clean up the product string
                product = product.strip()
                if not product and description:
                    # If no product but we have description with period
                    parts = rest.split('.')
                    if len(parts) > 1:
                        description = parts[0]
                        product = '.'.join(parts[1:])
            else:
                description = "Unknown"
                product = full_description
            
            # Create the item dictionary
            item = {
                'Purchased': purchased,
                'Received': received,
                'Code1': code1,
                'Code2': code2,
                'Brand': brand,
                'Description': description,  # Just the type (F S, Flo, Diges, etc)
                'Product': product,         # The actual product description
                'CostPerPacket': cost_per_packet,
                'TotalCost': total_cost,
                'BarInParanthesis': bar,
                'UnitCost': round(cost_per_packet / bar, 2) if bar > 0 else None
            }
            
            print(f"Match found for line: {original_line}")
            print(f"Extracted item: {item}\n")
            items.append(item)
            
        except (ValueError, IndexError) as e:
            print(f"Error processing line: {original_line}")
            print(f"Error: {str(e)}\n")
            continue
    
    return items

# Main function to process PDF and export to Excel
def invoice_pdf_to_excel(pdf_path, output_excel_path):
    # Step 1: Extract text from PDF
    text = extract_text_from_pdf(pdf_path)
    
    # Step 2: Parse the text to extract structured data
    invoice_items = parse_invoice_text(text)
    
    # Step 3: Create DataFrame and export to Excel
    if invoice_items:
        df = pd.DataFrame(invoice_items)
        
        print("\nDebug: Initial DataFrame columns:")
        print(df.columns.tolist())
        
        # Reorder columns to match the image format
        column_order = [
            'Purchased', 'Received', 'Code1', 'Code2', 'Brand', 'Description',
            'Product', 'CostPerPacket', 'TotalCost', 'BarInParanthesis', 'UnitCost'
        ]
        
        print("\nDebug: Checking for missing columns:")
        for col in column_order:
            if col not in df.columns:
                print(f"Missing column: {col}")
        
        # Make sure all columns exist before reordering
        for col in column_order:
            if col not in df.columns:
                df[col] = None
        
        df = df[column_order]
        
        print("\nDebug: Final DataFrame columns:")
        print(df.columns.tolist())
        
        # Apply Excel formatting
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoice Details')
            workbook = writer.book
            worksheet = writer.sheets['Invoice Details']
            
            # Format currency columns
            currency_columns = ['CostPerPacket', 'UnitCost', 'TotalCost']
            for col in currency_columns:
                col_idx = column_order.index(col) + 1  # +1 because Excel columns start at 1
                for row in range(2, len(df) + 2):  # +2 to account for header and 1-based indexing
                    cell = worksheet.cell(row=row, column=col_idx)
                    cell.number_format = '#,##0.00'
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

        print(f"\nData successfully exported to {output_excel_path}")
        print(f"Number of records processed: {len(df)}")
    else:
        print("No invoice data could be extracted from the PDF.")

# Create the main GUI window
def create_gui():
    def select_pdf():
        file_path = filedialog.askopenfilename(
            title="Select PDF Invoice",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file_path:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, file_path)
            
            # Auto-generate output path with timestamp
            directory = os.path.dirname(file_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_output = os.path.join(directory, f"invoice_data_{timestamp}.xlsx")
            output_entry.delete(0, tk.END)
            output_entry.insert(0, default_output)
    
    def select_save_location():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            initialfile=f"invoice_data_{timestamp}.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, file_path)
    
    def process_file():
        input_path = input_entry.get()
        output_path = output_entry.get()
        
        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input PDF and output Excel file location.")
            return
            
        try:
            # Make sure you have Tesseract OCR installed and in your PATH
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            
            # Process the file
            invoice_pdf_to_excel(input_path, output_path)
            
            messagebox.showinfo("Success", f"File processed successfully!\nSaved to: {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    
    # Create main window
    root = tk.Tk()
    root.title("Invoice PDF to Excel Converter")
    root.geometry("600x250")
    
    # Add padding and configure grid
    for i in range(4):
        root.grid_rowconfigure(i, pad=10)
    root.grid_columnconfigure(1, weight=1)
    
    # Input PDF selection
    tk.Label(root, text="Select PDF File:").grid(row=0, column=0, padx=5, sticky="e")
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=5, sticky="ew")
    tk.Button(root, text="Browse...", command=select_pdf).grid(row=0, column=2, padx=5)
    
    # Output Excel location
    tk.Label(root, text="Save Excel As:").grid(row=1, column=0, padx=5, sticky="e")
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=1, column=1, padx=5, sticky="ew")
    tk.Button(root, text="Browse...", command=select_save_location).grid(row=1, column=2, padx=5)
    
    # Process button
    process_button = tk.Button(root, text="Process PDF", command=process_file, width=20)
    process_button.grid(row=2, column=1, pady=20)
    
    # Status label
    status_label = tk.Label(root, text="Ready to process files...", fg="gray")
    status_label.grid(row=3, column=0, columnspan=3)
    
    return root

# Modified main section
if __name__ == "__main__":
    root = create_gui()
    root.mainloop()