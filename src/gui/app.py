"""
GUI module for the invoice conversion application
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import pytesseract


def create_gui(process_callback):
    """
    Create the main GUI window
    
    Args:
        process_callback (function): Callback function to process the PDF file
        
    Returns:
        tk.Tk: The main window object
    """
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
            title="Save Excel As",
            defaultextension=".xlsx",
            initialfile=f"invoice_data_{timestamp}.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, file_path)
    
    def log_message(message):
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)  # Scroll to the end
        root.update_idletasks()  # Update the UI
    
    def process_file():
        input_path = input_entry.get()
        output_path = output_entry.get()
        
        if not input_path or not output_path:
            messagebox.showerror("Error", "Please select both input PDF and output Excel file location.")
            return
            
        try:
            # Update status and disable process button
            status_label.config(text="Processing...", foreground="blue")
            process_button.config(state="disabled")
            log_text.delete(1.0, tk.END)  # Clear previous logs
            log_message("Starting PDF processing...")
            
            # Make sure you have Tesseract OCR installed and in your PATH
            pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            
            # Process the file with logging callback
            process_callback(input_path, output_path, log_callback=log_message)
            
            # Update status and re-enable process button
            status_label.config(text="File processed successfully!", foreground="green")
            process_button.config(state="normal")
            
            messagebox.showinfo("Success", f"File processed successfully!\nSaved to: {output_path}")
        except Exception as e:
            # Update status and re-enable process button
            status_label.config(text="Error occurred!", foreground="red")
            process_button.config(state="normal")
            log_message(f"Error: {str(e)}")
            
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
    
    # Create main window
    root = tk.Tk()
    root.title("Invoice PDF to Excel Converter")
    root.geometry("800x600")  # Increased height for log section
    root.minsize(800, 600)  # Set minimum window size
    
    # Configure style
    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat")
    style.configure("TLabel", padding=6)
    style.configure("TEntry", padding=6)
    
    # Create main frame with padding
    main_frame = ttk.Frame(root, padding="20")
    main_frame.grid(row=0, column=0, sticky="nsew")
    
    # Configure grid
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    
    # Title/Header
    title_label = ttk.Label(main_frame, text="Invoice PDF to Excel Converter", font=("Helvetica", 16, "bold"))
    title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
    
    # Input PDF selection
    ttk.Label(main_frame, text="Select PDF File:").grid(row=1, column=0, padx=5, sticky="e")
    input_entry = ttk.Entry(main_frame, width=50)
    input_entry.grid(row=1, column=1, padx=5, sticky="ew")
    input_button = ttk.Button(main_frame, text="Browse...", command=select_pdf)
    input_button.grid(row=1, column=2, padx=5)
    
    # Output Excel location
    ttk.Label(main_frame, text="Save Excel As:").grid(row=2, column=0, padx=5, sticky="e")
    output_entry = ttk.Entry(main_frame, width=50)
    output_entry.grid(row=2, column=1, padx=5, sticky="ew")
    output_button = ttk.Button(main_frame, text="Browse...", command=select_save_location)
    output_button.grid(row=2, column=2, padx=5)
    
    # Process button
    process_button = ttk.Button(main_frame, text="Process PDF", command=process_file, width=20)
    process_button.grid(row=3, column=1, pady=20)
    
    # Status label
    status_label = ttk.Label(main_frame, text="Ready to process files...", foreground="gray")
    status_label.grid(row=4, column=0, columnspan=3)
    
    # Log section
    log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
    log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=10)
    
    # Create text widget for logs with scrollbar
    log_text = tk.Text(log_frame, height=10, width=80, wrap=tk.WORD)
    log_scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    log_text.configure(yscrollcommand=log_scrollbar.set)
    
    # Grid the log widgets
    log_text.grid(row=0, column=0, sticky="nsew")
    log_scrollbar.grid(row=0, column=1, sticky="ns")
    
    # Configure log frame grid
    log_frame.grid_columnconfigure(0, weight=1)
    log_frame.grid_rowconfigure(0, weight=1)
    
    # Add tooltips
    def create_tooltip(widget, text):
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            
            label = ttk.Label(tooltip, text=text, justify='left',
                            background="#ffffe0", relief='solid', borderwidth=1)
            label.pack()
            
            def hide_tooltip():
                tooltip.destroy()
            
            widget.tooltip = tooltip
            widget.bind('<Leave>', lambda e: hide_tooltip())
        
        widget.bind('<Enter>', show_tooltip)
    
    # Add tooltips to buttons
    create_tooltip(input_button, "Select the PDF invoice file to process")
    create_tooltip(output_button, "Choose where to save the Excel file")
    create_tooltip(process_button, "Start processing the PDF file")
    
    return root 