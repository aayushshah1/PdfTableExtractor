import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import sys
from extract_transactions_simple import extract_transactions_simple

class PDFToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Set up the main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Application title
        title_label = ttk.Label(main_frame, text="PDF Table Extractor", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # Description
        description = "This tool extracts transaction data from PDF files and saves it to Excel."
        desc_label = ttk.Label(main_frame, text=description, wraplength=500)
        desc_label.pack(pady=5)
        
        # Frame for file selection
        file_frame = ttk.LabelFrame(main_frame, text="Select PDF File", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        # PDF file selection
        self.pdf_path = tk.StringVar()
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, width=50)
        pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = ttk.Button(file_frame, text="Browse", command=self.browse_pdf)
        browse_btn.pack(side=tk.RIGHT)
        
        # Frame for output selection
        output_frame = ttk.LabelFrame(main_frame, text="Output Excel File (Optional)", padding="10")
        output_frame.pack(fill=tk.X, pady=10)
        
        # Output file selection
        self.output_path = tk.StringVar()
        output_entry = ttk.Entry(output_frame, textvariable=self.output_path, width=50)
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        output_btn = ttk.Button(output_frame, text="Browse", command=self.browse_output)
        output_btn.pack(side=tk.RIGHT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=100, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=10)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, wraplength=500)
        status_label.pack(pady=5)
        
        # Convert button
        convert_btn = ttk.Button(main_frame, text="Convert to Excel", command=self.start_conversion)
        convert_btn.pack(pady=10)
        
        # Credits
        credits = "Created for data extraction from financial transaction PDFs"
        credits_label = ttk.Label(main_frame, text=credits, font=("Arial", 8))
        credits_label.pack(side=tk.BOTTOM, pady=5)
    
    def browse_pdf(self):
        filetypes = [("PDF files", "*.pdf")]
        filename = filedialog.askopenfilename(title="Select PDF File", filetypes=filetypes)
        if filename:
            self.pdf_path.set(filename)
            
            # Automatically set a default output path
            base_name = os.path.splitext(os.path.basename(filename))[0]
            output_path = os.path.join(os.path.dirname(filename), f"{base_name}_extraction.xlsx")
            self.output_path.set(output_path)
    
    def browse_output(self):
        filetypes = [("Excel files", "*.xlsx")]
        filename = filedialog.asksaveasfilename(
            title="Save Excel File As",
            filetypes=filetypes,
            defaultextension=".xlsx"
        )
        if filename:
            self.output_path.set(filename)
    
    def start_conversion(self):
        pdf_path = self.pdf_path.get()
        output_path = self.output_path.get() if self.output_path.get() else None
        
        if not pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first!")
            return
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", f"PDF file not found: {pdf_path}")
            return
        
        # Start progress bar
        self.progress.start()
        self.status_var.set("Converting... Please wait.")
        self.root.update()
        
        # Run conversion in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self.run_conversion, args=(pdf_path, output_path))
        thread.daemon = True
        thread.start()
    
    def run_conversion(self, pdf_path, output_path):
        try:
            # Redirect stdout to capture console output
            original_stdout = sys.stdout
            from io import StringIO
            captured_output = StringIO()
            sys.stdout = captured_output
            
            # Run the extraction
            result = extract_transactions_simple(pdf_path, output_path)
            
            # Get back console output
            sys.stdout = original_stdout
            log_output = captured_output.getvalue()
            
            # Update GUI with results
            self.root.after(0, self.update_status, result, output_path, log_output)
            
        except Exception as e:
            self.root.after(0, self.handle_error, str(e))
    
    def update_status(self, result, output_path, log_output):
        self.progress.stop()
        
        if result is not None:
            row_count = len(result)
            if output_path is None:
                base_name = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
                output_path = f"{base_name}_extraction.xlsx"
                
            success_message = f"Successfully extracted {row_count} rows to {output_path}"
            self.status_var.set(success_message)
            
            messagebox.showinfo("Success", success_message)
            
            # Ask if user wants to open the Excel file
            if messagebox.askyesno("Open Excel File", "Would you like to open the Excel file now?"):
                self.open_excel_file(output_path)
        else:
            error_msg = "Extraction failed. Check if the PDF contains transaction tables."
            self.status_var.set(error_msg)
            messagebox.showerror("Extraction Failed", f"{error_msg}\n\nDetails:\n{log_output}")
    
    def handle_error(self, error_message):
        self.progress.stop()
        self.status_var.set(f"Error: {error_message}")
        messagebox.showerror("Error", f"An error occurred during conversion:\n\n{error_message}")
    
    def open_excel_file(self, filepath):
        """Open the Excel file with the default application"""
        import platform
        import subprocess
        
        try:
            if platform.system() == 'Darwin':  # macOS
                subprocess.call(('open', filepath))
            elif platform.system() == 'Windows':
                os.startfile(filepath)
            else:  # Linux
                subprocess.call(('xdg-open', filepath))
        except Exception as e:
            messagebox.showerror("Error", f"Could not open the file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToExcelApp(root)
    root.mainloop()
