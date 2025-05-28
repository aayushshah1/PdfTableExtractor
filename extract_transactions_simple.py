import pandas as pd
import pdfplumber
import os
import re
from datetime import datetime
import sys

def extract_transactions_simple(pdf_path, output_excel_path=None):
    """
    Extract transaction data from PDF tables using a simple approach:
    - Keep rows that have data in more than 2 columns
    - Skip rows that only have data in first 2 columns (like Scrip_Symbol rows)
    
    Args:
        pdf_path (str): Path to the PDF file
        output_excel_path (str, optional): Path to save Excel file. If None, uses PDF name + '_extraction.xlsx'
    
    Returns:
        pd.DataFrame: DataFrame containing all extracted transaction data
    """
    print(f"Processing PDF: {pdf_path}")
    
    # Set default output path if not provided
    if output_excel_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_excel_path = f"{base_name}_extraction.xlsx"
    
    all_transactions = []
    column_names = None
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Process each page
            for page_num, page in enumerate(pdf.pages):
                print(f"Processing page {page_num+1} of {len(pdf.pages)}")
                
                # Extract tables from the page
                tables = page.extract_tables()
                
                for table_idx, table in enumerate(tables):
                    if not table or len(table) < 1:
                        continue
                    
                    print(f"Found table {table_idx+1} on page {page_num+1} with {len(table)} rows")
                    
                    # Get column names from the first table's first row
                    if column_names is None and page_num == 0 and table_idx == 0:
                        # Get the column names and filter out None values
                        column_names = []
                        for i, col in enumerate(table[0]):
                            if col is not None and col.strip():  # Only add non-empty columns
                                column_names.append(col)
                            # Skip None columns - don't add "Column_2" etc.
                        
                        print(f"Using column names: {column_names}")
                        
                        # Skip the header row for the first table
                        start_row = 1
                    else:
                        # For subsequent tables, process all rows
                        start_row = 0
                    
                    # Process rows
                    for row_idx, row in enumerate(table[start_row:], start_row):
                        # Check if this is a row with more than 2 columns of data
                        non_empty_cols = sum(1 for cell in row if cell is not None and str(cell).strip())
                        
                        if non_empty_cols > 2:
                            # This is a transaction row (not a Scrip_Symbol row)
                            transaction = {}
                            
                            # Map columns by position, not by name (to avoid Column_2 issue)
                            col_index = 0
                            for i, cell in enumerate(row):
                                if cell is not None:
                                    # Only include columns with actual header names
                                    if col_index < len(column_names):
                                        transaction[column_names[col_index]] = cell
                                        col_index += 1
                            
                            all_transactions.append(transaction)
    
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return None
    
    # Convert to DataFrame
    if all_transactions:
        df = pd.DataFrame(all_transactions)
        
        # Clean data
        # 1. Clean numeric columns - remove commas and convert to numbers
        numeric_cols = ['B.Qty', 'B.Rate', 'S.Qty', 'S.Rate', 'N.Qty', 'N.Rate', 'N.Amt']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].astype(str)
                df[col] = df[col].str.replace(',', '', regex=False)
                df[col] = df[col].apply(
                    lambda x: re.sub(r'[^\d.-]', '', str(x)) if pd.notna(x) and str(x).strip() else '')
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 2. Clean and standardize dates
        if 'Date' in df.columns:
            def standardize_date(date_str):
                if pd.isna(date_str) or not str(date_str).strip():
                    return None
                
                date_str = str(date_str).strip()
                try:
                    # Try parsing with different formats
                    formats = [
                        '%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%Y/%m/%d',
                        '%d-%b-%Y', '%d %b %Y', '%b %d, %Y', '%B %d, %Y',
                        '%d-%m-%y', '%d/%m/%y', '%y-%m-%d', '%y/%m/%d'
                    ]
                    
                    for fmt in formats:
                        try:
                            return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
                        except ValueError:
                            continue
                            
                    return date_str  # Return original if no format matches
                except Exception:
                    return date_str
                    
            df['Date'] = df['Date'].apply(standardize_date)
        
        # 3. Clean up multi-line text in cells
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].str.replace('\n', ' ', regex=False)
        
        # 4. Remove duplicate header rows that might have been extracted as data
        if 'Company' in df.columns and 'Date' in df.columns:
            df = df[~((df['Company'] == 'Company') & (df['Date'] == 'Date'))]
        
        # Save to Excel
        df.to_excel(output_excel_path, index=False)
        print(f"\nSuccessfully extracted {len(df)} rows and saved to {output_excel_path}")
        
        # Show sample of extracted data
        if len(df) > 0:
            print("\nSample of extracted data:")
            print(df.head().to_string())
        
        return df
    else:
        print("No transaction data found in the PDF.")
        return None

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter path to PDF file: ")
        
        if not pdf_path:
            pdf_path = "Data/Main.PDF"  # Default path
    
    extract_transactions_simple(pdf_path)
