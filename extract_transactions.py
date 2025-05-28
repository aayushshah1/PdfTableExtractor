import camelot
import pandas as pd
import os
import re
from datetime import datetime

def extract_transactions_from_pdf(pdf_path, output_excel_path=None):
    """
    Extract transaction data (Company, Date, N.Amt) from PDF and save to Excel.
    
    Args:
        pdf_path (str): Path to the PDF file
        output_excel_path (str, optional): Path to save the Excel file. If None, creates path based on PDF name.
    
    Returns:
        pd.DataFrame: DataFrame containing the extracted transaction data
    """
    print(f"Processing {pdf_path}...")
    
    # Set default output path if not provided
    if output_excel_path is None:
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_excel_path = f"{base_name}_transactions.xlsx"
    
    # Extract tables from the PDF using camelot
    # Note: flavor='lattice' works well for tables with visible lines
    # You might need to try 'stream' if lattice doesn't capture all data
    tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
    
    if len(tables) == 0:
        print("No tables found. Trying with stream flavor...")
        tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        
    if len(tables) == 0:
        print("No tables found in the PDF.")
        return None
    
    # Process and combine all tables
    all_data = []
    
    for i, table in enumerate(tables):
        df = table.df
        print(f"Table {i+1} has {len(df)} rows and {len(df.columns)} columns")
        
        # Skip tables that don't have enough columns
        if len(df.columns) < 3:
            continue
            
        # Check if this is a transaction table by looking for header keywords
        headers = df.iloc[0].str.lower()
        
        # If headers don't contain key terms, try to find header row
        header_row_idx = 0
        for idx, row in df.iterrows():
            row_text = ' '.join(row.astype(str)).lower()
            if 'company' in row_text and ('date' in row_text or 'dt' in row_text) and ('n.amt' in row_text or 'amount' in row_text):
                header_row_idx = idx
                break
        
        # Get the column indices for Company, Date, and N.Amt
        if header_row_idx > 0:
            # Use the identified header row
            df.columns = df.iloc[header_row_idx]
            df = df.iloc[header_row_idx+1:].reset_index(drop=True)
        
        # Identify columns by name (case-insensitive)
        company_col = None
        date_col = None
        amount_col = None
        
        for col_idx, col_name in enumerate(df.columns):
            col_name_lower = str(col_name).lower().strip()
            if 'company' in col_name_lower:
                company_col = col_idx
            elif 'date' in col_name_lower or 'dt' in col_name_lower:
                date_col = col_idx
            elif 'n.amt' in col_name_lower or 'net' in col_name_lower and 'amount' in col_name_lower:
                amount_col = col_idx
        
        # Skip table if required columns are not found
        if company_col is None or date_col is None or amount_col is None:
            print(f"Skipping table {i+1} - couldn't identify required columns")
            continue
        
        # Extract relevant columns
        transactions = pd.DataFrame({
            'Company': df.iloc[:, company_col],
            'Date': df.iloc[:, date_col],
            'N.Amt': df.iloc[:, amount_col]
        })
        
        # Clean data
        # Remove rows where company is empty or contains header-like terms
        transactions = transactions[transactions['Company'].str.strip() != '']
        transactions = transactions[~transactions['Company'].str.contains('company|date|n.amt', case=False)]
        
        # Clean amount values - remove commas, spaces, etc.
        transactions['N.Amt'] = transactions['N.Amt'].astype(str)
        transactions['N.Amt'] = transactions['N.Amt'].str.replace(',', '')
        transactions['N.Amt'] = transactions['N.Amt'].apply(
            lambda x: re.sub(r'[^\d.-]', '', str(x)) if pd.notna(x) and str(x).strip() else '')
        
        # Convert to numeric, coercing errors to NaN
        transactions['N.Amt'] = pd.to_numeric(transactions['N.Amt'], errors='coerce')
        
        # Clean and standardize dates
        def standardize_date(date_str):
            if pd.isna(date_str) or not str(date_str).strip():
                return None
            
            date_str = str(date_str).strip()
            try:
                # Try parsing with different formats - expanded list
                formats = [
                    '%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y/%m/%d',
                    '%d-%b-%Y', '%d %b %Y', '%b %d, %Y', '%B %d, %Y',
                    '%d-%m-%y', '%d/%m/%y', '%y-%m-%d', '%y/%m/%d',
                    '%d-%B-%Y', '%d %B %Y', '%d.%m.%Y', '%d.%m.%y'
                ]
                
                for fmt in formats:
                    try:
                        return datetime.strptime(date_str, fmt).strftime('%Y-%m-%d')
                    except ValueError:
                        continue
                
                # Try to handle numeric-only dates (like 20220131)
                if re.match(r'^\d{8}$', date_str):
                    return datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')
                
                return date_str  # Return original if no format matches
            except Exception:
                return date_str
                
        transactions['Date'] = transactions['Date'].apply(standardize_date)
        
        # Add to our collection
        all_data.append(transactions)
    
    if not all_data:
        print("No valid transaction tables found.")
        return None
    
    # Combine all data
    result_df = pd.concat(all_data, ignore_index=True)
    
    # Filter out rows with missing critical data
    result_df = result_df.dropna(subset=['Company', 'Date', 'N.Amt'])
    
    # Save to Excel
    result_df.to_excel(output_excel_path, index=False)
    print(f"Extracted {len(result_df)} transactions and saved to {output_excel_path}")
    
    return result_df

if __name__ == "__main__":
    # Example usage
    pdf_path = './Data/Main.PDF'
    try:
        # For debugging, print out more details about the tables
        print("Attempting to extract tables using camelot...")
        tables = camelot.read_pdf(pdf_path, pages='all')
        
        for i, table in enumerate(tables):
            df = table.df
            print(f"\nTable {i+1} contents (first few rows):")
            print(df.head().to_string())
            
            # Print all column headers to help with debugging
            print(f"\nColumn headers for Table {i+1}:")
            for col_idx, col in enumerate(df.columns):
                print(f"Column {col_idx}: '{col}'")
            
            # Try to identify more flexibly
            print("\nLooking for key columns in Table contents:")
            for row_idx, row in df.iterrows():
                row_str = ' '.join(row.astype(str)).lower()
                if 'company' in row_str or 'date' in row_str or 'n.amt' in row_str or 'amount' in row_str:
                    print(f"Row {row_idx} contains key terms: {row.to_list()}")
        
        # Now proceed with the regular extraction
        extract_transactions_from_pdf(pdf_path)
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        print("\nTrying alternative approach with different parse options...")
        try:
            # Try with different parameters for camelot
            print("Trying with different table area detection...")
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream', 
                                     table_areas=['0,0,100%,100%'], 
                                     columns=['10%', '30%', '45%', '60%', '75%', '90%'])
            
            if len(tables) > 0:
                print(f"Found {len(tables)} tables with custom parameters")
                # Try processing with custom column mapping
                for i, table in enumerate(tables):
                    df = table.df
                    print(f"\nTable {i+1} with {len(df)} rows detected using custom parameters")
                    print(df.head().to_string())
            
            # Continue with the rest of the fallback approaches
            # ...existing code for the rest of the fallbacks...
