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
                        
                        # Add "Scrip_Symbol" as the first column name
                        column_names.insert(0, "Scrip_Symbol")
                        
                        print(f"Using column names: {column_names}")
                        
                        # Skip the header row for the first table
                        start_row = 1
                    else:
                        # For subsequent tables, process all rows
                        start_row = 0
                    
                    # Process rows
                    current_scrip_symbol = None  # To track the current script symbol
                    
                    for row_idx, row in enumerate(table[start_row:], start_row):
                        # Check if this is a Scrip_Symbol row
                        if row and len(row) > 2 and row[0] == 'Scrip_Symbol :' and row[2] is not None:
                            # Extract the scrip symbol (e.g., "500116 IDBI - MITHIL DEEPAK KOTWAL")
                            current_scrip_symbol = row[2]
                            continue  # Skip this row from the final output
                        
                        # Check if this is a row with more than 2 columns of data (transaction row)
                        non_empty_cols = sum(1 for cell in row if cell is not None and str(cell).strip())
                        
                        # Only process rows with sufficient data (more than 2 columns)
                        if non_empty_cols > 2:
                            # This is a transaction row
                            transaction = {}
                            
                            # Only add the current scrip symbol if the row has substantial data (8+ columns)
                            # This prevents filling scrip symbol for empty/near-empty rows
                            filled_cols_count = sum(1 for cell in row if cell is not None and str(cell).strip())
                            
                            if current_scrip_symbol and filled_cols_count >= 8:
                                transaction[column_names[0]] = current_scrip_symbol
                            else:
                                transaction[column_names[0]] = "Unknown"
                            
                            # Map the rest of the columns
                            col_index = 1  # Start from 1 since we've already added Scrip_Symbol
                            for cell in row:
                                if cell is not None and col_index < len(column_names):
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
        
        # 5. Clean up the Scrip_Symbol column - remove the "Scrip_Symbol :" prefix if present
        # and extract just the company name without numbers and client name
        if 'Scrip_Symbol' in df.columns:
            df['Scrip_Symbol'] = df['Scrip_Symbol'].astype(str)
            
            def clean_scrip_symbol(symbol):
                # First remove any "Scrip_Symbol :" prefix
                symbol = symbol.replace('Scrip_Symbol :', '').strip()
                
                # If there's a dash, take only the part before it
                if ' - ' in symbol:
                    symbol = symbol.split(' - ')[0].strip()
                
                # Handle special case like "BSE BSE"
                if symbol.startswith('BSE '):
                    return 'BSE'
                
                # Remove leading numbers and spaces
                symbol = re.sub(r'^\d+\s+', '', symbol)
                
                return symbol
            
            df['Scrip_Symbol'] = df['Scrip_Symbol'].apply(clean_scrip_symbol)
        
        # Before creating portfolio summary, remove rows with very few filled columns
        # This will prevent empty lines after the transaction table
        if 'N.Qty' in df.columns:
            # Count non-empty cells in each row
            df['filled_columns'] = df.apply(lambda row: sum(pd.notna(val) and str(val).strip() != '' for val in row), axis=1)
            # Keep only rows with sufficient data (8+ filled columns)
            df = df[df['filled_columns'] >= 8]
            # Drop the helper column
            df = df.drop('filled_columns', axis=1)
        
        # Instead of using multiple sheets, we'll place everything in one sheet
        # Create a summary portfolio dataframe
        portfolio_df = None
        if 'Scrip_Symbol' in df.columns and 'N.Qty' in df.columns:
            # Group by Scrip_Symbol and sum N.Qty
            portfolio_df = df.groupby('Scrip_Symbol')['N.Qty'].sum().reset_index()
            # Add Current_Price column (will add formula later)
            portfolio_df['Current_Price'] = ''
            print(f"Created Portfolio summary with {len(portfolio_df)} unique securities")
        
        # Save to Excel with transactions and portfolio in the same sheet
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            # Write transactions to the first part of the sheet
            df.to_excel(writer, sheet_name='Transactions', index=False)
            
            # If portfolio data exists, add it below with a 1-row gap (was 3 before)
            if portfolio_df is not None:
                # Calculate the starting row for portfolio (transactions rows + header + 1 blank row)
                portfolio_start_row = len(df) + 1 + 1
                
                # Write a header for the portfolio section
                worksheet = writer.sheets['Transactions']
                worksheet.cell(row=portfolio_start_row, column=1, value="PORTFOLIO SUMMARY")
                
                # Write column headers for the portfolio section
                worksheet.cell(row=portfolio_start_row + 1, column=1, value="Scrip_Symbol")
                worksheet.cell(row=portfolio_start_row + 1, column=2, value="Total_Quantity")
                worksheet.cell(row=portfolio_start_row + 1, column=3, value="Current_Price")
                
                # Write portfolio data
                for i, row in portfolio_df.iterrows():
                    row_idx = portfolio_start_row + 2 + i  # +2 for portfolio header and column headers
                    
                    # Write Scrip_Symbol
                    worksheet.cell(row=row_idx, column=1, value=row['Scrip_Symbol'])
                    
                    # Write N.Qty
                    worksheet.cell(row=row_idx, column=2, value=row['N.Qty'])
                    
                    # Add GOOGLEFINANCE formula for Current_Price
                    formula = f'=GOOGLEFINANCE("{row["Scrip_Symbol"]}")'
                    worksheet.cell(row=row_idx, column=3, value=formula)
                
                # Add TOTAL row
                total_row = portfolio_start_row + 2 + len(portfolio_df)
                worksheet.cell(row=total_row, column=1, value="TOTAL")
                
                # Add SUM formula for the Total_Quantity column
                first_qty_cell = worksheet.cell(row=portfolio_start_row+2, column=2).coordinate
                last_qty_cell = worksheet.cell(row=total_row-1, column=2).coordinate
                worksheet.cell(row=total_row, column=2, value=f"=SUM({first_qty_cell}:{last_qty_cell})")
        
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
