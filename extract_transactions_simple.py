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
                        
                        # Add "Scrip_Symbol" as the first column name (KEEP ORIGINAL LOGIC)
                        column_names.insert(0, "Scrip_Symbol")
                        
                        print(f"Using column names: {column_names}")

                        # Skip the header row for the first table
                        start_row = 1
                    else:
                        # For subsequent tables, process all rows
                        start_row = 0
                    
                    # Process rows
                    current_scrip_symbol = None  # To track the current script symbol
                    current_bom_id = None  # To track the current BOM ID
                    
                    for row_idx, row in enumerate(table[start_row:], start_row):
                        # Check if this is a Scrip_Symbol row
                        if row and len(row) > 2 and row[0] == 'Scrip_Symbol :' and row[2] is not None:
                            # Extract the scrip symbol and BOM ID (e.g., "500116 IDBI - MITHIL DEEPAK KOTWAL")
                            full_symbol = row[2]
                            
                            # Extract BOM ID (first set of digits)
                            bom_match = re.match(r'^(\d+)', full_symbol.strip())
                            if bom_match:
                                current_bom_id = bom_match.group(1)
                            else:
                                current_bom_id = None
                            
                            current_scrip_symbol = full_symbol
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
                                # REMOVE BOM_ID handling here - keep original logic
                            else:
                                transaction[column_names[0]] = "Unknown"

                            # Map the rest of the columns (KEEP ORIGINAL LOGIC)
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
        
        # REMOVE dynamic BOM_ID column creation
        # if 'BOM_ID' not in df.columns:
        #     df.insert(1, 'BOM_ID', None)

        # REMOVE dynamic column_names update
        # if 'BOM_ID' not in column_names:
        #     column_names.insert(1, 'BOM_ID')

        # Before creating portfolio summary, fix Unknown scrip symbols by propagating the last valid symbol
        if 'Scrip_Symbol' in df.columns:
            last_valid_symbol = None
            for idx, row in df.iterrows():
                current = row['Scrip_Symbol']
                if current != 'Unknown' and pd.notna(current) and str(current).strip():
                    last_valid_symbol = current
                elif last_valid_symbol is not None:
                    df.at[idx, 'Scrip_Symbol'] = last_valid_symbol

        # 5. Clean up the Scrip_Symbol column and extract BOM IDs (ENHANCED LOGIC)
        if 'Scrip_Symbol' in df.columns:
            df['Scrip_Symbol'] = df['Scrip_Symbol'].astype(str)
            
            def clean_scrip_symbol_and_extract_bom(symbol):
                # First remove any "Scrip_Symbol :" prefix
                symbol = symbol.replace('Scrip_Symbol :', '').strip()
                
                # Extract BOM ID if present (first set of digits before space)
                bom_id = None
                bom_match = re.match(r'^(\d+)', symbol)
                if bom_match:
                    bom_id = bom_match.group(1)  # Extract "544325"
                    # Remove the BOM ID from symbol
                    symbol = re.sub(r'^\d+\s*', '', symbol)  # Remove "544325 "
                
                # If there's a dash, take only the part before it
                if ' - ' in symbol:
                    symbol = symbol.split(' - ')[0].strip()  # "ITCHOTELS - NAME" -> "ITCHOTELS"
                
                # Handle duplicate symbols (e.g., "CDSL CDSL" -> "CDSL")
                words = symbol.split()
                if len(words) == 2 and words[0] == words[1]:
                    symbol = words[0]
                elif len(words) > 1:
                    # Remove spaces for single-word symbols
                    symbol = ''.join(words) if all(word.isalpha() for word in words) else symbol
                
                # Handle special cases
                if symbol.startswith('BSE '):
                    symbol = 'BSE'
                
                return symbol, bom_id
            
            # Apply cleaning and BOM extraction
            cleaned_data = df['Scrip_Symbol'].apply(clean_scrip_symbol_and_extract_bom)
            df['Scrip_Symbol'] = [item[0] for item in cleaned_data]
            
            # ADD BOM_ID column with extracted values
            df.insert(1, 'BOM_ID', [item[1] for item in cleaned_data])

        # Remove rows with very few filled columns - this will prevent empty lines after the transaction table
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
            # Group by Scrip_Symbol and sum N.Qty, also get the first BOM_ID for each symbol
            portfolio_df = df.groupby('Scrip_Symbol').agg({
                'N.Qty': 'sum',
                'BOM_ID': 'first'  # Take first BOM_ID for each symbol
            }).reset_index()

            # Add Current_Price column (will add formula later)
            portfolio_df['Current_Price'] = ''
            # Add Value column (will add formula later)
            portfolio_df['Value'] = ''
            print(f"Created Portfolio summary with {len(portfolio_df)} unique securities")
        
        # Save to Excel with transactions and portfolio in the same sheet
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            # Write transactions to the first part of the sheet
            df.to_excel(writer, sheet_name='Transactions', index=False)
            
            # Get worksheet reference
            worksheet = writer.sheets['Transactions']
            
            # Find the actual last row with transaction data by examining the worksheet
            last_transaction_row = 0
            for row in range(len(df), 0, -1):
                # Check if row has content in key columns
                row_in_sheet = row + 1  # +1 for header
                if worksheet.cell(row=row_in_sheet, column=1).value:  # Check Scrip_Symbol
                    last_transaction_row = row_in_sheet
                    break
            
            if last_transaction_row == 0:
                # Fallback: use length + header
                last_transaction_row = len(df) + 1
            
            # Add Portfolio_Value row after the last actual transaction row
            portfolio_value_row = last_transaction_row + 1
            
            # Add Portfolio Value row with TODAY() in column 4 (Date column is now D)
            worksheet.cell(row=portfolio_value_row, column=1, value="Portfolio_Value")
            worksheet.cell(row=portfolio_value_row, column=4, value="=TODAY()")  # TODAY() in Date column (4)
            
            # If portfolio data exists, add it below with a 1-row gap
            if portfolio_df is not None:
                # Calculate the starting row for portfolio (Portfolio_Value row + 1 blank row)
                portfolio_start_row = portfolio_value_row + 2
                
                # Write a header for the portfolio section
                worksheet.cell(row=portfolio_start_row, column=1, value="PORTFOLIO SUMMARY")
                
                # Write column headers for the portfolio section
                worksheet.cell(row=portfolio_start_row + 1, column=1, value="BOM_ID")
                worksheet.cell(row=portfolio_start_row + 1, column=2, value="Scrip_Symbol")
                worksheet.cell(row=portfolio_start_row + 1, column=3, value="Total_Quantity")
                worksheet.cell(row=portfolio_start_row + 1, column=4, value="Current_Price")
                worksheet.cell(row=portfolio_start_row + 1, column=5, value="Value")
                
                # Write portfolio data
                for i, row in portfolio_df.iterrows():
                    row_idx = portfolio_start_row + 2 + i  # +2 for portfolio header and column headers
                    
                    # Format BOM ID with BOM: prefix if available
                    bom_formatted = f"BOM:{row['BOM_ID']}" if pd.notna(row['BOM_ID']) else ""
                    worksheet.cell(row=row_idx, column=1, value=bom_formatted)
                    
                    # Format Scrip_Symbol with NSE: prefix
                    nse_symbol = f"NSE:{row['Scrip_Symbol']}"
                    worksheet.cell(row=row_idx, column=2, value=nse_symbol)
                    
                    # Write N.Qty
                    worksheet.cell(row=row_idx, column=3, value=row['N.Qty'])
                    
                    # Enhanced GOOGLEFINANCE formula with time-based logic and NSE/BOM fallback
                    nse_ref = worksheet.cell(row=row_idx, column=2).coordinate
                    bom_ref = worksheet.cell(row=row_idx, column=1).coordinate
                    
                    # Time-based formula with NSE/BOM fallback
                    price_formula = f'''=IFERROR(
IF(TIME(HOUR(NOW()),MINUTE(NOW()),0)>TIME(15,30,0),
INDEX(GOOGLEFINANCE({nse_ref},"close",WORKDAY(TODAY(),0),WORKDAY(TODAY(),0)),2,2),
INDEX(GOOGLEFINANCE({nse_ref},"close",WORKDAY(TODAY(),-1),WORKDAY(TODAY(),-1)),2,2)),
IFERROR(
IF(TIME(HOUR(NOW()),MINUTE(NOW()),0)>TIME(15,30,0),
INDEX(GOOGLEFINANCE({bom_ref},"close",WORKDAY(TODAY(),0),WORKDAY(TODAY(),0)),2,2),
INDEX(GOOGLEFINANCE({bom_ref},"close",WORKDAY(TODAY(),-1),WORKDAY(TODAY(),-1)),2,2)),
"Add manually"))'''
                    
                    worksheet.cell(row=row_idx, column=4, value=price_formula)
                    
                    # Add Value formula (quantity × price)
                    qty_ref = worksheet.cell(row=row_idx, column=3).coordinate
                    price_ref = worksheet.cell(row=row_idx, column=4).coordinate
                    worksheet.cell(row=row_idx, column=5, value=f"={qty_ref}*{price_ref}")
                
                # Add TOTAL row
                total_row = portfolio_start_row + 2 + len(portfolio_df)
                worksheet.cell(row=total_row, column=2, value="TOTAL")
                
                # Add SUM formulas for the Total_Quantity and Value columns
                first_qty_cell = worksheet.cell(row=portfolio_start_row+2, column=3).coordinate
                last_qty_cell = worksheet.cell(row=total_row-1, column=3).coordinate
                worksheet.cell(row=total_row, column=3, value=f"=SUM({first_qty_cell}:{last_qty_cell})")
                
                # Add SUM formula for the Value column
                first_value_cell = worksheet.cell(row=portfolio_start_row+2, column=5).coordinate
                last_value_cell = worksheet.cell(row=total_row-1, column=5).coordinate
                total_value_cell = worksheet.cell(row=total_row, column=5, value=f"=SUM({first_value_cell}:{last_value_cell})")
                
                # Now update the Portfolio_Value row's N.Amt cell with reference to total portfolio value
                total_value_ref = total_value_cell.coordinate
                n_amt_col = None
                for col_idx, col_name in enumerate(df.columns, 1):
                    if col_name == "N.Amt":
                        n_amt_col = col_idx
                        break
                
                if n_amt_col:
                    worksheet.cell(row=portfolio_value_row, column=n_amt_col, value=f"={total_value_ref}")
                
                # Add Portfolio XIRR row after a blank row
                xirr_row = total_row + 2  # +2 for one blank row
                worksheet.cell(row=xirr_row, column=1, value="Portfolio XIRR")
                
                # Add XIRR formula that uses transaction dates and amounts
                # Find the N.Amt column (usually column 11)
                n_amt_col = None
                for col_idx, col_name in enumerate(df.columns, 1):
                    if col_name == "N.Amt":
                        n_amt_col = col_idx
                        break
                
                if n_amt_col:
                    # Create XIRR formula referencing dates (column 4) and N.Amt (column n_amt_col)
                    # Transaction rows start from row 2 (after header) and go through the Portfolio_Value row
                    date_range = f"D2:D{portfolio_value_row}"
                    amount_range = f"{chr(64+n_amt_col)}2:{chr(64+n_amt_col)}{portfolio_value_row}"
                    # Fix: correct parameter order - values first, then dates
                    xirr_formula = f"=XIRR({amount_range},{date_range})"
                    worksheet.cell(row=xirr_row, column=2, value=xirr_formula)
                
                # Add Portfolio XIRR Percentage row
                xirr_pct_row = xirr_row + 1
                worksheet.cell(row=xirr_pct_row, column=1, value="Portfolio XIRR Percentage")
                worksheet.cell(row=xirr_pct_row, column=2, value=f"={worksheet.cell(row=xirr_row, column=2).coordinate}*100")
        
        print(f"\nSuccessfully extracted {len(df)} rows and saved to {output_excel_path}")
        
        # Show sample of extracted data
        if len(df) > 0:
            print("\nSample of extracted data:")
            print(df.head().to_string())
        
        return df
    else:
        print("No transaction data found in the PDF.")
        return None

# if __name__ == "__main__":
#     if len(sys.argv) > 1:
#         pdf_path = sys.argv[1]
#     else:
#         pdf_path = input("Enter path to PDF file: ")
        
#         if not pdf_path:
#             pdf_path = "Data/Main.PDF"  # Default path
    
#     extract_transactions_simple(pdf_path)
        
#         print(f"\nSuccessfully extracted {len(df)} rows and saved to {output_excel_path}")
        
#         # Show sample of extracted data
#         if len(df) > 0:
#             print("\nSample of extracted data:")
#             print(df.head().to_string())
        
#         return df
#     else:
#         print("No transaction data found in the PDF.")
#         return None

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter path to PDF file: ")
        
        if not pdf_path:
            pdf_path = "Data/Main.PDF"  # Default path
    
    extract_transactions_simple(pdf_path)