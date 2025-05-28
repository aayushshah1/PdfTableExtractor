import camelot
import pandas as pd
import os
import sys

def explore_pdf(pdf_path):
    """
    Explore a PDF file and extract details about its tables to help with debugging
    extraction issues.
    """
    print(f"\n==== Exploring PDF: {pdf_path} ====\n")
    
    # Try using lattice first
    print("Attempting with lattice flavor...")
    try:
        tables_lattice = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
        print(f"Found {len(tables_lattice)} tables using lattice method")
        
        for i, table in enumerate(tables_lattice):
            print(f"\nTable {i+1} (lattice) - Accuracy: {table.accuracy}")
            print(f"Dimensions: {table.shape[0]} rows × {table.shape[1]} columns")
            print(f"Table areas: {table.table_areas}")
            print("\nSample of table content:")
            print(table.df.head().to_string())
    except Exception as e:
        print(f"Error with lattice extraction: {str(e)}")
    
    # Try using stream
    print("\nAttempting with stream flavor...")
    try:
        tables_stream = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
        print(f"Found {len(tables_stream)} tables using stream method")
        
        for i, table in enumerate(tables_stream):
            print(f"\nTable {i+1} (stream) - Accuracy: {table.accuracy}")
            print(f"Dimensions: {table.shape[0]} rows × {table.shape[1]} columns")
            print(f"Table areas: {table.table_areas}")
            print("\nSample of table content:")
            print(table.df.head().to_string())
    except Exception as e:
        print(f"Error with stream extraction: {str(e)}")
    
    # Try with pdfplumber
    print("\nAttempting with pdfplumber...")
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                print(f"\nPage {i+1} dimensions: {page.width}x{page.height}")
                tables = page.extract_tables()
                print(f"Found {len(tables)} tables on page {i+1} using pdfplumber")
                
                for j, table in enumerate(tables):
                    if table and len(table) > 0:
                        print(f"\nTable {j+1} on Page {i+1}")
                        print(f"Dimensions: {len(table)} rows × {len(table[0])} columns")
                        print("\nSample of table content:")
                        for row in table[:5]:  # Show first 5 rows
                            print(row)
    except ImportError:
        print("pdfplumber not installed. Install with: pip install pdfplumber")
    except Exception as e:
        print(f"Error with pdfplumber extraction: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter the path to the PDF file: ")
        if not pdf_path:
            pdf_path = './Data/Main.PDF'  # Default
    
    explore_pdf(pdf_path)
    print("\nExploration completed. Use the information above to adjust extraction parameters.")
