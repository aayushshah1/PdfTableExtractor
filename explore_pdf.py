import pdfplumber
import os
import sys

def explore_pdf(pdf_path):
    """
    Explore a PDF file and extract details about its tables to help with debugging
    extraction issues.
    """
    print(f"\n==== Exploring PDF: {pdf_path} ====\n")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Process each page
            for page_num, page in enumerate(pdf.pages):
                print(f"\nPage {page_num+1} dimensions: {page.width}x{page.height}")
                
                tables = page.extract_tables()
                print(f"Found {len(tables)} tables on page {page_num+1}")
                
                for table_idx, table in enumerate(tables):
                    if not table:
                        continue
                        
                    print(f"\nTable {table_idx+1} on Page {page_num+1}")
                    print(f"Dimensions: {len(table)} rows × {len(table[0]) if table[0] else 0} columns")
                    print("\nSample of table content (first 5 rows):")
                    
                    # Print header row differently
                    if table and len(table) > 0:
                        print("\nHeader row:")
                        print(table[0])
                        
                        # Print up to 4 more rows
                        if len(table) > 1:
                            print("\nData rows:")
                            for row in table[1:6]:  # up to 5 data rows
                                print(row)
    
    except Exception as e:
        print(f"Error exploring PDF: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter the path to the PDF file: ")
        
    if not pdf_path:
        print("No PDF file specified.")
        sys.exit(1)
        
    if not os.path.exists(pdf_path):
        print(f"Error: File not found at {pdf_path}")
        sys.exit(1)
        
    explore_pdf(pdf_path)
    print("\nExploration complete. Use this information to customize the extraction process if needed.")
