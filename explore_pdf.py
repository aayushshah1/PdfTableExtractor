"""
Utility script to explore the structure of PDF files.
Helps examine tables, text, and other elements to help with customizing extraction.
"""
import pdfplumber
import sys
import os

def explore_pdf(pdf_path):
    """
    Explore the structure of a PDF file to help with extraction.
    Prints tables, their dimensions, and other useful information.
    """
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        return

    try:
        print(f"\nExploring PDF: {pdf_path}\n")
        print("=" * 80)
        
        with pdfplumber.open(pdf_path) as pdf:
            # Get basic PDF information
            print(f"Number of pages: {len(pdf.pages)}")
            
            # Explore each page
            for page_num, page in enumerate(pdf.pages):
                print(f"\nPAGE {page_num + 1}:")
                print("-" * 40)
                
                # Extract tables
                tables = page.extract_tables()
                if tables:
                    print(f"Found {len(tables)} tables on page {page_num + 1}")
                    
                    # Print details of each table
                    for i, table in enumerate(tables):
                        print(f"\n  Table {i + 1}:")
                        print(f"  - Rows: {len(table)}")
                        print(f"  - Columns: {len(table[0]) if table and table[0] else 0}")
                        
                        # Display header row if available
                        if table and len(table) > 0:
                            print("\n  - Header row:")
                            for j, cell in enumerate(table[0]):
                                print(f"    Col {j}: {cell}")
                        
                        # Show a sample of data rows
                        print("\n  - Data sample:")
                        max_rows = min(3, len(table) - 1) if len(table) > 1 else 0
                        for row in range(1, max_rows + 1):
                            print(f"    Row {row}:", end=" ")
                            row_content = []
                            for cell in table[row]:
                                cell_str = str(cell).replace('\n', ' ')[:20]
                                if len(cell_str) == 20:
                                    cell_str += "..."
                                row_content.append(cell_str)
                            print(" | ".join(row_content))
                else:
                    print("No tables found on this page")
                    
                # Get page text for troubleshooting
                text = page.extract_text()
                text_preview = text[:200] + "..." if text and len(text) > 200 else text
                print(f"\n  Text preview:\n  {text_preview}")

    except Exception as e:
        print(f"Error exploring PDF: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = input("Enter path to PDF file to explore: ")
        
    explore_pdf(pdf_path)
