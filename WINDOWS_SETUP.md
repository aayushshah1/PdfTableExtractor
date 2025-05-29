# PDF Table Extractor - Windows Setup Guide

This guide will help you install and run the PDF Table Extractor application on a Windows computer that doesn't have any development tools installed.

## Option 1: Standalone Executable (Recommended for non-technical users)

1. **Download the application**
   - Download the `PDFTableExtractor.zip` file from [this link](https://yourcompany.com/downloads/PDFTableExtractor.zip)
   - Right-click on the downloaded zip file and select "Extract All..."
   - Choose a location where you want to extract the files (e.g., your Desktop)

2. **Run the application**
   - Open the extracted folder
   - Double-click on `PDFTableExtractor.exe` to start the application
   - That's it! No installation needed.

3. **Using the application**
   - Click "Browse" to select your PDF file
   - The output Excel file location will be set automatically
   - Click "Convert to Excel" to process the PDF
   - When complete, you can open the Excel file directly

## Option 2: Manual Installation (For IT Support)

If the standalone executable doesn't work in your environment, follow these steps:

1. **Install Python 3.10**
   - Download Python 3.10 from [python.org](https://www.python.org/downloads/windows/)
   - **Important**: During installation, check the box "Add Python to PATH"
   - Complete the installation

2. **Download the application source code**
   - Download the application folder
   - Extract it to a location of your choice (e.g., C:\PDFTableExtractor)

3. **Open Command Prompt**
   - Press Win+R
   - Type "cmd" and press Enter

4. **Install dependencies**
   - Navigate to the application folder:
     ```
     cd C:\path\to\PDFTableExtractor
     ```
   - Install required packages:
     ```
     pip install -r requirements.txt
     ```

5. **Run the application**
   - In the same command prompt window:
     ```
     python pdf_to_excel_app.py
     ```
   - Or double-click on `run_app.bat` in the File Explorer

## Troubleshooting

**Application won't start**
- Make sure you've extracted all files from the zip archive
- Try running as Administrator (right-click, "Run as Administrator")

**Error during conversion**
- Make sure your PDF is not password protected
- Check if the PDF contains actual tables (not just images)

**Need additional help?**
- Contact support at: support@yourcompany.com
- Include any error messages you see
