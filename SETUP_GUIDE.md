# PDF Table Extractor - Complete Setup Guide

This guide provides detailed instructions for setting up and running PDF Table Extractor on any computer.

## Getting Started

### Option 1: Using the Pre-built Windows Executable (Easiest)

1. **Download the application**
   - Download the `PDFTableExtractor.zip` file from the [GitHub releases page](https://github.com/aayushshah1/PdfTableExtractor/releases)
   - Right-click on the downloaded zip file and select "Extract All..."
   - Choose a location where you want to extract the files (e.g., your Desktop)

2. **Run the application**
   - Open the extracted folder
   - Double-click on `PDFTableExtractor.exe` to start the application
   - That's it! No installation or setup required!

3. **Process your PDF files**
   - Click "Browse" to select your PDF file
   - The output Excel file location will be set automatically
   - Click "Convert to Excel" to process the file
   - When complete, you can open the generated Excel file directly

### Option 2: Manual Installation (Windows)

1. **Install Python 3.9 or 3.10** (recommended versions)
   - Download Python from [python.org](https://www.python.org/downloads/windows/)
   - Run the installer and check "Add Python to PATH" during installation
   - To verify installation, open Command Prompt and type: `python --version`

2. **Download the application source code**
   - Download the ZIP from [GitHub repository](https://github.com/aayushshah1/PdfTableExtractor)
   - Extract it to a location of your choice (e.g., `C:\PDFTableExtractor`)

3. **Run the application**
   - Double-click on `run_app.bat` in the extracted folder
   - This script will automatically:
     - Create a virtual environment
     - Install all required dependencies
     - Launch the application

### Option 3: Manual Installation (macOS)

1. **Install Python 3.9 or 3.10**
   - Download Python from [python.org](https://www.python.org/downloads/macos/)
   - Run the installer and follow the instructions
   - To verify installation, open Terminal and type: `python3 --version`

2. **Download the application source code**
   - Download the ZIP from [GitHub repository](https://github.com/aayushshah1/PdfTableExtractor)
   - Extract it to a location of your choice (e.g., your Documents folder)

3. **Make the run script executable**
   - Open Terminal and navigate to the application folder:
     ```
     cd /path/to/PDFTableExtractor
     ```
   - Make the script executable:
     ```
     chmod +x run_app.command
     ```

4. **Run the application**
   - Double-click on `run_app.command` in Finder
   - Alternatively, run it from Terminal:
     ```
     ./run_app.command
     ```

### Option 4: Manual Installation (Linux)

1. **Install Python 3.9 or 3.10**
   ```
   sudo apt update
   sudo apt install python3 python3-pip python3-venv
   ```

2. **Download the application source code**
   ```
   git clone https://github.com/aayushshah1/PdfTableExtractor
   cd PdfTableExtractor
   ```

3. **Set up and run the application**
   ```
   chmod +x run_app.command
   ./run_app.command
   ```

## Building the Windows Executable Yourself

If you need to create your own Windows executable:

1. **Set up a Windows machine with Python 3.9 or 3.10**
   - Follow the steps in Option 2 above to install Python

2. **Install PyInstaller**
   ```
   pip install pyinstaller==4.10
   ```

3. **Run the build script**
   ```
   python build_windows_exe.py
   ```

4. **Locate the executable**
   - After building completes, you'll find `PDFTableExtractor.zip` in the project folder
   - This file contains the standalone executable and everything needed to run it

## Troubleshooting

### Windows

**Missing DLLs or modules**
- Make sure you have the Microsoft Visual C++ Redistributable installed
- Download from [Microsoft's website](https://aka.ms/vs/17/release/vc_redist.x64.exe)

**Application crashes immediately**
- Try running the application from Command Prompt to see error messages:
  ```
  cd C:\path\to\PDFTableExtractor
  PDFTableExtractor.exe
  ```

**Application won't start**
- Make sure you've extracted all files from the zip archive
- Try running as Administrator (right-click, "Run as Administrator")

**Error during conversion**
- Make sure your PDF is not password protected
- Check if the PDF contains actual tables (not just images)

### macOS

**"App is from an unidentified developer"**
- Right-click on the application and select "Open" instead of double-clicking
- Then click "Open" in the dialog box

**Permission denied running scripts**
- Make sure you've made the script executable: `chmod +x run_app.command`
- If using the built app: `chmod +x PDFTableExtractor.app/Contents/MacOS/PDFTableExtractor`

### Linux

**Missing dependencies**
- Install additional system libraries:
  ```
  sudo apt install python3-tk libpango-1.0-0 libharfbuzz0b libpangoft2-1.0-0
  ```

**Display issues**
- Make sure you have a desktop environment installed and running

## Common Issues Across All Platforms

**PDF not processing correctly**
- Make sure the PDF contains actual text and tables, not just images
- Check if the PDF is password-protected (remove protection before processing)

**Excel formula errors**
- The GOOGLEFINANCE formulas require Google Sheets to work
- Either import the Excel file into Google Sheets or replace with appropriate Excel formulas

**Dataset too large**
- For very large PDFs, increase your system's available memory or process the PDF in smaller chunks

## Further Assistance

If you continue to experience issues:
- Open an issue on the [GitHub repository](https://github.com/aayushshah1/PdfTableExtractor/issues)
- Include detailed information about your system and the error messages you're seeing
