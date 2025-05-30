"""
Script to build a standalone Windows executable using PyInstaller.
Run this on a Windows system with PyInstaller installed.
"""
import os
import subprocess
import shutil
import sys
import platform

def check_compatibility():
    """Check if the current Python environment is compatible for building."""
    # Check Python version
    python_version = platform.python_version_tuple()
    major, minor = int(python_version[0]), int(python_version[1])
    
    if major != 3 or minor < 7 or minor > 10:
        print(f"WARNING: This script is tested with Python 3.7-3.10. You are using Python {major}.{minor}.")
        print("This might cause compatibility issues with PyInstaller.")
        
        if not input("Do you want to continue anyway? (y/n): ").lower().startswith('y'):
            print("Build aborted.")
            sys.exit(1)
    
    # Check if running on Windows
    if platform.system() != "Windows":
        print("WARNING: This build script is designed for Windows systems.")
        print(f"You are running on {platform.system()}.")
        
        if not input("Do you want to continue anyway? (y/n): ").lower().startswith('y'):
            print("Build aborted.")
            sys.exit(1)

def build_executable():
    """Build the standalone Windows executable."""
    print("Building standalone Windows executable...")
    
    # First check compatibility
    check_compatibility()
    
    # Install PyInstaller if not already installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.call(['pip', 'install', 'pyinstaller==4.10'])
    
    # Create the spec file with carefully selected dependencies
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Explicitly include necessary packages
added_files = [
    ('requirements.txt', '.'),
]

# Add your icon if it exists
if os.path.exists('icon.ico'):
    icon = 'icon.ico'
else:
    icon = None

a = Analysis(['pdf_to_excel_app.py'],
             pathex=[],
             binaries=[],
             datas=added_files,
             hiddenimports=['pandas', 'openpyxl', 'pdfplumber', 'PIL', 'numpy'],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,  
          [],
          name='PDFTableExtractor',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None,
          icon=icon)
"""
    
    with open("pdf_table_extractor.spec", "w") as f:
        f.write(spec_content)
    
    # Run PyInstaller with clean cache
    print("Running PyInstaller (this may take a few minutes)...")
    subprocess.call(['pyinstaller', '--clean', 'pdf_table_extractor.spec'])
    
    # Create distribution zip including sample files
    print("Creating distribution zip file...")
    if os.path.exists("dist/PDFTableExtractor.exe"):
        # Create a dist folder for packaging
        if not os.path.exists("dist/PDFTableExtractor"):
            os.makedirs("dist/PDFTableExtractor")
        
        # Copy the executable
        shutil.copy("dist/PDFTableExtractor.exe", "dist/PDFTableExtractor/")
        
        # Add sample data folder if it exists
        if os.path.exists("sample_data"):
            if not os.path.exists("dist/PDFTableExtractor/sample_data"):
                os.makedirs("dist/PDFTableExtractor/sample_data")
            for file in os.listdir("sample_data"):
                if file.endswith(".pdf"):
                    shutil.copy(f"sample_data/{file}", f"dist/PDFTableExtractor/sample_data/")
        
        # Add README for users
        with open("dist/PDFTableExtractor/README.txt", "w") as f:
            f.write("PDF Table Extractor\n")
            f.write("===================\n\n")
            f.write("1. Double-click PDFTableExtractor.exe to start the application\n")
            f.write("2. Click 'Browse' to select your PDF file\n")
            f.write("3. Click 'Convert to Excel' to process the PDF\n")
            f.write("4. When complete, you can open the Excel file directly\n\n")
            f.write("Find sample PDFs in the 'sample_data' folder, if included.\n\n")
            f.write("For more information, visit: https://github.com/aayushshah1/PDFTableExtractor\n")
        
        # Create the zip file
        shutil.make_archive("PDFTableExtractor", 'zip', "dist/PDFTableExtractor")
        print(f"Created: {os.path.abspath('PDFTableExtractor.zip')}")
    else:
        print("ERROR: Build failed. Could not find PDFTableExtractor.exe in the dist folder.")
        print("Check the PyInstaller output above for any errors.")
    
    print("\nBuild complete!")
    print("Distribute PDFTableExtractor.zip to your users.")

if __name__ == "__main__":
    build_executable()
