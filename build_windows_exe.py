"""
Script to build a standalone Windows executable using PyInstaller.
Run this on a Windows system with PyInstaller installed.
"""
import os
import subprocess
import shutil

def build_executable():
    print("Building standalone Windows executable...")
    
    # Install PyInstaller if not already installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.call(['pip', 'install', 'pyinstaller'])
    
    # Create the spec file
    spec_content = """
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['pdf_to_excel_app.py'],
             pathex=[],
             binaries=[],
             datas=[],
             hiddenimports=['pandas', 'openpyxl'],
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
          entitlements_file=None , icon='icon.ico')
"""
    
    with open("pdf_table_extractor.spec", "w") as f:
        f.write(spec_content)
    
    # Run PyInstaller
    print("Running PyInstaller (this may take a few minutes)...")
    subprocess.call(['pyinstaller', 'pdf_table_extractor.spec', '--clean'])
    
    # Create distribution zip
    print("Creating distribution zip file...")
    if os.path.exists("dist/PDFTableExtractor"):
        shutil.make_archive("PDFTableExtractor", 'zip', "dist/PDFTableExtractor")
        print(f"Created: {os.path.abspath('PDFTableExtractor.zip')}")
    
    print("\nBuild complete!")
    print("Distribute PDFTableExtractor.zip to your users.")

if __name__ == "__main__":
    build_executable()
