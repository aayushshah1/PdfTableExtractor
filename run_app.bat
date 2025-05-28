@echo off
echo Starting PDF Table Extractor...
cd /d "%~dp0"
call venv\Scripts\activate.bat
python pdf_to_excel_app.py
if %ERRORLEVEL% NEQ 0 (
    echo Error running the application.
    echo Please make sure you've installed all requirements by running:
    echo pip install -r requirements.txt
    pause
) else (
    exit
)
