@echo off
echo Starting PDF Table Extractor...

cd /d "%~dp0"

REM Check if Python is installed
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please visit https://www.python.org/downloads/ to install Python,
    echo or use the standalone executable version instead.
    pause
    exit /b 1
)

REM Check if virtual environment exists
if exist venv\ (
    REM Activate virtual environment
    call venv\Scripts\activate.bat
) else (
    REM Create virtual environment
    python -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
    call venv\Scripts\activate.bat
    
    REM Install dependencies
    pip install -r requirements.txt
    if %ERRORLEVEL% neq 0 (
        echo Failed to install dependencies.
        pause
        exit /b 1
    )
)

REM Run the application
python pdf_to_excel_app.py
if %ERRORLEVEL% neq 0 (
    echo Error running the application.
    echo Please check if all requirements are installed correctly.
    pause
    exit /b 1
)

REM Deactivate virtual environment
call venv\Scripts\deactivate.bat
exit /b 0
