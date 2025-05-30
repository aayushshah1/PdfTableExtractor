#!/bin/bash

# Change to the directory of this script
cd "$(dirname "$0")"

# Check if virtual environment exists
if [ -d "venv" ]; then
    # Activate virtual environment
    source venv/bin/activate
else
    # Create virtual environment
    python3 -m venv venv
    source venv/bin/activate
    
    # Install dependencies
    pip install -r requirements.txt
fi

# Run the application
python3 pdf_to_excel_app.py

# Deactivate virtual environment
deactivate