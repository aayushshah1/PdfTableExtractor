#!/bin/bash

# Change to the directory of this script
cd "$(dirname "$0")"

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python is not installed."
    echo "Please visit https://www.python.org/downloads/ to install Python 3."
    read -p "Press enter to exit..."
    exit 1
fi

# Check if virtual environment exists
if [ -d "venv" ]; then
    # Activate virtual environment
    source venv/bin/activate
    
    # Check if pandas is installed (key dependency)
    python3 -c "import pandas" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "Dependencies not found in existing virtual environment."
        echo "Installing dependencies..."
        pip install --upgrade pip
        pip install -r requirements.txt
        if [ $? -ne 0 ]; then
            echo "Failed to install dependencies."
            read -p "Press enter to exit..."
            exit 1
        fi
    fi
else
    # Create virtual environment
    echo "Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "Failed to create virtual environment."
        read -p "Press enter to exit..."
        exit 1
    fi
    source venv/bin/activate
    
    # Install dependencies
    echo "Installing dependencies..."
    pip install --upgrade pip
    pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "Failed to install dependencies."
        read -p "Press enter to exit..."
        exit 1
    fi
fi

# Run the application
echo "Starting PDF Table Extractor..."
python3 pdf_to_excel_app.py
if [ $? -ne 0 ]; then
    echo "Error running the application."
    echo "Please check if all requirements are installed correctly."
    read -p "Press enter to exit..."
    exit 1
fi

# Deactivate virtual environment
deactivate