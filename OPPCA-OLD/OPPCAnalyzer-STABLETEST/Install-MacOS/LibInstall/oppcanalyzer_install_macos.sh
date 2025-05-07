#!/bin/bash

# Target Application Folder (Create if it doesn't exist)
APP_FOLDER="OPPCAnalyzer"
mkdir -p "$APP_FOLDER"

# Install Homebrew (Will only execute if actually missing)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install the latest available Python
echo "Installing the latest Python..."
brew install python

# Create virtual environment inside the application folder
echo "Creating virtual environment..."
/usr/local/bin/python3 -m venv "$APP_FOLDER/oppca_venv"

# Activate virtual environment
source "$APP_FOLDER/oppca_venv/bin/activate"

# Upgrade pip
echo "Upgrading pip..."
python3 -m pip install --upgrade pip

# Install or upgrade packages
for pkg in tcl tk openpyxl pandas xlsxwriter; do
    if pip3 show $pkg &> /dev/null; then
        echo "Upgrading $pkg..."
        pip3 install --upgrade $pkg
    else
        echo "Installing $pkg..."
        pip3 install $pkg
    fi
done