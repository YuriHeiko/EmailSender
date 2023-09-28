#!/bin/bash

# Set the name of your virtual environment
venv_name="my_venv"

# Check if the virtual environment exists; if not, create it
if [ ! -d "$venv_name" ]; then
    echo "Creating a virtual environment: $venv_name"
    python3 -m venv "$venv_name"
fi

# Activate the virtual environment
source "$venv_name/bin/activate"

# Install Python packages from requirements.txt
echo "Installing Python packages from requirements.txt"
pip install -r requirements.txt

# Run your Python script (replace 'script.py' with your actual script name)
echo "Running your Python script: script.py"
python script.py

# Deactivate the virtual environment
deactivate

echo "Virtual environment deactivated."
