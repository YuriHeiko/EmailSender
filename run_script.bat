@echo off

:: Set the name of your virtual environment
set "venv_name=my_venv"

:: Check if the virtual environment exists; if not, create it
if not exist "%venv_name%\Scripts\activate.bat" (
    echo Creating a virtual environment: %venv_name%
    python -m venv "%venv_name%"
)

:: Activate the virtual environment
call "%venv_name%\Scripts\activate"

:: Install Python packages from requirements.txt
echo Installing Python packages from requirements.txt
pip install -r requirements.txt

:: Run your Python script (replace 'script.py' with your actual script name)
echo Running your Python script: script.py
python script.py

:: Deactivate the virtual environment
deactivate

echo Virtual environment deactivated.
