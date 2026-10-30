@echo off
REM ==============================================================================
REM Local Build Script for Question Categorizer Desktop App (Windows)
REM ==============================================================================

echo === 1. Checking Python Environment ===
python --version
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH.
    exit /b 1
)

echo === 2. Setting up build virtual environment ===
if not exist "build_venv" (
    python -m venv build_venv
)
call build_venv\Scripts\activate.bat

echo === 3. Installing dependencies ===
python -m pip install --upgrade pip
if exist "requirements.txt" (
    pip install -r requirements.txt
) else if exist "..\Streamlit approach\requirements.txt" (
    pip install -r "..\Streamlit approach\requirements.txt"
)
pip install pyinstaller

echo === 4. Building with PyInstaller ===
pyinstaller --clean desktop.spec

echo === 5. Compressing into ZIP ===
powershell -Command "Compress-Archive -Path dist\QuestionCategorizer -DestinationPath QuestionCategorizer-Windows.zip -Force"

echo === Build Complete: QuestionCategorizer-Windows.zip ===
pause
