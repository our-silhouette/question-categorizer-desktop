#!/bin/bash
# ==============================================================================
# Local Build Script for Question Categorizer Desktop App (macOS / Linux)
# ==============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== 1. Checking Python Environment ==="
if ! command -v python3 &> /dev/null; then
    echo "Error: Python 3 is not installed or not in PATH."
    exit 1
fi

echo "Python version: $(python3 --version)"

echo "=== 2. Setting up virtual environment ==="
if [ ! -d "build_venv" ]; then
    python3 -m venv build_venv
fi
source build_venv/bin/activate

echo "=== 3. Installing dependencies ==="
pip install --upgrade pip
if [ -f "requirements.txt" ]; then
    pip install -r requirements.txt
elif [ -f "../Streamlit approach/requirements.txt" ]; then
    pip install -r "../Streamlit approach/requirements.txt"
fi
pip install pyinstaller

echo "=== 4. Building with PyInstaller ==="
pyinstaller --clean desktop.spec

echo "=== 5. Packaging into ZIP ==="
cd dist
if [ -d "QuestionCategorizer.app" ]; then
    zip -r -y ../QuestionCategorizer-macOS.zip QuestionCategorizer.app
    echo "Successfully built: QuestionCategorizer-macOS.zip"
elif [ -d "QuestionCategorizer" ]; then
    zip -r -y ../QuestionCategorizer.zip QuestionCategorizer
    echo "Successfully built: QuestionCategorizer.zip"
fi
cd ..

echo "=== Build Complete! ==="
echo "You can now test the app by opening: dist/QuestionCategorizer.app"
