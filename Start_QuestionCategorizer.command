#!/bin/bash
# ==============================================================================
# Question Categorizer - macOS Quick Launcher
# ==============================================================================
# This script removes the web quarantine flag and launches the app directly.

echo "=================================================================="
echo "  📚 Launching Question Categorizer..."
echo "=================================================================="

# 1. Determine script directory
DIR="$(cd "$(dirname "$0")" && pwd)"

# 2. Locate QuestionCategorizer.app
APP=""
if [ -d "$DIR/QuestionCategorizer.app" ]; then
    APP="$DIR/QuestionCategorizer.app"
elif [ -d "/Applications/QuestionCategorizer.app" ]; then
    APP="/Applications/QuestionCategorizer.app"
fi

if [ -z "$APP" ]; then
    echo "❌ Error: Could not locate QuestionCategorizer.app."
    echo "Please ensure Start_QuestionCategorizer.command is in the same folder as QuestionCategorizer.app."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "[1/2] Clearing macOS security quarantine..."
xattr -cr "$APP" 2>/dev/null || true
chmod +x "$APP/Contents/MacOS/QuestionCategorizer" 2>/dev/null || true

echo "[2/2] Launching application..."
open "$APP"

echo ""
echo "✅ Success! Question Categorizer is starting."
echo "Your default web browser will open in a few moments."
echo "You can close this terminal window."
echo ""
sleep 2
exit 0
