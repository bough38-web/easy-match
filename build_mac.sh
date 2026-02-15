#!/bin/bash

APP_NAME="ExcelMatcher_v4.8"

echo "[BUILD] Starting macOS Build for $APP_NAME..."

# Check PyInstaller
if ! command -v pyinstaller &> /dev/null
then
    echo "[ERROR] PyInstaller could not be found. Please install: pip install pyinstaller"
    exit 1
fi

# Clean
rm -rf dist build *.spec

# Build
# --windowed: Creates .app bundle
# --onedir: Generally better for macOS .app bundles than --onefile for startup speed, 
# but --onefile is cleaner for distribution if not signing. 
# Let's use --windowed (which implies .app)

echo "[BUILD] Running PyInstaller..."
pyinstaller --noconfirm --windowed --clean \
    --name "$APP_NAME" \
    --hidden-import "pandas" \
    --hidden-import "xlwings" \
    --hidden-import "openpyxl" \
    --hidden-import "xlsxwriter" \
    main.py

if [ $? -eq 0 ]; then
    echo "[SUCCESS] Build finished!"
    echo "App bundle: dist/$APP_NAME.app"
    
    # Optional: Create DMG (requires create-dmg)
    # echo "You can now package dist/$APP_NAME.app into a DMG."
else
    echo "[ERROR] Build failed."
    exit 1
fi
