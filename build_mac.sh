#!/bin/bash

APP_NAME="ExcelMatcher_v1.0.7"

echo "[BUILD] $APP_NAME macOS 빌드 시작..."

# PyInstaller 확인
if ! command -v pyinstaller &> /dev/null
then
    echo "[오류] PyInstaller를 찾을 수 없습니다. 다음 명령어로 설치하세요: pip install pyinstaller"
    exit 1
fi

# 정리
rm -rf dist build *.spec

# Build
# --windowed: Creates .app bundle
# --onedir: Generally better for macOS .app bundles than --onefile for startup speed, 
# but --onefile is cleaner for distribution if not signing. 
# Let's use --windowed (which implies .app)

# 0. Pre-check & Install Critical Dependencies
echo "[SETUP] Installing critical dependencies..."
pip install requests --upgrade
pip install Pillow --upgrade
pip install pyinstaller --upgrade

# Build using CLI arguments (avoiding Spec file Unicode path issues)
echo "[BUILD] PyInstaller 실행 중 (CLI Mode)..."
pyinstaller --noconfirm --windowed --clean \
    --name "ExcelMatcher_v1.0.7" \
    --add-data "usage_guide.html:." \
    --add-data "assets:assets" \
    --add-data "presets.json:." \
    --add-data "replacements.json:." \
    --hidden-import "pandas" \
    --hidden-import "xlwings" \
    --hidden-import "openpyxl" \
    --hidden-import "xlsxwriter" \
    --hidden-import "requests" \
    --hidden-import "PIL" \
    --hidden-import "PIL.Image" \
    --hidden-import "PIL.ImageTk" \
    --collect-all "Pillow" \
    --exclude-module "PyQt5" \
    --exclude-module "PyQt6" \
    --exclude-module "qtpy" \
    --exclude-module "QtPy" \
    --exclude-module "jupyter" \
    --exclude-module "notebook" \
    --exclude-module "scipy" \
    --exclude-module "matplotlib" \
    --exclude-module "IPython" \
    --exclude-module "sympy" \
    --exclude-module "astropy" \
    main.py

if [ $? -eq 0 ]; then
    echo "[성공] 빌드 완료!"
    echo "앱 번들: dist/ExcelMatcher_v1.0.7.app"
    
    # Optional: Create DMG (requires create-dmg)
    # echo "You can now package dist/$APP_NAME.app into a DMG."
else
    echo "[오류] 빌드 실패."
    exit 1
fi
