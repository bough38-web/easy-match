#!/bin/bash

APP_NAME="ExcelMatcher_v1.0.0"

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
pip install requests Pillow pyinstaller pandas openpyxl xlsxwriter xlwings rapidfuzz python-calamine tkinterdnd2 --upgrade

# Build using CLI arguments (avoiding Spec file Unicode path issues)
echo "[BUILD] PyInstaller 실행 중 (CLI Mode)..."
pyinstaller --noconfirm --windowed --clean \
    --name "ExcelMatcher_v1.0.0" \
    --add-data "seller_assets/user_manual_v1.0.0.html:seller_assets" \
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
    --hidden-import "rapidfuzz" \
    --hidden-import "calamine" \
    --hidden-import "tkinterdnd2" \
    --collect-all "Pillow" \
    --collect-all "tkinterdnd2" \
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
    echo "앱 번들: dist/$APP_NAME.app"
    
    # Create DMG using hdiutil (standard on macOS)
    echo "[BUILD] DMG 파일 제작 중 (Staging structure)..."
    DMG_NAME="${APP_NAME}.dmg"
    rm -f "dist/$DMG_NAME"
    
    # Breakthrough: Use a staging folder so the .app itself is in the DMG root
    rm -rf "dist/dmg_staging"
    mkdir -p "dist/dmg_staging"
    cp -R "dist/$APP_NAME.app" "dist/dmg_staging/"
    
    hdiutil create -volname "$APP_NAME" -srcfolder "dist/dmg_staging" -ov -format UDZO "dist/$DMG_NAME"
    
    rm -rf "dist/dmg_staging"
    
    if [ $? -eq 0 ]; then
        echo "[성공] 배포용 DMG 제작 완료: dist/$DMG_NAME"
    else
        echo "[경고] DMG 제작 실패."
    fi
else
    echo "[오류] 빌드 실패."
    exit 1
fi
