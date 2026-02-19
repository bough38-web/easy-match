@echo off
chcp 65001
echo [BUILD] ExcelMatcher Windows 빌드 시작...

REM PyInstaller 및 주요 의존성 확인/설치
echo [SETUP] Installing critical dependencies...
pip install requests Pillow pyinstaller pandas openpyxl xlsxwriter xlwings rapidfuzz python-calamine --upgrade

REM 이전 빌드 정리
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build Command
REM --noconfirm: overwrite existing
REM --onefile: single exe
REM --windowed: hide terminal
REM --name: output filename
REM --add-data: include additional files
REM --hidden-import: ensure dependencies are included



REM Build using CLI arguments (avoiding Spec file Unicode path issues)
echo [BUILD] PyInstaller 실행 중 (CLI Mode)...
pyinstaller --noconfirm --onedir --windowed ^
    --name "ExcelMatcher_v1.0.0" ^
    --add-data "seller_assets;seller_assets" ^
    --add-data "assets;assets" ^
    --add-data "presets.json;." ^
    --add-data "replacements.json;." ^
    --hidden-import "pandas" ^
    --hidden-import "xlwings" ^
    --hidden-import "openpyxl" ^
    --hidden-import "xlsxwriter" ^
    --hidden-import "requests" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageTk" ^
    --hidden-import "rapidfuzz" ^
    --hidden-import "calamine" ^
    --collect-all "Pillow" ^
    --exclude-module "PyQt5" ^
    --exclude-module "PyQt6" ^
    --exclude-module "qtpy" ^
    --exclude-module "QtPy" ^
    --exclude-module "jupyter" ^
    --exclude-module "notebook" ^
    --exclude-module "scipy" ^
    --exclude-module "matplotlib" ^
    --exclude-module "IPython" ^
    --exclude-module "sympy" ^
    --exclude-module "astropy" ^
    --icon "assets\logo.png" ^
    main.py

if %errorlevel% neq 0 (
    echo [오류] 빌드 실패!
    pause
    exit /b 1
)

echo [성공] 빌드 완료! 'dist/ExcelMatcher' 폴더를 확인하세요.
echo.
echo 이제 EXE 파일을 사용자에게 배포할 수 있습니다.
echo EXE 파일은 독립 실행형이며 모든 의존성을 포함합니다.
pause
