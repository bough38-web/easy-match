@echo off
chcp 65001
echo [BUILD] Starting Windows Build for ExcelMatcher...

REM Check PyInstaller
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller not found. Installing...
    pip install pyinstaller
)

REM Clean previous build
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"

REM Build Command
REM --noconfirm: overwrite existing
REM --onefile: single exe
REM --windowed: hide terminal
REM --name: output filename
REM --add-data: include additional files
REM --hidden-import: ensure dependencies are included

echo [BUILD] Running PyInstaller...
pyinstaller --noconfirm --onefile --windowed ^
    --name "EasyMatch_v1.0" ^
    --hidden-import "pandas" ^
    --hidden-import "xlwings" ^
    --hidden-import "openpyxl" ^
    --hidden-import "xlsxwriter" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL.Image" ^
    --hidden-import "PIL.ImageTk" ^
    --add-data "assets;assets" ^
    --add-data "presets.json;." ^
    --add-data "replacements.json;." ^
    --add-data "usage_guide.html;." ^
    --icon "assets\app.ico" ^
    main.py

if %errorlevel% neq 0 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo [SUCCESS] Build complete! Check 'dist/EasyMatch_v1.0.exe'
echo.
echo You can now distribute the EXE file to users.
echo The EXE is standalone and includes all dependencies.
pause
