@echo off
chcp 65001
echo [BUILD] ExcelMatcher Windows 빌드 시작...

REM PyInstaller 확인
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] PyInstaller가 설치되지 않았습니다. 설치를 시작합니다...
    pip install pyinstaller
)

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
    --hidden-import "pandas" ^
    --hidden-import "xlwings" ^
    --hidden-import "openpyxl" ^
    --hidden-import "xlsxwriter" ^
    --hidden-import "requests" ^
    --collect-all "Pillow" ^
    --add-data "assets;assets" ^
    --add-data "seller_assets\user_manual_v1.0.0.html;seller_assets" ^
    --icon "assets\app.ico" ^
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
