@echo off
setlocal

:: ── Speedy Scandocs — Windows Build Script ───────────────────────────────
:: Produces: build\windows\Output\SpeedyScandocsSetup.exe
::
:: Requirements (run once before using this script):
::   pip install pyinstaller
::   Install Inno Setup 6 from https://jrsoftware.org/isinfo.php
::   Download Tesseract installer to build\windows\Installers\tesseract-ocr-w64-setup.exe
::     from https://github.com/UB-Mannheim/tesseract/wiki

:: Move to repo root (two levels up from this script)
cd /d "%~dp0..\.."

echo.
echo ════════════════════════════════════════════════════
echo  Step 1: Creating clean build environment
echo ════════════════════════════════════════════════════
if not exist "build\.venv\Scripts\python.exe" (
    python -m venv build\.venv
)
build\.venv\Scripts\pip install -q -r requirements.txt pyinstaller

echo.
echo ════════════════════════════════════════════════════
echo  Step 2: PyInstaller — bundling app
echo ════════════════════════════════════════════════════
if exist "dist" rd /s /q dist
build\.venv\Scripts\python -m PyInstaller "build\windows\scandocs.spec" --clean -y
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller failed. See output above.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════
echo  Step 3: Checking Tesseract installer
echo ════════════════════════════════════════════════════
if not exist "build\windows\Installers\tesseract-ocr-w64-setup.exe" (
    echo.
    echo ERROR: Tesseract installer not found.
    echo Please download it from:
    echo   https://github.com/UB-Mannheim/tesseract/wiki
    echo and place it at:
    echo   build\windows\Installers\tesseract-ocr-w64-setup.exe
    pause
    exit /b 1
)
echo Tesseract installer found.

echo.
echo ════════════════════════════════════════════════════
echo  Step 4: Inno Setup — creating installer wizard
echo ════════════════════════════════════════════════════
set ISCC="%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if not exist %ISCC% (
    echo ERROR: Inno Setup not found. Install from https://jrsoftware.org/isinfo.php
    pause
    exit /b 1
)
%ISCC% "build\windows\installer.iss"
if errorlevel 1 (
    echo.
    echo ERROR: Inno Setup failed. See output above.
    pause
    exit /b 1
)

echo.
echo ════════════════════════════════════════════════════
echo  BUILD COMPLETE
echo  Installer: build\windows\Output\SpeedyScandocsSetup.exe
echo ════════════════════════════════════════════════════
echo.
pause
