@echo off
title Traffic Volume Analyzer
color 1F

:: ── Move to the folder where this .bat file lives ────────────────────────────
cd /d "%~dp0"

echo ============================================================
echo   Traffic Volume Analyzer — Setup ^& Launch
echo ============================================================
echo.

:: ── Find Python ──────────────────────────────────────────────────────────────
set PYTHON=

python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    goto :found_python
)

py --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py
    goto :found_python
)

echo  ERROR: Python was not found on this computer.
echo.
echo  Please install Python first:
echo    1. Go to https://www.python.org/downloads/
echo    2. Download the latest version (3.x)
echo    3. Run the installer
echo    4. IMPORTANT: Check the box "Add Python to PATH"
echo    5. Click Install Now
echo    6. Once done, close this window and double-click run.bat again
echo.
pause
exit /b 1

:found_python
for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do set PY_VER=%%v
echo  Found: %PY_VER%
echo.

:: ── Install required packages ────────────────────────────────────────────────
echo  Installing required packages (this may take a moment the first time)...
echo.

%PYTHON% -m pip install openpyxl matplotlib --quiet --upgrade
if errorlevel 1 (
    echo.
    echo  WARNING: Package installation may have had issues.
    echo  The program will still try to launch.
)

echo  Packages ready.
echo.

:: ── Launch the GUI ───────────────────────────────────────────────────────────
echo  Starting Traffic Volume Analyzer...
echo  (Close this window after the application opens, or keep it open to see errors)
echo.

%PYTHON% gui.py

if errorlevel 1 (
    echo.
    echo  The program closed with an error.
    echo  Please take a screenshot of this window and send it for support.
    echo.
    pause
)
