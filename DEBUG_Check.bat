@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   Store Health - Diagnostic Check
echo ============================================
echo.
cd /d "%~dp0"
echo [1] Current directory: %CD%
echo.

set "PYTHON=C:\Users\Howard\AppData\Local\Python\bin\python.exe"

echo [2] Python path: %PYTHON%
if exist "%PYTHON%" (
    echo     -> File exists: YES
) else (
    echo     -> File exists: NO  *** THIS IS THE PROBLEM ***
    echo.
    echo Please check your Python installation path.
    pause
    exit /b 1
)

echo.
echo [3] Python version:
"%PYTHON%" --version
echo.

echo [4] Checking dependencies...
"%PYTHON%" -c "import selenium; print('    selenium: OK')"
"%PYTHON%" -c "import pandas; print('    pandas: OK')"
"%PYTHON%" -c "import openpyxl; print('    openpyxl: OK')"
"%PYTHON%" -c "import webdriver_manager; print('    webdriver_manager: OK')"
"%PYTHON%" -c "import numpy; print('    numpy: OK')"
echo.

echo [5] Quick import test (store_health_core)...
"%PYTHON%" -c "import store_health_core; print('    store_health_core: OK')"
if %ERRORLEVEL% NEQ 0 (
    echo     *** IMPORT ERROR in store_health_core.py ***
)
echo.

echo [6] Quick import test (store_health_html)...
"%PYTHON%" -c "import store_health_html; print('    store_health_html: OK')"
if %ERRORLEVEL% NEQ 0 (
    echo     *** IMPORT ERROR in store_health_html.py ***
)
echo.

echo [7] Checking for existing data files...
if exist "StoreHealth_*.xlsx" (
    echo     Found data files:
    dir /b StoreHealth_*.xlsx
) else (
    echo     No StoreHealth data files found (will need to download)
)
echo.

echo [8] Checking docs folder...
if exist "docs\index.html" (
    echo     docs\index.html exists
) else (
    echo     docs\index.html not found (will be created on first run)
)

echo.
echo ============================================
echo   Diagnostic complete!
echo ============================================
echo.
echo If all checks show OK, try running:
echo   RUN_StoreHealth_Auto.bat
echo.
echo If there's an import error, copy the error message
echo and send it to Claude for help.
echo.
pause
