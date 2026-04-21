@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   Store Health Check - Auto Mode
echo ============================================
echo.
cd /d "%~dp0"

set "PYTHON=C:\Users\Howard\AppData\Local\Python\bin\python.exe"

if not exist "%PYTHON%" (
    echo [ERROR] Python not found at: %PYTHON%
    echo Please update PYTHON path in this .bat file.
    pause
    exit /b 1
)

echo [CHECK] Verifying Python and dependencies...
"%PYTHON%" -c "import selenium; import pandas; import openpyxl; print('[OK] All dependencies found')"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Missing Python dependencies. Install with:
    echo   "%PYTHON%" -m pip install selenium pandas openpyxl webdriver-manager
    echo.
    pause
    exit /b 1
)

echo.
echo [RUN] Starting store_health_auto.py ...
echo.
"%PYTHON%" store_health_auto.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
    echo ============================================
)

echo.
echo Press any key to close...
pause >nul
