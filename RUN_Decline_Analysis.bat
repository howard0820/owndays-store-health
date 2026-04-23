@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   Inventory Health Decline Analysis
echo   (Uses already-downloaded data files)
echo ============================================
echo.
cd /d "%~dp0"

set "PYTHON=C:\Users\Howard\AppData\Local\Python\bin\python.exe"

if not exist "%PYTHON%" (
    echo [ERROR] Python not found at: %PYTHON%
    pause
    exit /b 1
)

"%PYTHON%" analyze_decline.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
)

echo.
echo Press any key to close...
pause >nul
