@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   Monthly Store Rebalance
echo ============================================
echo.
cd /d "%~dp0"

set "PYTHON=C:\Users\Howard\AppData\Local\Python\bin\python.exe"

if not exist "%PYTHON%" (
    echo [ERROR] Python not found at: %PYTHON%
    pause
    exit /b 1
)

echo [CHECK] Dependencies...
"%PYTHON%" -c "import pandas; import openpyxl; print('[OK]')"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Missing dependencies.
    pause
    exit /b 1
)

echo.
"%PYTHON%" store_rebalance.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
)

echo.
echo Press any key to close...
pause >nul
