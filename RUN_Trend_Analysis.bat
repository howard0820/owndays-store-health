@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   Store Health - Trend Analysis
echo   (Jan/Feb/Mar/Apr 1st + Apr 20th)
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
"%PYTHON%" -c "import selenium; import pandas; import openpyxl; print('[OK] Ready')"
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Missing dependencies.
    pause
    exit /b 1
)

echo.
echo This will download data for 5 dates and analyze them.
echo Already-downloaded files will be reused automatically.
echo Estimated time: ~5-10 min per new download.
echo.
echo Press any key to start, or Ctrl+C to cancel...
pause >nul

echo.
"%PYTHON%" store_health_trend.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
    echo ============================================
)

echo.
echo Press any key to close...
pause >nul
