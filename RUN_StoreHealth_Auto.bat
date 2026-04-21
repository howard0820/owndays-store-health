@echo off
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

"%PYTHON%" store_health_auto.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
)
pause
