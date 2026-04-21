@echo off
echo ============================================
echo   Shanghai Additional Replenishment Wishlist
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

"%PYTHON%" shanghai_wishlist.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Script exited with error code: %ERRORLEVEL%
)
pause
