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
    echo.
    echo Press any key to close...
    pause >nul
    exit /b 1
)

REM ---- Auto push to GitHub Pages ----
if exist "docs\index.html" (
    echo.
    echo ============================================
    echo   Pushing to GitHub Pages...
    echo ============================================

    git add docs/index.html >nul 2>&1
    git add store_health_core.py store_health_auto.py store_health_html.py store_health_insight.py >nul 2>&1
    git add .gitignore *.bat >nul 2>&1

    for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a/%%b/%%c
    for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
    git commit -m "Update Store Health Dashboard - %TODAY% %NOW%" >nul 2>&1

    if %ERRORLEVEL% EQU 0 (
        git push -u origin main >nul 2>&1
        if %ERRORLEVEL% EQU 0 (
            echo   [OK] Dashboard pushed to GitHub Pages
        ) else (
            echo   [WARN] Push failed - check credentials
        )
    ) else (
        echo   [SKIP] No changes to commit
    )
) else (
    echo.
    echo [SKIP] docs\index.html not found, skipping push
)

echo.
echo ============================================
echo   All done!
echo ============================================
echo.
echo Press any key to close...
pause >nul
