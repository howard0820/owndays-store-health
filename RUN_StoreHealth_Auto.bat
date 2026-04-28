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
REM IMPORTANT: only the published output (docs/index.html) is auto-committed.
REM Source code (.py / .bat) must be committed manually — past incident 2026-04-27:
REM auto-staging .py files pushed an empty store_health_auto.py to origin/main.
if exist "docs\index.html" (
    echo.
    echo ============================================
    echo   Pushing to GitHub Pages...
    echo ============================================

    REM Sanity check — refuse to push if the dashboard is suspiciously small.
    for %%A in ("docs\index.html") do set "INDEX_SIZE=%%~zA"
    if %INDEX_SIZE% LSS 100000 (
        echo   [ABORT] docs\index.html is only %INDEX_SIZE% bytes ^(min 100000^).
        echo   Refusing to push a broken dashboard.
        goto :skip_push
    )

    git add docs/index.html >nul 2>&1

    for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a/%%b/%%c
    for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
    git commit -m "Update Store Health Dashboard - %TODAY% %NOW%"

    if %ERRORLEVEL% EQU 0 (
        echo   Pushing to origin/main...
        git push -u origin main
        if %ERRORLEVEL% EQU 0 (
            echo.
            echo   Push successful! GitHub Pages will update in ~1 min.
            echo   https://howard0820.github.io/owndays-store-health/
        ) else (
            echo.
            echo   [ERROR] Push failed. Check your git credentials.
        )
    ) else (
        echo   No changes to commit ^(dashboard unchanged^).
    )
)

:skip_push
echo.
echo ============================================
echo   Done!
echo ============================================
pause