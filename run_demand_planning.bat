@echo off
chcp 65001 >nul
title OWNDAYS Demand Planning

echo ============================================================
echo   OWNDAYS Taiwan Demand Planning
echo   Step 1: Download 4 weekly inventory snapshots
echo   Step 2: Run analysis model
echo ============================================================
echo.

cd /d "%~dp0"

set PYTHON="C:\Users\Howard\Desktop\Python\python.exe"

echo [Check] Python path...
%PYTHON% --version
if errorlevel 1 (
    echo.
    echo ERROR: Python not found at %PYTHON%
    echo Please edit this .bat file and fix the PYTHON path.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Step 1: Downloading 4 inventory snapshots...
echo ============================================================
echo.

%PYTHON% inventory_download_4periods.py
if errorlevel 1 (
    echo.
    echo ERROR: Download failed. Check your network / credentials.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Step 2: Finding downloaded files...
echo ============================================================
echo.

REM Find the 4 most recent inventory files
set W1=
set W2=
set W3=
set W4=

for /f "delims=" %%f in ('dir /b /o-d TW_在庫表_W1_*.xlsx 2^>nul') do if not defined W1 set "W1=%%f"
for /f "delims=" %%f in ('dir /b /o-d TW_在庫表_W2_*.xlsx 2^>nul') do if not defined W2 set "W2=%%f"
for /f "delims=" %%f in ('dir /b /o-d TW_在庫表_W3_*.xlsx 2^>nul') do if not defined W3 set "W3=%%f"
for /f "delims=" %%f in ('dir /b /o-d TW_在庫表_W4_*.xlsx 2^>nul') do if not defined W4 set "W4=%%f"

if not defined W1 (
    echo ERROR: W1 file not found. Download may have failed.
    pause
    exit /b 1
)
if not defined W2 (
    echo ERROR: W2 file not found.
    pause
    exit /b 1
)
if not defined W3 (
    echo ERROR: W3 file not found.
    pause
    exit /b 1
)
if not defined W4 (
    echo ERROR: W4 file not found.
    pause
    exit /b 1
)

echo   W1: %W1%
echo   W2: %W2%
echo   W3: %W3%
echo   W4: %W4%
echo   BP: buying plan_simplified.xlsx

echo.
echo ============================================================
echo   Step 3: Running Demand Planning Model...
echo ============================================================
echo.

%PYTHON% demand_planning_model.py --w1 "%W1%" --w2 "%W2%" --w3 "%W3%" --w4 "%W4%" --bp "buying plan_simplified.xlsx" --out .

if errorlevel 1 (
    echo.
    echo ERROR: Analysis failed. See error above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   DONE! Check the output files:
echo.
dir /b TW_Demand_Planning_*.xlsx TW_Demand_Planning_*.html 2>nul
echo.
echo   Press any key to close...
echo ============================================================
pause
