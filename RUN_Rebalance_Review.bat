@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ============================================================
echo   OWNDAYS Taiwan - Rebalance Effectiveness Review
echo ============================================================
echo.

REM --- Find rebalance file ---
set "RB_FILE="
for /f "delims=" %%F in ('dir /b /o-d Rebalance_20*.xlsx 2^>nul') do (
    if not defined RB_FILE set "RB_FILE=%%F"
)
if not defined RB_FILE (
    echo ERROR: No Rebalance_YYYY-MM-DD.xlsx found!
    echo Please ensure the rebalance plan Excel is in this folder.
    pause
    exit /b 1
)
echo   Rebalance plan: %RB_FILE%

REM --- Find post-rebalance inventory file ---
set "POST_FILE="
for /f "delims=" %%F in ('dir /b /o-d StoreHealth_frames_*.xlsx 2^>nul') do (
    if not defined POST_FILE set "POST_FILE=%%F"
)
if not defined POST_FILE (
    echo ERROR: No StoreHealth_frames_*.xlsx found!
    echo Please run Store Health first to get the latest inventory data.
    pause
    exit /b 1
)
echo   Post-rebalance: %POST_FILE%
echo.

REM --- Run review ---
"C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe" rebalance_review.py --rebalance "%RB_FILE%" --post "%POST_FILE%"

echo.
if %ERRORLEVEL% EQU 0 (
    echo   Review complete! Check the output Excel file.
) else (
    echo   ERROR: Review failed. Check the error messages above.
)
echo.
pause
