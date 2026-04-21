@echo off
echo ============================================
echo   Push Store Health to GitHub Pages
echo ============================================
echo.
cd /d "%~dp0"

REM Check if docs/index.html exists
if not exist "docs\index.html" (
    echo [ERROR] docs\index.html not found!
    echo Please run RUN_StoreHealth_Auto.bat first.
    pause
    exit /b 1
)

REM Check if git is initialized
if not exist ".git" (
    echo [SETUP] Initializing git repo...
    git init
    git branch -M main

    echo.
    echo [SETUP] Creating .gitignore...
    (
        echo *.xlsx
        echo *.xls
        echo *.pyc
        echo __pycache__/
        echo .env
        echo *.crdownload
    ) > .gitignore

    echo.
    echo ======================================================
    echo   FIRST TIME SETUP - Follow these steps:
    echo ======================================================
    echo.
    echo   1. Go to https://github.com/new
    echo   2. Create a new repo named: owndays-store-health
    echo      - Set to PRIVATE
    echo      - Do NOT add README or .gitignore
    echo   3. Copy the repo URL, then run:
    echo.
    echo      git remote add origin https://github.com/YOUR_USERNAME/owndays-store-health.git
    echo.
    echo   4. Run this script again.
    echo ======================================================
    pause
    exit /b 0
)

REM Check if remote exists
git remote -v | findstr "origin" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] No remote 'origin' configured.
    echo Run: git remote add origin https://github.com/YOUR_USERNAME/owndays-store-health.git
    pause
    exit /b 1
)

echo [1/3] Staging files...
git add docs/index.html
git add store_health_core.py store_health_auto.py store_health_interactive.py store_health_html.py
git add .gitignore
git add *.bat

echo [2/3] Committing...
for /f "tokens=1-3 delims=/ " %%a in ('date /t') do set TODAY=%%a/%%b/%%c
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
git commit -m "Update Store Health Dashboard - %TODAY% %NOW%"

echo [3/3] Pushing to GitHub...
git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo   Push successful!
    echo.
    echo   After first push, enable GitHub Pages:
    echo   1. Go to your repo on GitHub
    echo   2. Settings -> Pages
    echo   3. Source: Deploy from a branch
    echo   4. Branch: main, Folder: /docs
    echo   5. Save
    echo.
    echo   Your dashboard will be at:
    echo   https://YOUR_USERNAME.github.io/owndays-store-health/
    echo ============================================
) else (
    echo.
    echo [ERROR] Push failed. Check your credentials.
    echo If this is your first push, you may need to:
    echo   1. Install Git: https://git-scm.com/download/win
    echo   2. Login: gh auth login (if using GitHub CLI)
    echo   3. Or use: git config credential.helper manager
)

pause
