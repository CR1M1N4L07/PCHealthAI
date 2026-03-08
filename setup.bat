@echo off
title PC Health AI – Setup
color 0B

REM ── Switch to the folder where this .bat file lives ──
cd /d "%~dp0"

echo.
echo  =====================================================
echo    PC Health AI – One-Time Setup
echo  =====================================================
echo.

REM ── Check Python ──────────────────────────────────────
python --version >NUL 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  [ERROR] Python is not installed or not in PATH.
    echo.
    echo  Please install Python 3.11 or newer from:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANT: During installation tick the box that says
    echo             "Add Python to PATH" before clicking Install.
    echo.
    pause
    exit /b 1
)

echo  [OK] Python found:
python --version
echo.

REM ── Create virtual environment ────────────────────────
echo  Creating virtual environment in "venv" folder...
python -m venv venv
if %ERRORLEVEL% NEQ 0 (
    echo  [ERROR] Failed to create virtual environment.
    pause
    exit /b 1
)
echo  [OK] Virtual environment created.
echo.

REM ── Activate venv ─────────────────────────────────────
call venv\Scripts\activate.bat

REM ── Upgrade pip silently ──────────────────────────────
echo  Upgrading pip...
python -m pip install --upgrade pip --quiet
echo  [OK] pip upgraded.
echo.

REM ── Install requirements ──────────────────────────────
echo  Installing required packages (this may take 1-2 minutes)...
echo.
pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERROR] Package installation failed.
    echo  Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo  =====================================================
echo    Setup Complete!
echo  =====================================================
echo.
echo  To launch the app, double-click  start.bat
echo.
pause
