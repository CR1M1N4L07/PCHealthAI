@echo off
title PC Health AI - Full Build
color 0B
echo.
echo  ================================================
echo    PC Health AI - Full Build System
echo    Step 1: PyInstaller EXE
echo    Step 2: Inno Setup Installer
echo    Created by CRiMiNAL
echo  ================================================
echo.

cd /d "%~dp0"

echo  Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found. Install Python 3.11+ and add it to PATH.
    pause
    exit /b 1
)
echo  [OK] Python found.

echo.
echo  [1/5] Installing Python dependencies...
pip install --upgrade pyinstaller customtkinter psutil wmi pywin32 pillow >nul 2>&1
echo        Done.

echo.
echo  [2/5] Checking icon.ico exists...
if not exist "icon.ico" (
    echo  [ERROR] icon.ico not found. Place icon.ico in this folder then re-run.
    pause
    exit /b 1
)
echo        [OK] Found.

echo.
echo  [3/5] Checking config.json has no API key...
findstr /C:"sk-ant" config.json >nul 2>&1
if not errorlevel 1 (
    echo  [ERROR] config.json contains an API key!
    echo          Edit config.json so it only contains: {}
    echo          Then re-run this script.
    pause
    exit /b 1
)
echo        [OK] config.json is clean.

echo.
echo  [4/5] Building PCHealthAI.exe with PyInstaller...
echo        This takes 1-3 minutes, please wait...
echo.
pyinstaller PCHealthAI.spec --clean --noconfirm
if errorlevel 1 (
    echo  [ERROR] PyInstaller failed. See output above.
    pause
    exit /b 1
)
echo.
echo        [OK] EXE ready: dist\PCHealthAI.exe

echo.
echo  [5/5] Looking for Inno Setup...

set ISCC_PATH=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC_PATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set ISCC_PATH=C:\Program Files\Inno Setup 6\ISCC.exe

if "%ISCC_PATH%"=="" (
    echo.
    echo  [WARN] Inno Setup not installed.
    echo         Download it free from: jrsoftware.org/isdl.php
    echo         Install it, then re-run this script to build the installer.
    echo.
    echo         Your EXE is ready at: dist\PCHealthAI.exe
    echo         You can share that file directly without an installer.
    pause
    exit /b 0
)
echo        [OK] Inno Setup found.

if not exist "installer_output" mkdir installer_output
"%ISCC_PATH%" PCHealthAI_Installer.iss
if errorlevel 1 (
    echo  [ERROR] Inno Setup failed. Check PCHealthAI_Installer.iss and try again.
    pause
    exit /b 1
)

echo.
echo  ================================================
echo    BUILD COMPLETE!
echo.
echo    EXE:       dist\PCHealthAI.exe
echo    Installer: installer_output\PCHealthAI-Setup-v1.0.0.exe
echo.
echo    Upload the installer to your GitHub Release.
echo  ================================================
echo.

start explorer installer_output
pause
