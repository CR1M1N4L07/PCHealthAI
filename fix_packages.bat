@echo off
title PC Health AI – Quick Fix
cd /d "%~dp0"
call venv\Scripts\activate.bat

echo.
echo  Fixing incompatible package versions...
echo.

pip install --upgrade anthropic httpx

echo.
echo  Done! Now double-click start.bat to launch the app.
echo.
pause
