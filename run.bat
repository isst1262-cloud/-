@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found in PATH.
    echo Please install Python from https://python.org and try again.
    pause
    exit /b 1
)

python menu.py

echo.
echo Program closed. Press any key to close this window.
pause >nul
