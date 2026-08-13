@echo off
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found in PATH.
    echo Please install Python from https://python.org and try again.
    pause
    exit /b 1
)

where cloudflared >nul 2>nul
if errorlevel 1 (
    echo [ERROR] cloudflared was not found in PATH.
    echo Try closing and reopening this window first.
    echo If that does not help, install it: winget install Cloudflare.cloudflared
    pause
    exit /b 1
)

python run_mobile.py
pause
