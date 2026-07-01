@echo off
chcp 65001 >nul
cls

echo ========================================
echo   Cunwei Management System Starting...
echo ========================================

cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ first.
    echo https://www.python.org/downloads/
    pause
    exit /b
)

python -c "import flask" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install -r requirements.txt
)

echo.
echo Starting server...
echo Local:    http://localhost:5000
@echo Network:  http://%COMPUTERNAME%:5000
@echo Admin:    admin / admin123456
@echo Press Ctrl+C to stop
@echo ========================================

python app.py
pause
