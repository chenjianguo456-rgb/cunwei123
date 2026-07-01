@echo off
cd /d "%~dp0"

echo ==========================================
echo   Starting Cunwei Server + Internet...
echo ==========================================

start "Server" cmd /c "python app.py"

echo Waiting 5 seconds for server...
timeout /t 5 /nobreak >nul

echo ==========================================
echo   Starting Internet Tunnel...
echo   (Copy the HTTPS URL below to phone)
echo ==========================================

cloudflared.exe tunnel --url http://localhost:5000

echo ==========================================
echo   Tunnel closed. Press any key to exit.
pause
