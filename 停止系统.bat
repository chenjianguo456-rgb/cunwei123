@echo off
chcp 65001 >nul 2>&1
title 停止村委系统
color 0C

echo ═══════════════════════════════════════════════════════════
echo   停止村委资料共享系统
echo ═══════════════════════════════════════════════════════════
echo.

:: 停止SSH隧道
echo [1/2] 断开公网隧道...
taskkill /F /IM ssh.exe >nul 2>&1
echo     ✓ 已断开

:: 停止Flask
echo [2/2] 停止村委系统程序...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
    echo     ✓ 已停止进程 PID: %%a
)

echo.
echo ═══════════════════════════════════════════════════════════
echo  系统已完全停止
echo ═══════════════════════════════════════════════════════════
pause
