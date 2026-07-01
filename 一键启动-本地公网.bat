@echo off
chcp 65001 >nul 2>&1
title 村委资料共享系统 - 本地部署+公网访问
color 0A

echo ═══════════════════════════════════════════════════════════
echo   村委资料共享系统 - 本地部署 + 公网访问 一键启动
echo ═══════════════════════════════════════════════════════════
echo.

:: 检查是否已有Flask在运行
echo [1/3] 检查程序运行状态...
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo     ✓ 程序已在运行，跳过启动
    goto :start_tunnel
)

:: 启动Flask程序
echo [2/3] 正在启动村委系统...
cd /d "%~dp0"
start /b "" "C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe" app.py > flask_run.log 2>&1

:: 等待程序启动
echo     等待程序启动...
set /a count=0
:wait_start
timeout /t 1 /nobreak >nul
set /a count+=1
netstat -ano | findstr ":5000" | findstr "LISTENING" >nul 2>&1
if %errorlevel%==0 (
    echo     ✓ 程序启动成功！
    goto :start_tunnel
)
if %count% lss 15 goto :wait_start
echo     ✗ 程序启动超时，请检查 flask_run.log
pause
exit /b 1

:start_tunnel
echo [3/3] 正在建立公网访问隧道...
echo.
echo ───────────────────────────────────────────────────────────
echo  正在连接公网隧道服务（localhost.run）...
echo  请勿关闭此窗口！关闭窗口将断开公网访问。
echo ───────────────────────────────────────────────────────────
echo.

:: 启动SSH隧道（localhost.run免费服务，无需注册）
ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=60 -R 80:localhost:5000 nokey@localhost.run

echo.
echo ═══════════════════════════════════════════════════════════
echo  公网隧道已断开！
echo  如需重新连接，请再次运行此脚本。
echo ═══════════════════════════════════════════════════════════
pause
