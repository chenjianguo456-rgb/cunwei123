@echo off
chcp 65001 >nul 2>&1
title 上传数据到 GitHub
color 0B

echo ═══════════════════════════════════════════════════════════
echo   上传本地数据到 GitHub（备份）
echo ═══════════════════════════════════════════════════════════
echo.

cd /d "%~dp0"

echo [1/4] 检查数据库文件...
if not exist "instance\cunwei.db" (
    echo     ✗ 未找到数据库文件 instance\cunwei.db
    echo     请先运行程序并录入数据
    pause
    exit /b 1
)
echo     ✓ 找到数据库文件

echo [2/4] 暂存所有文件...
git add -A
if %errorlevel% neq 0 (
    echo     ✗ Git 暂存失败
    pause
    exit /b 1
)
echo     ✓ 已暂存

echo [3/4] 提交更改...
git commit -m "备份数据: %date% %time%"
if %errorlevel% neq 0 (
    echo     没有新的更改需要提交
    goto :push
)
echo     ✓ 已提交

:push
echo [4/4] 推送到 GitHub...
git push origin main
if %errorlevel% neq 0 (
    echo     ✗ 推送失败，请检查网络连接
    pause
    exit /b 1
)

echo.
echo ═══════════════════════════════════════════════════════════
echo  ✓ 数据已成功上传到 GitHub！
echo  仓库地址: https://github.com/chenjianguo456-rgb/cunwei123
echo ═══════════════════════════════════════════════════════════
pause
