@echo off
chcp 65001 >nul
echo ==========================================
echo   村委资料共享管理系统 - 公网访问启动工具
echo ==========================================
echo.

REM 检查cpolar是否存在
if not exist "cpolar.exe" (
    echo [提示] 未找到 cpolar.exe
    echo.
    echo 请按以下步骤操作：
    echo 1. 访问 https://www.cpolar.com/download
    echo 2. 下载 Windows 版本（cpolar-windows-amd64.zip）
    echo 3. 解压后将 cpolar.exe 放到此目录
    echo 4. 重新运行此脚本
    echo.
    pause
    start https://www.cpolar.com/download
    exit /b 1
)

echo [1/2] 启动村委管理系统...
start "村委系统" cmd /c "python app.py"

echo [2/2] 等待5秒后启动公网隧道...
timeout /t 5 /nobreak >nul

echo.
echo ==========================================
echo   公网访问地址将在下方显示
echo   请将 https://xxxx.cpolar.io 发送给用户
echo ==========================================
echo.

cpolar.exe http 5000

pause
