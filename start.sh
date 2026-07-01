#!/bin/bash
# 村委资料共享系统一键启动脚本
# 适用于 Linux / Mac / Git Bash

cd "$(dirname "$0")"

echo "========================================"
echo "  🏘️ 村委资料共享系统 启动中..."
echo "========================================"

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未检测到 python3，请先安装 Python 3.10+"
    exit 1
fi

# 检查依赖
if ! python3 -c "import flask" 2>/dev/null; then
    echo "📦 正在安装依赖..."
    pip3 install -r requirements.txt
fi

echo "🚀 启动服务..."
echo "💻 本机访问：http://localhost:5000"
echo "📱 局域网访问：http://$(hostname -I | awk '{print $1}'):5000"
echo ""
echo "🔐 管理员账户：admin / admin123456"
echo "⚠️  按 Ctrl+C 停止服务"
echo "========================================"

python3 app.py
