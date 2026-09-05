#!/bin/bash
# ── 星易查 - 围串标风险识别系统 启动脚本 ──

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PORT=5001
URL="http://localhost:${PORT}"

cd "$PROJECT_DIR" || exit 1

# 检查是否已在运行
if lsof -ti:"$PORT" &>/dev/null; then
    echo "✅ 服务器已在运行中 (端口 $PORT)"
    echo "🌐 正在打开浏览器..."
    open "$URL"
    exit 0
fi

# 启动服务器
clear
echo "╔══════════════════════════════════════════╗"
echo "║                                          ║"
echo "║     🚀  星 易 查                        ║"
echo "║     围串标风险识别分析系统               ║"
echo "║                                          ║"
echo "╚══════════════════════════════════════════╝"
echo ""
echo "  📡 正在启动服务器..."
echo "  🌐 端口: $PORT"
echo "  🔗 地址: $URL"
echo ""
echo "  ⚠️  关闭此窗口将停止服务器"
echo ""

# 后台启动并等待就绪后打开浏览器
./venv/bin/python3 app.py "$PORT" &
SERVER_PID=$!

# 等待服务器就绪
for i in $(seq 1 30); do
    if curl -s -o /dev/null "$URL" 2>/dev/null; then
        echo "✅ 服务器启动成功！"
        echo "🌐 正在打开浏览器..."
        open "$URL"
        break
    fi
    sleep 1
done

# 前台等待服务器进程
wait $SERVER_PID
