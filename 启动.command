#!/bin/bash
cd "$(dirname "$0")"

echo "正在启动 PDF 转 Word 工具..."

# 检查依赖
if [ ! -f ".deps_installed" ]; then
    echo "首次运行，正在安装依赖（约需几分钟）..."
    pip3 install flask pdf2docx pymupdf python-docx easyocr Pillow werkzeug -q
    touch .deps_installed
    echo "依赖安装完成！"
fi

# 启动服务
python3 app.py &
SERVER_PID=$!

# 等待服务启动后打开浏览器
sleep 2
open http://localhost:8765

# 等待服务结束
wait $SERVER_PID
