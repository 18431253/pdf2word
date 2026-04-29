@echo off
chcp 65001 >nul
title PDF 转 Word

echo 正在启动 PDF 转 Word 工具...

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo 未检测到 Python，请先安装 Python：
    echo https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b
)

:: 安装依赖（首次运行）
if not exist ".deps_installed" (
    echo 首次运行，正在安装依赖（约需几分钟）...
    pip install flask pdf2docx pymupdf python-docx easyocr Pillow werkzeug PyQt6 -q
    echo. > .deps_installed
    echo 依赖安装完成！
)

:: 启动服务并打开浏览器
echo 正在启动服务...
start "" http://localhost:8765
python -c "
import threading, webbrowser, time
def open_browser():
    time.sleep(1.5)
    webbrowser.open('http://localhost:8765')
threading.Thread(target=open_browser, daemon=True).start()
" >nul 2>&1

python app.py
pause
