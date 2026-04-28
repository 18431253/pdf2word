"""
PDF转Word 一键启动器
双击此文件即可运行
"""
import os
import sys
import webbrowser
import threading
import time

# 隐藏控制台窗口 (Windows)
import subprocess
if sys.platform == 'win32':
    import ctypes
    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    SW_HIDE = 0
    hInstance = kernel32.GetModuleHandleW(None)
    ClassName = 'PDF2WordConsole'
    wc = kernel32.GetConsoleWindow()

def open_browser():
    """延迟打开浏览器"""
    time.sleep(1.5)
    webbrowser.open('http://localhost:8765')

if __name__ == '__main__':
    # 打开浏览器
    threading.Thread(target=open_browser, daemon=True).start()
    
    # 启动Flask服务
    from app import app
    print("=" * 50)
    print("PDF转Word 服务已启动！")
    print("浏览器将自动打开...")
    print("关闭时直接关闭此窗口即可")
    print("=" * 50)
    app.run(debug=False, host="0.0.0.0", port=8765)
