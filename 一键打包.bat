@echo off
chcp 65001 >nul
echo ========================================
echo   PDF转Word 一键打包工具
echo ========================================
echo.

echo [1/3] 正在检查并安装Python依赖...
pip install flask pdf2docx pymupdf python-docx Pillow werkzeug pyinstaller -q
if errorlevel 1 (
    echo 安装失败，请确保已安装Python：https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo [2/3] 正在打包（请耐心等待，首次可能需要3-5分钟）...
pyinstaller --onefile --windowed --name "PDF转Word" --add-data "templates;templates" --add-data "converter.py;." launcher.py

echo.
echo [3/3] 完成！
echo.
echo ========================================
echo   打包完成！
echo   exe文件在 dist\PDF转Word.exe
echo   双击即可运行！
echo ========================================
pause
