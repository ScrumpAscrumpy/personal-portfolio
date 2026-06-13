@echo off
chcp 65001 >nul
title 打包 DocFiller 为独立 .exe

pip install pyinstaller python-docx -q

pyinstaller --onefile --windowed ^
    --name "DocFiller" ^
    --icon=NONE ^
    docfiller.py

echo.
echo 打包完成！exe 文件在 dist\DocFiller.exe
echo 双击即可运行，无需安装 Python
pause
