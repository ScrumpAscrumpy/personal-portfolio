@echo off
chcp 65001 >nul
title DocFiller

:: Check Python
where python >nul 2>&1
if errorlevel 1 (
    echo Python 未安装，请从 https://python.org 下载安装 Python 3.8+
    pause
    exit
)

:: Install dependency silently if missing
python -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖 python-docx ...
    pip install python-docx -q
)

:: Launch app
python "%~dp0docfiller.py"
