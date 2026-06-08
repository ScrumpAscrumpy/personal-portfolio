#!/bin/bash
# DocFiller launcher for macOS / Linux

cd "$(dirname "$0")"

# Check Python3
if ! command -v python3 &>/dev/null; then
    osascript -e 'display alert "请先安装 Python 3.8+" as critical'
    exit 1
fi

# Install dependency if missing
python3 -c "import docx" 2>/dev/null || pip3 install python-docx -q

# Launch
python3 docfiller.py
