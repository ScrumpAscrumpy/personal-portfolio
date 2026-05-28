#!/usr/bin/env python3
"""DocFiller - 批量文档填写系统 入口"""
import os
import sys

# Make sure the app package is importable
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app.ui.main_window import MainWindow


def main():
    app = MainWindow()
    app.run()


if __name__ == '__main__':
    main()
