#!/bin/bash
# ============================================
#  MG 数据生成工具 — macOS 打包脚本
#  依赖: pip3 install pyinstaller pandas openpyxl numpy xlrd
#  输出: dist/MG数据生成工具
# ============================================

set -e

echo "[1/3] 检查 PyInstaller..."
pip3 show pyinstaller > /dev/null 2>&1 || pip3 install pyinstaller

echo "[2/3] 开始打包..."
pyinstaller --onefile --windowed --name "MG数据生成工具" \
    --add-data "generate_tables.py:." \
    --hidden-import openpyxl \
    --hidden-import openpyxl.cell \
    --hidden-import pandas \
    --hidden-import numpy \
    --hidden-import xlrd \
    --clean \
    --noconfirm \
    generate_tables_gui.py

echo "[3/3] 打包完成!"
echo ""
echo "输出文件: dist/MG数据生成工具"
echo "(在 Finder 中可右键 → 打开，首次运行需要允许来自非 App Store 的应用)"
