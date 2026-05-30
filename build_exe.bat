@echo off
REM ============================================
REM  MG 数据生成工具 — Windows 打包脚本
REM  依赖: pip install pyinstaller pandas openpyxl numpy xlrd
REM  输出: dist\MG数据生成工具.exe
REM ============================================

echo [1/3] 检查 PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
)

echo [2/3] 开始打包...
pyinstaller --onefile --windowed --name "MG数据生成工具" ^
    --add-data "generate_tables.py;." ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell ^
    --hidden-import pandas ^
    --hidden-import numpy ^
    --hidden-import xlrd ^
    --clean ^
    --noconfirm ^
    generate_tables_gui.py

if %errorlevel% equ 0 (
    echo [3/3] 打包完成!
    echo.
    echo 输出文件: dist\MG数据生成工具.exe
    explorer /select, dist\MG数据生成工具.exe
) else (
    echo 打包失败，请检查错误信息
    pause
)
