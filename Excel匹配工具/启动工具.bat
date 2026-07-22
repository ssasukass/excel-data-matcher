@echo off
chcp 65001 >nul
title Excel 数据匹配填充工具

echo ============================================
echo  Excel 数据匹配填充工具 - 环境准备
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python 已检测到

:: 安装依赖
echo.
echo 正在安装依赖（首次运行需要联网下载）...
pip install -r "%~dp0requirements.txt" -q
if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装失败，但仍可尝试运行
) else (
    echo [OK] 依赖安装完成
)

echo.
echo 启动图形界面...
start "" python "%~dp0match_excel_gui.py"
timeout /t 2 >nul
exit
