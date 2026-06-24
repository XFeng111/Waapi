@echo off
chcp 65001 >nul
title 批量重命名工具
echo ========================================
echo   批量重命名工具 v2.0
echo ========================================
echo.

cd /d "%~dp0"

python "Vo_读表重命名工具.py"

if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序运行失败，请检查 Python 环境和依赖库。
    echo 如未安装 openpyxl，请执行: pip install openpyxl
)

@REM pause
