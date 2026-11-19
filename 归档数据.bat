@echo off
chcp 65001 >nul
echo ==============================
echo   考级数据看板 - 数据归档
echo ==============================
echo.
echo 请确保所有Excel文件已关闭！
echo.
pause
echo.
echo 正在执行归档...
echo.
cd /d "%~dp0"
python archive_data.py
echo.
pause
