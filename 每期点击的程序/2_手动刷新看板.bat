@echo off
chcp 65001 >nul
cd /d "%~dp0.."
echo ============================================================
echo  手动刷新看板
echo  说明：用现有数据重新生成看板（不下载BI），用于验证配置效果
echo ============================================================
echo.
python generate_dashboard.py
if %errorlevel%==0 (
    echo.
    echo [成功] 看板已刷新，请打开 dashboard.html 查看
) else (
    echo.
    echo [失败] 请检查上方错误信息
)
pause
