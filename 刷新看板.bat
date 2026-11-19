@echo off
chcp 65001 >nul
echo ==============================
echo   考级数据看板 - 数据刷新
echo ==============================
echo.
echo 正在重新生成看板数据...
echo.
cd /d "%~dp0"
python generate_dashboard.py
echo.
if %errorlevel% equ 0 (
    echo [成功] 看板已更新，请在浏览器中刷新页面查看最新数据。
) else (
    echo [失败] 生成过程出错，请检查数据源文件。
)
echo.
pause
