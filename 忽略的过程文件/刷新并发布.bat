@echo off
chcp 65001 >nul
echo ========================================
echo   刷新看板数据 + 发布到云端
echo ========================================
echo.
echo [1/2] 正在重新生成看板...
python "%~dp0generate_dashboard.py"
if %errorlevel% neq 0 (
    echo 生成失败，请检查数据源文件！
    pause
    exit /b 1
)
echo.
echo [2/2] 正在推送到 GitHub...
cd /d "%~dp0"
git add dashboard.html
git commit -m "update data %date% %time:~0,5%"
git push
echo.
echo ========================================
echo   完成！线上看板将在1-2分钟内更新
echo   链接：https://lindachen2026.github.io/art-exam-dashboard/dashboard.html
echo ========================================
pause
