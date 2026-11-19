@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo   考级数据 - 自动下载 + 刷新 + 发布
echo ========================================
echo.
echo 时间: %date% %time:~0,5%
echo.

echo [1/3] 从BI系统下载考级学员信息表...
python "%~dp0auto_download_bi.py"
if %errorlevel% neq 0 (
    echo [失败] BI下载出错，流程中止。
    goto :end
)
echo.

echo [2/3] 重新生成看板...
python "%~dp0generate_dashboard.py"
if %errorlevel% neq 0 (
    echo [失败] 看板生成出错，流程中止。
    goto :end
)
echo.

echo [3/3] 推送到 GitHub...
git add dashboard.html
git commit -m "update data %date% %time:~0,5%"
git push
echo.

echo ========================================
echo   全部完成！线上看板将在1-2分钟内更新
echo   https://lindachen2026.github.io/art-exam-dashboard/dashboard.html
echo ========================================

:end
echo.
