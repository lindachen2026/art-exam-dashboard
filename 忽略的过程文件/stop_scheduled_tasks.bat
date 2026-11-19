@echo off
chcp 65001 >nul
echo ========================================
echo 停止考级数据自动更新定时任务
echo ========================================
echo.

REM 删除定时任务
echo 正在停止定时任务...
schtasks /delete /tn "考级数据看板自动更新" /f 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [成功] 已删除任务 "考级数据看板自动更新"
) else (
    echo [警告] 任务未找到或删除失败
)

schtasks /delete /tn "考级数据自动更新" /f 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [成功] 已删除任务 "考级数据自动更新"
)

echo.
echo ========================================
echo 所有定时任务已停止
echo ========================================
echo.
pause
