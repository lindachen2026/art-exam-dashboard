@echo off
REM ============================================
REM 配置考级数据看板定时任务
REM 自动化频次：
REM   - 报名期间：每天 10:50, 14:00, 18:00, 23:59
REM   - 报名结束后次日：10:00（最后一次）
REM ============================================

echo ============================================
echo 考级数据看板 - 定时任务配置
echo ============================================
echo.

set TASK_NAME=考级数据看板自动更新
set SCRIPT_PATH=%~dp0auto_update_and_publish.py
set PYTHON_EXE=python

echo 任务名称: %TASK_NAME%
echo 脚本路径: %SCRIPT_PATH%
echo Python解释器: %PYTHON_EXE%
echo.

REM ============================================
REM 删除已有的同名任务（如果存在）
REM ============================================
echo [1/5] 删除旧任务...
schtasks /Query /TN "%TASK_NAME%-10:50" >nul 2>&1
if %errorlevel%==0 (
    schtasks /Delete /TN "%TASK_NAME%-10:50" /F
)
schtasks /Query /TN "%TASK_NAME%-14:00" >nul 2>&1
if %errorlevel%==0 (
    schtasks /Delete /TN "%TASK_NAME%-14:00" /F
)
schtasks /Query /TN "%TASK_NAME%-18:00" >nul 2>&1
if %errorlevel%==0 (
    schtasks /Delete /TN "%TASK_NAME%-18:00" /F
)
schtasks /Query /TN "%TASK_NAME%-23:59" >nul 2>&1
if %errorlevel%==0 (
    schtasks /Delete /TN "%TASK_NAME%-23:59" /F
)
echo   已删除旧任务
echo.

REM ============================================
REM 创建新任务
REM ============================================
echo [2/5] 创建定时任务...
echo.

REM 任务1: 每天上午 10:50
echo   创建任务: 每天上午 10:50
schtasks /Create /TN "%TASK_NAME%-10:50" ^
    /TR "%PYTHON_EXE% \"%SCRIPT_PATH%\"" ^
    /SC DAILY /ST 10:50 ^
    /RL HIGHEST ^
    /F
if %errorlevel%==0 (
    echo   ✓ 已创建: 每天上午 10:50
) else (
    echo   ✗ 创建失败: 每天上午 10:50
)
echo.

REM 任务2: 每天下午 14:00
echo   创建任务: 每天下午 14:00
schtasks /Create /TN "%TASK_NAME%-14:00" ^
    /TR "%PYTHON_EXE% \"%SCRIPT_PATH%\"" ^
    /SC DAILY /ST 14:00 ^
    /RL HIGHEST ^
    /F
if %errorlevel%==0 (
    echo   ✓ 已创建: 每天下午 14:00
) else (
    echo   ✗ 创建失败: 每天下午 14:00
)
echo.

REM 任务3: 每天下午 18:00
echo   创建任务: 每天下午 18:00
schtasks /Create /TN "%TASK_NAME%-18:00" ^
    /TR "%PYTHON_EXE% \"%SCRIPT_PATH%\"" ^
    /SC DAILY /ST 18:00 ^
    /RL HIGHEST ^
    /F
if %errorlevel%==0 (
    echo   ✓ 已创建: 每天下午 18:00
) else (
    echo   ✗ 创建失败: 每天下午 18:00
)
echo.

REM 任务4: 每天晚上 23:59
echo   创建任务: 每天晚上 23:59
schtasks /Create /TN "%TASK_NAME%-23:59" ^
    /TR "%PYTHON_EXE% \"%SCRIPT_PATH%\"" ^
    /SC DAILY /ST 23:59 ^
    /RL HIGHEST ^
    /F
if %errorlevel%==0 (
    echo   ✓ 已创建: 每天晚上 23:59
) else (
    echo   ✗ 创建失败: 每天晚上 23:59
)
echo.

REM ============================================
REM 查看已创建的任务
REM ============================================
echo [3/5] 验证任务...
schtasks /Query /TN "%TASK_NAME%-10:50" /FO LIST
schtasks /Query /TN "%TASK_NAME%-14:00" /FO LIST
schtasks /Query /TN "%TASK_NAME%-18:00" /FO LIST
schtasks /Query /TN "%TASK_NAME%-23:59" /FO LIST
echo.

REM ============================================
REM 测试运行
REM ============================================
echo [4/5] 是否立即测试运行?
set /p CONFIRM=输入 y 立即测试，其他键跳过:
if /i "%CONFIRM%"=="y" (
    echo   正在测试运行...
    schtasks /Run /TN "%TASK_NAME%-10:50"
    echo   已触发测试，请查看输出日志
) else (
    echo   跳过测试
)
echo.

REM ============================================
REM 使用说明
REM ============================================
echo [5/5] 配置完成
echo ============================================
echo.
echo ✓ 已创建4个定时任务:
echo   - 每天上午 10:50
echo   - 每天下午 14:00
echo   - 每天下午 18:00
echo   - 每天晚上 23:59
echo.
echo 任务将在报名期间自动运行
echo.
echo 管理命令:
echo   - 启用任务: schtasks /Change /TN "任务名" /ENABLE
echo   - 禁用任务: schtasks /Change /TN "任务名" /DISABLE
echo   - 删除任务: schtasks /Delete /TN "任务名" /F
echo   - 手动运行: schtasks /Run /TN "任务名"
echo   - 查看任务: taskschd.msc
echo.
echo 注意事项:
echo   1. 报名期间自动运行，无需手动干预
echo   2. 报名结束后，请手动禁用定时任务
echo   3. 报名结束次日10:00需手动运行一次最后更新
echo   4. 然后运行 archive_current_period.py 归档
echo.
echo ============================================
pause
