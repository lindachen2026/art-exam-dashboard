#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
启动本期定时任务
自动从"数据源/本期考级目标.xlsx"读取配置，创建定时任务
"""
import os
import subprocess
import sys
from datetime import datetime, timedelta
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "数据源", "本期考级目标.xlsx")
PYTHON_EXE = sys.executable

print("=" * 60)
print("考级数据看板 - 启动本期定时任务")
print("=" * 60)

# ===== 1. 读取配置 =====
print("\n[1/4] 读取配置...")
config = {}
try:
    wb = load_workbook(CONFIG_FILE)
    ws = wb.active
    for row in ws.iter_rows():
        if row[0].value and row[1].value:
            config[str(row[0].value).strip()] = str(row[1].value).strip()
    wb.close()
except Exception as e:
    print(f"  [错误] 无法读取配置文件: {e}")
    print(f"  请确认文件存在: {CONFIG_FILE}")
    input("\n按Enter退出...")
    sys.exit(1)

date_start  = config.get("报名开始日期", "")
date_end    = config.get("报名结束日期", "")
time_points = [t.strip() for t in config.get("每日播报时间", "10:50,14:00,18:00,23:59").split(",")]
quarter     = config.get("季度", "")

if not date_start or not date_end:
    print("  [错误] 配置文件中未找到'报名开始日期'或'报名结束日期'")
    print("  请检查: 数据源/本期考级目标.xlsx")
    input("\n按Enter退出...")
    sys.exit(1)

print(f"  季度:         {quarter}")
print(f"  报名开始日期: {date_start}")
print(f"  报名结束日期: {date_end}")
print(f"  每日播报时间: {', '.join(time_points)}")

try:
    dt_end   = datetime.strptime(date_end.split()[0], "%Y-%m-%d")
    dt_next  = dt_end + timedelta(days=1)
    ed_param = dt_end.strftime("%m/%d/%Y")     # schtasks /ED 格式
    sd_next  = dt_next.strftime("%m/%d/%Y")    # 次日日期
    next_date_str = dt_next.strftime("%Y-%m-%d")
except ValueError:
    print(f"  [错误] 日期格式错误，应为 YYYY-MM-DD，当前值: '{date_end}'")
    input("\n按Enter退出...")
    sys.exit(1)

# ===== 2. 清除旧任务 =====
print(f"\n[2/4] 清除旧任务...")
task_prefix = "考级数据看板自动更新"
old_names = [f"{task_prefix}-{t.replace(':', '')}" for t in time_points]
old_names.append("考级数据收尾归档")

for name in old_names:
    r = subprocess.run(["schtasks", "/delete", "/tn", name, "/f"],
                       capture_output=True)
    if r.returncode == 0:
        print(f"  已删除旧任务: {name}")

# ===== 3. 创建每日定时任务 =====
print(f"\n[3/4] 创建每日定时任务（截止 {date_end}）...")
script_main = os.path.join(BASE_DIR, "auto_update_and_publish.py")
success_count = 0

for t in time_points:
    task_name = f"{task_prefix}-{t.replace(':', '')}"
    cmd = [
        "schtasks", "/create",
        "/tn", task_name,
        "/tr", f'"{PYTHON_EXE}" "{script_main}"',
        "/sc", "DAILY",
        "/st", t,
        "/ed", ed_param,
        "/rl", "HIGHEST",
        "/f"
    ]
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode == 0:
        print(f"  [成功] 每天 {t}（截止 {date_end}）")
        success_count += 1
    else:
        err = r.stderr.strip().split('\n')[0] if r.stderr else "未知错误"
        print(f"  [失败] 每天 {t}: {err[:80]}")

# ===== 4. 创建次日收尾归档任务（一次性） =====
print(f"\n[4/4] 创建次日收尾归档任务（{next_date_str} 10:00）...")
script_final = os.path.join(BASE_DIR, "finalize_period.py")
cmd = [
    "schtasks", "/create",
    "/tn", "考级数据收尾归档",
    "/tr", f'"{PYTHON_EXE}" "{script_final}"',
    "/sc", "ONCE",
    "/sd", sd_next,
    "/st", "10:00",
    "/rl", "HIGHEST",
    "/f"
]
r = subprocess.run(cmd, capture_output=True, text=True,
                   encoding="utf-8", errors="replace")
if r.returncode == 0:
    print(f"  [成功] {next_date_str} 10:00 自动执行最后更新+归档")
else:
    err = r.stderr.strip().split('\n')[0] if r.stderr else "未知错误"
    print(f"  [失败]: {err[:80]}")

print("\n" + "=" * 60)
print("[完成] 定时任务配置完成！")
print("=" * 60)
print(f"\n  本期报名区间:  {date_start}  ~  {date_end}")
print(f"  每日播报时间:  {', '.join(time_points)}")
print(f"  收尾归档时间:  {next_date_str} 10:00（自动执行，无需手动）")
print(f"\n  查看任务: Win+R 输入 taskschd.msc")
print(f"  共创建 {success_count}/{len(time_points)} 个每日任务")
input("\n按Enter关闭...")
