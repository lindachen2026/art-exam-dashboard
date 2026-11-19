#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
本期收尾留档：最后一次更新看板 + 归档数据
报名结束次日 10:00 自动触发，也可手动双击运行
"""
import subprocess
import sys
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

print("=" * 60)
print("考级数据 - 本期收尾留档")
print(f"执行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("=" * 60)

# 步骤1：最后一次完整更新（下载BI + 生成看板 + 推送GitHub）
print("\n[1/2] 执行最后一次完整更新...")
r = subprocess.run([sys.executable, "auto_update_and_publish.py"], cwd=BASE_DIR)
if r.returncode == 0:
    print("  [成功] 最后一次看板更新完成")
else:
    print("  [警告] 看板更新失败，继续执行归档（归档将使用当前数据）")

# 步骤2：归档本期数据
print("\n[2/2] 归档本期数据...")
r = subprocess.run([sys.executable, "archive_current_period.py"], cwd=BASE_DIR)
if r.returncode == 0:
    print("  [成功] 本期数据归档完成")
else:
    print("  [警告] 归档脚本执行异常，请手动检查并运行 archive_current_period.py")

print("\n" + "=" * 60)
print("[完成] 本期收尾留档完成！")
print("=" * 60)
print("\n  已生成: 数据源/历史归档/本期归档文件")
print("  已更新: 数据源/往期数据.xlsx")
print("  下期启动前请更新: 数据源/本期考级目标.xlsx 和 目标明细.xlsx")
input("\n按Enter关闭...")
