#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
自动更新看板并推送到GitHub
流程: 下载BI数据 → 生成看板 → 推送GitHub
"""

import subprocess
import sys
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

print("=" * 80)
print("考级数据看板自动更新")
print("=" * 80)
print(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print(f"目录: {BASE_DIR}")
print()

# ===== 步骤1: 下载BI数据 =====
print("[步骤 1/3] 下载BI数据...")
print("-" * 80)
try:
    result = subprocess.run(
        [sys.executable, "auto_download_bi.py"],
        timeout=300,  # 5分钟超时
        cwd=BASE_DIR
    )
    if result.returncode == 0:
        print("[成功] BI数据下载完成")
    else:
        print("[警告] 下载失败，继续使用现有数据...")
except subprocess.TimeoutExpired:
    print("[警告] 下载超时(5分钟)，继续使用现有数据...")
except Exception as e:
    print(f"[警告] 下载异常，继续使用现有数据...")

# ===== 步骤2: 生成看板 =====
print("\n[步骤 2/3] 生成看板...")
print("-" * 80)
try:
    result = subprocess.run(
        [sys.executable, "generate_dashboard.py"],
        timeout=60,
        cwd=BASE_DIR
    )
    if result.returncode == 0:
        print("[成功] 看板生成成功")
    else:
        print("[错误] 看板生成失败")
        sys.exit(1)
except Exception as e:
    print(f"[错误] 看板生成异常")
    sys.exit(1)

# ===== 步骤3: 推送GitHub =====
print("\n[步骤 3/3] 推送到GitHub...")
print("-" * 80)

# 3.1 检查是否有变更
result = subprocess.run(
    ["git", "status", "--porcelain"],
    capture_output=True,
    text=True,
    encoding='utf-8',
    errors='ignore'
)
if not result.stdout.strip():
    print("[成功] 无数据变更，跳过推送")
    print("\n" + "=" * 80)
    print("更新完成（无变更）")
    print("=" * 80)
    sys.exit(0)

# 3.2 添加所有变更
try:
    # 只添加dashboard.html（数据源文件被gitignore忽略，无需推送）
    if os.path.exists("dashboard.html"):
        subprocess.run(["git", "add", "dashboard.html"], check=True)
    else:
        print("[警告] dashboard.html 不存在")
        sys.exit(1)

    # 3.3 提交
    commit_msg = f"update data {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    subprocess.run(
        ["git", "commit", "-m", commit_msg],
        check=True,
        capture_output=True
    )
    print(f"[成功] 已提交: {commit_msg}")

    # 3.4 推送
    subprocess.run(["git", "push", "origin", "main"], check=True, capture_output=True)
    print("[成功] 已推送到GitHub")

except subprocess.CalledProcessError as e:
    print(f"[警告] Git操作失败: {e}")
    print("请检查Git配置和网络连接")

print("\n" + "=" * 80)
print("[完成] 全部完成")
print("=" * 80)
print(f"看板地址: https://lindachen2026.github.io/art-exam-dashboard/dashboard.html")
