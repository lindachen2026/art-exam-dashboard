#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
停止所有考级相关定时任务
"""
import subprocess
import sys
sys.stdout.reconfigure(encoding='utf-8')

print("=" * 50)
print("停止考级定时任务")
print("=" * 50)

# 查询所有考级相关任务
result = subprocess.run(
    ['schtasks', '/query', '/fo', 'CSV', '/nh'],
    capture_output=True, text=True, encoding='utf-8', errors='replace'
)

tasks = []
for line in result.stdout.split('\n'):
    line = line.strip()
    if line and '考级' in line:
        name = line.split(',')[0].strip('"')
        if name:
            tasks.append(name)

if not tasks:
    print("\n未找到任何考级相关定时任务（已全部停止或未创建）")
else:
    print(f"\n找到 {len(tasks)} 个任务，正在删除...\n")
    for name in tasks:
        r = subprocess.run(['schtasks', '/delete', '/tn', name, '/f'],
                           capture_output=True)
        status = '[成功]' if r.returncode == 0 else '[失败]'
        print(f"  {status} {name}")

print(f"\n共处理 {len(tasks)} 个任务")
print("=" * 50)
input("\n按Enter关闭...")
