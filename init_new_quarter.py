#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
新季度初始化配置脚本
用途: 在每期报名启动前，快速配置新季度的所有必要参数
运行: python init_new_quarter.py
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, timedelta
import os

print("=" * 80)
print("新季度初始化配置向导")
print("=" * 80)

# ===== 1. 收集新季度基础信息 =====
print("\n[步骤 1/5] 请输入新季度基础信息")
print("-" * 80)

quarter = input("季度标识 (格式: 2026Q3): ").strip()
target_count = int(input("目标人数: "))
target_hours = int(input("目标课时: "))
start_date_str = input("报名开始日期 (格式: 2026-07-03): ").strip()
end_date_str = input("报名结束日期 (格式: 2026-07-12): ").strip()
exam_start_str = input("考试开始日期 (格式: 2026-08-22): ").strip()
exam_end_str = input("考试结束日期 (格式: 2026-08-23): ").strip()

# 解析日期
start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
reg_days = (end_date - start_date).days + 1

print(f"\n✓ 报名天数: {reg_days}天")

# ===== 2. 配置MTD流速分布 =====
print("\n[步骤 2/5] 配置MTD每日流速")
print("-" * 80)
print(f"需要为 {reg_days} 天配置流速分布（总和必须=1.0）")
print("\n建议配置参考:")
print("  - 第一天(开放日): 0.15-0.20")
print("  - 工作日: 0.07-0.10")
print("  - 周末: 0.03-0.05")
print("  - 最后1-2天(截止冲刺): 0.10-0.15")

use_manual = input("\n是否手动输入每日流速? (y/n, 默认n使用均匀分布): ").strip().lower()

mtd_flows = []
if use_manual == 'y':
    print(f"\n请依次输入 {reg_days} 天的流速:")
    current = start_date
    for i in range(reg_days):
        weekday = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][current.weekday()]
        flow = float(input(f"  {current.strftime('%m-%d')} {weekday}: "))
        mtd_flows.append(flow)
        current += timedelta(days=1)
else:
    # 均匀分布
    flow_per_day = 1.0 / reg_days
    mtd_flows = [flow_per_day] * reg_days
    print(f"\n✓ 使用均匀分布，每天流速: {flow_per_day:.4f}")

# 验证流速总和
flow_sum = sum(mtd_flows)
if abs(flow_sum - 1.0) > 0.001:
    print(f"\n⚠️  警告: 流速总和 = {flow_sum:.4f}, 不等于1.0")
    adjust = input("是否自动调整最后一天使总和=1.0? (y/n): ").strip().lower()
    if adjust == 'y':
        mtd_flows[-1] += (1.0 - flow_sum)
        print(f"✓ 已调整，最后一天流速: {mtd_flows[-1]:.4f}")
    else:
        print("请手动调整流速配置")
        exit(1)

# ===== 3. 国内海外目标分配 =====
print("\n[步骤 3/5] 配置国内/海外目标分配")
print("-" * 80)
domestic_target = int(input(f"国内目标人数 (总计{target_count}): "))
overseas_target = target_count - domestic_target
print(f"✓ 海外目标人数: {overseas_target}")

# ===== 4. 更新配置文件 =====
print("\n[步骤 4/5] 更新配置文件...")
print("-" * 80)

# 4.1 更新 本期考级目标.xlsx
FILE_TARGET = "数据源/本期考级目标.xlsx"
wb_target = openpyxl.load_workbook(FILE_TARGET)
ws_target = wb_target.active

# 找到各维度对应的行并更新
for row in ws_target.iter_rows(min_row=2, max_row=20):
    dimension = row[0].value
    if dimension == "季度":
        row[1].value = quarter
    elif dimension == "目标人数":
        row[1].value = target_count
    elif dimension == "目标课时":
        row[1].value = target_hours
    elif dimension == "报名时间":
        row[1].value = f"{start_date.strftime('%#m.%#d')}-{end_date.strftime('%#m.%#d')}"
    elif dimension == "考试时间":
        exam_start = datetime.strptime(exam_start_str, "%Y-%m-%d")
        exam_end = datetime.strptime(exam_end_str, "%Y-%m-%d")
        row[1].value = f"{exam_start.strftime('%#m.%#d')}-{exam_end.strftime('%#m.%#d')}"
    elif dimension == "国内":
        row[1].value = domestic_target
    elif dimension == "海外":
        row[1].value = overseas_target

wb_target.save(FILE_TARGET)
print(f"✓ 已更新: {FILE_TARGET}")

# 4.2 更新 目标明细.xlsx - MTD sheet
FILE_DETAIL = "数据源/目标明细.xlsx"
wb_detail = openpyxl.load_workbook(FILE_DETAIL)

# 删除旧的MTD sheet（如果存在）
if "MTD" in wb_detail.sheetnames:
    del wb_detail["MTD"]

# 创建新的MTD sheet
ws_mtd = wb_detail.create_sheet("MTD", 0)

# 写入表头
ws_mtd['A1'] = "MTD报名趋势目标人数"
ws_mtd['A1'].font = Font(size=14, bold=True)
ws_mtd['A2'] = "日期"
ws_mtd['B2'] = "星期"
ws_mtd['C2'] = "报名人数"
ws_mtd['D2'] = "每日流速"

# 设置表头样式
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
for col in ['A2', 'B2', 'C2', 'D2']:
    ws_mtd[col].fill = header_fill
    ws_mtd[col].font = header_font
    ws_mtd[col].alignment = Alignment(horizontal='center')

# 写入每日数据
current = start_date
for i in range(reg_days):
    row_idx = i + 3
    weekday = ['周一', '周二', '周三', '周四', '周五', '周六', '周日'][current.weekday()]
    flow = mtd_flows[i]

    ws_mtd.cell(row_idx, 1, current)  # 日期
    ws_mtd.cell(row_idx, 2, weekday)   # 星期
    ws_mtd.cell(row_idx, 3, f"=D{row_idx}*{target_count}")  # 报名人数公式
    ws_mtd.cell(row_idx, 4, flow)      # 流速

    # 设置日期格式
    ws_mtd.cell(row_idx, 1).number_format = 'yyyy-mm-dd'
    # 设置流速格式
    ws_mtd.cell(row_idx, 4).number_format = '0.00'

    current += timedelta(days=1)

# 调整列宽
ws_mtd.column_dimensions['A'].width = 15
ws_mtd.column_dimensions['B'].width = 10
ws_mtd.column_dimensions['C'].width = 12
ws_mtd.column_dimensions['D'].width = 12

wb_detail.save(FILE_DETAIL)
print(f"✓ 已更新: {FILE_DETAIL} (MTD sheet)")

# ===== 5. 清空本期报名数据 =====
print("\n[步骤 5/5] 清空本期报名数据...")
print("-" * 80)

FILE_CURRENT = "数据源/本期各维度报名明细.xlsx"
clear_confirm = input(f"是否清空 {FILE_CURRENT} 的所有实际数据? (y/n): ").strip().lower()

if clear_confirm == 'y':
    wb_current = openpyxl.load_workbook(FILE_CURRENT)

    # 清空各个sheet（保留表头）
    sheets_to_clear = ["本期报名总表", "科目统计", "科目级别统计", "学员分层", "复报首报"]
    for sheet_name in sheets_to_clear:
        if sheet_name in wb_current.sheetnames:
            ws = wb_current[sheet_name]
            # 删除第2行及以后的所有数据
            ws.delete_rows(2, ws.max_row)
            print(f"  ✓ 已清空: {sheet_name}")

    wb_current.save(FILE_CURRENT)
    print(f"✓ 已清空: {FILE_CURRENT}")
else:
    print("⊗ 跳过清空数据")

# ===== 6. 生成初始看板验证 =====
print("\n[步骤 6/6] 生成初始看板验证...")
print("-" * 80)

generate_confirm = input("是否运行 generate_dashboard.py 生成初始看板? (y/n): ").strip().lower()
if generate_confirm == 'y':
    import subprocess
    result = subprocess.run(["python", "generate_dashboard.py"], capture_output=True, text=True)
    if result.returncode == 0:
        print("✓ 看板生成成功")
    else:
        print("⚠️  看板生成失败，请检查错误信息:")
        print(result.stderr)

# ===== 完成总结 =====
print("\n" + "=" * 80)
print("✅ 新季度配置完成!")
print("=" * 80)
print(f"\n已配置内容:")
print(f"  • 季度: {quarter}")
print(f"  • 目标人数: {target_count} (国内{domestic_target} + 海外{overseas_target})")
print(f"  • 目标课时: {target_hours}")
print(f"  • 报名时间: {start_date_str} ~ {end_date_str} ({reg_days}天)")
print(f"  • 考试时间: {exam_start_str} ~ {exam_end_str}")
print(f"  • MTD流速: 已配置{reg_days}天，总和={sum(mtd_flows):.4f}")

print(f"\n更新的文件:")
print(f"  ✓ {FILE_TARGET}")
print(f"  ✓ {FILE_DETAIL}")
if clear_confirm == 'y':
    print(f"  ✓ {FILE_CURRENT} (已清空)")

print(f"\n后续步骤:")
print(f"  1. 检查生成的 dashboard.html，验证初始状态")
print(f"  2. 每日从报名系统导出数据，更新 {FILE_CURRENT}")
print(f"  3. 运行 python generate_dashboard.py 生成最新看板")
print(f"  4. 报名结束后运行 python archive_current_period.py 归档")
print("=" * 80)
