#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
考级数据归档脚本（增强版）
用途: 在每期报名结束后，一键生成完整的历史归档文件
参考: 2026Q1_考级统计归档.xlsx 的结构
运行: python archive_current_period.py
"""

import pandas as pd
import json
import re
from datetime import datetime
import os

# ===== 文件路径配置 =====
FILE_CURRENT_TARGET = "数据源/本期考级目标.xlsx"
FILE_STUDENT_INFO = "数据源/考级学员信息表.xlsx"
FILE_FOLLOW = "数据源/目标明细.xlsx"
FILE_HISTORY = "数据源/往期数据.xlsx"
FILE_DASHBOARD = "dashboard.html"
ARCHIVE_DIR = "数据源/历史归档"

os.makedirs(ARCHIVE_DIR, exist_ok=True)

print("=" * 80)
print("考级数据归档工具（增强版 - 参考Q1结构）")
print("=" * 80)

# ===== 0. 从dashboard提取DATA =====
print("\n[0/5] 从dashboard提取JSON数据...")
dashboard_data = {}
try:
    with open(FILE_DASHBOARD, 'r', encoding='utf-8') as f:
        content = f.read()
    match = re.search(r'const DATA = ({.*?});', content, re.DOTALL)
    if match:
        dashboard_data = json.loads(match.group(1))
        print("  [成功] 已成功读取dashboard数据")
except Exception as e:
    print(f"  [警告] 读取dashboard失败 {e}")

# ===== 1. 读取本期基础信息 =====
print("\n[1/5] 读取本期基础信息...")
df_target = pd.read_excel(FILE_CURRENT_TARGET, sheet_name="Sheet1")
quarter_label = ""
total_target = 0
target_hours = 0

for _, row in df_target.iterrows():
    if str(row.iloc[0]).strip() == '季度':
        quarter_label = str(row.iloc[1]).strip()
    elif str(row.iloc[0]).strip() == '目标人数':
        total_target = int(row.iloc[1])
    elif str(row.iloc[0]).strip() == '目标课时':
        target_hours = int(row.iloc[1])

archive_label = quarter_label
print(f"  季度: {quarter_label}")
print(f"  目标人数: {total_target}")
print(f"  目标课时: {target_hours}")

archive_filename = f"{ARCHIVE_DIR}/{archive_label}_考级统计归档.xlsx"

# ===== 2. 读取本期实际完成数据 =====
print("\n[2/5] 读取本期实际完成数据...")

# 读取报名信息
df_reg = pd.read_excel(FILE_STUDENT_INFO, sheet_name="考级学员信息表")
df_reg = df_reg[df_reg['用户id'].notna()]
df_reg['用户id'] = df_reg['用户id'].astype(int)

total_reg_count = df_reg['用户id'].nunique()
print(f"  总报名人数: {total_reg_count}")

# 课时统计
df_history = pd.read_excel(FILE_HISTORY, sheet_name="与往期总体对比", header=1)
current_hours = 0
for _, row in df_history.iterrows():
    if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == archive_label:
        if pd.notna(row.iloc[3]):
            current_hours = int(row.iloc[3])
            break

if current_hours == 0:
    current_hours = target_hours

print(f"  总报名课时: {current_hours}")

# 国内海外统计
domestic_count = len(df_reg[df_reg['学员地区'].str.contains('国内|北京|上海|广州|成都', na=False)])
overseas_count = total_reg_count - domestic_count
print(f"  国内: {domestic_count}, 海外: {overseas_count}")

# 首报复报统计
df_follow = pd.read_excel(FILE_FOLLOW, sheet_name="跟进学员名单（在班学员）")
df_follow['学员ID'] = df_follow['学员ID'].astype(int)
df_follow['累计报考次数'] = pd.to_numeric(df_follow['累计报考次数'], errors='coerce')

registered_ids = set(df_reg['用户id'].unique())
first_count = 0
repeat_count = 0

for rid in registered_ids:
    mask = df_follow['学员ID'] == rid
    if mask.any():
        count = df_follow[mask]['累计报考次数'].values[0]
        if pd.isna(count) or count == 0:
            first_count += 1
        else:
            repeat_count += 1
    else:
        first_count += 1

print(f"  首报: {first_count}, 复报: {repeat_count}")

# 报考率 = 池内已报名人数 / 池内总人数（与看板算法一致）
pool_total = len(df_follow)
pool_registered_count = len(registered_ids & set(df_follow['学员ID'].unique()))
total_reg_rate = pool_registered_count / pool_total if pool_total > 0 else 0

# ===== 3. 更新往期数据.xlsx =====
print("\n[3/5] 更新往期数据.xlsx...")

df_history_overall = pd.read_excel(FILE_HISTORY, sheet_name="与往期总体对比", header=0)
df_history_region = pd.read_excel(FILE_HISTORY, sheet_name="国内海外对比", header=0)
df_history_repeat = pd.read_excel(FILE_HISTORY, sheet_name="复报首报往期对比", header=0)

# 更新"与往期总体对比"
if archive_label not in df_history_overall['时间'].values:
    new_row = pd.DataFrame([{
        '时间': archive_label,
        '报名人数': total_reg_count,
        '报考率': total_reg_rate,
        '预扣课时': current_hours,
        '人数目标': total_target,
        '目标完成率': total_reg_count / total_target if total_target > 0 else 0
    }])
    df_history_overall = pd.concat([df_history_overall, new_row], ignore_index=True)
else:
    idx = df_history_overall[df_history_overall['时间'] == archive_label].index[0]
    df_history_overall.loc[idx, '报名人数'] = total_reg_count
    df_history_overall.loc[idx, '报考率'] = total_reg_rate
    df_history_overall.loc[idx, '预扣课时'] = current_hours
    df_history_overall.loc[idx, '人数目标'] = total_target
    df_history_overall.loc[idx, '目标完成率'] = total_reg_count / total_target if total_target > 0 else 0

# 更新"国内海外对比"（报考率 = 池内已报名人数 / 池内总人数，与看板算法一致）
df_follow['地区'] = df_follow['地区'].astype(str)
domestic_pool = df_follow[df_follow['地区'] == '国内']
overseas_pool = df_follow[df_follow['地区'] != '国内']
registered_id_set = set(df_reg['用户id'].unique())
domestic_pool_reg = domestic_pool['学员ID'].isin(registered_id_set).sum()
overseas_pool_reg = overseas_pool['学员ID'].isin(registered_id_set).sum()
domestic_reg_rate = domestic_pool_reg / len(domestic_pool) if len(domestic_pool) > 0 else 0
overseas_reg_rate = overseas_pool_reg / len(overseas_pool) if len(overseas_pool) > 0 else 0
df_history_region.loc[df_history_region['地区'] == '合计', '上期报考率'] = total_reg_rate
df_history_region.loc[df_history_region['地区'] == '国内', '上期报考率'] = domestic_reg_rate
df_history_region.loc[df_history_region['地区'] == '海外', '上期报考率'] = overseas_reg_rate

# 更新"复报首报往期对比"
if archive_label not in df_history_repeat['季度'].values:
    new_row = pd.DataFrame([{
        '季度': archive_label,
        '首报人数': first_count,
        '首报占比': first_count / total_reg_count if total_reg_count > 0 else 0,
        '复报人数': repeat_count,
        '复报占比': repeat_count / total_reg_count if total_reg_count > 0 else 0
    }])
    df_history_repeat = pd.concat([df_history_repeat, new_row], ignore_index=True)
else:
    idx = df_history_repeat[df_history_repeat['季度'] == archive_label].index[0]
    df_history_repeat.loc[idx, '首报人数'] = first_count
    df_history_repeat.loc[idx, '首报占比'] = first_count / total_reg_count if total_reg_count > 0 else 0
    df_history_repeat.loc[idx, '复报人数'] = repeat_count
    df_history_repeat.loc[idx, '复报占比'] = repeat_count / total_reg_count if total_reg_count > 0 else 0

print("  [成功] 已更新往期数据")

# ===== 4. 保存增强版归档文件 =====
print(f"\n[4/5] 生成增强版归档文件...")

with pd.ExcelWriter(archive_filename, engine='openpyxl') as writer:
    # 1. 总体概览（15列，参考Q1）
    df_overview = pd.DataFrame([{
        '季度': archive_label,
        '报名人数': total_reg_count,
        '目标人数': total_target,
        '完成率': total_reg_count / total_target if total_target > 0 else 0,
        '报考率': total_reg_rate,
        '预扣课时': current_hours,
        '课时目标': target_hours,
        '首报人数': first_count,
        '首报占比': first_count / total_reg_count if total_reg_count > 0 else 0,
        '复报人数': repeat_count,
        '复报占比': repeat_count / total_reg_count if total_reg_count > 0 else 0,
        '国内报名': domestic_count,
        '国内报考率': domestic_count / total_reg_count if total_reg_count > 0 else 0,
        '海外报名': overseas_count,
        '海外报考率': overseas_count / total_reg_count if total_reg_count > 0 else 0
    }])
    df_overview.to_excel(writer, sheet_name='总体概览', index=False)
    print("  [成功] 总体概览")

    # 2. 科目统计
    if dashboard_data and 'subject_stats' in dashboard_data:
        df_subject = pd.DataFrame(dashboard_data['subject_stats'])
        df_subject = df_subject[['考试科目名称', '报名人数', '课时数']].copy()
        df_subject.to_excel(writer, sheet_name='科目统计', index=False)
        print(f"  [成功] 科目统计 ({len(df_subject)}个)")

    # 3. 科目级别统计
    if dashboard_data and 'subject_level_ranking' in dashboard_data:
        df_level = pd.DataFrame(dashboard_data['subject_level_ranking'])
        df_level = df_level[['name', 'current']].copy()
        df_level.columns = ['科目级别', '报名人数']
        df_level.to_excel(writer, sheet_name='科目级别统计', index=False)
        print(f"  [成功] 科目级别统计 ({len(df_level)}个)")

    # 4. 学员分层
    if dashboard_data and 'tier_stats' in dashboard_data:
        tier_data = []
        for t in dashboard_data['tier_stats']:
            tier_data.append({
                '分层': t.get('tier'),
                '池内人数': t.get('pool_count'),
                '报名人数': t.get('reg_count'),
                '报考率': t.get('rate')
            })
        df_tier = pd.DataFrame(tier_data)
        df_tier.to_excel(writer, sheet_name='学员分层', index=False)
        print(f"  [成功] 学员分层 ({len(df_tier)}层)")

    # 5. MTD每日流速
    if dashboard_data and 'mtd' in dashboard_data:
        mtd_data = []
        for m in dashboard_data['mtd']:
            mtd_data.append({
                '日期': m.get('short_label'),
                '星期': m.get('weekday'),
                '目标人数': m.get('daily_target'),
                '实际人数': m.get('actual_daily', 0),
                '目标流速': m.get('target_flow'),
                '实际流速': m.get('actual_flow', 0)
            })
        df_mtd = pd.DataFrame(mtd_data)
        df_mtd.to_excel(writer, sheet_name='MTD每日流速', index=False)
        print(f"  [成功] MTD每日流速 ({len(df_mtd)}天)")

# ===== 5. 保存往期数据.xlsx =====
print("\n[5/5] 保存往期数据.xlsx...")
with pd.ExcelWriter(FILE_HISTORY, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_history_overall.to_excel(writer, sheet_name='与往期总体对比', index=False)
    df_history_region.to_excel(writer, sheet_name='国内海外对比', index=False)
    df_history_repeat.to_excel(writer, sheet_name='复报首报往期对比', index=False)

print("\n" + "=" * 80)
print("[完成] 数据归档成功!")
print("=" * 80)
print(f"成功生成: {archive_filename}")
print(f"包含维度: 总体概览、科目统计、科目级别、学员分层、MTD流速")
print(f"往期数据已更新: {FILE_HISTORY}")
print()
