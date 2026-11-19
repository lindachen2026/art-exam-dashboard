# -*- coding: utf-8 -*-
"""
考级数据归档脚本
功能：
1. 将本期统计结果保存到 数据源/历史归档/ 目录
2. 回填 往期数据.xlsx 中本期数据的空列
"""

import pandas as pd
import numpy as np
import os
import json
import shutil
from datetime import datetime, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "数据源")
ARCHIVE_DIR = os.path.join(DATA_DIR, "历史归档")

FILE_STUDENT_INFO = os.path.join(DATA_DIR, "考级学员信息表.xlsx")
FILE_TARGET = os.path.join(DATA_DIR, "目标明细.xlsx")
FILE_HOURS_STD = os.path.join(DATA_DIR, "考级各科目级别课时标准.xlsx")
FILE_HISTORY = os.path.join(DATA_DIR, "往期数据.xlsx")
FILE_OVERALL_TARGET = os.path.join(DATA_DIR, "本期考级目标.xlsx")

# 确保归档目录存在
os.makedirs(ARCHIVE_DIR, exist_ok=True)

print("=" * 50)
print("  考级数据归档")
print("=" * 50)

# 读取季度标识
df_overall_target = pd.read_excel(FILE_OVERALL_TARGET)
quarter_label = ""
for _, row in df_overall_target.iterrows():
    if str(row.iloc[0]).strip() == '季度':
        quarter_label = str(row.iloc[1]).strip()
        break

if not quarter_label:
    print("[错误] 无法读取季度标识，请检查本期考级目标.xlsx")
    input("按回车退出...")
    exit(1)

print(f"\n当前季度: {quarter_label}")
print(f"归档目录: {ARCHIVE_DIR}")

# ========== 读取并计算本期数据 ==========
print("\n正在读取数据...")

# 考级学员信息表
df_reg = pd.read_excel(FILE_STUDENT_INFO, sheet_name="考级学员信息表")
df_reg.columns = ['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']
df_reg['用户id'] = df_reg['用户id'].astype(int)

# 课时标准
df_hours_raw = pd.read_excel(FILE_HOURS_STD, sheet_name=0, header=None)
hours_map = {}
for _, row in df_hours_raw.iterrows():
    level_name = row.iloc[3]
    hours = row.iloc[4]
    if pd.notna(level_name) and pd.notna(hours):
        try:
            hours_map[str(level_name).strip()] = int(float(hours))
        except (ValueError, TypeError):
            continue

def calc_hours(row):
    subject = str(row['考试科目名称']).strip()
    level = str(row['考试等级']).strip()
    if '花鸟' in subject or subject == '国画':
        key = f"国画花鸟{level}"
    else:
        key = f"{subject}{level}"
    return hours_map.get(key, 0)

df_reg['课时'] = df_reg.apply(calc_hours, axis=1)

# 跟进学员名单
df_follow = pd.read_excel(FILE_TARGET, sheet_name="跟进学员名单（在班学员）")
df_follow.columns = [
    '老师ID', '老师工号', '老师艺名', '学区', '团队', '小组', '班级ID', '班级名称',
    '班级课表', '阶段', '星期', '上课时间', '学员ID', '学员昵称', '年龄', '地区',
    '有效续费次数', '是否师生在群', '是否有有效群', '班主任ID', '班主任姓名', '班主任艺名',
    '班主任大区', '班主任团队', '大V渠道', '累计报考次数', '25Q4报考科目', '25Q4成绩',
    '25Q3报考科目', '25Q3成绩', '25Q1报考科目', '25Q1成绩', '24Q4报考科目', '24Q4成绩',
    '24Q3报考科目', '24Q3成绩', '24Q1报考科目', '24Q1成绩', '儿童画', '儿童画成绩',
    '漫画', '漫画成绩', '素描', '素描成绩', '国画', '国画成绩', '速写', '速写成绩',
    '推荐儿童画报考级别', '推荐漫画报考级别', '推荐素描报考级别', '推荐国画报考级别',
    '推荐速写报考级别', '漫画是否跳级', '素描是否跳级', '分层', '剩余课时',
    '最新国画专题课课表名称', '最新国画专题课上课日期'
]
df_follow['学员ID'] = df_follow['学员ID'].astype(int)

registered_ids = set(df_reg['用户id'].unique())
follow_ids = set(df_follow['学员ID'].unique())
df_follow['是否本期报名'] = df_follow['学员ID'].isin(registered_ids).astype(int)

# ========== 计算本期关键指标 ==========
print("正在计算本期指标...")

total_reg_count = df_reg['用户id'].nunique()
total_hours = df_reg['课时'].sum()
total_target = 4600
target_hours = 16000
pool_total = len(df_follow)
pool_registered = registered_ids & follow_ids
reg_rate = len(pool_registered) / pool_total
completion_rate = total_reg_count / total_target

# 首报/复报
df_follow_registered = df_follow[df_follow['是否本期报名'] == 1].copy()
df_follow_registered['报名类型'] = df_follow_registered['累计报考次数'].apply(
    lambda x: '首报' if (pd.isna(x) or x == 0) else '复报'
)
outside_pool_ids = registered_ids - follow_ids
outside_first_count = len(outside_pool_ids)
first_report_count = len(df_follow_registered[df_follow_registered['报名类型'] == '首报']) + outside_first_count
repeat_report_count = len(df_follow_registered[df_follow_registered['报名类型'] == '复报'])
total_report_count = first_report_count + repeat_report_count

# 国内/海外
domestic_pool = df_follow[df_follow['地区'] == '国内']
overseas_pool = df_follow[df_follow['地区'] != '国内']
domestic_reg_count = int(domestic_pool['是否本期报名'].sum())
overseas_reg_count = int(overseas_pool['是否本期报名'].sum())
domestic_rate = domestic_reg_count / len(domestic_pool) if len(domestic_pool) > 0 else 0
overseas_rate = overseas_reg_count / len(overseas_pool) if len(overseas_pool) > 0 else 0

# 科目统计
subject_stats = df_reg.groupby('考试科目名称').agg(
    报名人数=('用户id', 'nunique'),
    课时数=('课时', 'sum')
).reset_index()
subject_stats = subject_stats.sort_values('报名人数', ascending=False)

# 科目级别统计
subject_level_stats = df_reg.groupby(['考试科目名称', '考试等级']).agg(
    报名人数=('用户id', 'nunique')
).reset_index()
subject_level_stats['科目级别'] = subject_level_stats['考试科目名称'] + subject_level_stats['考试等级']

# 学员分层
tier_stats = {}
for tier in ['A层', 'B层', 'C层', 'D层']:
    pool = df_follow[df_follow['分层'] == tier]
    pool_count = len(pool)
    rc = pool['是否本期报名'].sum()
    rate = rc / pool_count if pool_count > 0 else 0
    tier_stats[tier] = {'pool': pool_count, 'reg': int(rc), 'rate': rate}

# ========== 1. 保存归档文件 ==========
print(f"\n正在保存归档文件...")

archive_filename = f"{quarter_label}_考级统计归档.xlsx"
archive_path = os.path.join(ARCHIVE_DIR, archive_filename)

with pd.ExcelWriter(archive_path, engine='openpyxl') as writer:
    # 总体概览
    overview_data = pd.DataFrame([{
        '季度': quarter_label,
        '报名人数': total_reg_count,
        '目标人数': total_target,
        '完成率': round(completion_rate, 4),
        '报考率': round(reg_rate, 4),
        '预扣课时': int(total_hours),
        '课时目标': target_hours,
        '首报人数': first_report_count,
        '首报占比': round(first_report_count / total_report_count, 4) if total_report_count > 0 else 0,
        '复报人数': repeat_report_count,
        '复报占比': round(repeat_report_count / total_report_count, 4) if total_report_count > 0 else 0,
        '国内报名': domestic_reg_count,
        '国内报考率': round(domestic_rate, 4),
        '海外报名': overseas_reg_count,
        '海外报考率': round(overseas_rate, 4),
    }])
    overview_data.to_excel(writer, sheet_name='总体概览', index=False)

    # 科目统计
    subject_stats.to_excel(writer, sheet_name='科目统计', index=False)

    # 科目级别统计
    subject_level_stats[['科目级别', '报名人数']].to_excel(writer, sheet_name='科目级别统计', index=False)

    # 学员分层
    tier_df = pd.DataFrame([{'分层': k, '池内人数': v['pool'], '报名人数': v['reg'], '报考率': v['rate']} for k, v in tier_stats.items()])
    tier_df.to_excel(writer, sheet_name='学员分层', index=False)

    # MTD每日流速
    df_mtd_raw = pd.read_excel(FILE_TARGET, sheet_name="MTD", header=None)
    df_reg['报名日期'] = pd.to_datetime(df_reg['报名时间']).dt.date
    daily_reg = df_reg.groupby('报名日期')['用户id'].nunique().reset_index()
    daily_reg.columns = ['日期', '报名人数']
    daily_reg_map = {str(r['日期']): int(r['报名人数']) for _, r in daily_reg.iterrows()}

    mtd_rows = []
    for week_start_row in [1, 10]:
        for col in range(1, 8):
            date_serial = df_mtd_raw.iloc[week_start_row, col]
            if pd.notna(date_serial):
                actual_date = datetime(1899, 12, 30) + timedelta(days=int(date_serial))
                date_str = actual_date.strftime('%Y-%m-%d')
                weekday = str(df_mtd_raw.iloc[week_start_row + 1, col])
                daily_target = int(df_mtd_raw.iloc[week_start_row + 3, col]) if pd.notna(df_mtd_raw.iloc[week_start_row + 3, col]) else 0
                actual_daily = daily_reg_map.get(date_str, 0)
                target_flow = daily_target / total_target if total_target > 0 else 0
                actual_flow = actual_daily / total_target if total_target > 0 else 0
                mtd_rows.append({
                    '日期': date_str,
                    '星期': weekday,
                    '目标人数': daily_target,
                    '实际人数': actual_daily,
                    '目标流速': round(target_flow, 6),
                    '实际流速': round(actual_flow, 6)
                })
    if mtd_rows:
        mtd_df = pd.DataFrame(mtd_rows)
        mtd_df.to_excel(writer, sheet_name='MTD每日流速', index=False)

print(f"  → 归档文件已保存: {archive_path}")

# ========== 2. 回填往期数据.xlsx ==========
print(f"\n正在回填往期数据...")

try:
    from openpyxl import load_workbook
    wb = load_workbook(FILE_HISTORY)

    # 回填"科目级别上期报名数据"中本期列
    if '科目级别上期报名数据' in wb.sheetnames:
        ws = wb['科目级别上期报名数据']
        # 本期数据列: G(合计), H(国内), I(海外), J(国内占比), K(海外占比)
        # 第2行是子表头，第3行开始是数据
        # 先填合计行 (row 3)
        ws.cell(row=3, column=7, value=total_reg_count)
        ws.cell(row=3, column=8, value=domestic_reg_count)
        ws.cell(row=3, column=9, value=overseas_reg_count)
        dom_ratio = domestic_reg_count / total_reg_count if total_reg_count > 0 else 0
        ovs_ratio = overseas_reg_count / total_reg_count if total_reg_count > 0 else 0
        ws.cell(row=3, column=10, value=round(dom_ratio, 6))
        ws.cell(row=3, column=11, value=round(ovs_ratio, 6))

        # 按科目填充 (row 5-9)
        subject_names_map = {'漫画': 5, '素描': 6, '儿童画': 7, '速写': 8, '国画花鸟': 9}
        for subj, row_idx in subject_names_map.items():
            subj_data = df_reg[df_reg['考试科目名称'] == subj]
            s_total = subj_data['用户id'].nunique()
            s_dom = subj_data[subj_data['学员地区'] == '国内']['用户id'].nunique()
            s_ovs = subj_data[subj_data['学员地区'] != '国内']['用户id'].nunique()
            ws.cell(row=row_idx, column=7, value=s_total)
            ws.cell(row=row_idx, column=8, value=s_dom)
            ws.cell(row=row_idx, column=9, value=s_ovs)
            ws.cell(row=row_idx, column=10, value=round(s_dom / s_total, 6) if s_total > 0 else 0)
            ws.cell(row=row_idx, column=11, value=round(s_ovs / s_total, 6) if s_total > 0 else 0)

        # 按科目级别填充 (row 11+)
        for i, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=1)):
            cell_val = row[0].value
            if cell_val and str(cell_val).strip() != '':
                sl_name = str(cell_val).strip()
                sl_data = df_reg[df_reg.apply(lambda r: (str(r['考试科目名称']).strip() + str(r['考试等级']).strip()) == sl_name, axis=1)]
                s_total = sl_data['用户id'].nunique()
                s_dom = sl_data[sl_data['学员地区'] == '国内']['用户id'].nunique()
                s_ovs = sl_data[sl_data['学员地区'] != '国内']['用户id'].nunique()
                r_idx = 11 + i
                ws.cell(row=r_idx, column=7, value=s_total)
                ws.cell(row=r_idx, column=8, value=s_dom)
                ws.cell(row=r_idx, column=9, value=s_ovs)
                ws.cell(row=r_idx, column=10, value=round(s_dom / s_total, 6) if s_total > 0 else 0)
                ws.cell(row=r_idx, column=11, value=round(s_ovs / s_total, 6) if s_total > 0 else 0)
        print("  → 科目级别上期报名数据 本期列已回填")

    wb.save(FILE_HISTORY)
    print(f"  → 往期数据.xlsx 已更新保存")

    # 另存一份往期数据到历史归档目录
    history_backup = os.path.join(ARCHIVE_DIR, f"往期数据_{quarter_label}.xlsx")
    shutil.copy2(FILE_HISTORY, history_backup)
    print(f"  → 往期数据已备份到: {history_backup}")

except Exception as e:
    print(f"  [警告] 回填往期数据时出错: {e}")
    print("  请确保往期数据.xlsx未被其他程序打开")

print("\n" + "=" * 50)
print("  归档完成！")
print("=" * 50)
input("\n按回车退出...")
