# -*- coding: utf-8 -*-
"""
考级数据可视化看板生成脚本
运行方式：python generate_dashboard.py
输出：dashboard.html
"""

import pandas as pd
import numpy as np
import json
import os
from datetime import datetime, timedelta

# ========== 路径配置 ==========
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "数据源")

FILE_STUDENT_INFO = os.path.join(DATA_DIR, "考级学员信息表.xlsx")
FILE_TARGET = os.path.join(DATA_DIR, "目标明细.xlsx")
FILE_HOURS_STD = os.path.join(DATA_DIR, "考级各科目级别课时标准.xlsx")
FILE_HISTORY = os.path.join(DATA_DIR, "往期数据.xlsx")
FILE_OVERALL_TARGET = os.path.join(DATA_DIR, "本期考级目标.xlsx")

OUTPUT_FILE = os.path.join(BASE_DIR, "dashboard.html")
PASSWORD = "2026Q4"

# ========== 1. 读取数据 ==========
print("正在读取数据...")

# 1.1 考级学员信息表
df_reg = pd.read_excel(FILE_STUDENT_INFO, sheet_name="考级学员信息表")
# 如果是旧格式（8列），按位置赋列名；如果是新格式（46+列），按列名选取
if len(df_reg.columns) == 8:
    df_reg.columns = ['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']
else:
    # 新格式：按列名选取需要的8列
    df_reg = df_reg[['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']]
df_reg = df_reg[df_reg['用户id'].notna()]
df_reg['用户id'] = df_reg['用户id'].astype(int)
df_reg['报名日期'] = pd.to_datetime(df_reg['报名时间']).dt.date

# 1.2 课时标准
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

# 1.3 跟进学员名单
df_follow = pd.read_excel(FILE_TARGET, sheet_name="跟进学员名单（在班学员）")
df_follow.columns = [
    '老师ID', '老师工号', '老师艺名', '学区', '团队', '小组', '班级ID', '班级名称',
    '班级课表', '阶段', '星期', '上课时间', '学员ID', '学员昵称', '年龄', '地区',
    '有效续费次数', '是否师生在群', '是否有有效群', '班主任ID', '班主任姓名', '班主任艺名',
    '班主任大区', '班主任团队', '大V渠道', '累计报考次数', '26Q3报考科目', '26Q3成绩',
    '26Q1报考科目', '26Q1成绩', '25Q4报考科目', '25Q4成绩', '25Q3报考科目', '25Q3成绩',
    '25Q1报考科目', '25Q1成绩', '24Q4报考科目', '24Q4成绩', '24Q3报考科目', '24Q3成绩',
    '24Q1报考科目', '24Q1成绩', '儿童画', '儿童画成绩', '漫画', '漫画成绩', '素描', '素描成绩',
    '国画', '国画成绩', '速写', '速写成绩', '推荐儿童画报考级别', '推荐漫画报考级别',
    '推荐素描报考级别', '推荐国画报考级别', '推荐速写报考级别', '漫画-是否跳级', '素描-是否跳级',
    '分层', '剩余课时', '最新国画专题课课表名称', '最新国画专题课上课日期'
]
df_follow['学员ID'] = df_follow['学员ID'].astype(int)
# 将累计报考次数转换为数值类型，非数值转为NaN
df_follow['累计报考次数'] = pd.to_numeric(df_follow['累计报考次数'], errors='coerce')

# 满级判定：推荐漫画报考级别或推荐素描报考级别为"已满级"
df_follow['是否满级'] = df_follow.apply(
    lambda r: '是' if (str(r['推荐漫画报考级别']).strip() == '已满级' or str(r['推荐素描报考级别']).strip() == '已满级') else '否',
    axis=1
)

# 1.4 团队目标数据
df_team_target_raw = pd.read_excel(FILE_TARGET, sheet_name="团队", header=None)

# 解析学区目标 (Row 8-14)
district_targets = {}
for i in range(8, 14):
    row = df_team_target_raw.iloc[i]
    name = str(row.iloc[0]).strip()
    if name and name != '总计' and name != 'nan':
        district_targets[name] = {
            'target': int(row.iloc[2]),
            'pool_count': int(row.iloc[3])
        }

# 解析团队目标 (Row 18-29)
team_targets = {}
for i in range(18, 30):
    row = df_team_target_raw.iloc[i]
    district = str(row.iloc[0]).strip()
    team = str(row.iloc[1]).strip()
    if district != '总计' and team != 'nan' and team != 'None':
        team_targets[f"{district}_{team}"] = {
            'district': district,
            'team': team,
            'target': int(row.iloc[3]),
            'pool_count': int(row.iloc[4])
        }

# 解析小组目标 (Row 33-64；表头在第32行 学区|团队|小组|目标学员数_精准|学员数|报名率，比团队区少“向上取整”列)
group_targets = {}
for i in range(33, 65):
    row = df_team_target_raw.iloc[i]
    district = str(row.iloc[0]).strip()
    team = str(row.iloc[1]).strip()
    group = str(row.iloc[2]).strip()
    if district != '总计' and group != 'nan' and group != 'None':
        group_targets[f"{district}_{team}_{group}"] = {
            'district': district,
            'team': team,
            'group': group,
            'target': int(row.iloc[3]),
            'pool_count': int(row.iloc[4])
        }

# 1.5 老师目标
df_teacher = pd.read_excel(FILE_TARGET, sheet_name="老师")
# 取前9列
df_teacher = df_teacher.iloc[:, :9]
df_teacher.columns = ['老师ID', '老师艺名', '学区', '团队', '小组', '目标学员数_精准', '目标学员数_向上取整', '学员数', '报名率']
df_teacher = df_teacher[['老师ID', '老师艺名', '学区', '团队', '小组', '目标学员数_向上取整', '学员数']]
df_teacher.rename(columns={'目标学员数_向上取整': '目标'}, inplace=True)

# 1.6 MTD数据
# 新格式: 每行一天，列为：日期、星期、目标人数、每日流速
df_mtd = pd.read_excel(FILE_TARGET, sheet_name="MTD", header=1)
mtd_data = []
for _, row in df_mtd.iterrows():
    date_val = row.iloc[0]  # 日期列
    weekday = str(row.iloc[1])  # 星期列
    daily_target = round(row.iloc[2]) if pd.notna(row.iloc[2]) else 0  # 目标人数列（四舍五入）
    daily_flow = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0  # 每日流速列
    if pd.notna(date_val):
        # date_val已经是datetime对象
        if isinstance(date_val, str):
            actual_date = pd.to_datetime(date_val)
        else:
            actual_date = date_val
        mtd_data.append({
            'date': actual_date,
            'weekday': weekday,
            'daily_target': daily_target,
            'daily_flow': daily_flow
        })

# 1.7 往期数据
df_history_overall = pd.read_excel(FILE_HISTORY, sheet_name="与往期总体对比", header=0)
df_history_region = pd.read_excel(FILE_HISTORY, sheet_name="国内海外对比", header=0)
df_history_repeat = pd.read_excel(FILE_HISTORY, sheet_name="复报首报往期对比", header=0)
df_history_subject = pd.read_excel(FILE_HISTORY, sheet_name="科目往期对比", header=None)
df_history_tier = pd.read_excel(FILE_HISTORY, sheet_name="学员分层的往期对比", header=0)

# 1.8 科目级别上期报名数据
df_subject_level_last = pd.read_excel(FILE_HISTORY, sheet_name="科目级别上期报名数据")
# 解析：第0行是子表头（合计/国内/海外），从第1行开始是数据
# 结构: 报名科目 | 上期报名人数(合计,国内,海外) | 上期人数占比(国内,海外) | 本期报名人数(合计,国内,海外) | 本期人数占比(国内,海外)
subject_level_last_map = {}
for _, row in df_subject_level_last.iterrows():
    name = str(row.iloc[0]).strip()
    if name == 'nan' or name == '合计' or name == 'NaN':
        continue
    total = row.iloc[1]
    if pd.notna(total):
        subject_level_last_map[name] = int(total)

# 1.9 读取本期考级目标 - 季度字段
df_overall_target = pd.read_excel(FILE_OVERALL_TARGET)
quarter_label = ""
for _, row in df_overall_target.iterrows():
    if str(row.iloc[0]).strip() == '季度':
        quarter_label = str(row.iloc[1]).strip()
        break

# ========== 2. 数据计算 ==========
print("正在计算指标...")

# 2.1 基础指标
total_reg_count = df_reg['用户id'].nunique()  # 报名人数（去重）

# 从"本期考级目标.xlsx"读取目标值
df_current_target = pd.read_excel(FILE_OVERALL_TARGET, sheet_name='Sheet1')
total_target = int(df_current_target[df_current_target['维度'] == '目标人数']['目标'].values[0])
target_hours = int(df_current_target[df_current_target['维度'] == '目标课时']['目标'].values[0])

pool_total = len(df_follow)  # 跟进学员总数

# 课时达成计算
def calc_hours(row):
    subject = str(row['考试科目名称']).strip()
    level = str(row['考试等级']).strip()
    # 构造匹配key
    if '花鸟' in subject or subject == '国画':
        key = f"国画花鸟{level}"
    else:
        key = f"{subject}{level}"
    return hours_map.get(key, 0)

df_reg['课时'] = df_reg.apply(calc_hours, axis=1)
total_hours_achieved = df_reg['课时'].sum()

# GMV
total_gmv = df_reg['考级资格消耗'].sum()

# 报名完成率
completion_rate = total_reg_count / total_target

# 课时完成率
hours_completion_rate = total_hours_achieved / target_hours

# 报名率 = 池内已报名人数 / 池内总人数
registered_ids = set(df_reg['用户id'].unique())
follow_ids = set(df_follow['学员ID'].unique())
pool_registered = registered_ids & follow_ids
reg_rate = len(pool_registered) / pool_total

# 2.2 匹配报名学员到跟进学员名单
df_follow['是否本期报名'] = df_follow['学员ID'].isin(registered_ids).astype(int)

# 2.3 首报/复报判定
# 复报：累计报考次数>=1 且本期有报名
# 首报：累计报考次数为空或0 且本期有报名
df_follow_registered = df_follow[df_follow['是否本期报名'] == 1].copy()
df_follow_registered['报名类型'] = df_follow_registered['累计报考次数'].apply(
    lambda x: '首报' if (pd.isna(x) or x == 0) else '复报'
)

# 池外报名学员（在考级学员信息表中但不在跟进学员名单中的）
outside_pool_ids = registered_ids - follow_ids
df_reg_outside = df_reg[df_reg['用户id'].isin(outside_pool_ids)]
# 池外学员全部视为首报
outside_first_count = df_reg_outside['用户id'].nunique()

first_report_count = len(df_follow_registered[df_follow_registered['报名类型'] == '首报']) + outside_first_count
repeat_report_count = len(df_follow_registered[df_follow_registered['报名类型'] == '复报'])
total_report_count = first_report_count + repeat_report_count

# 2.4 每日报名趋势
daily_reg = df_reg.groupby('报名日期')['用户id'].nunique().reset_index()
daily_reg.columns = ['日期', '报名人数']
daily_reg = daily_reg.sort_values('日期')
daily_reg['累计报名'] = daily_reg['报名人数'].cumsum()

# MTD累计目标
mtd_cumulative = []
cum = 0
for item in mtd_data:
    cum += item['daily_target']
    # 匹配实际报名人数
    date_str = item['date'].strftime('%Y-%m-%d')
    actual_daily = 0
    for _, row in daily_reg.iterrows():
        if str(row['日期']) == date_str:
            actual_daily = int(row['报名人数'])
            break
    # 使用MTD配置的流速
    target_flow = item['daily_flow']
    actual_flow = actual_daily / total_target if total_target > 0 else 0
    mtd_cumulative.append({
        'date': date_str,
        'label': f"{item['date'].month}.{item['date'].day}\n{item['weekday'].replace('星期', '周')}",
        'short_label': f"{item['date'].month}.{item['date'].day}",
        'weekday': item['weekday'].replace('星期', '周'),
        'daily_target': item['daily_target'],
        'cum_target': cum,
        'actual_daily': actual_daily,
        'target_flow': round(target_flow, 6),
        'actual_flow': round(actual_flow, 6)
    })

# 2.5 学区统计
districts = ['教学一区', '教学二区', '教学三区', 'C区', 'DEMO特战队']
df_follow['学区_clean'] = df_follow['学区'].apply(lambda x: x if x in districts else '无归属')

district_stats = []
# 统计各学区（基于跟进学员名单中有明确学区的学员）
for d in districts:
    pool = df_follow[df_follow['学区_clean'] == d]
    pool_count = len(pool)
    reg_count = pool['是否本期报名'].sum()
    target = district_targets.get(d, {}).get('target', 0)
    rate = reg_count / pool_count if pool_count > 0 else 0
    comp_rate = reg_count / target if target > 0 else 0
    # 中国大陆 / 中国港澳台 / 海外
    domestic = pool[pool['地区'] == '国内']
    hkmotw = pool[pool['地区'] == '港澳台']
    overseas = pool[~pool['地区'].isin(['国内', '港澳台'])]
    dom_reg = domestic['是否本期报名'].sum()
    hkm_reg = hkmotw['是否本期报名'].sum()
    ovs_reg = overseas['是否本期报名'].sum()
    district_stats.append({
        'district': d,
        'pool_count': pool_count,
        'reg_count': int(reg_count),
        'target': target,
        'completion_rate': comp_rate,
        'reg_rate': rate,
        'domestic_reg': int(dom_reg),
        'hkmotw_reg': int(hkm_reg),
        'overseas_reg': int(ovs_reg)
    })

# 无归属 = 考级学员信息表中报名学员，扣除已归属到5个学区的学员
# 已归属学区的学员ID集合
attributed_ids = set()
for d in districts:
    pool = df_follow[df_follow['学区_clean'] == d]
    attributed_ids.update(pool[pool['是否本期报名'] == 1]['学员ID'].tolist())

# 无归属报名学员 = 所有报名学员 - 已归属学区的
all_reg_ids = set(df_reg['用户id'].unique())
unattributed_ids = all_reg_ids - attributed_ids
df_reg_unattributed = df_reg[df_reg['用户id'].isin(unattributed_ids)]
unattr_reg_count = df_reg_unattributed['用户id'].nunique()
unattr_dom = df_reg_unattributed[df_reg_unattributed['学员地区'] == '国内']['用户id'].nunique()
unattr_hkm = df_reg_unattributed[df_reg_unattributed['学员地区'] == '港澳台']['用户id'].nunique()
unattr_ovs = df_reg_unattributed[~df_reg_unattributed['学员地区'].isin(['国内', '港澳台'])]['用户id'].nunique()

district_stats.append({
    'district': '无归属',
    'pool_count': 0,
    'reg_count': int(unattr_reg_count),
    'target': 0,
    'completion_rate': 0,
    'reg_rate': 0,
    'domestic_reg': int(unattr_dom),
    'hkmotw_reg': int(unattr_hkm),
    'overseas_reg': int(unattr_ovs)
})

# 按完成率排序
district_stats.sort(key=lambda x: x['completion_rate'], reverse=True)

# 学区汇总行
district_summary = {
    'district': '学区汇总',
    'pool_count': sum(d['pool_count'] for d in district_stats if d['district'] != '无归属'),
    'reg_count': sum(d['reg_count'] for d in district_stats if d['district'] != '无归属'),
    'target': sum(d['target'] for d in district_stats if d['district'] != '无归属'),
    'completion_rate': 0,
    'reg_rate': 0,
    'domestic_reg': sum(d['domestic_reg'] for d in district_stats if d['district'] != '无归属'),
    'hkmotw_reg': sum(d['hkmotw_reg'] for d in district_stats if d['district'] != '无归属'),
    'overseas_reg': sum(d['overseas_reg'] for d in district_stats if d['district'] != '无归属')
}
district_summary['completion_rate'] = district_summary['reg_count'] / district_summary['target'] if district_summary['target'] > 0 else 0
district_summary['reg_rate'] = district_summary['reg_count'] / district_summary['pool_count'] if district_summary['pool_count'] > 0 else 0

# 2.6 中国大陆/中国港澳台/海外统计
domestic_pool = df_follow[df_follow['地区'] == '国内']
hkmotw_pool = df_follow[df_follow['地区'] == '港澳台']
overseas_pool = df_follow[~df_follow['地区'].isin(['国内', '港澳台'])]
domestic_reg_count = domestic_pool['是否本期报名'].sum()
hkmotw_reg_count = hkmotw_pool['是否本期报名'].sum()
overseas_reg_count = overseas_pool['是否本期报名'].sum()
domestic_rate = domestic_reg_count / len(domestic_pool) if len(domestic_pool) > 0 else 0
hkmotw_rate = hkmotw_reg_count / len(hkmotw_pool) if len(hkmotw_pool) > 0 else 0
overseas_rate = overseas_reg_count / len(overseas_pool) if len(overseas_pool) > 0 else 0

# 中国大陆/港澳台/海外课时
df_reg_with_region = df_reg.copy()
domestic_hours = df_reg_with_region[df_reg_with_region['学员地区'] == '国内']['课时'].sum()
hkmotw_hours = df_reg_with_region[df_reg_with_region['学员地区'] == '港澳台']['课时'].sum()
overseas_hours = df_reg_with_region[~df_reg_with_region['学员地区'].isin(['国内', '港澳台'])]['课时'].sum()

region_data = {
    'domestic': {
        'pool': len(domestic_pool), 'reg': int(domestic_reg_count),
        'rate': domestic_rate, 'hours': int(domestic_hours),
        'last_rate': float(df_history_region.iloc[1]['上期报考率']) if pd.notna(df_history_region.iloc[1]['上期报考率']) else 0
    },
    'hkmotw': {
        'pool': len(hkmotw_pool), 'reg': int(hkmotw_reg_count),
        'rate': hkmotw_rate, 'hours': int(hkmotw_hours),
        'last_rate': None
    },
    'overseas': {
        'pool': len(overseas_pool), 'reg': int(overseas_reg_count),
        'rate': overseas_rate, 'hours': int(overseas_hours),
        'last_rate': float(df_history_region.iloc[2]['上期报考率']) if pd.notna(df_history_region.iloc[2]['上期报考率']) else 0
    }
}

# 2.7 科目汇总
subject_stats = df_reg.groupby('考试科目名称').agg(
    报名人数=('用户id', 'nunique'),
    课时数=('课时', 'sum')
).reset_index()
subject_stats = subject_stats.sort_values('报名人数', ascending=False)
subject_stats['人数占比'] = subject_stats['报名人数'] / subject_stats['报名人数'].sum()
subject_stats['课时占比'] = subject_stats['课时数'] / subject_stats['课时数'].sum()

# 2.7b 科目首报/复报拆分
# 池内已报名学员的首报/复报标记
df_reg_with_type = df_reg.merge(
    df_follow_registered[['学员ID', '报名类型']].rename(columns={'学员ID': '用户id'}),
    on='用户id', how='left'
)
# 池外学员标记为首报
df_reg_with_type['报名类型'] = df_reg_with_type['报名类型'].fillna('首报')

subject_first_repeat = df_reg_with_type.groupby(['考试科目名称', '报名类型']).agg(
    报名人数=('用户id', 'nunique')
).reset_index()
subject_first_stats = {}
for subj in subject_stats['考试科目名称'].tolist():
    first_c = subject_first_repeat[(subject_first_repeat['考试科目名称'] == subj) & (subject_first_repeat['报名类型'] == '首报')]['报名人数'].sum()
    repeat_c = subject_first_repeat[(subject_first_repeat['考试科目名称'] == subj) & (subject_first_repeat['报名类型'] == '复报')]['报名人数'].sum()
    total_c = first_c + repeat_c
    subject_first_stats[subj] = {
        'first': int(first_c),
        'repeat': int(repeat_c),
        'first_ratio': first_c / total_c if total_c > 0 else 0
    }

# 科目往期对比数据
subject_history = {}
subjects_list = ['漫画', '素描', '儿童画', '国画花鸟', '速写']
for i, subj in enumerate(subjects_list):
    row_idx = i + 4  # rows 5-9 in the sheet (0-indexed: 4-8)
    if row_idx < len(df_history_subject):
        row = df_history_subject.iloc[row_idx]
        subject_history[subj] = {
            'last_first_count': int(row.iloc[6]) if pd.notna(row.iloc[6]) else 0,
            'last_first_ratio': float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0
        }

# 2.8 科目级别分布
subject_level_stats = df_reg.groupby(['考试科目名称', '考试等级']).agg(
    报名人数=('用户id', 'nunique'),
    课时数=('课时', 'sum')
).reset_index()
subject_level_stats['科目级别'] = subject_level_stats['考试科目名称'] + subject_level_stats['考试等级']
subject_level_stats['人数占比'] = subject_level_stats['报名人数'] / subject_level_stats['报名人数'].sum()

# 2.8b 科目级别报名排名（本期 vs 上期）
subject_level_ranking = []
for _, row in subject_level_stats.iterrows():
    sl_name = row['科目级别']
    curr_count = int(row['报名人数'])
    last_count = subject_level_last_map.get(sl_name, 0)
    change = curr_count - last_count
    change_rate = change / last_count if last_count > 0 else 0
    subject_level_ranking.append({
        'name': sl_name,
        'current': curr_count,
        'last': last_count,
        'change': change,
        'change_rate': change_rate
    })
subject_level_ranking.sort(key=lambda x: x['current'], reverse=True)

# 2.9 学员分层统计
tier_stats = []
for tier in ['A层', 'B层', 'C层', 'D层', 'E层']:
    pool = df_follow[df_follow['分层'] == tier]
    pool_count = len(pool)
    reg_count = pool['是否本期报名'].sum()
    rate = reg_count / pool_count if pool_count > 0 else 0
    tier_stats.append({
        'tier': tier,
        'pool_count': pool_count,
        'reg_count': int(reg_count),
        'rate': rate
    })

# 匹配往期分层数据
tier_history_map = {}
for _, row in df_history_tier.iterrows():
    t = str(row.iloc[0]).strip()
    if t in ['A层', 'B层', 'C层', 'D层', 'E层', '总计']:
        tier_history_map[t] = {
            'last_rate': float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
            'last_target': int(row.iloc[7]) if pd.notna(row.iloc[7]) else 0,
            'last_reg': int(row.iloc[8]) if pd.notna(row.iloc[8]) else 0
        }

for ts in tier_stats:
    # 本期新增分层（如E层）无往期数据，环比留空
    if ts['tier'] in tier_history_map:
        ts['last_rate'] = tier_history_map[ts['tier']]['last_rate']
        ts['rate_change'] = ts['rate'] - ts['last_rate']
        ts['has_history'] = True
    else:
        ts['last_rate'] = None
        ts['rate_change'] = None
        ts['has_history'] = False

# 2.10 阶段分布
stage_order = ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', '高阶', '院派', '大师', '不在班']
stage_stats = []
for stage in stage_order:
    pool = df_follow[df_follow['阶段'] == stage]
    pool_count = len(pool)
    reg_count = pool['是否本期报名'].sum()
    rate = reg_count / pool_count if pool_count > 0 else 0
    stage_stats.append({
        'stage': stage,
        'pool_count': pool_count,
        'reg_count': int(reg_count),
        'rate': rate
    })

# 没有匹配到阶段的学员归入"不在班"
unmatched_stage = df_follow[~df_follow['阶段'].isin(stage_order)]
if len(unmatched_stage) > 0:
    for ss in stage_stats:
        if ss['stage'] == '不在班':
            ss['pool_count'] += len(unmatched_stage)
            ss['reg_count'] += int(unmatched_stage['是否本期报名'].sum())
            ss['rate'] = ss['reg_count'] / ss['pool_count'] if ss['pool_count'] > 0 else 0

# 过滤掉池内人数和报名人数都为0的阶段
stage_stats = [s for s in stage_stats if s['pool_count'] > 0 or s['reg_count'] > 0]

# 2.11 满级学员统计
df_follow_mj = df_follow[df_follow['是否满级'] == '是']
mj_total = len(df_follow_mj)
mj_reg = df_follow_mj['是否本期报名'].sum()
mj_rate = mj_reg / mj_total if mj_total > 0 else 0

mj_by_district = []
for d in districts + ['无归属']:
    pool = df_follow_mj[df_follow_mj['学区_clean'] == d]
    pc = len(pool)
    rc = pool['是否本期报名'].sum()
    rate = rc / pc if pc > 0 else 0
    mj_by_district.append({'district': d, 'total': pc, 'reg': int(rc), 'rate': rate})
mj_by_district.sort(key=lambda x: x['rate'], reverse=True)

# 2.12 往期总体对比 - 根据当前季度动态处理
history_data = []
for _, row in df_history_overall.iterrows():
    time_label = str(row.iloc[0]).strip()
    if time_label in ['nan', '环比', '同比', '']:
        continue
    if time_label == quarter_label:  # 当前季度由下方实时数据追加，跳过Excel中的旧值
        continue
    # 所有历史数据都从Excel读取（包括26年Q1的完成数据）
    history_data.append({
        'time': time_label,
        'reg_count': int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
        'reg_rate': float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
        'hours': int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
        'target': int(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
        'completion_rate': float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
    })

# 追加当前季度的实时数据（用于趋势图）
history_data.append({
    'time': quarter_label.replace('年', '年'),  # 使用本期考级目标中的季度标识
    'reg_count': total_reg_count,
    'reg_rate': reg_rate,
    'hours': total_hours_achieved,
    'target': total_target,
    'completion_rate': completion_rate
})

# 计算环比同比 - 根据当前季度动态查找对比基准
# 提取当前季度的年份和季度号
import re
match = re.search(r'(\d+)年?Q(\d)', quarter_label)
if match:
    current_year = int(match.group(1))
    current_quarter = int(match.group(2))

    # 转换为两位数年份（与Excel格式匹配）
    year_short = current_year % 100

    # 环比：上一个季度（Q3->Q1, Q4->Q3等）
    if current_quarter == 1:
        last_year_short = (current_year - 1) % 100
        last_period_label = f'{last_year_short}年Q4'
    elif current_quarter == 3:
        last_period_label = f'{year_short}年Q1'
    else:
        last_period_label = f'{year_short}年Q{current_quarter-1}'

    # 同比：去年同季度
    same_year_short = (current_year - 1) % 100
    same_period_label = f'{same_year_short}年Q{current_quarter}'

    current = history_data[-1]  # 当前季度数据（最后一项）
    last_period = next((h for h in history_data if h['time'] == last_period_label), None)
    same_period = next((h for h in history_data if h['time'] == same_period_label), None)
else:
    # 兼容旧逻辑（如果解析失败）
    current = next((h for h in history_data if '26' in h['time'] and 'Q3' in h['time']), None)
    last_period = next((h for h in history_data if '26' in h['time'] and 'Q1' in h['time']), None)
    same_period = next((h for h in history_data if '25' in h['time'] and 'Q3' in h['time']), None)

def calc_change(curr, prev):
    if prev and prev != 0:
        return (curr - prev) / prev
    return 0

history_changes = {}
if current and last_period:
    history_changes['环比'] = {
        'reg_count': calc_change(current['reg_count'], last_period['reg_count']),
        'reg_rate': current['reg_rate'] - last_period['reg_rate'],
        'hours': calc_change(current['hours'], last_period['hours']),
        'completion_rate': calc_change(current['completion_rate'], last_period['completion_rate'])
    }
if current and same_period:
    history_changes['同比'] = {
        'reg_count': calc_change(current['reg_count'], same_period['reg_count']),
        'reg_rate': current['reg_rate'] - same_period['reg_rate'],
        'hours': calc_change(current['hours'], same_period['hours']),
        'completion_rate': calc_change(current['completion_rate'], same_period['completion_rate'])
    }

# 2.13 首报复报往期对比 - 所有历史从Excel读取，当前季度追加
repeat_history = []
seen_labels = set()  # 用于去重（归一化：去掉"年"字，避免"2026Q3"与"2026年Q3"重复）
def norm_label(s):
    return str(s).replace('年', '').strip()
for _, row in df_history_repeat.iterrows():
    label = str(row.iloc[0]).strip()
    # 过滤掉空值、同比、环比标签
    if label in ['nan', ''] or '同比' in label or '环比' in label:
        continue
    if pd.notna(row.iloc[1]) and norm_label(label) not in seen_labels:
        repeat_history.append({
            'time': label,
            'first_count': int(row.iloc[1]),
            'first_ratio': float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
            'repeat_count': int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
            'repeat_ratio': float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0
        })
        seen_labels.add(norm_label(label))
# 追加当前季度
current_label = f"{quarter_label[:4]}年{quarter_label[4:]}"
if norm_label(current_label) not in seen_labels:
    repeat_history.append({
        'time': current_label,
        'first_count': first_report_count,
        'first_ratio': first_report_count / total_report_count if total_report_count > 0 else 0,
        'repeat_count': repeat_report_count,
        'repeat_ratio': repeat_report_count / total_report_count if total_report_count > 0 else 0
    })

# 按时间排序（从旧到新）
def parse_time_label(label):
    """解析时间标签，返回(年份, 季度)用于排序"""
    import re
    match = re.search(r'(\d+)年?Q(\d)', label)
    if match:
        year = int(match.group(1))
        # 处理两位年份
        if year < 100:
            year = 2000 + year
        quarter = int(match.group(2))
        return (year, quarter)
    return (0, 0)  # 无法解析的放最前面

repeat_history.sort(key=lambda x: parse_time_label(x['time']))

# 2.14 团队页面数据准备
def get_team_data(district_name):
    """获取某学区下的所有团队/小组/老师数据"""
    if district_name == '无归属':
        pool = df_follow[df_follow['学区_clean'] == '无归属']
    else:
        pool = df_follow[df_follow['学区'] == district_name]

    result = {'teams': [], 'groups': [], 'teachers_top3': []}

    # 团队排名
    teams = pool.groupby('团队').agg(
        pool_count=('学员ID', 'count'),
        reg_count=('是否本期报名', 'sum')
    ).reset_index()
    teams = teams[teams['团队'].notna() & (teams['团队'] != '-')]

    for _, t in teams.iterrows():
        key = f"{district_name}_{t['团队']}"
        target = team_targets.get(key, {}).get('target', 0)
        comp = t['reg_count'] / target if target > 0 else 0
        result['teams'].append({
            'name': t['团队'],
            'reg_count': int(t['reg_count']),
            'target': target,
            'completion_rate': comp
        })
    result['teams'].sort(key=lambda x: x['completion_rate'], reverse=True)

    # 小组排名
    groups = pool.groupby(['团队', '小组']).agg(
        pool_count=('学员ID', 'count'),
        reg_count=('是否本期报名', 'sum')
    ).reset_index()
    groups = groups[groups['小组'].notna() & (groups['小组'] != '-')]

    for _, g in groups.iterrows():
        key = f"{district_name}_{g['团队']}_{g['小组']}"
        target = group_targets.get(key, {}).get('target', 0)
        comp = g['reg_count'] / target if target > 0 else 0
        result['groups'].append({
            'name': g['小组'],
            'team': g['团队'],
            'reg_count': int(g['reg_count']),
            'target': target,
            'completion_rate': comp
        })
    result['groups'].sort(key=lambda x: x['completion_rate'], reverse=True)

    # 老师TOP3
    teacher_pool = df_teacher[df_teacher['学区'] == district_name] if district_name != '无归属' else pd.DataFrame()
    if district_name != '无归属' and len(teacher_pool) > 0:
        teacher_reg = pool.groupby('老师艺名')['是否本期报名'].sum().reset_index()
        teacher_reg.columns = ['老师艺名', '报名人数']
        teacher_merged = teacher_reg.merge(teacher_pool[['老师艺名', '目标']], on='老师艺名', how='left')
        teacher_merged['完成率'] = teacher_merged.apply(
            lambda r: r['报名人数'] / r['目标'] if pd.notna(r['目标']) and r['目标'] > 0 else 0, axis=1
        )
        teacher_merged = teacher_merged.sort_values('报名人数', ascending=False).head(3)
        for _, tr in teacher_merged.iterrows():
            result['teachers_top3'].append({
                'name': tr['老师艺名'],
                'reg_count': int(tr['报名人数']),
                'target': int(tr['目标']) if pd.notna(tr['目标']) else 0,
                'completion_rate': tr['完成率']
            })

    # 科目级别占比（本学区）
    pool_ids = set(pool['学员ID'].unique())
    district_reg = df_reg[df_reg['用户id'].isin(pool_ids)]
    sl_stats = district_reg.groupby(['考试科目名称', '考试等级']).agg(
        报名人数=('用户id', 'nunique')
    ).reset_index()
    sl_stats['科目级别'] = sl_stats['考试科目名称'] + sl_stats['考试等级']
    total_district_reg = sl_stats['报名人数'].sum()
    sl_stats['占比'] = sl_stats['报名人数'] / total_district_reg if total_district_reg > 0 else 0
    result['subject_level'] = sl_stats[['科目级别', '报名人数', '占比']].to_dict('records')

    # 阶段报名率（本学区）
    stage_district = []
    for stage in stage_order:
        sp = pool[pool['阶段'] == stage]
        pc = len(sp)
        rc = sp['是否本期报名'].sum()
        rate = rc / pc if pc > 0 else 0
        stage_district.append({'stage': stage, 'pool_count': pc, 'reg_count': int(rc), 'rate': rate})
    result['stages'] = stage_district

    # 中国大陆/中国港澳台/海外报名率（本学区）
    dom = pool[pool['地区'] == '国内']
    hkm = pool[pool['地区'] == '港澳台']
    ovs = pool[~pool['地区'].isin(['国内', '港澳台'])]
    result['region'] = {
        'domestic': {'pool': len(dom), 'reg': int(dom['是否本期报名'].sum()),
                     'rate': dom['是否本期报名'].sum() / len(dom) if len(dom) > 0 else 0},
        'hkmotw': {'pool': len(hkm), 'reg': int(hkm['是否本期报名'].sum()),
                   'rate': hkm['是否本期报名'].sum() / len(hkm) if len(hkm) > 0 else 0},
        'overseas': {'pool': len(ovs), 'reg': int(ovs['是否本期报名'].sum()),
                     'rate': ovs['是否本期报名'].sum() / len(ovs) if len(ovs) > 0 else 0}
    }

    # 学员分层报名率（本学区）
    tier_district = []
    for tier in ['A层', 'B层', 'C层', 'D层', 'E层']:
        tp = pool[pool['分层'] == tier]
        pc = len(tp)
        rc = tp['是否本期报名'].sum()
        rate = rc / pc if pc > 0 else 0
        tier_district.append({'tier': tier, 'pool_count': pc, 'reg_count': int(rc), 'rate': rate})
    result['tiers'] = tier_district

    # 满级报名率（本学区）
    mj_pool = pool[pool['是否满级'] == '是']
    mj_pc = len(mj_pool)
    mj_rc = mj_pool['是否本期报名'].sum()
    mj_r = mj_rc / mj_pc if mj_pc > 0 else 0
    non_mj_pool = pool[pool['是否满级'] != '是']
    non_mj_pc = len(non_mj_pool)
    non_mj_rc = non_mj_pool['是否本期报名'].sum()
    non_mj_r = non_mj_rc / non_mj_pc if non_mj_pc > 0 else 0
    result['mj'] = {'mj_rate': mj_r, 'non_mj_rate': non_mj_r, 'mj_count': mj_pc, 'mj_reg': int(mj_rc)}

    # 首报/复报报名率（本学区）
    pool_first = pool[(pool['累计报考次数'].isna()) | (pool['累计报考次数'] == 0)]
    pool_repeat = pool[(pool['累计报考次数'].notna()) & (pool['累计报考次数'] >= 1)]
    first_pc = len(pool_first)
    first_rc = pool_first['是否本期报名'].sum()
    repeat_pc = len(pool_repeat)
    repeat_rc = pool_repeat['是否本期报名'].sum()
    result['first_repeat'] = {
        'first': {'pool': first_pc, 'reg': int(first_rc), 'rate': first_rc / first_pc if first_pc > 0 else 0},
        'repeat': {'pool': repeat_pc, 'reg': int(repeat_rc), 'rate': repeat_rc / repeat_pc if repeat_pc > 0 else 0}
    }

    return result

# 大盘数据（用于对比）
overall_benchmarks = {
    'stages': {s['stage']: s['rate'] for s in stage_stats},
    'region': {'domestic': domestic_rate, 'hkmotw': hkmotw_rate, 'overseas': overseas_rate},
    'tiers': {t['tier']: t['rate'] for t in tier_stats},
    'mj': {'mj_rate': mj_rate, 'non_mj_rate': (total_reg_count - mj_reg) / (pool_total - mj_total) if (pool_total - mj_total) > 0 else 0},
    'first_repeat': {
        'first': len(df_follow[(df_follow['累计报考次数'].isna()) | (df_follow['累计报考次数'] == 0)]),
        'first_reg': df_follow[(df_follow['累计报考次数'].isna()) | (df_follow['累计报考次数'] == 0)]['是否本期报名'].sum(),
        'repeat': len(df_follow[(df_follow['累计报考次数'].notna()) & (df_follow['累计报考次数'] >= 1)]),
        'repeat_reg': df_follow[(df_follow['累计报考次数'].notna()) & (df_follow['累计报考次数'] >= 1)]['是否本期报名'].sum()
    },
    'subject_level': subject_level_stats[['科目级别', '人数占比']].set_index('科目级别')['人数占比'].to_dict()
}
# 补充大盘首报复报rate
fp = overall_benchmarks['first_repeat']['first']
fr = overall_benchmarks['first_repeat']['first_reg']
rp = overall_benchmarks['first_repeat']['repeat']
rr = overall_benchmarks['first_repeat']['repeat_reg']
overall_benchmarks['first_repeat']['first_rate'] = fr / fp if fp > 0 else 0
overall_benchmarks['first_repeat']['repeat_rate'] = rr / rp if rp > 0 else 0

# 2.15 荣誉榜数据 & 全部老师个人排名数据
# 加权公式：首报×1.5 + 复报×1.0
teacher_reg_all = df_follow[df_follow['是否本期报名'] == 1].copy()
teacher_reg_all = teacher_reg_all[teacher_reg_all['老师艺名'].notna() & (teacher_reg_all['老师艺名'] != '-')]
# 判定首报/复报
teacher_reg_all['报名类型'] = teacher_reg_all['累计报考次数'].apply(
    lambda x: '首报' if (pd.isna(x) or x == 0) else '复报'
)
# 按老师汇总报名与首报
teacher_reg_agg = teacher_reg_all.groupby('老师艺名').agg(
    报名人数=('学员ID', 'count'),
    首报人数=('报名类型', lambda x: (x == '首报').sum())
).reset_index()

# 以全部老师表(df_teacher)为基准合并，涵盖所有老师（含0报名老师）
teacher_all_df = df_teacher[['老师艺名', '学区', '团队', '小组', '目标']].rename(columns={'目标': '目标学员数'}).copy()
teacher_all_df = teacher_all_df.merge(teacher_reg_agg, on='老师艺名', how='left')
teacher_all_df['报名人数'] = teacher_all_df['报名人数'].fillna(0).astype(int)
teacher_all_df['首报人数'] = teacher_all_df['首报人数'].fillna(0).astype(int)
teacher_all_df['复报人数'] = teacher_all_df['报名人数'] - teacher_all_df['首报人数']
teacher_all_df['加权报名人数'] = teacher_all_df['首报人数'] * 1.5 + teacher_all_df['复报人数'] * 1.0
teacher_all_df['目标学员数'] = teacher_all_df['目标学员数'].fillna(0).astype(int)
teacher_all_df['完成率'] = teacher_all_df.apply(
    lambda r: r['报名人数'] / r['目标学员数'] if r['目标学员数'] > 0 else 0, axis=1
)
# 按加权报名人数降序，并列时按完成率降序
teacher_all_df = teacher_all_df.sort_values(['加权报名人数', '完成率', '报名人数'], ascending=[False, False, False]).reset_index(drop=True)

all_teachers_ranking = []
for i, row in teacher_all_df.iterrows():
    all_teachers_ranking.append({
        'rank': i + 1,
        'name': str(row['老师艺名']),
        'target': int(row['目标学员数']),
        'reg_count': int(row['报名人数']),
        'completion_rate': round(row['完成率'], 4),
        'first_count': int(row['首报人数']),
        'weighted_count': round(row['加权报名人数'], 1),
        'district': str(row['学区']) if pd.notna(row['学区']) else '-',
        'team': str(row['团队']) if pd.notna(row['团队']) else '-',
        'group': str(row['小组']) if pd.notna(row['小组']) else '-'
    })

# 个人TOP10（总览荣誉榜展示前10名）
honor_teachers = all_teachers_ranking[:10]

# 团队TOP4：按完成率降序（全局所有团队）
honor_teams = []
for key, val in team_targets.items():
    district = val['district']
    team_name = val['team']
    target = val['target']
    if target <= 0:
        continue
    pool = df_follow[(df_follow['学区'] == district) & (df_follow['团队'] == team_name)]
    reg_c = pool['是否本期报名'].sum()
    comp_rate = reg_c / target
    honor_teams.append({
        'name': team_name,
        'completion_rate': comp_rate,
        'district': district
    })
honor_teams.sort(key=lambda x: x['completion_rate'], reverse=True)
honor_teams = [{'rank': i+1, **t} for i, t in enumerate(honor_teams[:4])]

# 小组TOP5：按完成率降序（全局所有小组）
honor_groups = []
for key, val in group_targets.items():
    district = val['district']
    team_name = val['team']
    group_name = val['group']
    target = val['target']
    if target <= 0 or not group_name or group_name in ['-', 'nan', 'None', '']:
        continue
    pool = df_follow[(df_follow['学区'] == district) & (df_follow['团队'] == team_name) & (df_follow['小组'] == group_name)]
    reg_c = pool['是否本期报名'].sum()
    comp_rate = reg_c / target
    honor_groups.append({
        'name': group_name,
        'completion_rate': comp_rate,
        'district': district
    })
honor_groups.sort(key=lambda x: x['completion_rate'], reverse=True)
honor_groups = [{'rank': i+1, **g} for i, g in enumerate(honor_groups[:5])]

# 计算各学区数据
all_districts = districts + ['无归属']
team_page_data = {}
for d in all_districts:
    team_page_data[d] = get_team_data(d)

print("数据计算完成。")

# ========== 3. 生成分析结论 ==========
print("正在生成分析结论...")

def gen_overview_conclusion():
    """总体往期对比结论"""
    parts = []
    parts.append(f"本期报名{total_reg_count}人，完成率{completion_rate*100:.1f}%。")
    if last_period:
        chg = history_changes.get('环比', {})
        parts.append(f"环比上期（{last_period['reg_count']}人）{'增长' if chg.get('reg_count',0)>=0 else '下降'}{abs(chg.get('reg_count',0))*100:.1f}%。")
    if same_period:
        chg = history_changes.get('同比', {})
        parts.append(f"同比去年同期（{same_period['reg_count']}人）{'增长' if chg.get('reg_count',0)>=0 else '下降'}{abs(chg.get('reg_count',0))*100:.1f}%。")
    return ''.join(parts)

def gen_daily_conclusion():
    """每日趋势结论"""
    if len(daily_reg) == 0:
        return "暂无报名数据。"
    last_day = daily_reg.iloc[-1]
    total_cum = int(daily_reg['累计报名'].iloc[-1])
    last_mtd_target = mtd_cumulative[-1]['cum_target'] if mtd_cumulative else total_target
    progress = total_cum / last_mtd_target if last_mtd_target > 0 else 0
    return f"截至{last_day['日期']}，累计报名{total_cum}人，达MTD终期目标{progress*100:.1f}%。最后一日新增{int(last_day['报名人数'])}人。"

def gen_district_conclusion():
    """学区完成情况结论"""
    if not district_stats:
        return ""
    valid_districts = [d for d in district_stats if d['district'] != '无归属']
    if not valid_districts:
        return ""
    best = valid_districts[0]
    worst = valid_districts[-1]
    return f"完成率最高：{best['district']} {best['completion_rate']*100:.1f}%；最低：{worst['district']} {worst['completion_rate']*100:.1f}%。整体完成率{completion_rate*100:.1f}%。"

def gen_region_conclusion():
    """中国大陆/中国港澳台/海外结论"""
    dom_chg = region_data['domestic']['rate'] - region_data['domestic']['last_rate']
    ovs_chg = region_data['overseas']['rate'] - region_data['overseas']['last_rate']
    return f"中国大陆报名率{region_data['domestic']['rate']*100:.2f}%，环比上期{'上升' if dom_chg>=0 else '下降'}{abs(dom_chg)*100:.2f}个百分点；中国港澳台报名率{region_data['hkmotw']['rate']*100:.2f}%（本期新增细分）；海外报名率{region_data['overseas']['rate']*100:.2f}%，环比上期{'上升' if ovs_chg>=0 else '下降'}{abs(ovs_chg)*100:.2f}个百分点。"

def gen_subject_conclusion():
    """科目汇总结论"""
    if len(subject_stats) == 0:
        return ""
    top = subject_stats.iloc[0]
    return f"报名人数最多科目：{top['考试科目名称']} {int(top['报名人数'])}人，占比{top['人数占比']*100:.1f}%。"

def gen_tier_conclusion():
    """学员分层结论"""
    parts = []
    for ts in tier_stats:
        if ts.get('rate_change') is None:
            parts.append(f"{ts['tier']}报考率{ts['rate']*100:.2f}%")
        else:
            chg_str = f"{'↑' if ts['rate_change']>=0 else '↓'}{abs(ts['rate_change'])*100:.2f}%"
            parts.append(f"{ts['tier']}报考率{ts['rate']*100:.2f}%({chg_str})")
    return '；'.join(parts) + '。'

def gen_stage_conclusion():
    """阶段分布结论"""
    max_count = max(stage_stats, key=lambda x: x['reg_count'])
    max_rate = max(stage_stats, key=lambda x: x['rate'])
    return f"报名人数最多阶段：{max_count['stage']}（{max_count['reg_count']}人）；报名率最高：{max_rate['stage']}（{max_rate['rate']*100:.2f}%）。"

def gen_mj_conclusion():
    """满级学员结论"""
    best = mj_by_district[0] if mj_by_district else None
    base = f"满级学员整体报名率{mj_rate*100:.2f}%。"
    if best:
        base += f"报名率最高：{best['district']} {best['rate']*100:.2f}%"
    return base

def gen_repeat_conclusion():
    """首报复报结论"""
    if total_report_count == 0:
        return "暂无数据。"
    first_ratio = first_report_count / total_report_count * 100
    # 找上期数据对比（取排序后当前季度的前一期）
    cur_norm = f"{quarter_label[:4]}{quarter_label[4:]}"
    last_rh = None
    for r in repeat_history:
        if str(r['time']).replace('年', '').strip() == cur_norm:
            break
        last_rh = r
    chg = ""
    if last_rh:
        diff = first_ratio/100 - last_rh['first_ratio']
        chg = f"，环比上期{'上升' if diff>=0 else '下降'}{abs(diff)*100:.2f}%"
    return f"本期首报{first_report_count}人，占比{first_ratio:.1f}%{chg}；复报{repeat_report_count}人，占比{100-first_ratio:.1f}%。"

def gen_team_conclusion(district_name, data):
    """学区整体结论"""
    ds = next((d for d in district_stats if d['district'] == district_name), None)
    if not ds:
        return f"{district_name}暂无数据。"
    rank = district_stats.index(ds) + 1
    parts = [
        f"{district_name}本期报名{ds['reg_count']}人，完成率{ds['completion_rate']*100:.1f}%，排名第{rank}；",
        f"报名率{ds['reg_rate']*100:.2f}%，{'高于' if ds['reg_rate']>reg_rate else '低于'}大盘{abs(ds['reg_rate']-reg_rate)*100:.2f}个百分点。"
    ]
    if data.get('mj'):
        parts.append(f"满级学员报名率{data['mj']['mj_rate']*100:.2f}%。")
    return ''.join(parts)

# ========== 4. 生成HTML ==========
print("正在生成HTML...")

# 生成更新时间
update_time = datetime.now().strftime('%Y-%m-%d %H:%M')

# 准备JSON数据传入前端
dashboard_data = {
    'cards': {
        'reg_count': total_reg_count,
        'target': total_target,
        'completion_rate': round(completion_rate * 100, 1),
        'hours_achieved': int(total_hours_achieved),
        'hours_target': target_hours,
        'hours_completion': round(hours_completion_rate * 100, 1),
        'reg_rate': round(reg_rate * 100, 2),
        'gmv': int(total_gmv)
    },
    'history': history_data,
    'history_changes': {k: {kk: round(vv, 4) for kk, vv in v.items()} for k, v in history_changes.items()},
    'daily': {
        'dates': [str(d) for d in daily_reg['日期'].tolist()],
        'daily_count': daily_reg['报名人数'].tolist(),
        'cumulative': daily_reg['累计报名'].tolist()
    },
    'mtd': mtd_cumulative,
    'district_stats': district_stats,
    'district_summary': district_summary,
    'region': region_data,
    'subject_stats': subject_stats[['考试科目名称', '报名人数', '课时数', '人数占比', '课时占比']].to_dict('records'),
    'subject_first_stats': subject_first_stats,
    'subject_history': subject_history,
    'subject_level_stats': subject_level_stats[['科目级别', '报名人数', '课时数', '人数占比']].to_dict('records'),
    'subject_level_ranking': subject_level_ranking,
    'tier_stats': tier_stats,
    'stage_stats': stage_stats,
    'mj_by_district': mj_by_district,
    'mj_overall': {'total': mj_total, 'reg': int(mj_reg), 'rate': round(mj_rate, 4)},
    'repeat_history': repeat_history,
    'first_report': {'count': first_report_count, 'ratio': round(first_report_count/total_report_count, 4) if total_report_count > 0 else 0},
    'repeat_report': {'count': repeat_report_count, 'ratio': round(repeat_report_count/total_report_count, 4) if total_report_count > 0 else 0},
    'quarter_label': quarter_label,
    'honor': {
        'teachers': honor_teachers,
        'teams': honor_teams,
        'groups': honor_groups
    },
    'all_teachers': all_teachers_ranking,
    'conclusions': {
        'overview': gen_overview_conclusion(),
        'daily': gen_daily_conclusion(),
        'district': gen_district_conclusion(),
        'region': gen_region_conclusion(),
        'subject': gen_subject_conclusion(),
        'tier': gen_tier_conclusion(),
        'stage': gen_stage_conclusion(),
        'mj': gen_mj_conclusion(),
        'repeat': gen_repeat_conclusion()
    },
    'team_page': {},
    'benchmarks': {
        'stages': {s['stage']: round(s['rate'], 4) for s in stage_stats},
        'region': {'domestic': round(domestic_rate, 4), 'hkmotw': round(hkmotw_rate, 4), 'overseas': round(overseas_rate, 4)},
        'tiers': {t['tier']: round(t['rate'], 4) for t in tier_stats},
        'mj_rate': round(mj_rate, 4),
        'non_mj_rate': round(overall_benchmarks['mj']['non_mj_rate'], 4),
        'first_rate': round(overall_benchmarks['first_repeat']['first_rate'], 4),
        'repeat_rate': round(overall_benchmarks['first_repeat']['repeat_rate'], 4),
        'subject_level': {k: round(v, 4) for k, v in overall_benchmarks['subject_level'].items()}
    }
}

# 团队页面数据（不含无归属）
team_districts = [d for d in all_districts if d != '无归属']
for d in team_districts:
    data = team_page_data[d]
    ds = next((x for x in district_stats if x['district'] == d), None)
    rank = district_stats.index(ds) + 1 if ds else 0
    dashboard_data['team_page'][d] = {
        'teams': data['teams'],
        'groups': data['groups'],
        'teachers_top3': data['teachers_top3'],
        'subject_level': data['subject_level'],
        'stages': data['stages'],
        'region': data['region'],
        'tiers': data['tiers'],
        'mj': data['mj'],
        'first_repeat': data['first_repeat'],
        'conclusion': gen_team_conclusion(d, data),
        'kpi': {
            'reg_count': ds['reg_count'] if ds else 0,
            'target': ds['target'] if ds else 0,
            'completion_rate': round(ds['completion_rate'] * 100, 1) if ds else 0,
            'reg_rate': round(ds['reg_rate'] * 100, 2) if ds else 0,
            'pool_count': ds['pool_count'] if ds else 0,
            'rank': rank,
            'total_districts': len([x for x in district_stats if x['district'] != '无归属'])
        }
    }

# 序列化
data_json = json.dumps(dashboard_data, ensure_ascii=False, default=str)

# 生成完整HTML
html_content = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>考级数据看板 - ''' + quarter_label + '''</title>
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", Roboto, sans-serif; background: #F4F6F9; color: #1E293B; min-height: 100vh; }
#login-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: #F4F6F9; display: flex; align-items: center; justify-content: center; z-index: 9999; }
#login-box { background: #FFFFFF; border-radius: 14px; padding: 40px 48px; text-align: center; box-shadow: 0 12px 36px rgba(15,23,42,0.08); border: 1px solid #E2E8F0; }
#login-box h2 { color: #0F172A; margin-bottom: 20px; font-size: 22px; font-weight: 700; }
#login-box input { background: #F8FAFC; border: 1px solid #CBD5E1; border-radius: 8px; padding: 12px 16px; color: #0F172A; font-size: 16px; width: 240px; outline: none; transition: border-color 0.2s, box-shadow 0.2s; }
#login-box input:focus { border-color: #2563EB; background: #FFFFFF; box-shadow: 0 0 0 3px rgba(37,99,235,0.12); }
#login-box button { margin-top: 16px; background: #2563EB; color: #FFFFFF; border: none; border-radius: 8px; padding: 12px 36px; font-size: 16px; font-weight: 600; cursor: pointer; transition: background 0.2s; }
#login-box button:hover { background: #1D4ED8; }
#login-error { color: #DC2626; margin-top: 12px; font-size: 14px; font-weight: 500; display: none; }
.nav { background: #FFFFFF; border-bottom: 1px solid #E2E8F0; box-shadow: 0 2px 8px rgba(15,23,42,0.04); padding: 0; position: sticky; top: 0; z-index: 100; }
.nav-banner { background: linear-gradient(135deg, #EFF6FF 0%, #E0E7FF 50%, #F0F9FF 100%); padding: 28px 40px 16px; text-align: center; position: relative; overflow: hidden; border-bottom: 1px solid #E2E8F0; }
.nav-banner::before { content: ""; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: radial-gradient(ellipse at 50% 0%, rgba(37,99,235,0.08) 0%, transparent 70%); pointer-events: none; }
.nav-banner-title { font-size: 28px; font-weight: 800; color: #0F172A; letter-spacing: 2px; margin-bottom: 8px; position: relative; }
.nav-banner-sub { font-size: 14px; color: #475569; position: relative; }
.nav-banner-sub span { background: #2563EB; padding: 4px 14px; border-radius: 999px; margin: 0 6px; color: #FFFFFF; font-weight: 600; font-size: 13px; box-shadow: 0 2px 6px rgba(37,99,235,0.25); }
.nav-tabbar { display: flex; align-items: center; padding: 0 24px; height: 48px; background: #FFFFFF; justify-content: center; position: relative; }
.nav-tabs { display: flex; gap: 6px; }
.nav-tab { padding: 7px 22px; border-radius: 6px; cursor: pointer; font-size: 15px; font-weight: 600; color: #475569; transition: all 0.2s; }
.nav-tab:hover { color: #0F172A; background: #F1F5F9; }
.nav-tab.active { color: #FFFFFF; background: #2563EB; box-shadow: 0 2px 6px rgba(37,99,235,0.25); }
.nav-right { position: absolute; right: 24px; display: flex; align-items: center; height: 100%; }
.main { padding: 24px; max-width: 1400px; margin: 0 auto; }
.page { display: none; }
.page.active { display: block; }
.cards-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(155px, 1fr)); gap: 16px; margin-bottom: 24px; }
.card { background: #FFFFFF; border-radius: 10px; padding: 20px; border: 1px solid #E2E8F0; border-left: 4px solid #2563EB; box-shadow: 0 1px 4px rgba(15,23,42,0.03); }
.card-label { font-size: 13px; font-weight: 500; color: #475569; margin-bottom: 8px; letter-spacing: 0.5px; }
.card-value { font-size: 24px; font-weight: 700; color: #0F172A; }
.card-sub { font-size: 12px; color: #64748B; margin-top: 4px; }
.card:nth-child(2) { border-left-color: #7C3AED; }
.card:nth-child(3) { border-left-color: #059669; }
.card:nth-child(4) { border-left-color: #0891B2; }
.card:nth-child(5) { border-left-color: #D97706; }
.card:nth-child(6) { border-left-color: #DB2777; }
.card:nth-child(7) { border-left-color: #7C3AED; }
.card:nth-child(8) { border-left-color: #059669; }
.section { background: #FFFFFF; border-radius: 10px; padding: 24px; margin-bottom: 24px; border: 1px solid #E2E8F0; box-shadow: 0 1px 4px rgba(15,23,42,0.03); }
.section-title { font-size: 16px; font-weight: 700; color: #0F172A; margin-bottom: 10px; }
.section-conclusion { font-size: 13px; font-weight: 500; color: #B45309; background: #FFFBEB; border: 1px solid #FDE68A; border-left: 4px solid #F59E0B; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; line-height: 1.6; }
.chart-container { width: 100%; height: 350px; }
.chart-container-sm { width: 100%; height: 280px; }
.data-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
.data-table th { background: #F8FAFC; color: #334155; padding: 10px 12px; text-align: left; font-weight: 600; border-bottom: 2px solid #CBD5E1; border-top: 1px solid #E2E8F0; }
.data-table td { padding: 10px 12px; border-bottom: 1px solid #E2E8F0; color: #1E293B; position: relative; }
.data-table tr:hover td { background: #F0F7FF; }
.text-green { color: #059669; font-weight: 600; }
.text-red { color: #DC2626; font-weight: 600; }
.text-yellow { color: #D97706; font-weight: 600; }
.grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; }
@media (max-width: 900px) { .grid-2 { grid-template-columns: 1fr; } }
.district-tabs { display: flex; gap: 8px; margin-bottom: 24px; flex-wrap: wrap; }
.district-tab { padding: 8px 18px; border-radius: 6px; cursor: pointer; font-size: 13px; font-weight: 600; color: #475569; background: #FFFFFF; border: 1px solid #CBD5E1; transition: all 0.2s; }
.district-tab:hover { color: #2563EB; border-color: #2563EB; background: #EFF6FF; }
.district-tab.active { color: #FFFFFF; background: #2563EB; border-color: #2563EB; box-shadow: 0 2px 6px rgba(37,99,235,0.2); }
.top3 { display: flex; gap: 16px; margin-bottom: 24px; flex-wrap: wrap; }
.top3-item { background: #FFFFFF; border-radius: 10px; padding: 16px 24px; flex: 1; min-width: 200px; border: 1px solid #E2E8F0; box-shadow: 0 1px 4px rgba(15,23,42,0.03); }
.top3-item:nth-child(1) { background: linear-gradient(135deg, #FEF9C3 0%, #FFFBEB 100%); border-color: #FDE047; border-left: 4px solid #EAB308; }
.top3-item:nth-child(2) { background: linear-gradient(135deg, #F1F5F9 0%, #F8FAFC 100%); border-color: #CBD5E1; border-left: 4px solid #94A3B8; }
.top3-item:nth-child(3) { background: linear-gradient(135deg, #FFEDD5 0%, #FFF7ED 100%); border-color: #FDBA74; border-left: 4px solid #EA580C; }
.top3-rank { font-size: 24px; margin-right: 8px; }
.top3-name { font-size: 16px; font-weight: 700; color: #0F172A; display: inline; }
.top3-stats { font-size: 13px; color: #475569; margin-top: 6px; font-weight: 500; }
.district-kpi { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 24px; }
.district-kpi-card { background: #FFFFFF; border-radius: 10px; padding: 16px; text-align: center; border: 1px solid #E2E8F0; box-shadow: 0 1px 4px rgba(15,23,42,0.03); }
.district-kpi-label { font-size: 12px; font-weight: 500; color: #475569; margin-bottom: 6px; }
.district-kpi-value { font-size: 24px; font-weight: 700; color: #2563EB; }
.district-kpi-rank { font-size: 12px; font-weight: 600; color: #D97706; margin-top: 4px; }
</style>
</head>
<body>
<div id="login-overlay">
  <div id="login-box">
    <h2>考级数据看板</h2>
    <input type="password" id="pwd-input" placeholder="请输入访问密码" onkeydown="if(event.key==='Enter')checkPwd()">
    <br><button onclick="checkPwd()">进入</button>
    <div id="login-error">密码错误，请重试</div>
  </div>
</div>
<nav class="nav">
  <div class="nav-banner">
    <div class="nav-banner-title">美术考级数据看板</div>
    <div class="nav-banner-sub"><span>''' + quarter_label + '''</span></div>
  </div>
  <div class="nav-tabbar">
    <div class="nav-tabs">
      <div class="nav-tab active" onclick="switchPage('overview')">总览</div>
      <div class="nav-tab" onclick="switchPage('team')">团队</div>
      <div class="nav-tab" onclick="switchPage('personal')">个人</div>
    </div>
    <div class="nav-right" style="display:flex;align-items:center">
      <span style="font-size:12px;color:#475569;font-weight:500">更新时间：''' + update_time + '''</span>
    </div>
  </div>
</nav>
<div class="main">
<div id="page-overview" class="page active"></div>
<div id="page-team" class="page"></div>
<div id="page-personal" class="page"></div>
</div>

<script>
const DATA = ''' + data_json + ''';
const PASSWORD = "''' + PASSWORD + '''";

function checkPwd() {
  const v = document.getElementById("pwd-input").value;
  if (v === PASSWORD) {
    document.getElementById("login-overlay").style.display = "none";
    initDashboard();
  } else {
    document.getElementById("login-error").style.display = "block";
  }
}

function switchPage(page) {
  document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.getElementById("page-" + page).classList.add("active");
  event.target.classList.add("active");
  if (page === "team" && !window._teamInited) { initTeamPage(); window._teamInited = true; }
  if (page === "personal" && !window._personalInited) { initPersonalPage(); window._personalInited = true; }
}

function pct(v, digits=1) { return (v * 100).toFixed(digits) + "%"; }
function pct2(v) { return (v * 100).toFixed(2) + "%"; }
function tooltipFmt(params) {
  if (!Array.isArray(params)) params = [params];
  let html = params[0].axisValueLabel || params[0].name || '';
  params.forEach(p => {
    const v = p.value;
    const display = (typeof v === 'number' && Math.abs(v) < 10) ? pct2(v) : (typeof v === 'number' ? num(v) : v);
    html += '<br/>' + p.marker + ' ' + p.seriesName + ': ' + display;
  });
  return html;
}
function num(v) { return v.toLocaleString(); }
function changeClass(v) { return v >= 0 ? "text-green" : "text-red"; }
function changeStr(v, isRate=false) {
  if (isRate) return (v >= 0 ? "+" : "") + (v * 100).toFixed(2) + "%";
  return (v >= 0 ? "+" : "") + (v * 100).toFixed(2) + "%";
}
function kpiColor(rate) {
  if (rate >= 100) return "#059669";
  if (rate >= 80) return "#2563EB";
  return "#EA580C";
}

function initDashboard() {
  renderOverview();
}

function renderOverview() {
  const c = DATA.cards;
  let html = `
  <div class="cards-row">
    <div class="card"><div class="card-label">报名人数</div><div class="card-value" style="color:#059669;font-weight:700">${num(c.reg_count)}</div></div>
    <div class="card"><div class="card-label">目标人数</div><div class="card-value">${num(c.target)}</div></div>
    <div class="card"><div class="card-label">报名完成率</div><div class="card-value" style="color:${kpiColor(c.completion_rate)}">${c.completion_rate}%</div></div>
    <div class="card"><div class="card-label">课时达成</div><div class="card-value">${num(c.hours_achieved)}</div></div>
    <div class="card"><div class="card-label">课时目标</div><div class="card-value">${num(c.hours_target)}</div></div>
    <div class="card"><div class="card-label">报名率</div><div class="card-value">${c.reg_rate}%</div></div>
  </div>

  <div class="section">
    <div class="section-title">本期与往期对比</div>
    <div class="section-conclusion">${DATA.conclusions.overview}</div>
    <div id="chart-history" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>时间</th><th>报名人数</th><th>报考率</th><th>预扣课时</th><th>人数目标</th><th>目标完成率</th></tr></thead><tbody>`;

  DATA.history.forEach(h => {
    const compColor = h.completion_rate >= 1 ? 'color:#059669;font-weight:700' : (h.completion_rate >= 0.8 ? 'color:#2563EB;font-weight:600' : 'color:#EA580C;font-weight:600');
    html += `<tr><td>${h.time}</td><td>${num(h.reg_count)}</td><td>${pct2(h.reg_rate)}</td><td>${num(h.hours)}</td><td>${num(h.target)}</td><td style="${compColor}">${pct(h.completion_rate)}</td></tr>`;
  });
  // 环比同比
  if (DATA.history_changes['环比']) {
    const hb = DATA.history_changes['环比'];
    html += `<tr><td>环比</td><td class="${changeClass(hb.reg_count)}">${changeStr(hb.reg_count)}</td><td class="${changeClass(hb.reg_rate)}">${changeStr(hb.reg_rate,true)}</td><td class="${changeClass(hb.hours)}">${changeStr(hb.hours)}</td><td>-</td><td class="${changeClass(hb.completion_rate)}">${changeStr(hb.completion_rate)}</td></tr>`;
  }
  if (DATA.history_changes['同比']) {
    const tb = DATA.history_changes['同比'];
    html += `<tr><td>同比</td><td class="${changeClass(tb.reg_count)}">${changeStr(tb.reg_count)}</td><td class="${changeClass(tb.reg_rate)}">${changeStr(tb.reg_rate,true)}</td><td class="${changeClass(tb.hours)}">${changeStr(tb.hours)}</td><td>-</td><td class="${changeClass(tb.completion_rate)}">${changeStr(tb.completion_rate)}</td></tr>`;
  }
  html += `</tbody></table></div>

  <div class="section">
    <div class="section-title">每日报名趋势</div>
    <div class="section-conclusion">${DATA.conclusions.daily}</div>
    <div id="chart-daily" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>日期</th><th>星期</th><th>目标人数</th><th>实际人数</th><th>目标流速</th><th>实际流速</th></tr></thead><tbody>`;
  DATA.mtd.forEach(m => {
    const tFlow = (m.target_flow * 100).toFixed(1) + '%';
    const aFlow = m.actual_daily > 0 ? (m.actual_flow * 100).toFixed(1) + '%' : '-';
    const flowColor = m.actual_flow >= m.target_flow ? 'color:#059669' : (m.actual_daily > 0 ? 'color:#DC2626' : '');
    html += `<tr><td>${m.short_label}</td><td>${m.weekday}</td><td>${m.daily_target}</td><td>${m.actual_daily || '-'}</td><td>${tFlow}</td><td style="${flowColor};font-weight:600">${aFlow}</td></tr>`;
  });
  html += `</tbody></table></div>

  <div class="section">
    <div class="section-title">学区完成情况</div>
    <div class="section-conclusion">${DATA.conclusions.district}</div>
    <table class="data-table"><thead><tr><th>学区</th><th>目标</th><th>实际报名</th><th>完成率</th><th>报名率</th><th>中国大陆</th><th>中国港澳台</th><th>海外</th></tr></thead><tbody>`;
  const districtRows = DATA.district_stats.filter(d => d.district !== '无归属');
  const compRatesArr = districtRows.map(d => d.completion_rate);
  const regRatesArr = districtRows.map(d => d.reg_rate);
  const maxComp = Math.max(...compRatesArr), minComp = Math.min(...compRatesArr);
  const maxRegR = Math.max(...regRatesArr), minRegR = Math.min(...regRatesArr);
  districtRows.forEach(d => {
    const compColor = d.completion_rate === maxComp ? 'color:#059669;font-weight:700' : (d.completion_rate === minComp ? 'color:#DC2626;font-weight:700' : (d.completion_rate>=1?'color:#059669;font-weight:600':'color:#D97706;font-weight:600'));
    const regRColor = d.reg_rate === maxRegR ? 'color:#059669;font-weight:700' : (d.reg_rate === minRegR ? 'color:#DC2626;font-weight:700' : '');
    const barW = Math.min(d.completion_rate * 100, 100);
    const barColor = d.completion_rate >= 1 ? 'rgba(5,150,105,0.12)' : 'rgba(37,99,235,0.10)';
    html += `<tr><td>${d.district}</td><td>${d.target}</td><td>${d.reg_count}</td><td style="${compColor};position:relative"><span style="position:absolute;left:0;top:0;bottom:0;width:${barW}%;background:${barColor};border-radius:3px"></span><span style="position:relative">${pct(d.completion_rate)}</span></td><td style="${regRColor}">${pct(d.reg_rate)}</td><td>${d.domestic_reg}</td><td>${d.hkmotw_reg}</td><td>${d.overseas_reg}</td></tr>`;
  });
  // 学区汇总行
  const ds = DATA.district_summary;
  html += `<tr style="font-weight:700;background:#F8FAFC;border-top:2px solid #2563EB"><td>学区汇总</td><td>${ds.target}</td><td>${ds.reg_count}</td><td class="${ds.completion_rate>=1?'text-green':'text-yellow'}">${pct(ds.completion_rate)}</td><td>${pct(ds.reg_rate)}</td><td>${ds.domestic_reg}</td><td>${ds.hkmotw_reg}</td><td>${ds.overseas_reg}</td></tr>`;
  // 无归属行
  const ua = DATA.district_stats.find(d => d.district === '无归属');
  if (ua) {
    html += `<tr style="color:#64748B;border-top:1px dashed #CBD5E1"><td>无归属</td><td>-</td><td>${ua.reg_count}</td><td>-</td><td>-</td><td>${ua.domestic_reg}</td><td>${ua.hkmotw_reg}</td><td>${ua.overseas_reg}</td></tr>`;
  }
  html += `</tbody></table></div>

  <div class="grid-2">
    <div class="section">
      <div class="section-title">中国大陆/中国港澳台/海外对比</div>
      <div class="section-conclusion">${DATA.conclusions.region}</div>
      <div id="chart-region" class="chart-container-sm"></div>
    </div>
    <div class="section">
      <div class="section-title">中国大陆/中国港澳台/海外分布</div>
      <div id="chart-region-pie" class="chart-container-sm"></div>
    </div>
  </div>

  <div class="section">
    <div class="section-title">科目汇总</div>
    <div class="section-conclusion">${DATA.conclusions.subject}</div>
    <div id="chart-subject" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>科目</th><th>报名人数</th><th>课时</th><th>人数占比</th><th>本期首报</th><th>本期首报占比</th><th>本期复报</th><th>上期首报</th><th>首报人数环比</th></tr></thead><tbody>`;
  const subjRatios = DATA.subject_stats.map(s => s['人数占比']);
  const maxSubjR = Math.max(...subjRatios), minSubjR = Math.min(...subjRatios);
  DATA.subject_stats.forEach(s => {
    const hist = DATA.subject_history[s['考试科目名称']] || {};
    const fs = DATA.subject_first_stats[s['考试科目名称']] || {};
    const firstCount = fs.first || 0;
    const repeatCount = fs.repeat || 0;
    const firstRatio = fs.first_ratio || 0;
    const lastFirst = hist.last_first_count || 0;
    const firstChange = lastFirst > 0 ? ((firstCount - lastFirst) / lastFirst * 100).toFixed(1) + '%' : '-';
    const ratioStyle = s['人数占比'] === maxSubjR ? 'color:#059669;font-weight:700' : (s['人数占比'] === minSubjR ? 'color:#DC2626;font-weight:700' : '');
    html += `<tr><td>${s['考试科目名称']}</td><td>${s['报名人数']}</td><td>${s['课时数']}</td><td style="${ratioStyle}">${pct(s['人数占比'])}</td><td>${firstCount}</td><td>${pct(firstRatio)}</td><td>${repeatCount}</td><td>${lastFirst||'-'}</td><td class="${firstCount>=lastFirst?'text-green':'text-red'}">${firstChange}</td></tr>`;
  });
  html += `</tbody></table></div>

  <div class="section">
    <div class="section-title">首报与复报对比</div>
    <div class="section-conclusion">${DATA.conclusions.repeat}</div>
    <div id="chart-repeat" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>季度</th><th>首报人数</th><th>首报占比</th><th>复报人数</th><th>复报占比</th></tr></thead><tbody>`;
  const repeatRatios = DATA.repeat_history.map(r => r.first_ratio);
  const maxRepR = Math.max(...repeatRatios), minRepR = Math.min(...repeatRatios);
  DATA.repeat_history.forEach(r => {
    const frStyle = r.first_ratio === maxRepR ? 'color:#059669;font-weight:700' : (r.first_ratio === minRepR ? 'color:#DC2626;font-weight:700' : '');
    html += `<tr><td>${r.time}</td><td>${r.first_count}</td><td style="${frStyle}">${pct(r.first_ratio)}</td><td>${r.repeat_count}</td><td>${pct(r.repeat_ratio)}</td></tr>`;
  });
  html += `</tbody></table></div>

  <div class="section">
    <div class="section-title">学员分层对比</div>
    <div class="section-conclusion">${DATA.conclusions.tier}</div>
    <div id="chart-tier" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>分层</th><th>池内人数</th><th>报名人数</th><th>本次报考率</th><th>上期报考率</th><th>环比</th></tr></thead><tbody>`;
  const tierRates = DATA.tier_stats.map(t => t.rate);
  const maxTierR = Math.max(...tierRates), minTierR = Math.min(...tierRates);
  DATA.tier_stats.forEach(t => {
    const rateStyle = t.rate === maxTierR ? 'color:#059669;font-weight:700' : (t.rate === minTierR ? 'color:#DC2626;font-weight:700' : '');
    const hasHistory = t.last_rate !== null && t.last_rate !== undefined && t.rate_change !== null && t.rate_change !== undefined;
    if (hasHistory) {
      const chg = t.rate_change;
      const arrow = chg >= 0 ? "▲" : "▼";
      const chgStr = arrow + " " + Math.abs(chg * 100).toFixed(2) + "%";
      html += `<tr><td>${t.tier}</td><td>${num(t.pool_count)}</td><td>${t.reg_count}</td><td style="${rateStyle}">${pct(t.rate,2)}</td><td>${pct(t.last_rate,2)}</td><td class="${changeClass(chg)}">${chgStr}</td></tr>`;
    } else {
      html += `<tr><td>${t.tier}</td><td>${num(t.pool_count)}</td><td>${t.reg_count}</td><td style="${rateStyle}">${pct(t.rate,2)}</td><td>—</td><td style="color:#94A3B8">—</td></tr>`;
    }
  });
  html += `</tbody></table></div>

  <div class="grid-2">
    <div class="section">
      <div class="section-title">阶段分布</div>
      <div class="section-conclusion">${DATA.conclusions.stage}</div>
      <div id="chart-stage" class="chart-container-sm"></div>
    </div>
    <div class="section">
      <div class="section-title">满级学员报名情况</div>
      <div class="section-conclusion">${DATA.conclusions.mj}</div>
      <table class="data-table"><thead><tr><th>学区</th><th>满级人数</th><th>已报名</th><th>报名率</th></tr></thead><tbody>`;
  const mjRates = DATA.mj_by_district.filter(m => m.total > 0).map(m => m.rate);
  const maxMjR = Math.max(...mjRates), minMjR = Math.min(...mjRates);
  DATA.mj_by_district.forEach(m => {
    const rateStyle = m.total > 0 ? (m.rate === maxMjR ? 'color:#059669;font-weight:700' : (m.rate === minMjR ? 'color:#DC2626;font-weight:700' : '')) : '';
    html += `<tr><td>${m.district}</td><td>${m.total}</td><td>${m.reg}</td><td style="${rateStyle}">${pct(m.rate,2)}</td></tr>`;
  });
  html += `</tbody></table></div></div>`;

  // 荣誉榜
  const rankIcons = ["🥇","🥈","🥉"];
  const rankBg = ["linear-gradient(90deg, #FEF9C3 0%, #FFFBEB 100%)","linear-gradient(90deg, #F1F5F9 0%, #F8FAFC 100%)","linear-gradient(90deg, #FFEDD5 0%, #FFF7ED 100%)"];
  html += `<div class="section">
    <div class="section-title">🏆 荣誉榜</div>
    <div class="grid-2">
      <div>
        <div style="font-size:14px;font-weight:700;color:#B45309;margin-bottom:10px">▎团队 TOP4（完成率）</div>
        <table class="data-table"><thead><tr><th>排名</th><th>团队名</th><th>完成率</th><th>学区</th></tr></thead><tbody>`;
  DATA.honor.teams.forEach(t => {
    const icon = t.rank <= 3 ? rankIcons[t.rank-1] : t.rank;
    const bg = t.rank <= 3 ? `background:${rankBg[t.rank-1]}` : '';
    html += `<tr style="${bg}"><td>${icon}</td><td style="font-weight:${t.rank<=3?'700':'400'};color:#0F172A">${t.name}</td><td class="text-green">${pct(t.completion_rate)}</td><td>${t.district}</td></tr>`;
  });
  html += `</tbody></table></div>
      <div>
        <div style="font-size:14px;font-weight:700;color:#B45309;margin-bottom:10px">▎小组 TOP5（完成率）</div>
        <table class="data-table"><thead><tr><th>排名</th><th>小组名</th><th>完成率</th><th>学区</th></tr></thead><tbody>`;
  DATA.honor.groups.forEach(g => {
    const icon = g.rank <= 3 ? rankIcons[g.rank-1] : g.rank;
    const bg = g.rank <= 3 ? `background:${rankBg[g.rank-1]}` : '';
    html += `<tr style="${bg}"><td>${icon}</td><td style="font-weight:${g.rank<=3?'700':'400'};color:#0F172A">${g.name}</td><td class="text-green">${pct(g.completion_rate)}</td><td>${g.district}</td></tr>`;
  });
  html += `</tbody></table></div>
    </div>
  </div>`;

  // 科目级别报名排名（放最底部）
  html += `<div class="section">
    <div class="section-title">科目级别报名排名（本期 vs 上期）</div>
    <div id="chart-subject-level-rank" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>排名</th><th>科目级别</th><th>本期人数</th><th>上期人数</th><th>变化</th><th>变化率</th></tr></thead><tbody>`;
  DATA.subject_level_ranking.forEach((s, i) => {
    const chgClass = s.change >= 0 ? 'text-green' : 'text-red';
    const chgStr = s.change >= 0 ? '+' + s.change : s.change;
    const rateStr = s.last > 0 ? (s.change_rate >= 0 ? '+' : '') + (s.change_rate * 100).toFixed(1) + '%' : '-';
    const boldStyle = Math.abs(s.change_rate) >= 0.2 ? 'font-weight:700' : '';
    html += `<tr style="${boldStyle}"><td>${i+1}</td><td>${s.name}</td><td>${s.current}</td><td>${s.last||'-'}</td><td class="${chgClass}">${chgStr}</td><td class="${chgClass}">${rateStr}</td></tr>`;
  });
  html += `</tbody></table></div>`;

  document.getElementById("page-overview").innerHTML = html;

  // 渲染图表
  setTimeout(() => {
    renderHistoryChart();
    renderDailyChart();
    renderRegionChart();
    renderRegionPie();
    renderSubjectChart();
    renderSubjectLevelRankChart();
    renderRepeatChart();
    renderTierChart();
    renderStageChart();
  }, 100);
}

function renderHistoryChart() {
  const chart = echarts.init(document.getElementById("chart-history"));
  const times = DATA.history.map(h => h.time);
  const regCounts = DATA.history.map(h => h.reg_count);
  const rates = DATA.history.map(h => h.reg_rate);
  const compRates = DATA.history.map(h => h.completion_rate);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["报名人数", "报考率", "目标完成率"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 60, bottom: 30, top: 45 },
    xAxis: { type: "category", data: times, axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: [
      { type: "value", name: "人数", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
      { type: "value", name: "比率", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { show: false } }
    ],
    series: [
      { name: "报名人数", type: "bar", data: regCounts, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 26, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 10 } },
      { name: "报考率", type: "line", yAxisIndex: 1, data: rates, symbol: "circle", symbolSize: 6, itemStyle: { color: "#059669" }, lineStyle: { width: 2.5 }, smooth: true, label: { show: true, position: "right", offset: [4, -4], color: "#059669", backgroundColor: "#ECFDF5", borderColor: "#A7F3D0", borderWidth: 1, borderRadius: 4, padding: [2, 5], fontWeight: 700, fontSize: 10, formatter: p => pct(p.value) }, labelLayout: { hideOverlap: false } },
      { name: "目标完成率", type: "line", yAxisIndex: 1, data: compRates, itemStyle: { color: "#D97706" }, lineStyle: { width: 2.5 }, smooth: true }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderDailyChart() {
  const chart = echarts.init(document.getElementById("chart-daily"));
  const labels = DATA.mtd.map(m => m.label);
  const cumTargets = DATA.mtd.map(m => m.cum_target);
  // 匹配实际数据到MTD日期
  const actualCum = DATA.mtd.map(m => {
    const idx = DATA.daily.dates.findIndex(d => d === m.date.split(" ")[0]);
    if (idx >= 0) return DATA.daily.cumulative[idx];
    // 尝试匹配
    const dateStr = m.date;
    for (let i = 0; i < DATA.daily.dates.length; i++) {
      if (DATA.daily.dates[i] === dateStr) return DATA.daily.cumulative[i];
    }
    return null;
  });
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["实际累计", "MTD目标"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 30, bottom: 50, top: 40 },
    xAxis: { type: "category", data: labels, axisLabel: { color: "#475569", interval: 0, fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "实际累计", type: "line", data: actualCum, itemStyle: { color: "#2563EB" }, lineStyle: { width: 2.5 }, smooth: true, label: { show: true, color: "#1D4ED8", fontWeight: 600, fontSize: 10 }, areaStyle: { color: { type: "linear", x: 0, y: 0, x2: 0, y2: 1, colorStops: [{offset:0,color:"rgba(37,99,235,0.18)"},{offset:1,color:"rgba(37,99,235,0.01)"}] } } },
      { name: "MTD目标", type: "line", data: cumTargets, itemStyle: { color: "#D97706" }, lineStyle: { type: "dashed", width: 2 }, smooth: true }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderRegionChart() {
  const chart = echarts.init(document.getElementById("chart-region"));
  const r = DATA.region;
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["报名率", "上期报考率"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: ["中国大陆", "中国港澳台", "海外"], axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "报名率", type: "bar", data: [r.domestic.rate, r.hkmotw.rate, r.overseas.rate], itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 40, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, formatter: p => (p.value*100).toFixed(2)+"%" } },
      { name: "上期报考率", type: "bar", data: [r.domestic.last_rate, r.hkmotw.last_rate, r.overseas.last_rate], itemStyle: { color: "#94A3B8", borderRadius: [4, 4, 0, 0] }, barWidth: 40, label: { show: true, position: "top", color: "#475569", formatter: p => (p.value*100).toFixed(2)+"%" } }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderRegionPie() {
  const chart = echarts.init(document.getElementById("chart-region-pie"));
  const r = DATA.region;
  const total = r.domestic.reg + r.hkmotw.reg + r.overseas.reg;
  chart.setOption({
    tooltip: { trigger: "item" },
    series: [{
      type: "pie", radius: ["40%", "70%"], center: ["50%", "55%"],
      data: [
        { value: r.domestic.reg, name: "中国大陆 " + r.domestic.reg + "人", itemStyle: { color: "#2563EB" } },
        { value: r.hkmotw.reg, name: "中国港澳台 " + r.hkmotw.reg + "人", itemStyle: { color: "#7C3AED" } },
        { value: r.overseas.reg, name: "海外 " + r.overseas.reg + "人", itemStyle: { color: "#0891B2" } }
      ],
      label: { color: "#0F172A", fontWeight: 600, fontSize: 13, formatter: p => p.name + "\\n" + (p.value/total*100).toFixed(1) + "%" }
    }]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderSubjectChart() {
  const chart = echarts.init(document.getElementById("chart-subject"));
  const names = DATA.subject_stats.map(s => s['考试科目名称']);
  const firstCounts = names.map(n => (DATA.subject_first_stats[n] || {}).first || 0);
  const repeatCounts = names.map(n => (DATA.subject_first_stats[n] || {}).repeat || 0);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["首报", "复报"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: names, axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "首报", type: "bar", stack: "total", data: firstCounts, itemStyle: { color: "#0891B2" }, barWidth: 40, label: { show: true, position: "inside", color: "#FFFFFF", fontWeight: 600, fontSize: 11 } },
      { name: "复报", type: "bar", stack: "total", data: repeatCounts, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 40, label: { show: true, position: "inside", color: "#FFFFFF", fontWeight: 600, fontSize: 11 } }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderSubjectLevelRankChart() {
  const el = document.getElementById("chart-subject-level-rank");
  if (!el) return;
  const chart = echarts.init(el);
  const top10 = DATA.subject_level_ranking.slice(0, 10);
  const names = top10.map(s => s.name);
  const currents = top10.map(s => s.current);
  const lasts = top10.map(s => s.last);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本期", "上期"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 80, right: 30, bottom: 60, top: 40 },
    xAxis: { type: "category", data: names, axisLabel: { color: "#475569", fontSize: 10, rotate: 45 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本期", type: "bar", data: currents, itemStyle: { color: "#2563EB", borderRadius: [3, 3, 0, 0] }, barWidth: 14, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 10 } },
      { name: "上期", type: "bar", data: lasts, itemStyle: { color: "#94A3B8", borderRadius: [3, 3, 0, 0] }, barWidth: 14, label: { show: true, position: "top", color: "#475569", fontSize: 10 } }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderRepeatChart() {
  const chart = echarts.init(document.getElementById("chart-repeat"));
  const times = DATA.repeat_history.map(r => r.time);
  const firstCounts = DATA.repeat_history.map(r => r.first_count);
  const repeatCounts = DATA.repeat_history.map(r => r.repeat_count);
  const firstRatios = DATA.repeat_history.map(r => r.first_ratio);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["首报", "复报", "首报占比"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 60, bottom: 30, top: 40 },
    xAxis: { type: "category", data: times, axisLabel: { color: "#475569", fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: [
      { type: "value", name: "人数", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
      { type: "value", name: "占比", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { show: false } }
    ],
    series: [
      { name: "首报", type: "bar", stack: "total", data: firstCounts, itemStyle: { color: "#0891B2" }, label: { show: true, position: "inside", color: "#FFFFFF", fontWeight: 600, fontSize: 10 } },
      { name: "复报", type: "bar", stack: "total", data: repeatCounts, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, label: { show: true, position: "inside", color: "#FFFFFF", fontWeight: 600, fontSize: 10 } },
      { name: "首报占比", type: "line", yAxisIndex: 1, data: firstRatios, itemStyle: { color: "#D97706" }, lineStyle: { width: 2.5 }, smooth: true }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderTierChart() {
  const chart = echarts.init(document.getElementById("chart-tier"));
  const tiers = DATA.tier_stats.map(t => t.tier);
  const pools = DATA.tier_stats.map(t => t.pool_count);
  const regs = DATA.tier_stats.map(t => t.reg_count);
  const rates = DATA.tier_stats.map(t => t.rate);
  const lastRates = DATA.tier_stats.map(t => t.last_rate);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["池内人数", "报名人数", "本期报考率", "上期报考率"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 60, right: 60, bottom: 30, top: 40 },
    xAxis: { type: "category", data: tiers, axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: [
      { type: "value", name: "人数", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
      { type: "value", name: "报考率", nameTextStyle: { color: "#475569" }, axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { show: false } }
    ],
    series: [
      { name: "池内人数", type: "bar", data: pools, itemStyle: { color: "#CBD5E1", borderRadius: [4, 4, 0, 0] }, barWidth: 25 },
      { name: "报名人数", type: "bar", data: regs, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 25, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 10 } },
      { name: "本期报考率", type: "line", yAxisIndex: 1, data: rates, itemStyle: { color: "#059669" }, lineStyle: { width: 2.5 }, smooth: true, label: { show: true, color: "#059669", fontWeight: 600, fontSize: 10, formatter: p => pct(p.value) } },
      { name: "上期报考率", type: "line", yAxisIndex: 1, data: lastRates, itemStyle: { color: "#64748B" }, lineStyle: { type: "dashed", width: 2 }, smooth: true }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

function renderStageChart() {
  const chart = echarts.init(document.getElementById("chart-stage"));
  const stages = DATA.stage_stats.map(s => s.stage);
  const regs = DATA.stage_stats.map(s => s.reg_count);
  const rates = DATA.stage_stats.map(s => s.rate);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["报名人数", "报名率"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 50, bottom: 30, top: 40 },
    xAxis: { type: "category", data: stages, axisLabel: { color: "#475569", fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: [
      { type: "value", axisLabel: { color: "#475569" }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
      { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { show: false } }
    ],
    series: [
      { name: "报名人数", type: "bar", data: regs, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 20, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 9 } },
      { name: "报名率", type: "line", yAxisIndex: 1, data: rates, itemStyle: { color: "#059669" }, lineStyle: { width: 2.5 }, smooth: true, label: { show: true, color: "#059669", fontWeight: 600, fontSize: 9, formatter: p => pct(p.value) } }
    ]
  });
  window.addEventListener("resize", () => chart.resize());
}

// ====== 团队页面 ======
function initTeamPage() {
  const districts = ["教学一区", "教学二区", "教学三区", "C区", "DEMO特战队"];
  let html = `<div class="district-tabs">`;
  districts.forEach((d, i) => {
    html += `<div class="district-tab ${i===0?'active':''}" onclick="switchDistrict('${d}', this)">${d}</div>`;
  });
  html += `</div><div id="district-content"></div>`;
  document.getElementById("page-team").innerHTML = html;
  renderDistrict(districts[0]);
}

function switchDistrict(name, el) {
  document.querySelectorAll(".district-tab").forEach(t => t.classList.remove("active"));
  el.classList.add("active");
  renderDistrict(name);
}

function renderDistrict(name) {
  const d = DATA.team_page[name];
  if (!d) { document.getElementById("district-content").innerHTML = "<p>暂无数据</p>"; return; }
  const bm = DATA.benchmarks;
  const kpi = d.kpi || {};
  let html = `
  <div class="section"><div class="section-conclusion">${d.conclusion}</div></div>

  <div class="district-kpi">
    <div class="district-kpi-card"><div class="district-kpi-label">报名人数</div><div class="district-kpi-value" style="color:#059669;font-weight:700">${num(kpi.reg_count)}</div></div>
    <div class="district-kpi-card"><div class="district-kpi-label">目标人数</div><div class="district-kpi-value" style="color:#0F172A">${num(kpi.target)}</div></div>
    <div class="district-kpi-card"><div class="district-kpi-label">完成率</div><div class="district-kpi-value" style="color:${kpiColor(kpi.completion_rate)}">${kpi.completion_rate}%</div><div class="district-kpi-rank">排名 ${kpi.rank}/${kpi.total_districts}</div></div>
    <div class="district-kpi-card"><div class="district-kpi-label">报名率</div><div class="district-kpi-value">${kpi.reg_rate}%</div><div class="district-kpi-rank">排名 ${kpi.rank}/${kpi.total_districts}</div></div>
    <div class="district-kpi-card"><div class="district-kpi-label">池内人数</div><div class="district-kpi-value" style="color:#0F172A">${num(kpi.pool_count)}</div></div>
  </div>

  <div class="grid-2">
    <div class="section">
      <div class="section-title">团队报名排名（按完成率）</div>
      <div id="chart-team-rank" class="chart-container-sm"></div>
    </div>
    <div class="section">
      <div class="section-title">小组报名排名（按完成率）</div>
      <div id="chart-group-rank" class="chart-container-sm"></div>
    </div>
  </div>`;

  // TOP3
  if (d.teachers_top3.length > 0) {
    html += `<div class="top3">`;
    const ranks = ["🥇", "🥈", "🥉"];
    d.teachers_top3.forEach((t, i) => {
      html += `<div class="top3-item"><span class="top3-rank">${ranks[i]}</span><span class="top3-name">${t.name}</span><div class="top3-stats">${t.reg_count}人 / 目标${t.target}人 / 完成率${pct(t.completion_rate)}</div></div>`;
    });
    html += `</div>`;
  }

  html += `
  <div class="grid-2">
    <div class="section">
      <div class="section-title">阶段报名率（学区 vs 大盘）</div>
      <div id="chart-district-stage" class="chart-container-sm"></div>
    </div>
    <div class="section">
      <div class="section-title">中国大陆/中国港澳台/海外报名率（学区 vs 大盘）</div>
      <div id="chart-district-region" class="chart-container-sm"></div>
    </div>
  </div>
  <div class="grid-2">
    <div class="section">
      <div class="section-title">学员分层报名率（学区 vs 大盘）</div>
      <div id="chart-district-tier" class="chart-container-sm"></div>
    </div>
    <div class="section">
      <div class="section-title">满级学员报名率（学区 vs 大盘）</div>
      <div id="chart-district-mj" class="chart-container-sm"></div>
    </div>
  </div>
  <div class="section">
    <div class="section-title">首报/复报报名率（学区 vs 大盘）</div>
    <div id="chart-district-fr" class="chart-container-sm"></div>
  </div>
  <div class="section">
    <div class="section-title">科目级别报名占比（学区 vs 大盘）</div>
    <div id="chart-district-subject" class="chart-container"></div>
    <table class="data-table"><thead><tr><th>排名</th><th>科目级别</th><th>报名人数</th><th>学区占比</th><th>大盘占比</th><th>差异</th></tr></thead><tbody>`;
  const slSorted = (d.subject_level || []).sort((a,b) => b['报名人数'] - a['报名人数']);
  const slRatios = slSorted.map(s => s['占比']);
  const maxSlR = slRatios.length > 0 ? Math.max(...slRatios) : 0;
  const minSlR = slRatios.length > 0 ? Math.min(...slRatios) : 0;
  slSorted.forEach((s, idx) => {
    const bmVal = bm.subject_level[s['科目级别']] || 0;
    const diff = s['占比'] - bmVal;
    const diffClass = diff >= 0 ? 'text-green' : 'text-red';
    const ratioStyle = s['占比'] === maxSlR ? 'color:#059669;font-weight:700' : (s['占比'] === minSlR ? 'color:#DC2626;font-weight:700' : '');
    html += `<tr><td>${idx+1}</td><td>${s['科目级别']}</td><td>${s['报名人数']}</td><td style="${ratioStyle}">${pct(s['占比'],2)}</td><td>${pct(bmVal,2)}</td><td style="color:${diff>=0?'#059669':'#DC2626'};font-weight:600">${(diff>=0?'+':'')+(diff*100).toFixed(2)}%</td></tr>`;
  });
  html += `</tbody></table></div>`;

  document.getElementById("district-content").innerHTML = html;

  setTimeout(() => {
    renderTeamRank(d);
    renderGroupRank(d);
    renderDistrictSubject(d, bm);
    renderDistrictStage(d, bm);
    renderDistrictRegion(d, bm);
    renderDistrictTier(d, bm);
    renderDistrictMJ(d, bm);
    renderDistrictFR(d, bm);
  }, 100);
}

function renderTeamRank(d) {
  const el = document.getElementById("chart-team-rank");
  if (!el || d.teams.length === 0) return;
  const chart = echarts.init(el);
  const names = d.teams.map(t => t.name).reverse();
  const rates = d.teams.map(t => t.completion_rate).reverse();
  chart.setOption({
    tooltip: { trigger: "axis", formatter: p => p[0].name + ": " + pct2(p[0].value) },
    grid: { left: 100, right: 60, bottom: 20, top: 10 },
    xAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    yAxis: { type: "category", data: names, axisLabel: { color: "#1E293B", fontWeight: 500, fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    series: [{ type: "bar", data: rates, itemStyle: { color: "#2563EB", borderRadius: [0, 4, 4, 0] }, barWidth: 18, label: { show: true, position: "right", color: "#0F172A", fontWeight: 600, fontSize: 10, formatter: p => pct(p.value) } }]
  });
}

function renderGroupRank(d) {
  const el = document.getElementById("chart-group-rank");
  if (!el || d.groups.length === 0) return;
  const chart = echarts.init(el);
  const top10 = d.groups.slice(0, 10);
  const names = top10.map(g => g.name).reverse();
  const rates = top10.map(g => g.completion_rate).reverse();
  chart.setOption({
    tooltip: { trigger: "axis", formatter: p => p[0].name + ": " + pct2(p[0].value) },
    grid: { left: 100, right: 60, bottom: 20, top: 10 },
    xAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    yAxis: { type: "category", data: names, axisLabel: { color: "#1E293B", fontWeight: 500, fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    series: [{ type: "bar", data: rates, itemStyle: { color: "#0891B2", borderRadius: [0, 4, 4, 0] }, barWidth: 18, label: { show: true, position: "right", color: "#0F172A", fontWeight: 600, fontSize: 10, formatter: p => pct(p.value) } }]
  });
}

function renderDistrictSubject(d, bm) {
  const el = document.getElementById("chart-district-subject");
  if (!el || !d.subject_level || d.subject_level.length === 0) return;
  const chart = echarts.init(el);
  const top10 = d.subject_level.sort((a,b) => b['占比'] - a['占比']).slice(0, 10);
  const names = top10.map(s => s['科目级别']);
  const districtRatios = top10.map(s => s['占比']);
  const benchRatios = names.map(n => bm.subject_level[n] || 0);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 60, top: 40 },
    xAxis: { type: "category", data: names, axisLabel: { color: "#475569", fontSize: 10, rotate: 45 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: districtRatios, itemStyle: { color: "#2563EB", borderRadius: [3, 3, 0, 0] }, barWidth: 12, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 9, formatter: p => pct(p.value) } },
      { name: "大盘", type: "bar", data: benchRatios, itemStyle: { color: "#94A3B8", borderRadius: [3, 3, 0, 0] }, barWidth: 12 }
    ]
  });
}

function renderDistrictStage(d, bm) {
  const el = document.getElementById("chart-district-stage");
  if (!el) return;
  const chart = echarts.init(el);
  const validStages = d.stages.filter(s => s.pool_count > 0 || s.reg_count > 0);
  const stages = validStages.map(s => s.stage);
  const districtRates = validStages.map(s => s.rate);
  const benchRates = stages.map(s => bm.stages[s] || 0);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: stages, axisLabel: { color: "#475569", fontSize: 11 }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: districtRates, itemStyle: { color: "#2563EB", borderRadius: [3, 3, 0, 0] }, barWidth: 14, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 9, formatter: p => pct(p.value) } },
      { name: "大盘", type: "bar", data: benchRates, itemStyle: { color: "#94A3B8", borderRadius: [3, 3, 0, 0] }, barWidth: 14 }
    ]
  });
}

function renderDistrictRegion(d, bm) {
  const el = document.getElementById("chart-district-region");
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: ["中国大陆", "中国港澳台", "海外"], axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: [d.region.domestic.rate, d.region.hkmotw.rate, d.region.overseas.rate], itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 11, formatter: p => (p.value*100).toFixed(2)+"%" } },
      { name: "大盘", type: "bar", data: [bm.region.domestic, bm.region.hkmotw, bm.region.overseas], itemStyle: { color: "#94A3B8", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#475569", fontSize: 11, formatter: p => (p.value*100).toFixed(2)+"%" } }
    ]
  });
}

function renderDistrictTier(d, bm) {
  const el = document.getElementById("chart-district-tier");
  if (!el) return;
  const chart = echarts.init(el);
  const tiers = d.tiers.map(t => t.tier);
  const districtRates = d.tiers.map(t => t.rate);
  const benchRates = tiers.map(t => bm.tiers[t] || 0);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: tiers, axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: districtRates, itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 25, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 10, formatter: p => pct(p.value) } },
      { name: "大盘", type: "bar", data: benchRates, itemStyle: { color: "#94A3B8", borderRadius: [4, 4, 0, 0] }, barWidth: 25 }
    ]
  });
}

function renderDistrictMJ(d, bm) {
  const el = document.getElementById("chart-district-mj");
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: ["满级", "非满级"], axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: [d.mj.mj_rate, d.mj.non_mj_rate], itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 11, formatter: p => pct(p.value) } },
      { name: "大盘", type: "bar", data: [bm.mj_rate, bm.non_mj_rate], itemStyle: { color: "#94A3B8", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#475569", fontSize: 11, formatter: p => pct(p.value) } }
    ]
  });
}

function renderDistrictFR(d, bm) {
  const el = document.getElementById("chart-district-fr");
  if (!el) return;
  const chart = echarts.init(el);
  chart.setOption({
    tooltip: { trigger: "axis", formatter: tooltipFmt },
    legend: { data: ["本学区", "大盘"], textStyle: { color: "#475569", fontWeight: 500 }, top: 0 },
    grid: { left: 50, right: 30, bottom: 30, top: 40 },
    xAxis: { type: "category", data: ["首报", "复报"], axisLabel: { color: "#475569" }, axisLine: { lineStyle: { color: "#CBD5E1" } } },
    yAxis: { type: "value", axisLabel: { color: "#475569", formatter: v => pct(v) }, splitLine: { lineStyle: { color: "#E2E8F0", type: "dashed" } } },
    series: [
      { name: "本学区", type: "bar", data: [d.first_repeat.first.rate, d.first_repeat.repeat.rate], itemStyle: { color: "#2563EB", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#0F172A", fontWeight: 600, fontSize: 11, formatter: p => pct(p.value) } },
      { name: "大盘", type: "bar", data: [bm.first_rate, bm.repeat_rate], itemStyle: { color: "#94A3B8", borderRadius: [4, 4, 0, 0] }, barWidth: 30, label: { show: true, position: "top", color: "#475569", fontSize: 11, formatter: p => pct(p.value) } }
    ]
  });
}

// ====== 个人页面 ======
function initPersonalPage() {
  const allList = DATA.all_teachers || [];
  const totalTeachers = allList.length;
  const activeTeachers = allList.filter(t => t.reg_count > 0).length;
  const completedTeachers = allList.filter(t => t.completion_rate >= 1).length;
  const districts = ["全部", "教学一区", "教学二区", "教学三区", "C区", "DEMO特战队"];

  let html = `
  <div class="cards-row">
    <div class="card"><div class="card-label">老师总数</div><div class="card-value">${num(totalTeachers)}</div></div>
    <div class="card"><div class="card-label">已有报名老师数</div><div class="card-value" style="color:#2563EB">${num(activeTeachers)}</div></div>
    <div class="card"><div class="card-label">目标达成老师数</div><div class="card-value" style="color:#059669">${num(completedTeachers)}</div></div>
    <div class="card"><div class="card-label">零报名老师数</div><div class="card-value" style="color:#DC2626">${num(totalTeachers - activeTeachers)}</div></div>
  </div>
  <div class="section">
    <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:14px">
      <div class="section-title" style="margin-bottom:0">👤 全部老师个人排名（按加权报名人数）</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <div class="district-tabs" style="margin-bottom:0" id="personal-district-tabs">`;
  districts.forEach((d, i) => {
    html += `<div class="district-tab ${i===0?'active':''}" style="padding:6px 12px;font-size:12px" onclick="filterPersonalDistrict('${d}', this)">${d}</div>`;
  });
  html += `</div>
        <input type="text" id="personal-search" placeholder="搜索老师艺名..." oninput="renderPersonalTable()" style="padding:6px 12px;border:1px solid #CBD5E1;border-radius:6px;font-size:13px;outline:none;width:160px;color:#0F172A;background:#F8FAFC">
      </div>
    </div>
    <div class="section-conclusion">排名规则：按加权报名人数（首报×1.5 + 复报×1.0）降序排列，并列时按目标完成率降序排列。共 ${totalTeachers} 位老师。</div>
    <div id="personal-table-wrap"></div>
  </div>`;

  document.getElementById("page-personal").innerHTML = html;
  window._personalDistrict = "全部";
  renderPersonalTable();
}

function filterPersonalDistrict(d, el) {
  document.querySelectorAll("#personal-district-tabs .district-tab").forEach(t => t.classList.remove("active"));
  el.classList.add("active");
  window._personalDistrict = d;
  renderPersonalTable();
}

function renderPersonalTable() {
  const allList = DATA.all_teachers || [];
  const dist = window._personalDistrict || "全部";
  const kw = (document.getElementById("personal-search") ? document.getElementById("personal-search").value : "").trim().toLowerCase();

  const filtered = allList.filter(t => {
    if (dist !== "全部" && t.district !== dist) return false;
    if (kw && !t.name.toLowerCase().includes(kw)) return false;
    return true;
  });

  const rankIcons = ["🥇","🥈","🥉"];
  const rankBg = ["linear-gradient(90deg, #FEF9C3 0%, #FFFBEB 100%)","linear-gradient(90deg, #F1F5F9 0%, #F8FAFC 100%)","linear-gradient(90deg, #FFEDD5 0%, #FFF7ED 100%)"];

  let html = `<table class="data-table"><thead><tr><th>排名</th><th>艺名</th><th>目标学员数</th><th>报名人数</th><th>完成率</th><th>首报人数</th><th>加权报名人数</th><th>学区</th></tr></thead><tbody>`;
  if (filtered.length === 0) {
    html += `<tr><td colspan="8" style="text-align:center;color:#64748B;padding:24px">未找到匹配的老师数据</td></tr>`;
  } else {
    filtered.forEach(t => {
      const icon = t.rank <= 3 ? rankIcons[t.rank-1] : t.rank;
      const bg = t.rank <= 3 ? `background:${rankBg[t.rank-1]}` : '';
      const compClass = t.completion_rate >= 1 ? 'text-green' : (t.reg_count === 0 ? 'text-red' : '');
      html += `<tr style="${bg}"><td>${icon}</td><td style="font-weight:${t.rank<=3?'700':'500'};color:#0F172A">${t.name}</td><td>${t.target}</td><td>${t.reg_count}</td><td class="${compClass}">${pct(t.completion_rate)}</td><td>${t.first_count}</td><td style="font-weight:600;color:#2563EB">${t.weighted_count}</td><td>${t.district}</td></tr>`;
    });
  }
  html += `</tbody></table>`;
  document.getElementById("personal-table-wrap").innerHTML = html;
}
</script>
</body>
</html>'''

with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\\n生成完成！文件路径：{OUTPUT_FILE}")
print("请在浏览器中打开 dashboard.html 查看效果。")
