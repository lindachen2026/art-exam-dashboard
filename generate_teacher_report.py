# -*- coding: utf-8 -*-
"""
全部老师报名人数及目标完成率报表
统计口径与看板"荣誉榜-个人排名"完全一致，区别：
  1) 列出全部老师（含本期零报名的老师）
  2) 完成率 = 报名人数 / 目标学员数（不加权）
  3) 排序：加权报名人数降序，并列按完成率降序
输出：老师报名完成率报表.xlsx
"""

import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "数据源")
FILE_STUDENT_INFO = os.path.join(DATA_DIR, "考级学员信息表.xlsx")
FILE_TARGET = os.path.join(DATA_DIR, "目标明细.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "老师报名完成率报表.xlsx")

# ===== 1. 读取本期报名学员ID =====
df_reg = pd.read_excel(FILE_STUDENT_INFO, sheet_name="考级学员信息表")
if len(df_reg.columns) == 8:
    df_reg.columns = ['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']
else:
    df_reg = df_reg[['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']]
df_reg = df_reg[df_reg['用户id'].notna()]
df_reg['用户id'] = df_reg['用户id'].astype(int)
registered_ids = set(df_reg['用户id'].unique())

# ===== 2. 读取跟进学员名单 =====
df_follow = pd.read_excel(FILE_TARGET, sheet_name="跟进学员名单（在班学员）")
df_follow.columns = [
    '老师ID', '老师工号', '老师艺名', '学区', '团队', '小组', '班级ID', '班级名称',
    '班级课表', '阶段', '星期', '上课时间', '学员ID', '学员昵称', '年龄', '地区',
    '有效续费次数', '是否师生在群', '是否有有效群', '班主任ID', '班主任姓名', '班主任艺名',
    '班主任大区', '班主任团队', '大V渠道', '累计报考次数', '26Q1报考科目', '26Q1成绩',
    '25Q4报考科目', '25Q4成绩', '25Q3报考科目', '25Q3成绩', '25Q1报考科目', '25Q1成绩',
    '24Q4报考科目', '24Q4成绩', '24Q3报考科目', '24Q3成绩', '24Q1报考科目', '24Q1成绩',
    '儿童画', '儿童画成绩', '漫画', '漫画成绩', '素描', '素描成绩', '国画', '国画成绩',
    '速写', '速写成绩', '推荐儿童画报考级别', '推荐漫画报考级别', '推荐素描报考级别',
    '推荐国画报考级别', '推荐速写报考级别', '漫画-是否跳级', '素描-是否跳级', '分层',
    '剩余课时', '最新国画专题课课表名称', '最新国画专题课上课日期', '备注'
]
df_follow['学员ID'] = df_follow['学员ID'].astype(int)
df_follow['累计报考次数'] = pd.to_numeric(df_follow['累计报考次数'], errors='coerce')
df_follow['是否本期报名'] = df_follow['学员ID'].isin(registered_ids).astype(int)

# ===== 3. 读取老师目标（全部老师基准） =====
df_teacher = pd.read_excel(FILE_TARGET, sheet_name="老师")
df_teacher = df_teacher.iloc[:, :9]
df_teacher.columns = ['老师ID', '老师艺名', '学区', '团队', '小组', '目标学员数_精准', '目标学员数_向上取整', '学员数', '报名率']
df_teacher = df_teacher[['老师艺名', '学区', '团队', '小组', '目标学员数_向上取整']]
df_teacher.rename(columns={'目标学员数_向上取整': '目标学员数'}, inplace=True)

# ===== 4. 报名聚合（仅本期报名，口径同荣誉榜） =====
reg = df_follow[df_follow['是否本期报名'] == 1].copy()
reg = reg[reg['老师艺名'].notna() & (reg['老师艺名'] != '-')]
# 首报：累计报考次数为空或0；复报：>=1
reg['报名类型'] = reg['累计报考次数'].apply(lambda x: '首报' if (pd.isna(x) or x == 0) else '复报')
agg = reg.groupby('老师艺名').agg(
    报名人数=('学员ID', 'count'),
    首报人数=('报名类型', lambda x: (x == '首报').sum())
).reset_index()

# ===== 5. 以全部老师为基准左连接 =====
result = df_teacher.merge(agg, on='老师艺名', how='left')
result['报名人数'] = result['报名人数'].fillna(0).astype(int)
result['首报人数'] = result['首报人数'].fillna(0).astype(int)
result['复报人数'] = result['报名人数'] - result['首报人数']
result['加权报名人数'] = result['首报人数'] * 1.5 + result['复报人数'] * 1.0
result['目标学员数'] = result['目标学员数'].fillna(0).astype(int)
result['完成率'] = result.apply(
    lambda r: r['报名人数'] / r['目标学员数'] if r['目标学员数'] > 0 else 0, axis=1
)

# ===== 6. 排序（加权降序，并列按完成率降序）并编号 =====
result = result.sort_values(['加权报名人数', '完成率'], ascending=[False, False]).reset_index(drop=True)
result.insert(0, '排名', range(1, len(result) + 1))

# ===== 7. 整理输出列（与荣誉榜表头一致） =====
out = result[['排名', '老师艺名', '目标学员数', '报名人数', '完成率', '首报人数', '加权报名人数', '学区']].copy()
out.rename(columns={'老师艺名': '艺名'}, inplace=True)
out['完成率'] = (out['完成率'] * 100).round(1).astype(str) + '%'

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as w:
    out.to_excel(w, sheet_name='老师报名完成率', index=False)

print(f"老师总数: {len(out)}")
print(f"有报名的老师数: {(result['报名人数'] > 0).sum()}")
print(f"零报名老师数: {(result['报名人数'] == 0).sum()}")
print(f"报名人数合计: {result['报名人数'].sum()}")
print(f"输出文件: {OUTPUT_FILE}")
print("\n前5行预览:")
print(out.head(5).to_string(index=False))
