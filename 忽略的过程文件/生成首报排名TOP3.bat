@echo off
chcp 65001 >nul
echo ========================================
echo 首报排名TOP3统计工具
echo ========================================
echo.
echo 正在统计首报排名数据...
echo.

cd /d "%~dp0.."
python -c "import pandas as pd; import sys; from datetime import datetime; sys.stdout.reconfigure(encoding='utf-8'); df_reg = pd.read_excel('数据源/考级学员信息表.xlsx', sheet_name='考级学员信息表'); df_reg = df_reg[['用户id', '用户名', '考试科目名称', '考试等级', '报名时间', '考试日期', '学员地区', '考级资格消耗']]; df_reg['用户id'] = df_reg['用户id'].astype(int); df_follow = pd.read_excel('数据源/目标明细.xlsx', sheet_name='跟进学员名单（在班学员）'); df_follow.columns = ['老师ID', '老师工号', '老师艺名', '学区', '团队', '小组', '班级ID', '班级名称', '班级课表', '阶段', '星期', '上课时间', '学员ID', '学员昵称', '年龄', '地区', '有效续费次数', '是否师生在群', '是否有有效群', '班主任ID', '班主任姓名', '班主任艺名', '班主任大区', '班主任团队', '大V渠道', '累计报考次数', '26Q1报考科目', '26Q1成绩', '25Q4报考科目', '25Q4成绩', '25Q3报考科目', '25Q3成绩', '25Q1报考科目', '25Q1成绩', '24Q4报考科目', '24Q4成绩', '24Q3报考科目', '24Q3成绩', '24Q1报考科目', '24Q1成绩', '儿童画', '儿童画成绩', '漫画', '漫画成绩', '素描', '素描成绩', '国画', '国画成绩', '速写', '速写成绩', '推荐儿童画报考级别', '推荐漫画报考级别', '推荐素描报考级别', '推荐国画报考级别', '推荐速写报考级别', '漫画-是否跳级', '素描-是否跳级', '分层', '剩余课时', '最新国画专题课课表名称', '最新国画专题课上课日期', '备注']; df_follow['学员ID'] = df_follow['学员ID'].astype(int); df_follow['累计报考次数'] = pd.to_numeric(df_follow['累计报考次数'], errors='coerce'); df_teacher = pd.read_excel('数据源/目标明细.xlsx', sheet_name='老师'); registered_ids = set(df_reg['用户id'].unique()); df_follow['是否本期报名'] = df_follow['学员ID'].isin(registered_ids).astype(int); df_follow_registered = df_follow[df_follow['是否本期报名'] == 1].copy(); df_follow_registered['报名类型'] = df_follow_registered['累计报考次数'].apply(lambda x: '首报' if (pd.isna(x) or x == 0) else '复报'); teacher_first = df_follow_registered[df_follow_registered['报名类型'] == '首报'].groupby(['老师ID']).agg(首报人数=('学员ID', 'nunique')).reset_index(); teacher_reg = df_follow_registered.groupby(['老师ID']).agg(报名人数=('学员ID', 'nunique')).reset_index(); result = df_teacher[['老师ID', '老师艺名', '学区', '团队', '小组', '目标学员数_向上取整']].copy(); result = result.merge(teacher_first, on='老师ID', how='left'); result = result.merge(teacher_reg, on='老师ID', how='left'); result['首报人数'] = result['首报人数'].fillna(0).astype(int); result['报名人数'] = result['报名人数'].fillna(0).astype(int); result['达标率'] = (result['报名人数'] / result['目标学员数_向上取整'] * 100).round(1); result = result.sort_values(['首报人数', '达标率'], ascending=[False, False]); top3 = result.head(3).copy(); top3.insert(0, '排名', range(1, len(top3) + 1)); now = datetime.now().strftime('%%Y-%%m-%%d %%H:%%M'); top3['统计截止时间'] = now; output = top3[['排名', '老师艺名', '学区', '团队', '小组', '首报人数', '报名人数', '目标学员数_向上取整', '达标率', '统计截止时间']]; output.columns = ['排名', '老师艺名', '学区', '团队', '小组', '首报报名人数', '总报名人数', '目标人数', '达标率(%%)', '统计截止时间']; print('\n首报排名TOP3:\n'); print(output.to_string(index=False)); output.to_excel('首报排名/首报排名TOP3.xlsx', index=False, sheet_name='首报排名'); print('\n文件已保存：首报排名/首报排名TOP3.xlsx')"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo [成功] 统计完成！
    echo ========================================
) else (
    echo.
    echo ========================================
    echo [错误] 统计失败，请检查数据文件
    echo ========================================
)

echo.
pause
