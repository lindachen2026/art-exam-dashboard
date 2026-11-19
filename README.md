# 考级数据看板自动化系统

## 🎯 快速开始

### 新季度启动（3步）

1. **编辑配置**（5分钟）
   - `数据源/本期考级目标.xlsx` - 季度、目标、日期
   - `数据源/目标明细.xlsx` - MTD流速（总和=1.0）
   - `auto_download_bi.py` - 更新DATE_START和DATE_END

2. **配置定时任务**（1分钟）
   ```bash
   双击运行: setup_schedule.bat
   ```

3. **测试验证**（2分钟）
   ```bash
   python auto_update_and_publish.py
   ```
   访问：https://lindachen2026.github.io/art-exam-dashboard/dashboard.html

### 报名期间（自动运行）

✅ 无需手动操作，系统每天自动执行4次：
- 上午 10:50
- 下午 14:00
- 下午 18:00
- 晚上 23:59

每次自动完成：下载BI数据 → 生成看板 → 推送GitHub

### 报名结束后（2步）

1. **禁用定时任务**
   - Win+R → `taskschd.msc`
   - 找到4个任务，右键禁用

2. **归档历史数据**
   ```bash
   python archive_current_period.py
   ```

---

## 📂 文件说明

### 核心脚本
- `auto_download_bi.py` - BI数据自动下载
- `generate_dashboard.py` - 看板生成
- `auto_update_and_publish.py` - 主调度（下载→生成→推送）
- `archive_current_period.py` - 历史归档
- `setup_schedule.bat` - 定时任务配置

### 配置文件
- `数据源/本期考级目标.xlsx` - 季度基础配置
- `数据源/目标明细.xlsx` - MTD流速配置
- `数据源/往期数据.xlsx` - 历史对比数据（自动更新）

### 数据文件
- `数据源/考级学员信息表.xlsx` - BI下载原始数据（自动更新）
- `数据源/本期各维度报名明细.xlsx` - 统计数据（自动生成）
- `dashboard.html` - 生成的看板（自动推送GitHub）

---

## 🔧 常用命令

**手动测试更新**:
```bash
python auto_update_and_publish.py
```

**查看定时任务**:
```bash
taskschd.msc
```

**归档历史数据**:
```bash
python archive_current_period.py
```

**查看看板**:
- 本地：`dashboard.html`
- 云端：https://lindachen2026.github.io/art-exam-dashboard/dashboard.html

---

## 📖 详细文档

完整实施手册：[实施手册.md](实施手册.md)

---

## ⚠️ 注意事项

1. **MTD流速总和必须=1.0**
   - 打开 `数据源/目标明细.xlsx - MTD sheet`
   - 检查D列（每日流速）总和
   - C列公式会自动计算目标人数

2. **报名日期范围必须一致**
   - `本期考级目标.xlsx` 中的报名时间
   - `auto_download_bi.py` 中的 DATE_START/DATE_END
   - `目标明细.xlsx` MTD的日期范围

3. **报名结束后必须归档**
   - 运行 `archive_current_period.py`
   - 自动更新 `往期数据.xlsx`
   - 新季度启动时无需手动补录

4. **GitHub Pages自动部署**
   - 推送到main分支后自动部署
   - 约1-2分钟后生效
   - 确保仓库已开启GitHub Pages

---

## 🚨 故障排查

**BI下载失败**:
- 检查BI系统是否可访问
- 检查账号密码是否正确
- 查看错误截图：`BI导数据表/error_*.png`

**看板数据异常**:
- 检查 `考级学员信息表.xlsx` 是否为空
- 检查MTD流速总和是否=1.0
- 查看往期数据是否完整

**GitHub推送失败**:
- 检查网络连接
- 检查Git配置
- 手动执行 `git push origin main` 查看错误

**定时任务未执行**:
- 打开"任务计划程序"检查任务状态
- 检查"上次运行结果"是否为0x0（成功）
- 查看任务历史记录

---

## 📊 系统架构

```
季度启动前（手动配置）
    ↓
配置定时任务（运行setup_schedule.bat）
    ↓
报名期间（自动运行，每天4次）
    ├─ 10:50 → 下载BI → 生成看板 → 推送GitHub
    ├─ 14:00 → 下载BI → 生成看板 → 推送GitHub
    ├─ 18:00 → 下载BI → 生成看板 → 推送GitHub
    └─ 23:59 → 下载BI → 生成看板 → 推送GitHub
    ↓
报名结束后
    ├─ 禁用定时任务
    ├─ 最后一次手动更新
    └─ 运行归档脚本（archive_current_period.py）
```

---

## 版本信息

- **版本**: v2.0
- **更新日期**: 2026-07-02
- **适用季度**: 2026Q3及以后
- **GitHub仓库**: https://github.com/lindachen2026/art-exam-dashboard
