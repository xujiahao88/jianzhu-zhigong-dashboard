# 建材直供/出库量 日度跟踪看板 — GitHub Pages 自动部署

> 从 `钢材直供量统计.xlsx` 自动生成多年度叠加线图看板，部署到 GitHub Pages。  
> 你只需每天把新数据给 WorkBuddy，WorkBuddy 跑脚本 → 推 GitHub → 网站自动更新。

## 一、它能干什么

1. 读 `源数据-更新` sheet 的原始 A-M 列数据（2018-至今）
2. 用 pandas 自动算 MA5、贸易商拿货量、直供占比
3. matplotlib 画两组多年度叠加图（与 Excel 原图同款样式）：
   - **Page 1**：直供/线盘/螺纹/贸易商拿货量 × 4 区域 = 16 张子图
   - **Page 2**：建材直供占比 × 4 区域 = 4 张子图
4. 生成 `index.html` 整站（可独立打开或推 GitHub Pages）
5. GitHub Actions 自动检测 push → 重跑脚本 → 发布到 Pages

## 二、文件结构

```
项目根目录/
├── build_dashboard.py          # 核心生成脚本
├── update_zhigong.py           # 数据更新脚本（先把新数据追加到 Excel）
├── 钢材直供量统计.xlsx          # 数据源（必须包含「源数据-更新」sheet）
├── 复刻指南_建材直供更新.md     # 数据更新流程的复刻指南
├── .github/workflows/
│   └── deploy.yml              # GitHub Actions 自动部署
├── .gitignore
└── dashboard/                  # 输出目录（自动生成）
    ├── index.html
    ├── page1_直供出库.png
    └── page2_直供占比.png
```

## 三、一次性配置（10 分钟）

### 3.1 创建 GitHub repo

1. 登录 GitHub → New repository → 命名如 `steel-dashboard`
2. **不要**勾选 "Initialize with README"（保持空 repo）

### 3.2 本地初始化 + 推送

```bash
# 在项目根目录
git init
git add build_dashboard.py update_zhigong.py 钢材直供量统计.xlsx .github .gitignore 复刻指南_建材直供更新.md
git commit -m "init: dashboard + update scripts"
git branch -M main
git remote add origin https://github.com/<你的用户名>/steel-dashboard.git
git push -u origin main
```

### 3.3 启用 GitHub Pages

1. 进 repo → **Settings** → **Pages**
2. **Source** 选 **GitHub Actions**
3. 等待第一次 workflow 跑完（一般 1-2 分钟）
4. 完成后访问 `https://<你的用户名>.github.io/steel-dashboard/`

## 四、日常更新流程

**每天只需 2 步**：

```bash
# 1. 拖入新一期 MySteel 源 Excel 给 WorkBuddy，让它跑 update_zhigong.py
python update_zhigong.py "D:/path/新出库量.xlsx"

# 2. 把更新后的 Excel 推 GitHub
git add 钢材直供量统计.xlsx
git commit -m "update: 2026-08-11 直供数据"
git push
```

GitHub Actions 会在 1-2 分钟内：
1. 拉取最新代码
2. 跑 `build_dashboard.py` 重新生成 `dashboard/`
3. 发布到 GitHub Pages
4. 网站自动更新（**无需手动操作**）

## 五、依赖

本地运行需：

```bash
pip install pandas openpyxl matplotlib
```

GitHub Actions 自动装好这些。

## 六、常见问题

### Q1：本地跑 `build_dashboard.py` 中文显示方块

Windows 装个微软雅黑字体，或把脚本里的 `_CN_FONT` 列表加上你有的中文字体名。

### Q2：GitHub Actions 跑失败

进 repo → Actions 标签 → 看错误日志。一般是网络/依赖问题，重试即可。

### Q3：想换主题色 / 改图表布局

直接改 `build_dashboard.py` 里的：
- `INDICATORS` / `RATIO_INDICATORS`：增减指标
- `build_charts()` 函数：改 figsize/dpi/颜色

### Q4：想让部署更频繁（不用等 push）

编辑 `.github/workflows/deploy.yml`，加个定时触发：
```yaml
on:
  schedule:
    - cron: '0 8 * * *'  # 每天 8 点
```

## 七、本地预览

```bash
python build_dashboard.py
# 然后用浏览器打开 dashboard/index.html
```

预览模式不依赖 GitHub。
