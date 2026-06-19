# Euraxess 岗位追踪器

每天自动抓取 [Euraxess](https://euraxess.ec.europa.eu/jobs) 上匹配你关键词的岗位，去重后生成 `jobs.xlsx`，方便每天快速浏览新增的岗位制博士 / 博后 / 研究员职位。

## 工作原理

```
GitHub Actions 每天定时跑
    ↓
按 keywords.txt 逐个搜索 Euraxess
    ↓
解析岗位（标题/机构/国家/截止日期/研究领域...）
    ↓
与 jobs.jsonl 历史对比 → 只入库新增的
    ↓
全量重写 jobs.xlsx → commit 推回仓库
    ↓
你本地 git pull → 打开 jobs.xlsx 看最新结果
```

## 一次性配置（10 分钟）

### 1. 在 GitHub 上 Fork / 新建仓库

把这个项目文件夹推到 GitHub：
```bash
cd "/c/Users/Dong MY/Desktop/euraxess-tracker"
git init
git add .
git commit -m "init euraxess tracker"
git branch -M main
git remote add origin git@github.com:<你的用户名>/euraxess-tracker.git
git push -u origin main
```

### 2. 启用 GitHub Actions

打开你的仓库 → **Actions** 标签页 → 如果有提示，点击 **"I understand my workflows, go ahead and enable them"**。

> 默认情况下，定时任务（`schedule`）只在仓库活跃（最近 60 天有 commit）时才会触发。Fork 出来的仓库第一次需要去 Actions 页面手动 enable workflow。之后正常情况下每天 UTC 0:07（北京时间 8:07）会自动跑一次。

### 3. 手动触发一次，验证能跑通

Actions 页面 → 左侧选 **"Daily Euraxess scan"** → 右上角 **"Run workflow"** 按钮 → 等几分钟，看到绿勾即成功。

跑完后仓库里会多出两个文件：
- `jobs.jsonl` — 累积的所有岗位数据（去重用的状态文件）
- `jobs.xlsx` — 给你看的 Excel

## 日常使用

### 改关键词

编辑 `keywords.txt`（每行一个），commit 推上去，下一次跑就生效。

```bash
# 本地改完后
git add keywords.txt
git commit -m "update keywords"
git push
```

### 看每天的最新岗位

```bash
git pull
# 然后用 Excel / WPS 打开 jobs.xlsx
```

`jobs.xlsx` 字段说明：

| 列 | 含义 |
|---|---|
| First Seen | 脚本第一次发现这条岗位的日期 |
| Posted | Euraxess 上岗位的发布日期 |
| Deadline | 申请截止日期 |
| Country | 国家 |
| Organization | 招聘机构 |
| Title | 岗位标题 |
| Research Field | 研究领域分类 |
| Researcher Profile | 研究员等级（R1/R2/R3/R4） |
| Type | 类型（JOB / FUNDING） |
| Matched Keyword | 命中的搜索关键词 |
| URL | 详情页链接 |

**「First Seen」列是最关键的** —— 按它倒序排序，今天第一次出现的行就是今天的新岗位。

## 本地手动跑（可选）

如果想在你自己电脑上跑（比如调试、或者想多跑一次）：

```bash
cd "/c/Users/Dong MY/Desktop/euraxess-tracker"
python -m venv .venv
source .venv/Scripts/activate   # Windows Git Bash
pip install -r requirements.txt
python scraper.py
```

跑完会在本地生成 / 更新 `jobs.jsonl` 和 `jobs.xlsx`。

## 常见问题

**Q: GitHub Actions 不自动跑？**  
A: 仓库 60 天内没有 commit 时 GitHub 会暂停定时任务。每月随便 push 个 commit 即可，或者去 Actions 页面手动点 "Run workflow"。

**Q: 关键词抓到的结果太多/太少？**  
A: 多了 → 把关键词改更具体（如 `"reinforcement learning"` 而不是 `"learning"`）；少了 → 改更宽泛或加更多关键词。`MAX_PAGES_PER_KEYWORD`（在 `scraper.py` 顶部）控制每个关键词抓多少页。

**Q: 想加国家/领域过滤？**  
A: 关键词里加就行，比如 `"phd position Germany"`。更复杂的过滤（如「只要 R1 等级」）可以改 `scraper.py` 的 `main()` 在 `unique.values()` 后再加一层过滤。

**Q: 想加邮件/Telegram 推送？**  
A: 在 `.github/workflows/daily.yml` 的 "Run scraper" 步骤后加一个发通知的 step。可以让 Claude 帮你加。

**Q: 数据会被 Euraxess 封吗？**  
A: 不会。每天一次、几个关键词、几十个请求，对 Euraxess 是非常温和的流量。脚本里也有 `time.sleep(1)` 的礼貌延迟。
