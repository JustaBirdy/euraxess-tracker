"""Euraxess 岗位追踪器

每天抓取 Euraxess 上匹配关键词的岗位，去重后追加到 jobs.jsonl，
并把全量数据写成 jobs.xlsx 供本地查看。

设计要点：
- 状态保存在 jobs.jsonl（纯文本，每行一个 JSON 对象），git diff 友好
- jobs.xlsx 是从 jobs.jsonl 全量重写的「视图文件」，方便用 Excel 打开
- 关键词放在 keywords.txt，每行一个；空行和 # 开头的注释行被忽略
- 意向国家放在 countries.txt（英文国名，如 Germany/Switzerland/Sweden/Netherlands）；
  留空 = 不限国家。国家用 facet 过滤(多国 OR)，ID 运行时从搜索页动态解析
"""

from __future__ import annotations

import json
import re
import sys
import time
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path

import httpx
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

# Windows 控制台默认 GBK，遇到 é/ç/中文引号会崩；强制 UTF-8 输出
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

BASE_URL = "https://euraxess.ec.europa.eu"
SEARCH_URL = f"{BASE_URL}/jobs/search"

ROOT = Path(__file__).parent
KEYWORDS_FILE = ROOT / "keywords.txt"
COUNTRIES_FILE = ROOT / "countries.txt"
JSONL_PATH = ROOT / "jobs.jsonl"
EXCEL_PATH = ROOT / "jobs.xlsx"

# 每个关键词抓前 N 页。每页约 10 条，3 页 ≈ 30 条/关键词/天，足够覆盖每日新增
MAX_PAGES_PER_KEYWORD = 3

# 相邻两次请求之间的最小间隔（秒）。Euraxess 对高频访问会返回 429，需要放慢节奏。
# 注意：GitHub Actions 跑在共享 IP 上，更容易被限流，这个值宁可大一点。
REQUEST_DELAY = 3

USER_AGENT = "euraxess-tracker/1.0 (+https://github.com/your-username/euraxess-tracker)"
JOB_HREF_RE = re.compile(r"/jobs/(\d+)")


@dataclass
class Job:
    job_id: str
    title: str
    organization: str
    country: str
    type: str
    posted: str
    deadline: str
    research_field: str
    researcher_profile: str
    description: str
    url: str
    first_seen: str
    keyword: str  # 命中的搜索关键词，方便回溯


def load_keywords() -> list[str]:
    """读取 keywords.txt，忽略空行和注释。"""
    if not KEYWORDS_FILE.exists():
        return []
    kws = []
    for line in KEYWORDS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        kws.append(line)
    return kws


def load_countries() -> list[str]:
    """读取 countries.txt（意向国家英文名），忽略空行和注释。留空=不限国家。"""
    if not COUNTRIES_FILE.exists():
        return []
    cs = []
    for line in COUNTRIES_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        cs.append(line)
    return cs


def get_country_id_map() -> dict[str, str]:
    """GET 一次搜索页，解析国家下拉框 select[name='job_country[]']，
    返回 {国家名小写: facet ID}。Euraxess 用数字 ID 做 country facet（如
    Germany=794），不硬编码、运行时动态解析以适配官方 ID 变动。
    """
    r = httpx.get(
        SEARCH_URL, headers={"User-Agent": USER_AGENT}, timeout=30, follow_redirects=True
    )
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")
    sel = soup.find("select", {"name": "job_country[]"})
    m: dict[str, str] = {}
    if sel:
        for opt in sel.find_all("option"):
            val = opt.get("value")
            txt = opt.get_text(strip=True)
            if val and txt:
                m[txt.lower()] = val
    return m


def fetch_page(keyword: str, page: int, country_ids: list[str] | None = None) -> str | None:
    """抓取一页搜索结果 HTML。

    country_ids 为意向国家的 facet ID 列表（由 get_country_id_map 解析得到）；
    非空时叠加国家过滤（多国 OR），为空则不限国家。

    遇到 429 会优先尊重服务器的 Retry-After 头并指数退避重试；
    彻底失败时返回 None（由调用方决定是否跳过该页），避免单页失败
    让整个 run 崩掉、丢掉已抓到的数据。
    """
    # Euraxess 是 Drupal 站点：GET 的 ?keywords= 参数会被完全忽略——实测 nonsense
    # 词、空参数、以及候选参数名(keys/text/search_api_fulltext)都返回同一批「全站最新」
    # 岗位。真正生效的关键词过滤是 facet 参数 f[0]=keywords:<词>。已用对比实验验证：
    # 正常词返回相关岗位、nonsense 词返回 0 条、page 翻页有效。默认排序即按发布时间
    # 倒序(最新优先)，无需额外 sort 参数(原 sort_by=publication_date 本就是无效噪音)。
    # f[0] 给关键词；f[1..] 给国家 facet（同 facet 多值 = OR）。
    params: dict[str, str | int] = {
        "f[0]": f"keywords:{keyword}",
        "page": page,
    }
    for i, cid in enumerate(country_ids or []):
        params[f"f[{i + 1}]"] = f"job_country:{cid}"
    max_attempts = 5
    backoff = 5.0  # 指数退避起点；每次翻倍，上限 60s
    for attempt in range(1, max_attempts + 1):
        try:
            r = httpx.get(
                SEARCH_URL,
                params=params,
                headers={"User-Agent": USER_AGENT},
                timeout=30,
                follow_redirects=True,
            )
            # 429 单独处理：raise_for_status 抛的异常拿不到响应头，无法读 Retry-After
            if r.status_code == 429:
                retry_after = r.headers.get("Retry-After", "").strip()
                if retry_after.isdigit():
                    wait = min(int(retry_after), 60)
                else:
                    wait = backoff  # Retry-After 也可能是 HTTP-date，退回指数退避
                print(f"  429 limited (keywords={keyword!r} page={page}); "
                      f"sleep {wait}s, retry {attempt}/{max_attempts}")
                time.sleep(wait)
                backoff = min(backoff * 2, 60.0)
                continue
            r.raise_for_status()
            return r.text
        except httpx.HTTPError as e:
            if attempt == max_attempts:
                print(f"  !! give up keywords={keyword!r} page={page} after "
                      f"{max_attempts} attempts: {e}")
                return None
            print(f"  retry {attempt}/{max_attempts} after error: {e}; "
                  f"sleep {backoff:.0f}s")
            time.sleep(backoff)
            backoff = min(backoff * 2, 60.0)
    return None


def parse_jobs(html: str, keyword: str) -> list[Job]:
    """从搜索结果 HTML 解析岗位列表。"""
    soup = BeautifulSoup(html, "lxml")
    jobs: list[Job] = []
    today = datetime.now().strftime("%Y-%m-%d")

    for article in soup.select("article.ecl-content-item"):
        # 标题 + 详情链接
        title_a = article.select_one("h3.ecl-content-block__title a")
        if not title_a:
            continue
        href = title_a.get("href", "")
        m = JOB_HREF_RE.search(href)
        if not m:
            continue
        job_id = m.group(1)
        title = title_a.get_text(strip=True)
        url = BASE_URL + href

        # 顶部两个 label（类型 + 国家）在 article 的前置兄弟节点
        # label-thumbnail-wrapper 里。Euraxess 把它们放在 article 之外。
        teaser_wrapper = article.find_previous_sibling(
            "div", class_="label-thumbnail-wrapper"
        )
        labels: list[str] = []
        if teaser_wrapper:
            labels = [s.get_text(strip=True) for s in teaser_wrapper.select("span.ecl-label")]
        job_type = labels[0] if labels else ""
        country = labels[1] if len(labels) > 1 else ""

        # 主要 meta：机构 + Posted on
        primary_items = article.select(
            "ul.ecl-content-block__primary-meta-container li.ecl-content-block__primary-meta-item"
        )
        organization = ""
        posted = ""
        if primary_items:
            org_a = primary_items[0].find("a")
            organization = org_a.get_text(strip=True) if org_a else primary_items[0].get_text(strip=True)
        if len(primary_items) >= 2:
            posted = primary_items[1].get_text(strip=True).replace("Posted on:", "").strip()

        # 描述
        desc_div = article.select_one("div.ecl-content-block__description")
        description = desc_div.get_text(" ", strip=True) if desc_div else ""

        # 次级 meta：Research Field / Researcher Profile / Application Deadline
        # 外层 div 的第一个 class 以 "id-" 开头（如 "id-Research-Field"）
        # 用 CSS 属性选择器（bs4 的 class_ callable 接收单个 class 字符串，不是 list，易踩坑）
        research_field = ""
        researcher_profile = ""
        deadline = ""
        for container in article.select('div[class^="id-"]'):
            classes = container.get("class", [])
            label = classes[0] if classes else ""
            value_div = container.select_one("div.ecl-text-standard")
            if not value_div:
                continue
            if label == "id-Research-Field":
                research_field = value_div.get_text(" ", strip=True)
            elif label == "id-Researcher-Profile":
                researcher_profile = value_div.get_text(" ", strip=True)
            elif label == "id-Application-Deadline":
                t = value_div.find("time")
                deadline = t.get_text(strip=True) if t else value_div.get_text(strip=True)

        jobs.append(
            Job(
                job_id=job_id,
                title=title,
                organization=organization,
                country=country,
                type=job_type,
                posted=posted,
                deadline=deadline,
                research_field=research_field,
                researcher_profile=researcher_profile,
                description=description,
                url=url,
                first_seen=today,
                keyword=keyword,
            )
        )
    return jobs


def load_seen_ids() -> set[str]:
    """从 jobs.jsonl 读出所有已记录的 job_id。"""
    if not JSONL_PATH.exists():
        return set()
    seen = set()
    for line in JSONL_PATH.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        try:
            seen.add(json.loads(line)["job_id"])
        except (json.JSONDecodeError, KeyError):
            continue
    return seen


def append_new_jobs(new_jobs: list[Job]) -> None:
    """把新岗位追加到 jobs.jsonl。"""
    with JSONL_PATH.open("a", encoding="utf-8") as f:
        for j in new_jobs:
            f.write(json.dumps(asdict(j), ensure_ascii=False) + "\n")


def write_excel() -> None:
    """从 jobs.jsonl 全量重写 jobs.xlsx。"""
    if not JSONL_PATH.exists():
        return

    rows = []
    for line in JSONL_PATH.read_text(encoding="utf-8").splitlines():
        if line.strip():
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue

    # 排序：先按首次发现日期倒序，再按发布日期倒序
    rows.sort(key=lambda r: (r.get("first_seen", ""), r.get("posted", "")), reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"

    headers = [
        "First Seen", "Posted", "Deadline", "Country", "Organization",
        "Title", "Research Field", "Researcher Profile", "Type",
        "Matched Keyword", "URL",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    for r in rows:
        ws.append([
            r.get("first_seen", ""),
            r.get("posted", ""),
            r.get("deadline", ""),
            r.get("country", ""),
            r.get("organization", ""),
            r.get("title", ""),
            r.get("research_field", ""),
            r.get("researcher_profile", ""),
            r.get("type", ""),
            r.get("keyword", ""),
            r.get("url", ""),
        ])

    # 列宽（按列字母）
    col_widths = {
        "A": 12, "B": 16, "C": 32, "D": 14, "E": 30,
        "F": 60, "G": 28, "H": 26, "I": 10, "J": 22, "K": 50,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # 冻结首行
    ws.freeze_panes = "A2"

    wb.save(EXCEL_PATH)


def main() -> int:
    print(f"=== Euraxess tracker run at {datetime.now().isoformat(timespec='seconds')} ===")

    keywords = load_keywords()
    if not keywords:
        print("No keywords configured. Edit keywords.txt first.")
        return 1
    print(f"Keywords: {keywords}")
    print(f"Max pages per keyword: {MAX_PAGES_PER_KEYWORD}")

    # 意向国家：解析成 facet ID（Euraxess 用数字 ID，如 Germany=794）
    countries = load_countries()
    country_ids: list[str] = []
    allowed_country_names: set[str] = set()  # 小写，用于结果兜底过滤
    if countries:
        try:
            cmap = get_country_id_map()
        except httpx.HTTPError as e:
            print(f"  !! 无法获取国家下拉框({e})，本次将不限国家抓取")
            cmap = {}
        for c in countries:
            cid = cmap.get(c.lower())
            if cid:
                country_ids.append(cid)
                allowed_country_names.add(c.lower())
                print(f"  country {c} -> facet id {cid}")
            else:
                print(f"  !! 国家 {c!r} 在 Euraxess 下拉框里找不到，已跳过")
        if not country_ids:
            print("  没有解析到任何有效国家，将不限国家抓取")

    # 抓取
    all_jobs: list[Job] = []
    for kw in keywords:
        for page in range(MAX_PAGES_PER_KEYWORD):
            print(f"  fetching: keywords={kw!r} page={page}")
            html = fetch_page(kw, page=page, country_ids=country_ids)
            if html is None:
                # 单页彻底抓不到（如持续限流）就跳过，继续跑后续关键词，
                # 保住这一轮已经抓到的其它数据。
                print(f"    skipped (fetch failed)")
                continue
            jobs = parse_jobs(html, keyword=kw)
            print(f"    parsed {len(jobs)} jobs")
            all_jobs.extend(jobs)
            time.sleep(REQUEST_DELAY)  # 礼貌延迟，避免触发限流

    # job_id 去重（不同关键词可能抓到同一条；同关键词也可能跨页重复）
    unique: dict[str, Job] = {}
    for j in all_jobs:
        # 同一 job_id 第一次出现时记下；后续命中其他关键词时不覆盖
        unique.setdefault(j.job_id, j)
    print(f"Total unique jobs fetched: {len(unique)}")

    # 与历史对比，挑出真正新增的
    seen = load_seen_ids()
    new_jobs = [j for j in unique.values() if j.job_id not in seen]

    # 兜底：即便 facet 已按国家过滤，这里再按国家名筛一遍，保证数据干净
    if allowed_country_names:
        before = len(new_jobs)
        new_jobs = [j for j in new_jobs if j.country.lower() in allowed_country_names]
        if before != len(new_jobs):
            print(f"Country backstop filtered {before - len(new_jobs)} non-target jobs")

    print(f"New jobs (not in history): {len(new_jobs)}")

    if new_jobs:
        append_new_jobs(new_jobs)

    # 重写 Excel
    write_excel()
    print(f"Excel written: {EXCEL_PATH}")

    if new_jobs:
        print("\n--- New jobs today ---")
        for j in new_jobs:
            print(f"  [{j.country}] {j.title}")
            print(f"    {j.organization} | deadline: {j.deadline} | {j.url}")

    print("\nDone.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
