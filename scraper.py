"""Euraxess 岗位追踪器

每天抓取 Euraxess 上匹配关键词的岗位，去重后追加到 jobs.jsonl，
并把全量数据写成 jobs.xlsx 供本地查看。

设计要点：
- 状态保存在 jobs.jsonl（纯文本，每行一个 JSON 对象），git diff 友好
- jobs.xlsx 是从 jobs.jsonl 全量重写的「视图文件」，方便用 Excel 打开
- 关键词放在 keywords.txt，每行一个；空行和 # 开头的注释行被忽略
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
JSONL_PATH = ROOT / "jobs.jsonl"
EXCEL_PATH = ROOT / "jobs.xlsx"

# 每个关键词抓前 N 页。每页约 10 条，3 页 ≈ 30 条/关键词/天，足够覆盖每日新增
MAX_PAGES_PER_KEYWORD = 3

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


def fetch_page(keyword: str, page: int) -> str:
    """抓取一页搜索结果 HTML。"""
    params = {
        "keywords": keyword,
        "sort_by": "publication_date",
        "page": page,
    }
    # Euraxess 偶尔会慢，给点容错
    for attempt in range(3):
        try:
            r = httpx.get(
                SEARCH_URL,
                params=params,
                headers={"User-Agent": USER_AGENT},
                timeout=30,
                follow_redirects=True,
            )
            r.raise_for_status()
            return r.text
        except httpx.HTTPError as e:
            if attempt == 2:
                raise
            print(f"  retry {attempt+1}/3 after error: {e}")
            time.sleep(3)


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

    # 抓取
    all_jobs: list[Job] = []
    for kw in keywords:
        for page in range(MAX_PAGES_PER_KEYWORD):
            print(f"  fetching: keywords={kw!r} page={page}")
            html = fetch_page(kw, page=page)
            jobs = parse_jobs(html, keyword=kw)
            print(f"    parsed {len(jobs)} jobs")
            all_jobs.extend(jobs)
            time.sleep(1)  # 礼貌延迟

    # job_id 去重（不同关键词可能抓到同一条；同关键词也可能跨页重复）
    unique: dict[str, Job] = {}
    for j in all_jobs:
        # 同一 job_id 第一次出现时记下；后续命中其他关键词时不覆盖
        unique.setdefault(j.job_id, j)
    print(f"Total unique jobs fetched: {len(unique)}")

    # 与历史对比，挑出真正新增的
    seen = load_seen_ids()
    new_jobs = [j for j in unique.values() if j.job_id not in seen]
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
