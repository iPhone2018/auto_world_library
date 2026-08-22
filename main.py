#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
worldlibrary.ai 电子书信息采集工具

流程：
1. 启动时访问 https://worldlibrary.ai/ 获取首页完整 HTML，再从首页引用的 JS bundle 中
   解析出合集中 20 个站点的站点名称（中/英文）与跳转链接（站点 key 如 main/agriculture）
2. 界面选择站点名称（可配置开始/结束年份，按年逐次请求），点击"开始执行"后：
   - 调用 https://worldlibrary.ai/wl-api/search/search（POST）翻页采集，
     每页固定最多 20 条，每次调用接口后 sleep 间隔秒；请求体带 document_type=book
     过滤（服务端只返回电子书），可选 publication_year_from/to 按出版年份过滤
   - 接口限制 from+size <= 10000：站点总量超过 9980 时，自动按书名标题的词条字典序
     范围（title:[lo TO hi]）递归二分切分，保证全部书籍都被完整覆盖采集
   - 逐本调用 https://worldlibrary.ai/wl-api/ai/book-translate 将英文书名翻译为中文
     （标题已含中文则跳过翻译，翻译失败回退原文）
   - 输出：每个站点一个"当前"Excel（output/worldlibrary_{站点key}.xlsx），
     超过 MAX_ROWS_PER_FILE 行自动归档为 worldlibrary_{key}_{序号}.xlsx 另开新文件；
     跨文件去重用纯文本 worldlibrary_{key}_ids.txt（每行一个书籍ID，非数据库），
     启动时流式读入去重（txt 丢失时会自动扫描 Excel 重建）；每累计 500 条落盘一次
3. 结束执行按钮可随时中断

说明：
- 书籍封面链接按官网 JS 中的规律由书籍ID直接拼接（官网本身也不校验是否存在）
- 接口有请求频率限制（约十几秒内十几次），触发后自动等待 60 秒起（翻倍）重试
"""

import glob
import math
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from queue import Empty, Queue
from tkinter import scrolledtext, ttk
from urllib.parse import urljoin

import openpyxl
import requests
from openpyxl.styles import Alignment, Font, PatternFill

# ==================== 配置区域 ====================

HOME_URL = "https://worldlibrary.ai/"
SEARCH_API = "https://worldlibrary.ai/wl-api/search/search"
TRANSLATE_API = "https://worldlibrary.ai/wl-api/ai/book-translate"
# 站点搜索页（作为请求头 Referer），站点主页（站点跳转链接）
SEARCH_PAGE_TPL = "https://worldlibrary.ai/search/{key}?keyword=&way=search-ai"
HOME_PAGE_TPL = "https://worldlibrary.ai/home/{key}"
BOOK_URL_TPL = "https://worldlibrary.ai/book/{key}/{book_id}"

PAGE_SIZE = 20          # 接口每页固定最多返回 20 条
WINDOW_LIMIT = 9980     # 接口限制 from+size<=10000，单查询最多覆盖 9980+20 条
ALNUM = "0123456789abcdefghijklmnopqrstuvwxyz"

# 标题词条字典序切分的顶层范围（[lo TO hi]，lo 为 "*" 表示开区间下界）。
# 各范围首尾相接（上一范围末尾字符开头的词条归入下一范围），覆盖书名所有词条。
TOP_TITLE_RANGES = [
    ("*", "9"),            # 数字开头的词条（年份等）
    ("a", "z"),            # 拉丁字母
    ("À", "ÿ"),            # 拉丁扩展
    ("Ā", "⿿"),           # 希腊/西里尔/希伯来/阿拉伯/天城文/泰文等
    ("ぁ", "鿿"),          # 日文假名 + CJK
    ("ꀀ", "￿"),           # 彝文/谚文/兼容字符等
]

REQUEST_INTERVAL = 3  # 搜索接口每次调用后 sleep 的秒数（太短会频繁触发限流）
TRANSLATE_INTERVAL = 0.2  # 翻译接口每次调用后 sleep 的秒数
API_TIMEOUT = 60        # 单次请求超时秒
API_RETRY = 4           # 失败/限流重试次数
API_RETRY_SLEEP = 5     # 网络异常重试前等待秒
THROTTLE_SLEEP = 60     # 触发限流后首次等待秒（随重试翻倍）
FLUSH_ROWS = 500        # 每累计新增该条数，将 Excel 追加落盘一次（防意外丢失）
MAX_ROWS_PER_FILE = 200000  # 单个Excel超过该行数自动归档（另开新文件），避免文件过大加载变慢

# 代理：None 表示自动探测（环境变量 → 系统代理：Windows 注册表 / macOS scutil）；可手动指定如 "http://127.0.0.1:7897"
PROXY_OVERRIDE = None

OUTPUT_DIR = "output"

# 输出列（与《书籍信息采集模板.xlsx》一致）
EXCEL_COLUMNS = ["书籍ID", "书籍名称", "作者", "出版社", "出版时间", "ISBN",
                 "页数", "书籍封面链接", "书籍链接", "SSN号", "读秀号"]

# 搜索接口请求体字段
SOURCE_FIELDS = [
    "title", "author", "language", "publisher", "publication_year", "collection",
    "document_type", "brief_summary", "standard_summary", "subject", "keyword",
    "bisac_l1", "bisac_l2", "bisac_l3", "page", "file_size_bytes",
]
FACETS = ["author", "language", "publication_year", "publisher", "subject",
          "keyword", "document_type", "collection", "bisac", "indicator"]
FILTER_DIMS = ["author", "language", "publisher", "subject", "keyword",
               "document_type", "indicator"]

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36")

# 站点列表兜底（首页/JS 解析失败时使用）：(key, 英文名, 中文名)
FALLBACK_SITES = [
    ("main", "Main Site", "主站点"),
    ("asian", "Asian Studies", "亚洲研究"),
    ("socialscience", "Social Sciences", "社会科学"),
    ("business", "Business & Economics", "商业与经济"),
    ("agriculture", "Agriculture", "农业"),
    ("humanity", "Humanity", "人文学科"),
    ("engineering", "Engineering", "工程学"),
    ("africa", "African Studies", "非洲研究"),
    ("chinesestudies", "Chinese Studies", "中国研究"),
    ("environmental", "Environmental", "环境"),
    ("history", "History", "历史"),
    ("lifescience", "Life Science", "生命科学"),
    ("middleeast", "Middles East Studies", "中东研究"),
    ("healthscience", "Health Sciences", "健康科学"),
    ("nursing", "Nursing", "护理学"),
    ("draa", "DRAA Select", "DRAA精选集"),
    ("science", "The Sciences", "科学"),
    ("asean", "ASEAN", "东盟"),
    ("defense", "Defense", "国防"),
    ("southasia", "South Asia", "南亚"),
]

# ==================== 代理探测 ====================


def _windows_system_proxy():
    """读取 Windows 系统代理（注册表 Internet Settings，Clash/v2rayN 等客户端
    开启"系统代理"后会写入）。非 Windows 或未启用时返回 None"""
    if sys.platform != "win32":
        return None
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings") as k:
            enabled = winreg.QueryValueEx(k, "ProxyEnable")[0]
            server = winreg.QueryValueEx(k, "ProxyServer")[0]
    except Exception:
        return None
    if not enabled or not server:
        return None
    http = https = None
    # ProxyServer 形如 "127.0.0.1:7890"，或 "http=127.0.0.1:7890;https=...;socks=..."
    for part in str(server).split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            scheme, _, addr = part.partition("=")
            scheme = scheme.strip().lower()
        else:
            scheme, addr = "", part
        addr = addr.strip()
        if not addr or scheme == "socks":   # socks 需要额外安装 PySocks，忽略
            continue
        if "://" not in addr:
            addr = "http://" + addr
        if scheme == "http" and http is None:
            http = addr
        elif scheme == "https" and https is None:
            https = addr
        elif not scheme:
            if http is None:
                http = addr
            if https is None:
                https = addr
    if http or https:
        return {"http": http or https, "https": https or http}
    return None


def detect_proxy():
    """探测可用代理：环境变量 → 系统代理（Windows 注册表 / macOS scutil）→ 直连"""
    if PROXY_OVERRIDE:
        return {"http": PROXY_OVERRIDE, "https": PROXY_OVERRIDE}
    for name in ("https_proxy", "HTTPS_PROXY", "http_proxy", "HTTP_PROXY",
                 "all_proxy", "ALL_PROXY"):
        val = os.environ.get(name)
        if val:
            return {"http": val, "https": val}
    win = _windows_system_proxy()
    if win:
        return win
    try:
        out = subprocess.run(["scutil", "--proxy"], capture_output=True,
                             text=True, timeout=5).stdout
        enabled = re.search(r"(HTTP|HTTPS)Enable\s*:\s*1", out)
        host = re.search(r"(HTTP|HTTPS)Proxy\s*:\s*(\S+)", out)
        port = re.search(r"(HTTP|HTTPS)Port\s*:\s*(\d+)", out)
        if enabled and host and port:
            url = "http://%s:%s" % (host.group(2), port.group(2))
            return {"http": url, "https": url}
    except Exception:
        pass
    return None


PROXIES = detect_proxy()

# ==================== 全局停止控制 ====================
_stop_event = threading.Event()


class TaskStoppedException(Exception):
    pass


def check_stop():
    if _stop_event.is_set():
        raise TaskStoppedException()


def interval_sleep(seconds: float):
    """可中断的 sleep：等待期间收到停止信号则立即抛出"""
    if _stop_event.wait(timeout=seconds):
        raise TaskStoppedException()


# ==================== 线程安全日志队列 ====================
LOG_QUEUE = Queue(maxsize=2000)


def log_print(text):
    LOG_QUEUE.put(text)


# ==================== 接口调用 ====================

def build_headers(site_key: str) -> dict:
    """请求头（Referer 使用对应站点的搜索页）"""
    return {
        'User-Agent': UA,
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Origin': 'https://worldlibrary.ai',
        'Referer': SEARCH_PAGE_TPL.format(key=site_key),
    }


def build_body(site_key: str, query: str = "*", filters: dict = None,
               from_: int = None) -> dict:
    """构造搜索请求体（与官网前端一致：过滤生效时带 _type=filter）"""
    body = {
        "search_type": "keyword",
        "query": query,
        "size": PAGE_SIZE,
        "facet_size": 100,
        "source_fields": list(SOURCE_FIELDS),
        "facet": list(FACETS),
        "author": [], "language": [], "publisher": [], "subject": [],
        "keyword": [], "document_type": ["book"], "indicator": [],
        "sort_field": "publication_year",
        "sort_order": "desc",
    }
    if site_key != "main":          # 主站点不带 collection，其余站点按站点 key 过滤合集
        body["collection"] = [site_key]
    filters = filters or {}
    has_filter = True               # 请求体始终带 document_type=["book"] 过滤
    for dim in FILTER_DIMS:
        if filters.get(dim):
            body[dim] = list(filters[dim])
    if filters.get("publication_year_from") is not None or filters.get("publication_year_to") is not None:
        body["publication_year_from"] = filters.get("publication_year_from")
        body["publication_year_to"] = filters.get("publication_year_to")
    if has_filter:
        body["_type"] = "filter"
    if from_ is not None:
        body["from"] = from_
    return body


def fetch_json(url: str, payload: dict, headers: dict, timeout: int = API_TIMEOUT,
               retries: int = API_RETRY):
    """POST JSON 接口，带失败重试与限流退避。成功返回解析后的 dict，失败返回 None"""
    wait = THROTTLE_SLEEP
    for attempt in range(1, retries + 1):
        check_stop()
        try:
            res = requests.post(url, json=payload, headers=headers,
                                timeout=timeout, proxies=PROXIES)
        except requests.RequestException as e:
            log_print(f"[!] 请求失败（第{attempt}/{retries}次）: {e}")
            interval_sleep(API_RETRY_SLEEP)
            continue
        try:
            j = res.json()
        except ValueError:
            log_print(f"[!] 响应非JSON（第{attempt}/{retries}次）HTTP {res.status_code}: "
                      f"{res.text[:120]}")
            interval_sleep(API_RETRY_SLEEP)
            continue
        if (j.get("message") == "Request frequency too high"
                or j.get("i18nMsg") == "message.request-frequency-too-high"):
            log_print(f"[!] 触发接口限流，等待 {wait} 秒后重试（第{attempt}/{retries}次）")
            interval_sleep(wait)
            wait *= 2
            continue
        return j
    return None


def search_page(site_key: str, query: str = "*", filters: dict = None, from_: int = None):
    """调用搜索接口取一页数据（含 total 与 aggregations）。失败返回 None"""
    return fetch_json(SEARCH_API, build_body(site_key, query, filters, from_),
                      build_headers(site_key))


def translate_title(site_key: str, book_id: str, title: str) -> str:
    """调用翻译接口将书名翻译为中文。失败/无结果返回空字符串"""
    j = fetch_json(TRANSLATE_API,
                   {"id": book_id, "title": title, "summary": "", "lang": "zh-CN"},
                   build_headers(site_key))
    if j and j.get("code") == 200:
        data = j.get("data") or {}
        if data.get("title"):
            return str(data["title"])
    return ""


# ==================== 站点列表（首页 HTML + JS bundle 解析） ====================

def parse_sites(js: str) -> list:
    """从 JS bundle 中解析站点列表：key、英文名、中文名、跳转链接"""
    sites = []
    # 1) zh-CN 语言包中 header 段的中文名映射
    #    压缩后的 JS 里 key 有的带引号（"main-site"）有的不带（agriculture），需两遍匹配
    zh_map = {}
    pos = js.find('"main-site":"主站点"')
    if pos >= 0:
        start = js.rfind('header:{', 0, pos)
        end = js.find('}', pos)
        if 0 <= start < end:
            seg = js[start:end]
            for k, v in re.findall(r'"([a-z0-9-]+)":"([^"]*)"', seg):
                zh_map[k] = v
            for k, v in re.findall(r'[{,]([a-z0-9-]+):"([^"]*)"', seg):
                zh_map[k] = v
    # 2) 站点对象：route:"/home/{key}" 向前找 name / i18nKey
    seen = set()
    for m in re.finditer(r'route:"/home/([a-z]+)"', js):
        key = m.group(1)
        if key in seen:
            continue
        seg = js[max(0, m.start() - 900):m.start()]
        names = re.findall(r'name:"([^"]+)"', seg)
        i18n = re.findall(r'i18nKey:"header\.([^"]+)"', seg)
        if not names:
            continue
        seen.add(key)
        sites.append({
            "key": key,
            "name": names[-1],
            "name_zh": zh_map.get(i18n[-1], "") if i18n else "",
            "home_url": HOME_PAGE_TPL.format(key=key),
            "search_url": SEARCH_PAGE_TPL.format(key=key),
        })
    return sites


def fallback_sites() -> list:
    return [{"key": k, "name": en, "name_zh": zh,
             "home_url": HOME_PAGE_TPL.format(key=k),
             "search_url": SEARCH_PAGE_TPL.format(key=k)}
            for k, en, zh in FALLBACK_SITES]


def fetch_site_list() -> list:
    """访问首页获取完整 HTML，再解析 JS bundle 得到 20 个站点。失败回退内置列表"""
    headers = {"User-Agent": UA, "Accept-Language": "zh-CN,zh;q=0.9"}
    try:
        res = requests.get(HOME_URL, headers=headers, timeout=30, proxies=PROXIES)
        res.raise_for_status()
        html = res.text
        log_print(f"[*] 首页HTML获取成功: {HOME_URL}（长度 {len(html)}）")
        m = re.search(r'src="(/assets/index-[^"]+\.js)"', html)
        if not m:
            log_print("[!] 首页HTML中未找到 JS bundle 引用，使用内置站点列表")
            return fallback_sites()
        js_url = urljoin(HOME_URL, m.group(1))
        res2 = requests.get(js_url, headers=headers, timeout=60, proxies=PROXIES)
        res2.raise_for_status()
        js = res2.text
        log_print(f"[*] JS bundle 获取成功: {js_url}（长度 {len(js)}）")
        sites = parse_sites(js)
        if not sites:
            log_print("[!] JS bundle 中未解析出站点，使用内置站点列表")
            return fallback_sites()
        log_print(f"[*] 已解析出 {len(sites)} 个站点:")
        for s in sites:
            zh = f"（{s['name_zh']}）" if s["name_zh"] else ""
            log_print(f"    {s['name']}{zh} [{s['key']}] -> {s['home_url']}")
        return sites
    except Exception as e:
        log_print(f"[!] 站点列表获取失败: {e}，使用内置站点列表")
        return fallback_sites()


# ==================== 字段处理 ====================

def cover_url(book_id: str) -> str:
    """按官网规律由书籍ID拼接封面链接：
    wplbn9000184970 -> https://download.worldlibrary.ai/wplbn/900/018/497/0/{id}/{id}_250.png"""
    digits = book_id.replace("wplbn", "")
    if len(digits) != 10 or not digits.isdigit():
        return ""
    return (f"https://download.worldlibrary.ai/wplbn/"
            f"{digits[0:3]}/{digits[3:6]}/{digits[6:9]}/{digits[9:]}/"
            f"{book_id}/{book_id}_250.png")


def author_str(src: dict) -> str:
    author = src.get("author")
    if isinstance(author, list):
        return ", ".join(str(x) for x in author)
    return str(author) if author else ""


def has_cjk(text: str) -> bool:
    return bool(re.search(r'[一-鿿]', text))


def book_to_row(site_key: str, book_id: str, src: dict, title_zh: str) -> list:
    """单条书籍数据 → 模板一行（ISBN、SSN号、读秀号接口无字段，留空）"""
    title = src.get("title") or ""
    return [book_id,
            title_zh or title,
            author_str(src),
            src.get("publisher") or "",
            str(src.get("publication_year")) if src.get("publication_year") is not None else "",
            "",
            str(src.get("page")) if src.get("page") is not None else "",
            cover_url(book_id),
            BOOK_URL_TPL.format(key=site_key, book_id=book_id),
            "",
            ""]


# ==================== Excel 输出 + 跨运行去重 ====================

class BookStore:
    """每个站点一个"当前"Excel（output/worldlibrary_{站点key}.xlsx）：
    - 行数达到 MAX_ROWS_PER_FILE 时自动归档为 worldlibrary_{key}_{序号}.xlsx 并另开新文件
    - 跨文件去重用纯文本 worldlibrary_{key}_ids.txt（每行一个书籍ID，非数据库）：
      启动时流式读入内存；文件不存在时从当前Excel回填生成
    - 每累计 500 条落盘一次，落盘成功后同步把新ID追加到 ids.txt"""

    def __init__(self, site_key: str, flush_every: int = FLUSH_ROWS):
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        self.site_key = site_key
        self.flush_every = flush_every
        self.path = os.path.join(OUTPUT_DIR, f"worldlibrary_{site_key}.xlsx")
        self.ids_path = os.path.join(OUTPUT_DIR, f"worldlibrary_{site_key}_ids.txt")
        self.seen_ids = set()       # 已入库书籍ID（内存中以整数/字符串存储）
        self.pending_ids = []       # 本次新增待登记到 ids.txt 的ID

        header_ok = self._check_existing_header()
        # txt 注册表优先；无有效Excel时只读txt，避免把损坏文件内容当ID
        self._load_seen_ids(backfill_ok=header_ok and os.path.exists(self.path))
        if os.path.exists(self.path) and not header_ok:
            bad = self.path
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            self.path = os.path.join(OUTPUT_DIR, f"worldlibrary_{site_key}_{stamp}.xlsx")
            try:
                os.rename(bad, os.path.join(OUTPUT_DIR,
                                            f"worldlibrary_{site_key}_bad_{stamp}.xlsx"))
            except OSError:
                pass
            log_print(f"[!] 已有文件 {os.path.basename(bad)} 表头不匹配或打开失败，"
                      f"已改名备份，本次新建 {os.path.basename(self.path)}")
        if header_ok and os.path.exists(self.path):
            self.wb = openpyxl.load_workbook(self.path)   # 续写已有文件
            self.ws = self.wb.active
        else:
            self._create_new_workbook()

        self.row_count = self.ws.max_row - 1
        self.pending = 0
        self.added = 0      # 本次运行新增
        self.dup = 0        # 重复跳过（含历史已采集）
        self.non_book = 0   # 非 book 类型跳过
        if self.row_count >= MAX_ROWS_PER_FILE:
            self._rotate()  # 已有文件已超上限：先归档再续写

    @staticmethod
    def _id_key(book_id: str):
        """书籍ID的内存表示：wplbn+10位数字 → 整数（大幅省内存），其它格式保持字符串"""
        s = str(book_id).strip()
        if s.startswith("wplbn") and len(s) == 15 and s[5:].isdigit():
            return int(s[5:])
        return s

    def _load_seen_ids(self, backfill_ok: bool = True):
        """启动时加载历史ID：优先读 ids.txt（快）；txt 不存在/为空且Excel有效时，
        从当前Excel流式读取并把原始ID回填到txt"""
        if os.path.exists(self.ids_path) and os.path.getsize(self.ids_path) > 0:
            try:
                with open(self.ids_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            self.seen_ids.add(self._id_key(line))
                log_print(f"[*] 已从 {os.path.basename(self.ids_path)} "
                          f"加载 {len(self.seen_ids)} 条历史ID，本次续写去重")
                return
            except Exception as e:
                log_print(f"[!] 读取 {self.ids_path} 失败: {e}")
        if backfill_ok:
            # txt 丢失/为空时：扫描当前Excel + 所有归档/历史Excel（含旧版本带时间戳的文件，
            # 排除表头损坏的 _bad_ 备份），重建去重注册
            files = [self.path]
            for p in sorted(glob.glob(os.path.join(OUTPUT_DIR,
                                                   f"worldlibrary_{self.site_key}_*.xlsx"))):
                if p != self.path and "_bad_" not in os.path.basename(p):
                    files.append(p)
            raw_ids = []
            for fp in files:
                if not os.path.exists(fp):
                    continue
                try:
                    wb_ro = openpyxl.load_workbook(fp, read_only=True)
                    try:
                        rows = wb_ro.active.iter_rows(values_only=True)
                        next(rows, None)    # 跳过表头
                        for row in rows:
                            if row and row[0]:
                                raw_ids.append(str(row[0]).strip())
                    finally:
                        wb_ro.close()
                except Exception as e:
                    log_print(f"[!] 读取 {os.path.basename(fp)} 历史ID失败: {e}")
            for rid in raw_ids:
                self.seen_ids.add(self._id_key(rid))
            if raw_ids:
                with open(self.ids_path, "w", encoding="utf-8") as f:
                    for rid in raw_ids:
                        f.write(rid + "\n")
            log_print(f"[*] 已从站点Excel文件加载 {len(self.seen_ids)} 条历史ID，"
                      f"并生成去重文件 {os.path.basename(self.ids_path)}")

    def _check_existing_header(self) -> bool:
        """已有Excel表头与模板一致返回 True；文件不存在也返回 True（视为新建）"""
        if not os.path.exists(self.path):
            return True
        try:
            wb_ro = openpyxl.load_workbook(self.path, read_only=True)
            try:
                rows = wb_ro.active.iter_rows(values_only=True)
                header = next(rows, None)
                return list(header) == EXCEL_COLUMNS
            finally:
                wb_ro.close()
        except Exception:
            return False

    def _create_new_workbook(self):
        """新建带表头样式的空工作簿"""
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.title = "书籍数据"
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_align = Alignment(horizontal="center", vertical="center")
        for col_idx, col_name in enumerate(EXCEL_COLUMNS, 1):
            cell = self.ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align
        col_widths = [25, 45, 30, 30, 15, 20, 12, 60, 55, 15, 20]
        for i, width in enumerate(col_widths, 1):
            col_letter = openpyxl.utils.get_column_letter(i)
            self.ws.column_dimensions[col_letter].width = width

    def _rotate(self):
        """当前文件达到行数上限：归档为编号文件，另开新文件继续写"""
        seq = 1
        while os.path.exists(os.path.join(OUTPUT_DIR,
                                          f"worldlibrary_{self.site_key}_{seq:04d}.xlsx")):
            seq += 1
        arch = os.path.join(OUTPUT_DIR, f"worldlibrary_{self.site_key}_{seq:04d}.xlsx")
        try:
            self.wb.save(self.path)
            os.rename(self.path, arch)
            log_print(f"[滚动] {os.path.basename(self.path)} 已达 {MAX_ROWS_PER_FILE} 行上限，"
                      f"归档为 {os.path.basename(arch)}，另开新文件继续写入")
        except OSError as e:
            log_print(f"[!] 归档失败（{e}），继续写入当前文件")
            return
        self._create_new_workbook()
        self.row_count = 0
        self.pending = 0

    def add_book(self, item: dict, site_key: str):
        """单条 hit 入库：只取 book、翻译书名、按ID去重。"""
        src = item.get("source") or {}
        if src.get("document_type") != "book":
            self.non_book += 1
            return
        book_id = str(item.get("id") or "").strip()
        if not book_id:
            return
        if self._id_key(book_id) in self.seen_ids:
            self.dup += 1
            return
        self.seen_ids.add(self._id_key(book_id))
        self.pending_ids.append(book_id)
        title = src.get("title") or ""
        title_zh = ""
        if title and not has_cjk(title):
            title_zh = translate_title(site_key, book_id, title)
            interval_sleep(TRANSLATE_INTERVAL)
        self.append(book_to_row(site_key, book_id, src, title_zh))
        self.added += 1

    def append(self, row: list):
        self.ws.append(row)
        for cell in self.ws[self.ws.max_row]:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        self.row_count += 1
        self.pending += 1
        if self.pending >= self.flush_every or self.row_count >= MAX_ROWS_PER_FILE:
            self.save()
            self.pending = 0
            if self.row_count >= MAX_ROWS_PER_FILE:
                self._rotate()
            else:
                log_print(f"[落盘] 已保存 {self.row_count} 条")

    def _flush_ids(self):
        """把待登记的新ID追加写入 ids.txt（Excel 落盘成功后才调用，保证不登记未持久化的ID）"""
        if self.pending_ids:
            with open(self.ids_path, "a", encoding="utf-8") as f:
                for bid in self.pending_ids:
                    f.write(bid + "\n")
            self.pending_ids = []

    def save(self):
        try:
            self.wb.save(self.path)
        except PermissionError:
            log_print(f"[!] 保存 Excel 失败：文件被占用（请关闭 Excel/WPS 中打开的同名文件），"
                      f"稍后下次落盘会自动重试")
            return
        except Exception as e:
            log_print(f"[!] 保存 Excel 失败: {e}")
            return
        self._flush_ids()

    def close(self):
        self.save()


# ==================== 采集主流程 ====================

def process_hits(site_key: str, hits: list, store: BookStore):
    for item in hits:
        check_stop()
        store.add_book(item, site_key)


def paginate_chunk(site_key: str, query: str, store: BookStore,
                   first_page: dict = None, label: str = "",
                   filters: dict = None) -> bool:
    """对一个分片（数量应 <= WINDOW_LIMIT）完整翻页采集。返回是否完整完成"""
    if first_page is None:
        first_page = search_page(site_key, query=query, filters=filters)
        interval_sleep(REQUEST_INTERVAL)
    if first_page is None:
        log_print(f"[!] {label} 首次请求失败，该分片跳过")
        return False
    total = first_page.get("total") or 0
    if total > WINDOW_LIMIT:
        log_print(f"[!] {label} 分片总量 {total} 超过单查询上限，仅采集前 {WINDOW_LIMIT} 条")
        total = WINDOW_LIMIT
    if total == 0:
        return True
    pages = math.ceil(total / PAGE_SIZE)
    verbose = pages <= 10   # 小分片每页都打日志（每页含翻译，耗时约1分钟，避免看起来卡住）
    hits1 = first_page.get("hits") or []
    if verbose:
        log_print(f"[采集] {label} 第 1/{pages} 页 本页{len(hits1)}条 处理中...")
    process_hits(site_key, hits1, store)
    if verbose:
        log_print(f"[采集] {label} 第 1/{pages} 页 完成 累计新增{store.added}条 重复{store.dup}条")
    if pages == 1:
        return True
    for pg in range(2, pages + 1):
        check_stop()
        j = search_page(site_key, query=query, from_=(pg - 1) * PAGE_SIZE, filters=filters)
        if j is None:
            log_print(f"[!] {label} 第 {pg}/{pages} 页多次重试仍失败，该分片中断（已采集数据已保存）")
            return False
        hits = j.get("hits") or []
        if verbose:
            log_print(f"[采集] {label} 第 {pg}/{pages} 页 本页{len(hits)}条 处理中...")
        process_hits(site_key, hits, store)
        if verbose:
            log_print(f"[采集] {label} 第 {pg}/{pages} 页 完成 累计新增{store.added}条 重复{store.dup}条")
        elif pg % 10 == 0 or pg == pages:
            log_print(f"[采集] {label} 第 {pg}/{pages} 页 累计新增{store.added}条 重复{store.dup}条")
        interval_sleep(REQUEST_INTERVAL)
    return True


def _count_query(site_key: str, query: str, filters: dict = None):
    """查询某个 query 的总条数（顺带返回第 1 页数据供翻页复用）。失败返回 None"""
    j = search_page(site_key, query=query, filters=filters)
    interval_sleep(REQUEST_INTERVAL)
    return j


# 分片递归进度统计
_split_probe_count = 0


def _range_query(lo: str, hi: str) -> str:
    return f"title:[* TO {hi}]" if lo == "*" else f"title:[{lo} TO {hi}]"


def _probe(site_key: str, query: str, store: BookStore, filters: dict = None):
    """探测一个查询的总数，并把第 1 页的 20 条直接入库（去重防重复）。
    返回 (cnt, first_page)；查询失败返回 (None, None)"""
    global _split_probe_count
    j = _count_query(site_key, query, filters=filters)
    _split_probe_count += 1
    if _split_probe_count % 50 == 0:
        log_print(f"[分片] 已探测 {_split_probe_count} 个查询")
    if j is None:
        return None, None
    cnt = j.get("total") or 0
    if cnt:
        process_hits(site_key, j.get("hits") or [], store)
    return cnt, j


def split_prefix_level(site_key: str, p: str, store: BookStore, top_hi: str,
                       depth: int = 0, filters: dict = None) -> bool:
    """按前缀 p 枚举采集：[p0 TO p1]...[pz TO top_hi] 共 36 块，
    覆盖 p 开头的多字符词条 + 边界词条（无缺口）。超限的块递归到下一层字符。
    尾部块 [pz TO top_hi] 包含边界词条（其书由 top_hi 自己的区间覆盖），
    递归到深处仍未拆开时直接跳过，避免大量重复翻页"""
    check_stop()
    for i, c in enumerate(ALNUM):
        nxt = p + ALNUM[i + 1] if i + 1 < len(ALNUM) else top_hi
        q = f"title:[{p}{c} TO {nxt}]"
        cnt, j = _probe(site_key, q, store, filters=filters)
        if cnt is None:
            log_print(f"[!] 分片 {q} 查询失败，跳过")
            continue
        if cnt == 0:
            continue
        if cnt <= WINDOW_LIMIT:
            paginate_chunk(site_key, q, store, first_page=j, label=q, filters=filters)
        elif i + 1 < len(ALNUM):
            split_prefix_level(site_key, p + c, store, nxt, depth + 1, filters)
        elif depth < 20:
            # 尾部 [pz TO top_hi]：继续按下一层字符拆分（剩余主体为边界词条的书，
            # 它们已由 top_hi 区间覆盖；拆到最深层仍超限时跳过）
            split_prefix_level(site_key, p + c, store, top_hi, depth + 1, filters)
        else:
            log_print(f"[*] 分片 {q} 拆分到最深层仍超上限（{cnt} 条），"
                      f"其中边界词条书籍已由其它分片覆盖，跳过")
    return True


def split_range(site_key: str, lo: str, hi: str, store: BookStore,
                filters: dict = None) -> bool:
    """范围分片 [lo TO hi]（单字符边界，lo 可为 "*" 开区间），无缺口递归拆分：
    - 单字符范围按字符（或 ALNUM 下标）二分
    - 相邻字符 [x TO y] = 裸词条x + x开头多字符词条 + 边界词条y，用前缀枚举完整覆盖"""
    check_stop()
    q = _range_query(lo, hi)
    cnt, j = _probe(site_key, q, store, filters=filters)
    if cnt is None:
        log_print(f"[!] 范围 {q} 查询失败，跳过")
        return False
    if cnt == 0:
        return True
    if cnt <= WINDOW_LIMIT:
        return paginate_chunk(site_key, q, store, first_page=j, label=q, filters=filters)
    if lo == "*":
        return split_range(site_key, "0", hi, store, filters=filters)
    if ord(hi) - ord(lo) >= 2:
        if lo in ALNUM and hi in ALNUM:
            m = ALNUM[(ALNUM.index(lo) + ALNUM.index(hi)) // 2]
            return (split_range(site_key, lo, m, store, filters=filters)
                    and split_range(site_key, ALNUM[ALNUM.index(m) + 1], hi, store,
                                    filters=filters))
        mid = chr((ord(lo) + ord(hi)) // 2)
        return (split_range(site_key, lo, mid, store, filters=filters)
                and split_range(site_key, chr(ord(mid) + 1), hi, store, filters=filters))
    # 相邻字符 [x TO y]
    c0, j0 = _probe(site_key, f"title:[{lo} TO {lo}]", store, filters=filters)
    if c0:
        if c0 <= WINDOW_LIMIT:
            paginate_chunk(site_key, f"title:[{lo} TO {lo}]", store, first_page=j0,
                           label=f"title:{lo}", filters=filters)
        else:
            log_print(f"[*] 词条[{lo}]单独覆盖 {c0} 本（超过单查询上限），"
                      f"这些书将依赖其它词条分片覆盖，跳过")
    if lo in ALNUM:
        return split_prefix_level(site_key, lo, store, hi, filters=filters)
    # 非ASCII相邻区间：前缀兜底
    q2 = f"title:{lo}*"
    c2, j2 = _probe(site_key, q2, store, filters=filters)
    if c2 is None:
        return False
    if c2 <= WINDOW_LIMIT:
        return paginate_chunk(site_key, q2, store, first_page=j2, label=q2, filters=filters)
    # 超限：说明该字符被分析器折叠（如 À→a），对应书籍已由 a-z 等区间覆盖，跳过
    log_print(f"[*] 词条 {lo} 开头查询到 {c2} 本（超过单查询上限），"
              f"这些书应已由其它分片覆盖，跳过")
    return True


def _collect_query(site_key: str, filters: dict, store: BookStore, label: str,
                   page1: dict = None) -> bool:
    """对"站点+年份(可选)"的一个查询范围完整采集：数量不超限则直接翻页，
    超限则按标题词条范围递归分片。返回 True 表示完成"""
    if page1 is None:
        page1 = search_page(site_key, filters=filters)
        interval_sleep(REQUEST_INTERVAL)
    if page1 is None:
        log_print(f"[!] {label} 首次请求失败，跳过")
        return False
    total = page1.get("total") or 0
    if total == 0:
        log_print(f"[*] {label}: 无数据，跳过")
        return True
    if total <= WINDOW_LIMIT:
        log_print(f"[*] {label}: 共 {total} 本，{math.ceil(total / PAGE_SIZE)} 页，直接翻页采集")
        return paginate_chunk(site_key, "*", store, first_page=page1, label=label,
                              filters=filters)
    log_print(f"[*] {label}: 共 {total} 本，超过接口单查询上限（{WINDOW_LIMIT}），"
              f"按书名标题词条范围递归分片采集")
    ok = True
    for lo, hi in TOP_TITLE_RANGES:
        ok &= split_range(site_key, lo, hi, store, filters=filters)
    return ok


def collect_site(site: dict, year_start: int = None, year_end: int = None):
    """采集一个站点。year_start/year_end 为 None 表示全站采集；
    否则按年份从 year_start 到 year_end 一年一年请求"""
    global _split_probe_count
    _split_probe_count = 0
    key = site["key"]
    zh = f"（{site['name_zh']}）" if site.get("name_zh") else ""
    year_info = (f"  年份范围: {year_start} ~ {year_end}（按年逐次请求）"
                 if year_start is not None else "  年份范围: 不限（全站采集）")
    log_print("\n" + "=" * 65)
    log_print(f"  📚 开始采集站点: {site['name']}{zh} [{key}]")
    log_print(f"  搜索接口: {SEARCH_API}")
    log_print(f"  Referer: {site['search_url']}")
    log_print(year_info)
    log_print(f"  每页 {PAGE_SIZE} 条  间隔 {REQUEST_INTERVAL} 秒  只取 document_type=book")
    log_print("=" * 65 + "\n")

    store = BookStore(key)
    try:
        if year_start is None:
            _collect_query(key, None, store, site['name'])
        else:
            for y in range(year_start, year_end + 1):
                check_stop()
                f = {"publication_year_from": y, "publication_year_to": y}
                log_print(f"\n===== 年份 {y} =====")
                _collect_query(key, f, store, f"{y}年")
    finally:
        store.close()

    log_print("\n" + "=" * 65)
    log_print(f"[+] 采集结束：本次新增 {store.added} 条，跳过重复 {store.dup} 条，"
              f"非book类型 {store.non_book} 条，文件累计 {store.row_count} 条")
    log_print(f"[+] 输出文件: {os.path.abspath(store.path)}")
    log_print("=" * 65 + "\n")


# ==================== GUI 界面 ====================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("worldlibrary.ai 电子书信息采集工具")
        self.root.geometry("860x680")
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        input_frame = ttk.Frame(root, padding=10)
        input_frame.pack(fill=tk.X)

        ttk.Label(input_frame, text="站点名称:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.site_var = tk.StringVar()
        self.site_combo = ttk.Combobox(input_frame, textvariable=self.site_var,
                                       state="readonly", width=52)
        self.site_combo.grid(row=0, column=1, padx=5)
        self.refresh_btn = ttk.Button(input_frame, text="刷新站点列表", command=self.load_sites)
        self.refresh_btn.grid(row=0, column=2, padx=5)

        ttk.Label(input_frame, text="开始年份:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.year_start_var = tk.StringVar(value="")
        ttk.Entry(input_frame, textvariable=self.year_start_var, width=8).grid(
            row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(input_frame, text="结束年份:").grid(row=1, column=2, sticky=tk.W, pady=3)
        self.year_end_var = tk.StringVar(value="")
        ttk.Entry(input_frame, textvariable=self.year_end_var, width=8).grid(
            row=1, column=3, sticky=tk.W, padx=5)
        ttk.Label(input_frame, text="(格式: YYYY 如 2017；留空=全站)").grid(
            row=1, column=4, sticky=tk.W, padx=5)

        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=2, column=0, columnspan=5, pady=8)
        self.start_btn = ttk.Button(btn_frame, text="开始执行", command=self.start_task)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(btn_frame, text="结束执行", command=self.stop_task,
                                   state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        ttk.Label(root, text="运行日志:", font=("Arial", 10, "bold")).pack(
            anchor=tk.W, padx=10, pady=(10, 0))
        self.log_text = scrolledtext.ScrolledText(root, height=28, state=tk.NORMAL, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.sites = []
        self.site_map = {}
        self.running = False
        self.worker_thread = None
        self.consume_log_queue()
        self.load_sites()   # 启动时后台获取站点列表

    def consume_log_queue(self):
        try:
            while True:
                msg = LOG_QUEUE.get_nowait()
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
        except Empty:
            pass
        self.root.after(50, self.consume_log_queue)

    def load_sites(self):
        if self.running:
            return
        log_print("\n[*] 正在访问首页获取站点列表...")
        self.refresh_btn.config(state=tk.DISABLED)

        def worker():
            sites = fetch_site_list()
            self.root.after(0, lambda: self._apply_sites(sites))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_sites(self, sites):
        self.sites = sites
        self.site_map = {}
        labels = []
        for s in sites:
            zh = f"{s['name_zh']} " if s.get("name_zh") else ""
            label = f"{zh}({s['name']}) [{s['key']}]"
            self.site_map[label] = s
            labels.append(label)
        self.site_combo["values"] = labels
        if labels:
            self.site_combo.current(0)
        self.refresh_btn.config(state=tk.NORMAL)
        log_print(f"[*] 站点列表加载完成，共 {len(sites)} 个站点\n")

    def on_close(self):
        self.stop_task()
        self.root.destroy()

    def validate_years(self):
        """校验年份输入。合法返回 (start, end)（可为 (None, None) 表示不限），非法返回 None"""
        s = self.year_start_var.get().strip()
        e = self.year_end_var.get().strip()
        if not s and not e:
            return None, None
        if not s or not e:
            log_print("[!] 错误：开始年份与结束年份需同时填写（都留空表示不限年份）")
            return None
        if not re.match(r"^\d{4}$", s) or not re.match(r"^\d{4}$", e):
            log_print("[!] 错误：年份格式应为 YYYY（如 2017）")
            return None
        ys, ye = int(s), int(e)
        if ys > ye:
            log_print("[!] 错误：开始年份不能晚于结束年份")
            return None
        if ys < 0 or ye > 9999:
            log_print("[!] 错误：年份超出范围（0-9999）")
            return None
        if ye - ys > 100:
            log_print(f"[*] 提示：年份跨度 {ye - ys + 1} 年，将按年逐次请求，耗时较长")
        return ys, ye

    def start_task(self):
        if self.running:
            return
        site = self.site_map.get(self.site_var.get())
        if not site:
            log_print("[!] 错误：请先选择要采集的站点名称")
            return
        years = self.validate_years()
        if years is None:
            return
        self.running = True
        _stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.refresh_btn.config(state=tk.DISABLED)

        self.worker_thread = threading.Thread(
            target=self.run_loop, args=(site, years[0], years[1]), daemon=True)
        self.worker_thread.start()

    def stop_task(self):
        if not self.running:
            return
        self.running = False
        _stop_event.set()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        log_print("\n[!] 已发送结束信号，正在中断当前操作...")

    def run_loop(self, site, year_start=None, year_end=None):
        try:
            collect_site(site, year_start, year_end)
        except TaskStoppedException:
            log_print("\n[!] 任务已被用户结束\n")
        except Exception as e:
            log_print(f"[!] 执行异常: {e}")
            import traceback
            log_print(traceback.format_exc())
        finally:
            self.running = False
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.refresh_btn.config(state=tk.NORMAL)
            log_print("\n[*] 本次执行结束\n")


def main():
    root = tk.Tk()
    app = App(root)
    if PROXIES:
        log_print(f"[*] 当前使用代理: {PROXIES['https']}")
    else:
        log_print("[!] 未检测到代理，将直连访问；若网络受限，请在配置区设置 PROXY_OVERRIDE，"
                  "例如 PROXY_OVERRIDE = \"http://127.0.0.1:7890\"（端口看你的代理客户端）")
    root.mainloop()


if __name__ == "__main__":
    main()
