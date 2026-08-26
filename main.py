from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel
import requests
import re
import concurrent.futures
from lxml import etree
from urllib.parse import urlparse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

app = FastAPI()

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SitemapScraper/1.0)"}
NS = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

PAGE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# Per-batch settings. One /api/robots call must finish well inside the
# hosting timeout: worst case is (batch_size / ROBOTS_WORKERS) * ROBOTS_TIMEOUT.
# At 80 URLs / 20 workers / 8s that is ~32s worst case, ~4s typical.
ROBOTS_WORKERS = 20
ROBOTS_TIMEOUT = 8
MAX_BATCH = 150              # hard cap on URLs per /api/robots call
MAX_HTML_BYTES = 250_000
SKIP_TYPES = ("image/", "video/", "audio/", "font/",
              "application/pdf", "application/zip")

# A single JSON response must stay under the platform's ~4.5MB body limit.
# 19.6k rows is ~5.8MB, so /api/urls is paginated.
MAX_PAGE = 6000

LANG_NAMES = {
    "en": "English", "fr": "French", "de": "German", "es": "Spanish",
    "zh": "Chinese", "ko": "Korean", "pt": "Portuguese", "ru": "Russian",
    "tr": "Turkish", "ja": "Japanese", "it": "Italian", "nl": "Dutch",
}


class AnalyzeRequest(BaseModel):
    url: str


class SheetsRequest(BaseModel):
    url: str
    categories: list[str]


class UrlsRequest(BaseModel):
    url: str
    type: str
    lang: str = "en"
    offset: int = 0
    limit: int = MAX_PAGE


class RobotsRequest(BaseModel):
    urls: list[str]


# --------------------------------------------------------------- sitemap I/O

def fetch_sitemap(url):
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    content = resp.content
    if url.endswith(".gz"):
        import gzip
        content = gzip.decompress(content)
    return content


def guess_type(url):
    filename = url.rstrip("/").split("/")[-1]
    match = re.match(r"sitemap[_-]?(.+)\.xml", filename, re.IGNORECASE)
    if match:
        raw = match.group(1).strip("-_ ")
        return raw.replace("-", " ").replace("_", " ").title()
    return filename


def guess_lang(url):
    path = urlparse(url).path
    lang_match = re.match(r"^/([a-z]{2})/", path)
    return lang_match.group(1) if lang_match else "en"


def guess_keyword(url):
    slug = urlparse(url).path.rstrip("/").split("/")[-1]
    slug = re.sub(r"\d+$", "", slug)
    return re.sub(r"[-_]+", " ", slug).title().strip()


def extract_urls(url):
    content = fetch_sitemap(url)
    tree = etree.fromstring(content)
    entries = []
    for u in tree.findall("sm:url", NS):
        loc = u.find("sm:loc", NS)
        lastmod = u.find("sm:lastmod", NS)
        changefreq = u.find("sm:changefreq", NS)
        priority = u.find("sm:priority", NS)
        page_url = loc.text.strip() if loc is not None else ""
        entries.append({
            "url": page_url,
            "keyword": guess_keyword(page_url) if page_url else "",
            "lastmod": lastmod.text.strip() if lastmod is not None else "",
            "changefreq": changefreq.text.strip() if changefreq is not None else "",
            "priority": priority.text.strip() if priority is not None else "",
        })
    return entries


def get_categories(sitemap_url):
    content = fetch_sitemap(sitemap_url)
    index_tree = etree.fromstring(content)
    root_tag = etree.QName(index_tree.tag).localname

    if root_tag == "sitemapindex":
        child_locs = [loc.text.strip()
                      for loc in index_tree.findall("sm:sitemap/sm:loc", NS)]
    else:
        child_locs = [sitemap_url]

    categories = {}
    for child_url in child_locs:
        stype = guess_type(child_url)
        lang = guess_lang(child_url)
        categories.setdefault(stype, []).append({"url": child_url, "lang": lang})
    return categories


def sheet_children(sitemap_url, stype, lang):
    """Child sitemap URLs making up one sheet, in a stable order."""
    cats = get_categories(sitemap_url)
    return sorted(i["url"] for i in cats.get(stype, []) if i["lang"] == lang)


def sheet_rows(sitemap_url, stype, lang):
    rows = []
    for child in sheet_children(sitemap_url, stype, lang):
        rows.extend(extract_urls(child))
    return rows


# ================================================================
# ROBOTS DIRECTIVE EXTRACTION
#
# The two patterns below are built by concatenating LT ("<") rather
# than written as literal tags. If this file is ever copied out of a
# browser that rendered it as HTML, a literal tag would be silently
# swallowed and every URL would report "Not set" with no error.
# Keep it written this way.
# ================================================================

LT = "<"
META_TAG_RE = re.compile(LT + r"meta\b[^>]*>", re.IGNORECASE)
HEAD_END_RE = re.compile(LT + r"/head\s*>", re.IGNORECASE)
ATTR_RE = re.compile(r"""([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))""")


def parse_meta_robots(html):
    """Content of the robots meta tag, falling back to the googlebot one."""
    found = {}
    for tag in META_TAG_RE.finditer(html):
        attrs = {}
        for m in ATTR_RE.finditer(tag.group(0)):
            attrs[m.group(1).lower()] = m.group(2) or m.group(3) or m.group(4) or ""
        name = attrs.get("name", "").strip().lower()
        if name in ("robots", "googlebot") and name not in found:
            found[name] = attrs.get("content", "").strip()
    return found.get("robots") or found.get("googlebot") or ""


def fetch_robots_directive(session, url):
    try:
        resp = session.get(url, headers=PAGE_HEADERS, timeout=ROBOTS_TIMEOUT,
                           stream=True, allow_redirects=True)
    except requests.RequestException as e:
        return f"Error: {type(e).__name__}"

    try:
        if resp.status_code != 200:
            return f"HTTP {resp.status_code}"

        header_val = resp.headers.get("X-Robots-Tag", "").strip()

        meta_val = ""
        ctype = resp.headers.get("Content-Type", "").lower()
        if not any(t in ctype for t in SKIP_TYPES):
            size, chunks = 0, []
            for chunk in resp.iter_content(8192):
                chunks.append(chunk)
                size += len(chunk)
                if size >= MAX_HTML_BYTES:
                    break
            html = b"".join(chunks).decode("utf-8", errors="ignore")
            head = HEAD_END_RE.split(html, maxsplit=1)[0]
            meta_val = parse_meta_robots(head)
    finally:
        resp.close()

    if meta_val and header_val:
        return f"{meta_val} (X-Robots-Tag: {header_val})"
    if meta_val:
        return meta_val
    if header_val:
        return f"X-Robots-Tag: {header_val}"
    return "Not set"


def derive_status(directive):
    d = (directive or "").lower()
    if d.startswith(("http ", "error", "not checked")):
        return "Unknown", "Unknown"
    if d == "not set" or not d:
        return "Index (default)", "Follow (default)"
    is_none = bool(re.search(r"\bnone\b", d))
    index = "Noindex" if ("noindex" in d or is_none) else "Index"
    follow = "Nofollow" if ("nofollow" in d or is_none) else "Follow"
    return index, follow


def check_batch(urls):
    """Directive for each URL in one batch, fetched in parallel."""
    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=ROBOTS_WORKERS, pool_maxsize=ROBOTS_WORKERS, max_retries=0
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    out = {}
    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=ROBOTS_WORKERS) as ex:
            futures = {ex.submit(fetch_robots_directive, session, u): u
                       for u in urls if u}
            for fut in concurrent.futures.as_completed(futures):
                u = futures[fut]
                try:
                    out[u] = fut.result()
                except Exception as e:
                    out[u] = f"Error: {type(e).__name__}"
    finally:
        session.close()
    return out


# --------------------------------------------------------------- endpoints

@app.post("/api/analyze")
def analyze(req: AnalyzeRequest):
    try:
        categories = get_categories(req.url)
        data = []
        for stype, items in categories.items():
            langs = sorted(set(i["lang"] for i in items))
            data.append({
                "type": stype,
                "count": len(items),
                "langs": ", ".join(LANG_NAMES.get(l, l.upper()) for l in langs),
            })
        return {"categories": data}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/sheets")
def sheets(req: SheetsRequest):
    """One entry per output sheet, with its total row count."""
    try:
        cats = get_categories(req.url)
        out = []
        for stype in req.categories:
            if stype not in cats:
                continue
            langs = sorted(set(i["lang"] for i in cats[stype]))
            for lang in langs:
                total = sum(len(extract_urls(c))
                            for c in sheet_children(req.url, stype, lang))
                label = LANG_NAMES.get(lang, lang.upper())
                name = stype if len(langs) == 1 and lang == "en" \
                    else f"{stype} - {label}"
                out.append({"name": name[:31], "type": stype,
                            "lang": lang, "total": total})
        return {"sheets": out}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/urls")
def urls(req: UrlsRequest):
    """One page of rows for a sheet."""
    try:
        limit = max(1, min(req.limit, MAX_PAGE))
        rows = sheet_rows(req.url, req.type, req.lang)
        return {"rows": rows[req.offset:req.offset + limit],
                "total": len(rows), "offset": req.offset}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/robots")
def robots(req: RobotsRequest):
    """Robots directive for a batch of URLs."""
    try:
        batch = req.urls[:MAX_BATCH]
        if not batch:
            return {"results": {}}
        return {"results": check_batch(batch)}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
