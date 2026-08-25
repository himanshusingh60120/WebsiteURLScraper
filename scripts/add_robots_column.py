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

# ---------------------------------------------------------------- robots cfg
PAGE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}
ROBOTS_WORKERS = 12          # parallel page fetches
ROBOTS_TIMEOUT = 10          # seconds per page
MAX_HTML_BYTES = 250_000     # stop downloading a page after this much HTML
SKIP_TYPES = ("image/", "video/", "audio/", "font/", "application/pdf",
              "application/zip")

# Serverless functions have a hard execution-time wall. Each URL is one HTTP
# request, so a 19k-URL sitemap cannot be checked inside a single request no
# matter how high maxDuration goes. This caps the work per sheet; anything
# beyond it is written as "Not checked (limit reached)". For a full run, use
# the standalone add_robots_column.py script instead.
ROBOTS_LIMIT_DEFAULT = 300

LANG_NAMES = {
    "en": "English", "fr": "French", "de": "German", "es": "Spanish",
    "zh": "Chinese", "ko": "Korean", "pt": "Portuguese", "ru": "Russian",
    "tr": "Turkish", "ja": "Japanese", "it": "Italian", "nl": "Dutch"
}


class AnalyzeRequest(BaseModel):
    url: str


class ScrapeRequest(BaseModel):
    url: str
    categories: list[str]
    check_robots: bool = True
    robots_limit: int = ROBOTS_LIMIT_DEFAULT


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
            "URL": page_url,
            "Keyword": guess_keyword(page_url) if page_url else "",
            "Last Modified": lastmod.text.strip() if lastmod is not None else "",
            "Change Freq": changefreq.text.strip() if changefreq is not None else "",
            "Priority": priority.text.strip() if priority is not None else "",
        })
    return entries


# =========================================================================
# ROBOTS DIRECTIVE EXTRACTION
# =========================================================================

META_TAG_RE = re.compile(r"<meta\b[^>]*>", re.IGNORECASE)
ATTR_RE = re.compile(r"""([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+))""")


def parse_meta_robots(html):
    """Content of <meta name="robots">, falling back to name="googlebot"."""
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
    """Raw robots directive for a single URL."""
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
        # Parse unless the type is clearly non-HTML. Skipping on anything that
        # merely fails a "html in ctype" test silently reports "Not set" for
        # servers that send an unusual or generic Content-Type.
        if not any(t in ctype for t in SKIP_TYPES):
            size, chunks = 0, []
            for chunk in resp.iter_content(8192):
                chunks.append(chunk)
                size += len(chunk)
                if size >= MAX_HTML_BYTES:
                    break
            html = b"".join(chunks).decode("utf-8", errors="ignore")
            head = re.split(r"</head\s*>", html, maxsplit=1, flags=re.IGNORECASE)[0]
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
    """(Index Status, Follow Status) from a raw directive string."""
    d = (directive or "").lower()
    if d.startswith(("http ", "error", "not checked")):
        return "Unknown", "Unknown"
    if d == "not set" or not d:
        return "Index (default)", "Follow (default)"
    is_none = bool(re.search(r"\bnone\b", d))
    index = "Noindex" if ("noindex" in d or is_none) else "Index"
    follow = "Nofollow" if ("nofollow" in d or is_none) else "Follow"
    return index, follow


def annotate_robots(rows, limit=ROBOTS_LIMIT_DEFAULT):
    """Fill Robots / Index Status / Follow Status on each row, in parallel."""
    targets = rows[:limit] if limit and limit > 0 else rows

    for row in rows[len(targets):]:
        row["Robots"] = "Not checked (limit reached)"
        row["Index Status"] = "Unknown"
        row["Follow Status"] = "Unknown"

    if not targets:
        return rows

    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(
        pool_connections=ROBOTS_WORKERS, pool_maxsize=ROBOTS_WORKERS, max_retries=0
    )
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    try:
        with concurrent.futures.ThreadPoolExecutor(max_workers=ROBOTS_WORKERS) as ex:
            futures = {
                ex.submit(fetch_robots_directive, session, r["URL"]): r
                for r in targets if r.get("URL")
            }
            for fut in concurrent.futures.as_completed(futures):
                row = futures[fut]
                try:
                    directive = fut.result()
                except Exception as e:
                    directive = f"Error: {type(e).__name__}"
                row["Robots"] = directive
                row["Index Status"], row["Follow Status"] = derive_status(directive)
    finally:
        session.close()

    return rows


# =========================================================================
# EXCEL OUTPUT
# =========================================================================

# Add "Index Status" here (and a width below) to also surface Index/Noindex —
# the value is already computed on every row.
HDR = ["URL", "Robots", "Follow Status", "Keyword",
       "Last Modified", "Change Freq", "Priority"]
COL_WIDTHS = [80, 30, 14, 45, 22, 14, 10]


def add_sheet(wb, name, rows):
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    HEADER_FILL = PatternFill("solid", fgColor="2F5496")
    HEADER_ALIGN = Alignment(horizontal="center", vertical="center")
    CELL_FONT = Font(name="Arial", size=10)
    BAD_FONT = Font(name="Arial", size=10, color="C00000", bold=True)
    WARN_FONT = Font(name="Arial", size=10, color="808080", italic=True)
    THIN_BORDER = Border(bottom=Side(style="thin", color="D9D9D9"))

    name = name[:31]
    ws = wb.create_sheet(title=name)

    for ci, h in enumerate(HDR, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.alignment = HEADER_FONT, HEADER_FILL, HEADER_ALIGN

    for ri, row in enumerate(rows, 2):
        for ci, key in enumerate(HDR, 1):
            val = row.get(key, "")
            c = ws.cell(row=ri, column=ci, value=val)
            c.border = THIN_BORDER
            if key in ("Robots", "Index Status", "Follow Status"):
                low = str(val).lower()
                if "nofollow" in low or "noindex" in low or re.search(r"\bnone\b", low):
                    c.font = BAD_FONT
                elif low.startswith(("error", "http ", "unknown", "not checked")):
                    c.font = WARN_FONT
                else:
                    c.font = CELL_FONT
            else:
                c.font = CELL_FONT

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    last_col = get_column_letter(len(HDR))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{last_col}{len(rows) + 1}"


def get_categories(sitemap_url):
    content = fetch_sitemap(sitemap_url)
    index_tree = etree.fromstring(content)
    root_tag = etree.QName(index_tree.tag).localname

    if root_tag == "sitemapindex":
        child_locs = [loc.text.strip() for loc in index_tree.findall("sm:sitemap/sm:loc", NS)]
    else:
        child_locs = [sitemap_url]

    categories = {}
    for child_url in child_locs:
        stype = guess_type(child_url)
        lang = guess_lang(child_url)
        categories.setdefault(stype, []).append({"url": child_url, "lang": lang})
    return categories


@app.post("/api/analyze")
def analyze(req: AnalyzeRequest):
    try:
        categories = get_categories(req.url)
        response_data = []
        for stype, items in categories.items():
            langs = sorted(set(item["lang"] for item in items))
            lang_labels = ", ".join(LANG_NAMES.get(l, l.upper()) for l in langs)
            response_data.append({
                "type": stype,
                "count": len(items),
                "langs": lang_labels
            })
        return {"categories": response_data}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/scrape")
def scrape(req: ScrapeRequest):
    try:
        all_categories = get_categories(req.url)
        scraped_data = {}

        for stype in req.categories:
            if stype in all_categories:
                for item in all_categories[stype]:
                    urls = extract_urls(item["url"])
                    if req.check_robots:
                        annotate_robots(urls, limit=req.robots_limit)
                    scraped_data[(stype, item["lang"])] = urls

        wb = Workbook()
        wb.remove(wb.active)

        types_in_data = sorted(set(t for t, l in scraped_data.keys()))
        for stype in types_in_data:
            langs = sorted([l for (t, l) in scraped_data if t == stype])
            for lang in langs:
                lang_label = LANG_NAMES.get(lang, lang.upper())
                sheet_name = stype if len(langs) == 1 and lang == "en" else f"{stype} - {lang_label}"
                add_sheet(wb, sheet_name, scraped_data[(stype, lang)])

        output = BytesIO()
        wb.save(output)
        excel_data = output.getvalue()

        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=sitemap_links.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
