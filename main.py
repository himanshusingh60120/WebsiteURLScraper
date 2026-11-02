from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import requests
import re
from lxml import etree
from lxml import html as lxml_html
from urllib.parse import urlparse, urljoin, urldefrag
from concurrent.futures import ThreadPoolExecutor

app = FastAPI()

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; SitemapScraper/1.0)"}
NS = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}

LANG_NAMES = {
    "en": "English", "fr": "French", "de": "German", "es": "Spanish",
    "zh": "Chinese", "ko": "Korean", "pt": "Portuguese", "ru": "Russian",
    "tr": "Turkish", "ja": "Japanese", "it": "Italian", "nl": "Dutch"
}

# Optional per-page fields: key sent by the frontend -> the Excel column(s) it fills
FIELD_COLUMNS = {
    "meta_title": ["Meta Title"],
    "meta_description": ["Meta Description"],
    "h1": ["H1"],
    "h2": ["H2"],
    "h3": ["H3"],
    "inlinks": ["Inlinks Count", "Inlinks"],
}
FIELD_ORDER = list(FIELD_COLUMNS)
HEADING_FIELDS = ("h1", "h2", "h3")
EXCEL_CELL_LIMIT = 32_000          # Excel caps a cell at 32,767 characters

MAX_WORKERS = 20                   # parallel page fetches inside one batch
PAGE_TIMEOUT = 10                  # seconds per page
MAX_HTML_BYTES = 300_000           # stop reading a page after this much HTML
MAX_HTML_BYTES_INLINKS = 1_000_000 # content links can sit far down the page
MAX_BATCH = 120                    # hard cap on URLs accepted per /api/pages call
MAX_INLINKS_PER_PAGE = 1_000

SESSION = requests.Session()
SESSION.mount("https://", requests.adapters.HTTPAdapter(pool_maxsize=MAX_WORKERS))
SESSION.mount("http://", requests.adapters.HTTPAdapter(pool_maxsize=MAX_WORKERS))


class AnalyzeRequest(BaseModel):
    url: str


class UrlsRequest(BaseModel):
    url: str
    categories: list[str]


class PagesRequest(BaseModel):
    urls: list[str]
    fields: list[str] = []


# ---------------------------------------------------------------------------
# Sitemap parsing
# ---------------------------------------------------------------------------

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


def extract_urls(url):
    content = fetch_sitemap(url)
    tree = etree.fromstring(content)
    entries = []
    for u in tree.findall("sm:url", NS):
        loc = u.find("sm:loc", NS)
        lastmod = u.find("sm:lastmod", NS)
        changefreq = u.find("sm:changefreq", NS)
        priority = u.find("sm:priority", NS)
        entries.append({
            "URL": loc.text.strip() if loc is not None else "",
            "Last Modified": lastmod.text.strip() if lastmod is not None else "",
            "Change Freq": changefreq.text.strip() if changefreq is not None else "",
            "Priority": priority.text.strip() if priority is not None else "",
        })
    return entries


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


# ---------------------------------------------------------------------------
# Page-level extraction helpers
# ---------------------------------------------------------------------------

def clean(text):
    return re.sub(r"\s+", " ", (text or "")).strip()


def first_value(values):
    for value in values:
        cleaned = clean(value)
        if cleaned:
            return cleaned
    return ""


def empty_columns(fields):
    return {col: "" for f in fields for col in FIELD_COLUMNS[f]}


# ---------------------------------------------------------------------------
# Inlinks inside the content body
# ---------------------------------------------------------------------------

_LOWER = "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'"


def _class_token(name):
    return f"//*[contains(concat(' ', normalize-space(@class), ' '), ' {name} ')]"


# Most specific first. For each selector, the candidate with the most text wins.
CONTENT_CLASS_NAMES = [
    "entry-content", "post-content", "article-content", "article-body", "post-body",
    "blog-content", "single-content", "content-body", "w-richtext", "rich-text",
]
CONTENT_SELECTORS = (
    ["//*[@itemprop='articleBody']"]
    + [_class_token(n) for n in CONTENT_CLASS_NAMES]
    + ["//article", "//main", "//*[@role='main']"]
)
MIN_CONTENT_CHARS = 200

# Always removed from the content node.
SEMANTIC_BOILERPLATE = (
    ".//nav | .//header | .//footer | .//aside | .//script | .//style | .//noscript"
    " | .//template | .//*[@role='navigation'] | .//*[@role='banner']"
    " | .//*[@role='contentinfo'] | .//*[@role='complementary']"
)

# Removed only if they don't hold most of the content (guards against
# page-wide wrappers like ASP.NET's <form> or class="sidebar-layout").
BOILERPLATE_HINTS = [
    "breadcrumb", "sidebar", "widget", "share", "sharing", "social", "related",
    "comment", "author-bio", "author-box", "newsletter", "cookie", "pagination",
    "post-nav", "menu", "skip-link",
]


def _token_prefix(attr, hint):
    # matches a class/id token that *starts with* the hint, e.g. "related-posts"
    return f"contains(concat(' ', normalize-space(translate(@{attr}, {_LOWER}))), ' {hint}')"


GUARDED_BOILERPLATE = ".//form | .//*[" + " or ".join(
    _token_prefix(attr, h) for attr in ("class", "id") for h in BOILERPLATE_HINTS
) + "]"

SKIP_SCHEMES = ("mailto:", "tel:", "javascript:", "sms:", "data:", "ftp:")


def _site_host(url):
    host = (urlparse(url).hostname or "").lower()
    return host[4:] if host.startswith("www.") else host


def _page_key(url):
    p = urlparse(urldefrag(url)[0])
    return (_site_host(url), p.path.rstrip("/") or "/", p.query)


def _find_content_node(doc):
    for selector in CONTENT_SELECTORS:
        nodes = doc.xpath(selector)
        if not nodes:
            continue
        best = max(nodes, key=lambda n: len(n.text_content()))
        if len(clean(best.text_content())) >= MIN_CONTENT_CHARS:
            return best
    body = doc.find(".//body")
    return body if body is not None else doc


def _strip_boilerplate(node):
    for el in node.xpath(SEMANTIC_BOILERPLATE):
        if el.getparent() is not None:
            el.drop_tree()

    total = len(node.text_content()) or 1
    for el in node.xpath(GUARDED_BOILERPLATE):
        if el.getparent() is None:
            continue
        if len(el.text_content()) > total * 0.6:
            continue
        el.drop_tree()


def extract_inlinks(doc, page_url):
    base = page_url
    base_href = doc.xpath("//base/@href")
    if base_href and base_href[0].strip():
        base = urljoin(page_url, base_href[0].strip())

    site = _site_host(page_url)
    self_key = _page_key(page_url)

    node = _find_content_node(doc)
    _strip_boilerplate(node)

    links = []
    for a in node.xpath(".//a[@href]"):
        href = (a.get("href") or "").strip()
        if not href or href.startswith("#") or href.lower().startswith(SKIP_SCHEMES):
            continue
        try:
            target = urldefrag(urljoin(base, href))[0]
            if urlparse(target).scheme not in ("http", "https"):
                continue
            if _site_host(target) != site or _page_key(target) == self_key:
                continue
        except ValueError:
            continue

        anchor = clean(a.text_content())
        if not anchor and a.xpath(".//img"):
            anchor = f"[image] {first_value(a.xpath('.//img/@alt'))}".strip()

        links.append({
            "url": target,
            "anchor": anchor[:500],
            "rel": clean(a.get("rel")),
        })
        if len(links) >= MAX_INLINKS_PER_PAGE:
            break
    return links


# ---------------------------------------------------------------------------
# Page parsing
# ---------------------------------------------------------------------------

def parse_page_fields(html_bytes, fields, page_url):
    data = empty_columns(fields)
    if not html_bytes:
        return data

    doc = lxml_html.fromstring(html_bytes)

    if "meta_title" in fields:
        title = first_value(doc.xpath("//head/title//text()"))
        if not title:
            title = first_value(doc.xpath("//title//text()"))
        if not title:
            title = first_value(doc.xpath("//meta[@property='og:title']/@content"))
        data["Meta Title"] = title

    if "meta_description" in fields:
        desc = first_value(doc.xpath(
            "//meta[translate(@name, 'DESCRIPTION', 'description')='description']/@content"
        ))
        if not desc:
            desc = first_value(doc.xpath("//meta[@property='og:description']/@content"))
        data["Meta Description"] = desc

    for tag in HEADING_FIELDS:
        if tag in fields:
            headings = [clean(h.text_content()) for h in doc.xpath(f"//{tag}")]
            joined = " | ".join(h for h in headings if h)
            data[tag.upper()] = joined[:EXCEL_CELL_LIMIT]

    # Runs last: it removes nav/header/footer etc. from the parsed tree.
    if "inlinks" in fields:
        links = extract_inlinks(doc, page_url)
        unique = list(dict.fromkeys(link["url"] for link in links))
        data["Inlinks Count"] = len(unique)
        data["Inlinks"] = " | ".join(unique)[:EXCEL_CELL_LIMIT]
        data["_inlinks"] = links   # per-link detail for the "Content Inlinks" sheet

    return data


def fetch_page(url, fields):
    url = (url or "").strip()
    result = empty_columns(fields)
    result["Status"] = ""
    if not url:
        result["Status"] = "No URL"
        return result

    max_bytes = MAX_HTML_BYTES_INLINKS if "inlinks" in fields else MAX_HTML_BYTES

    try:
        resp = SESSION.get(
            url, headers=HEADERS, timeout=PAGE_TIMEOUT,
            stream=True, allow_redirects=True
        )
        result["Status"] = str(resp.status_code)
        content_type = resp.headers.get("Content-Type", "").lower()
        if resp.status_code != 200 or "html" not in content_type:
            resp.close()
            return result

        buffer = bytearray()
        for chunk in resp.iter_content(16384):
            buffer.extend(chunk)
            if len(buffer) >= max_bytes:
                break
        final_url = resp.url
        resp.close()

        result.update(parse_page_fields(bytes(buffer), fields, final_url))
    except Exception as exc:
        result["Status"] = f"Error: {type(exc).__name__}"
    return result


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------

@app.post("/api/analyze")
def analyze(req: AnalyzeRequest):
    """List sitemap categories so the user can pick."""
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


@app.post("/api/urls")
def urls(req: UrlsRequest):
    """Return every URL from the chosen sitemap categories, grouped into sheets."""
    try:
        all_categories = get_categories(req.url)

        collected = {}
        for stype in req.categories:
            for item in all_categories.get(stype, []):
                key = (stype, item["lang"])
                collected.setdefault(key, []).extend(extract_urls(item["url"]))

        sheets = []
        for stype in sorted(set(t for t, _ in collected)):
            langs = sorted(l for (t, l) in collected if t == stype)
            for lang in langs:
                lang_label = LANG_NAMES.get(lang, lang.upper())
                name = stype if len(langs) == 1 and lang == "en" else f"{stype} - {lang_label}"
                sheets.append({"name": name, "rows": collected[(stype, lang)]})

        total = sum(len(s["rows"]) for s in sheets)
        return {"sheets": sheets, "total": total}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/pages")
def pages(req: PagesRequest):
    """Fetch one batch of pages. Used by both sitemap mode and pasted-URL mode."""
    fields = [f for f in FIELD_ORDER if f in req.fields]
    if not fields:
        raise HTTPException(status_code=400, detail="No page fields requested")
    if len(req.urls) > MAX_BATCH:
        raise HTTPException(status_code=400, detail=f"Batch too large (max {MAX_BATCH})")
    if not req.urls:
        return {"results": []}

    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(req.urls))) as pool:
        results = list(pool.map(lambda u: fetch_page(u, fields), req.urls))
    return {"results": results}
