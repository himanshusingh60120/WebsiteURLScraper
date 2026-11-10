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
EXCEL_CELL_LIMIT = 32_000  # Excel caps a cell at 32,767 characters

MAX_WORKERS = 20                 # parallel page fetches inside one batch
PAGE_TIMEOUT = 10                # seconds per page
MAX_HTML_BYTES = 300_000         # stop reading a page after this much HTML
MAX_HTML_BYTES_INLINKS = 1_000_000  # content links can sit far down the page
MAX_BATCH = 120                  # hard cap on URLs accepted per /api/pages call
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


# ===========================================================================
# Inlinks inside the content body
#
# How a link ends up in the "Inlinks" column:
#   1. Code (scripts, styles, svg) is removed, then page chrome (nav, header,
#      footer, aside, dialogs, tab bars, menus) is removed from the whole page.
#   2. The main content block is located: per-site rule -> known content
#      markers -> <article> -> <main> -> text-density scoring.
#   3. Boilerplate inside that block (bylines, author boxes, share bars, tabs,
#      tags, related posts, comments, CTAs, ...) is removed.
#   4. Each remaining same-site link is checked against URL rules: other views
#      of the same page, author/profile/utility pages, per-site rules.
# Links dropped in steps 3 and 4 are returned with a reason so they appear in
# the "Excluded Inlinks" sheet and every result can be audited.
# ===========================================================================

# Per-site tweaks. Key = domain without "www." (subdomains match too).
#   content: class names (or XPath expressions) of the main content block,
#            tried before the generic detection.
#   exclude: extra path segments whose pages are never content interlinks.
# New sites normally need no entry. Add one only if the Excluded Inlinks
# sheet or a manual check shows a problem.
SITE_RULES = {
    "kingsresearch.com": {
        "content": ["report-content"],
        "exclude": ["toc", "snapshot", "license-variant"],
    },
}

_LOWER = "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'"


def _class_token(name):
    return f"//*[contains(concat(' ', normalize-space(@class), ' '), ' {name} ')]"


def _token_prefix(attr, hint):
    # matches a class/id token that *starts with* the hint, e.g. "related-posts"
    return f"contains(concat(' ', normalize-space(translate(@{attr}, {_LOWER}))), ' {hint}')"


def _as_xpath(selector):
    selector = selector.strip()
    return selector if selector.startswith(("/", "(", ".")) else _class_token(selector)


# Known content-body markers from common CMSs and site builders, most specific first.
CONTENT_CLASS_NAMES = [
    "entry-content", "post-content", "article-content", "article-body", "article__body",
    "article__content", "post-body", "post__content", "blog-content", "blog-post-content",
    "single-content", "content-body", "story-body", "td-post-content",
    "elementor-widget-theme-post-content", "w-richtext", "rich-text", "richtext",
    "markdown-body", "report-content", "prose",
]
# (xpath, refine): refine=True means the container is usually page-wide, so the
# densest text block inside it is used instead of the whole container.
CONTENT_SELECTORS = (
    [("//*[@itemprop='articleBody']", False)]
    + [(_class_token(n), False) for n in CONTENT_CLASS_NAMES]
    + [("//article", False), ("//main", True), ("//*[@role='main']", True)]
)
MIN_CONTENT_CHARS = 200
SECONDARY_BLOCK_SHARE = 0.25  # other matches of a selector are kept if >= 25% of the largest
GUARD_SHARE = 0.6             # never remove a block holding more than 60% of the text

# Always removed from the whole page: code, not content.
NON_CONTENT_XPATH = (
    "//script | //style | //noscript | //template | //svg | //iframe | //object | //embed"
)

# Removed from the whole page unless they hold most of the text
# (guards against sites that wrap everything in <header> or <aside>).
PAGE_CHROME_XPATH = (
    ".//nav | .//header | .//footer | .//aside | .//dialog"
    " | .//*[@role='navigation'] | .//*[@role='banner'] | .//*[@role='contentinfo']"
    " | .//*[@role='complementary'] | .//*[@role='search'] | .//*[@role='dialog']"
    " | .//*[@role='tablist'] | .//*[@role='menu'] | .//*[@role='menubar']"
)

# Removed from inside the content block, guarded the same way.
BOILERPLATE_HINTS = [
    "breadcrumb", "sidebar", "widget", "share", "sharing", "social", "related",
    "comment", "author", "byline", "reviewer", "post-meta", "entry-meta",
    "article-meta", "meta-info", "tags", "tag-list", "newsletter", "subscribe",
    "cookie", "consent", "pagination", "pager", "post-nav", "prev-next", "menu",
    "skip-link", "tabs", "toc", "table-of-contents", "cta", "promo", "popup", "modal",
]
GUARDED_BOILERPLATE = ".//form | .//*[" + " or ".join(
    _token_prefix(attr, h) for attr in ("class", "id") for h in BOILERPLATE_HINTS
) + "]"

SKIP_SCHEMES = ("mailto:", "tel:", "javascript:", "sms:", "data:", "ftp:")

# A link whose path contains one of these exact segments is never a content
# interlink (author/reviewer profiles, account and checkout pages, feeds...).
UTILITY_SEGMENTS = frozenset({
    "author", "authors", "reviewed-by", "reviewer", "reviewers",
    "contributor", "contributors", "profile", "profiles",
    "login", "log-in", "signin", "sign-in", "signup", "sign-up", "register", "logout",
    "account", "my-account", "cart", "checkout", "wp-admin", "wp-login.php",
    "feed", "print", "share", "amp", "search",
})


def _site_host(url):
    host = (urlparse(url).hostname or "").lower()
    return host[4:] if host.startswith("www.") else host


def _page_key(url):
    p = urlparse(urldefrag(url)[0])
    return (_site_host(url), p.path.rstrip("/") or "/", p.query)


def _rules_for(host):
    for domain, rules in SITE_RULES.items():
        if host == domain or host.endswith("." + domain):
            return rules
    return {}


def _text_len(node):
    return len(clean(node.text_content()))


def _link_density(node):
    total = _text_len(node)
    if not total:
        return 1.0
    linked = sum(len(clean(a.text_content())) for a in node.iter("a"))
    return min(linked / total, 1.0)


def _outermost(elements):
    """Drop elements nested inside another element of the same list."""
    chosen = set(elements)
    return [el for el in elements if not any(anc in chosen for anc in el.iterancestors())]


def _describe(el):
    parts = [el.tag if isinstance(el.tag, str) else "element"]
    for attr in ("class", "id"):
        value = clean(el.get(attr))
        if value:
            parts.append(f'{attr}="{value[:40]}"')
    return "<" + " ".join(parts) + ">"


def _remove_page_chrome(doc):
    for el in doc.xpath(NON_CONTENT_XPATH):
        if el.getparent() is not None:
            el.drop_tree()

    body = doc.find(".//body")
    root = body if body is not None else doc
    total = _text_len(root) or 1
    chrome = [el for el in root.xpath(PAGE_CHROME_XPATH) if _text_len(el) <= total * GUARD_SHARE]
    for el in _outermost(chrome):
        if el.getparent() is not None:
            el.drop_tree()


def _densest_blocks(root):
    """Readability-style fallback: the block(s) holding the most paragraph text."""
    scores = {}
    for el in root.iter("p", "pre", "blockquote", "li"):
        text = clean(el.text_content())
        if len(text) < 25:
            continue
        points = 1 + text.count(",") + min(len(text) // 100, 3)
        parent = el.getparent()
        if parent is None:
            continue
        scores[parent] = scores.get(parent, 0) + points
        if parent is not root:
            grand = parent.getparent()
            if grand is not None:
                scores[grand] = scores.get(grand, 0) + points / 2
    if not scores:
        return []

    adjusted = {node: score * (1 - _link_density(node)) for node, score in scores.items()}
    best = max(adjusted, key=adjusted.get)

    # Content split across sibling sections: keep siblings that score well too.
    blocks = [best]
    parent = best.getparent()
    if best is not root and parent is not None:
        threshold = max(10, adjusted[best] * 0.2)
        blocks = [sib for sib in parent if sib is best or adjusted.get(sib, 0) >= threshold]

    if sum(_text_len(b) for b in blocks) < MIN_CONTENT_CHARS:
        return []
    return blocks


def _find_content_nodes(doc, rules):
    selectors = [(_as_xpath(s), False) for s in rules.get("content", [])] + CONTENT_SELECTORS
    for selector, refine in selectors:
        try:
            nodes = [
                n for n in doc.xpath(selector)
                if isinstance(n, etree._Element) and isinstance(n.tag, str)
            ]
        except etree.XPathError:
            continue
        if not nodes:
            continue
        sizes = {n: _text_len(n) for n in nodes}
        largest = max(sizes.values())
        if largest < MIN_CONTENT_CHARS:
            continue
        picked = _outermost([n for n in nodes if sizes[n] >= largest * SECONDARY_BLOCK_SHARE])
        if refine and len(picked) == 1:
            blocks = _densest_blocks(picked[0])
            if blocks:
                return blocks
        return picked

    body = doc.find(".//body")
    root = body if body is not None else doc
    return _densest_blocks(root) or [root]


def _strip_boilerplate(nodes):
    """Remove boilerplate inside the content block(s); return what was removed."""
    removed = []
    for node in nodes:
        total = _text_len(node) or 1
        hits = [el for el in node.xpath(GUARDED_BOILERPLATE) if _text_len(el) <= total * GUARD_SHARE]
        for el in _outermost(hits):
            if el.getparent() is None:
                continue
            removed.append((el, _describe(el)))
            el.drop_tree()
    return removed


def _path_segments(path):
    return [s.lower() for s in path.split("/") if s]


def _variant_slug(url):
    """The page's own slug, if distinctive enough to spot other views of it."""
    segments = _path_segments(urlparse(url).path)
    if not segments:
        return ""
    slug = segments[-1]
    return slug if len(slug) >= 8 and ("-" in slug or "_" in slug) else ""


def _exclusion_reason(target, page_slug, site_segments):
    segments = _path_segments(urlparse(target).path)
    for seg in segments:
        if seg in site_segments:
            return f"Site rule: /{seg}/ pages"
        if seg in UTILITY_SEGMENTS:
            return f"Author/profile/utility page (/{seg}/)"
    if page_slug and segments and segments[-1] == page_slug:
        return "Another view of this same page (tab, TOC, AMP, print or translation)"
    return ""


def _resolve(href, base):
    href = (href or "").strip()
    if not href or href.startswith("#") or href.lower().startswith(SKIP_SCHEMES):
        return None
    try:
        target = urldefrag(urljoin(base, href))[0]
        if urlparse(target).scheme not in ("http", "https"):
            return None
    except ValueError:
        return None
    return target


def _anchor_text(a):
    anchor = clean(a.text_content())
    if not anchor and a.xpath(".//img"):
        anchor = f"[image] {first_value(a.xpath('.//img/@alt'))}".strip()
    return anchor[:500]


def extract_inlinks(doc, page_url):
    """Return (content inlinks, excluded links). Mutates doc: run it last."""
    base = page_url
    base_href = doc.xpath("//base/@href")
    if base_href and base_href[0].strip():
        base = urljoin(page_url, base_href[0].strip())

    site = _site_host(page_url)
    self_key = _page_key(page_url)
    rules = _rules_for(site)
    site_segments = {s.lower().strip("/") for s in rules.get("exclude", [])}
    page_slug = _variant_slug(page_url)

    def internal_target(a):
        target = _resolve(a.get("href"), base)
        if not target:
            return None
        try:
            if _site_host(target) != site or _page_key(target) == self_key:
                return None
        except ValueError:
            return None
        return target

    _remove_page_chrome(doc)
    nodes = _find_content_nodes(doc, rules)

    links, excluded = [], []

    for block, label in _strip_boilerplate(nodes):
        for a in block.iter("a"):
            target = internal_target(a)
            if target and len(excluded) < MAX_INLINKS_PER_PAGE:
                excluded.append({
                    "url": target,
                    "anchor": _anchor_text(a),
                    "reason": f"Inside boilerplate block {label}",
                })

    for node in nodes:
        for a in node.iter("a"):
            target = internal_target(a)
            if not target:
                continue
            reason = _exclusion_reason(target, page_slug, site_segments)
            if reason:
                if len(excluded) < MAX_INLINKS_PER_PAGE:
                    excluded.append({"url": target, "anchor": _anchor_text(a), "reason": reason})
                continue
            links.append({
                "url": target,
                "anchor": _anchor_text(a),
                "rel": clean(a.get("rel")),
            })
            if len(links) >= MAX_INLINKS_PER_PAGE:
                return links, excluded
    return links, excluded


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

    # Runs last: it removes scripts, nav, header, footer etc. from the parsed tree.
    if "inlinks" in fields:
        links, excluded = extract_inlinks(doc, page_url)
        unique = list(dict.fromkeys(link["url"] for link in links))
        data["Inlinks Count"] = len(unique)
        data["Inlinks"] = " | ".join(unique)[:EXCEL_CELL_LIMIT]
        data["_inlinks"] = links       # per-link detail for the "Content Inlinks" sheet
        data["_excluded"] = excluded   # per-link detail for the "Excluded Inlinks" sheet

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
