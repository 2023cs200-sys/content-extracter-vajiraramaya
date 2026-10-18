#!/usr/bin/env python3
"""Collect one website host and its navigation into a Word document.

Run locally where internet access is available:
    py -m pip install -r requirements.txt
    py collect_vajirarama.py

The crawler stays on the exact host in START_URL and follows navigation links
on that host only. It records source URLs and image URLs/captions; it does not
republish or bypass access controls.
"""
from __future__ import annotations
import time
from collections import deque
from datetime import datetime
from urllib.parse import urljoin, urlparse, urldefrag
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup, Tag
from docx import Document
from docx.shared import Inches, Pt, RGBColor

START_URL = "https://lib.vajirarama.lk/"
MAX_PAGES = 500
OUTPUT_FILE = "website_content.docx"
DELAY_SECONDS = 0.8
TIMEOUT_SECONDS = 25
USER_AGENT = "EducationalContentArchive/1.0 (polite crawler; contact site administrator if needed)"
NON_HTML_EXTENSIONS = {
    ".7z", ".avi", ".css", ".csv", ".doc", ".docx", ".gif", ".gz", ".ico",
    ".jpeg", ".jpg", ".js", ".json", ".mp3", ".mp4", ".png", ".pdf", ".ppt",
    ".pptx", ".rar", ".svg", ".tar", ".txt", ".webp", ".woff", ".woff2",
    ".xls", ".xlsx", ".xml", ".zip",
}

session = requests.Session()
session.headers.update({"User-Agent": USER_AGENT})


ALLOWED_HOST = (urlparse(START_URL).hostname or "").lower()


def is_internal_host(hostname: str) -> bool:
    return hostname.lower() == ALLOWED_HOST


def canonicalize(url: str) -> str | None:
    url, _ = urldefrag(url)
    p = urlparse(url)
    if p.scheme not in ("http", "https"):
        return None
    if not p.hostname or not is_internal_host(p.hostname):
        return None
    path = p.path or "/"
    suffix = "." + path.rsplit(".", 1)[-1].lower() if "." in path.rsplit("/", 1)[-1] else ""
    if suffix in NON_HTML_EXTENSIONS:
        return None
    if path != "/" and not path.endswith("/") and "." not in path.rsplit("/", 1)[-1]:
        path += "/"
    # Drop tracking/query strings and keep the exact configured host.
    return f"{p.scheme}://{ALLOWED_HOST}{path}"


def get_robots_parser(host: str) -> RobotFileParser:
    rp = RobotFileParser()
    rp.set_url(f"https://{host}/robots.txt")
    try:
        rp.read()
    except Exception:
        # If robots.txt cannot be read, don't treat the failure as permission
        # to evade restrictions; the user should check the site's policy.
        pass
    return rp


def can_fetch(url: str, robots_by_host: dict[str, RobotFileParser]) -> bool:
    host = (urlparse(url).hostname or "").lower()
    if host not in robots_by_host:
        robots_by_host[host] = get_robots_parser(host)
    return robots_by_host[host].can_fetch(USER_AGENT, url)


def select_content(soup: BeautifulSoup) -> Tag:
    selectors = [
        "main", "article", ".entry-content", ".post-content",
        ".elementor-widget-theme-post-content", ".site-main", "#content"
    ]
    for selector in selectors:
        node = soup.select_one(selector)
        if node and len(node.get_text(" ", strip=True)) > 120:
            # Prefer the deepest likely article content if available.
            if selector in ("main", ".site-main", "#content"):
                nested = node.select_one("article, .entry-content, .post-content, .elementor-widget-theme-post-content")
                if nested and len(nested.get_text(" ", strip=True)) > 120:
                    return nested
            return node
    return soup.body or soup


def clean_content(node: Tag) -> Tag:
    node = BeautifulSoup(str(node), "html.parser")
    for bad in node.select("script, style, noscript, iframe, nav, footer, header, form, .sharedaddy, .related-posts, .comments-area"):
        bad.decompose()
    return node


def page_title(soup: BeautifulSoup, fallback: str) -> str:
    h = soup.select_one("h1")
    if h and h.get_text(" ", strip=True):
        return h.get_text(" ", strip=True)
    if soup.title and soup.title.get_text(" ", strip=True):
        return soup.title.get_text(" ", strip=True).split("|")[0].strip()
    return fallback


def extract_page(url: str) -> dict:
    response = session.get(url, timeout=TIMEOUT_SECONDS)
    response.raise_for_status()
    if not response.url or not is_internal_host(urlparse(response.url).hostname or ""):
        raise ValueError(f"redirected outside the configured host: {response.url}")
    soup = BeautifulSoup(response.text, "html.parser")
    title = page_title(soup, url.rstrip("/").split("/")[-1].replace("-", " ").title())
    content = clean_content(select_content(soup))

    # Keep meaningful page content and collect nested links from the content area.
    blocks = []
    links = set()
    images = []
    for el in content.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "blockquote", "pre", "table", "img"], recursive=True):
        if el.name == "img":
            src = el.get("src") or el.get("data-src") or el.get("data-lazy-src") or ""
            if not src and el.get("srcset"):
                src = el.get("srcset").split(",")[0].strip().split(" ")[0]
            if src:
                img_url = urljoin(url, src)
                alt = (el.get("alt") or "").strip()
                # A nearby figure caption is more useful than an empty alt.
                parent = el.find_parent("figure")
                caption = parent.select_one("figcaption") if parent else None
                cap_text = caption.get_text(" ", strip=True) if caption else ""
                images.append({"url": img_url, "alt": alt, "caption": cap_text})
            continue
        if el.name == "table":
            rows = []
            for tr in el.select("tr"):
                cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"], recursive=False)]
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                blocks.append({"type": "table", "text": "\n".join(rows)})
            continue
        text = el.get_text(" ", strip=True)
        if not text:
            continue
        if el.name.startswith("h"):
            blocks.append({"type": el.name, "text": text})
        elif el.name == "li":
            blocks.append({"type": "li", "text": text})
        elif el.name == "blockquote":
            blocks.append({"type": "quote", "text": text})
        elif el.name == "pre":
            blocks.append({"type": "pre", "text": text})
        elif el.name == "p":
            blocks.append({"type": "p", "text": text})

    # Crawl links from the extracted content, including links to internal
    # subdomains and nested navigation pages.
    for a in content.select("a[href]"):
        target = canonicalize(urljoin(url, a.get("href", "")))
        if target and target != url:
            links.add(target)

    # If the content selector missed navigation-style links, inspect all body
    # anchors as a fallback. canonicalize() still constrains them to the site.
    for a in (soup.body or soup).select("a[href]"):
        target = canonicalize(urljoin(url, a.get("href", "")))
        if target and target != url:
            links.add(target)

    # Remove duplicate images while preserving order.
    seen_img, unique_images = set(), []
    for img in images:
        key = img["url"]
        if key not in seen_img:
            seen_img.add(key)
            unique_images.append(img)
    return {"url": response.url, "title": title, "blocks": blocks, "images": unique_images, "links": sorted(links)}


def add_hyperlink(paragraph, text: str, url: str):
    """Create a clickable hyperlink in a python-docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = __import__('docx').oxml.OxmlElement('w:hyperlink')
    hyperlink.set(__import__('docx').oxml.ns.qn('r:id'), r_id)
    new_run = __import__('docx').oxml.OxmlElement('w:r')
    rPr = __import__('docx').oxml.OxmlElement('w:rPr')
    color = __import__('docx').oxml.OxmlElement('w:color'); color.set(__import__('docx').oxml.ns.qn('w:val'), '0563C1'); rPr.append(color)
    underline = __import__('docx').oxml.OxmlElement('w:u'); underline.set(__import__('docx').oxml.ns.qn('w:val'), 'single'); rPr.append(underline)
    new_run.append(rPr)
    t = __import__('docx').oxml.OxmlElement('w:t'); t.text = text; new_run.append(t)
    hyperlink.append(new_run); paragraph._p.append(hyperlink)


def build_docx(pages: list[dict], errors: list[tuple[str, str]]):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.7); sec.bottom_margin = Inches(0.7)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.name = "Arial"
    styles["Title"].font.color.rgb = RGBColor(31, 78, 121)

    doc.add_heading("Website Content Archive", 0)
    p = doc.add_paragraph()
    p.add_run("Collected on: ").bold = True
    p.add_run(datetime.now().astimezone().strftime("%Y-%m-%d %H:%M %Z"))
    p = doc.add_paragraph()
    p.add_run("Starting page: ").bold = True
    add_hyperlink(p, START_URL, START_URL)
    doc.add_paragraph(
        "This document is an automatically collected reference copy of pages discovered on the starting site and its internal subdomains. "
        "Original page URLs and image URLs are preserved. Text, image captions, and page structure may not be identical to the live website. "
        "Please respect the website's copyright and terms when reusing content."
    )
    doc.add_heading("Contents", level=1)
    for page in pages:
        p = doc.add_paragraph(style="List Number")
        p.add_run(page["title"] + " — ")
        add_hyperlink(p, page["url"], page["url"])
    if errors:
        doc.add_paragraph(f"Pages that could not be fetched: {len(errors)}. See the collection notes at the end.")

    for page in pages:
        doc.add_page_break()
        doc.add_heading(page["title"], level=1)
        p = doc.add_paragraph()
        p.add_run("Source URL: ").bold = True
        add_hyperlink(p, page["url"], page["url"])
        if not page["blocks"]:
            doc.add_paragraph("No readable article text was extracted from this page. It may use dynamic content or a layout requiring manual review.")
        for block in page["blocks"]:
            kind, text = block["type"], block["text"]
            if kind in ("h1", "h2", "h3", "h4", "h5", "h6"):
                level = min(max(int(kind[1]), 2), 4)
                doc.add_heading(text, level=level)
            elif kind == "li":
                doc.add_paragraph(text, style="List Bullet")
            elif kind == "quote":
                doc.add_paragraph(text, style="Quote")
            elif kind == "pre":
                para = doc.add_paragraph()
                run = para.add_run(text); run.font.name = "Consolas"; run.font.size = Pt(9)
            elif kind == "table":
                for row in text.splitlines():
                    doc.add_paragraph(row)
            else:
                doc.add_paragraph(text)
        if page["images"]:
            doc.add_heading("Images referenced on this page", level=2)
            for img in page["images"]:
                p = doc.add_paragraph(style="List Bullet")
                label = img["caption"] or img["alt"] or "Image (no caption/alt text supplied)"
                p.add_run(label + " — ")
                add_hyperlink(p, img["url"], img["url"])
                if img["alt"] and img["caption"] and img["alt"] != img["caption"]:
                    p.add_run(" (alt text: " + img["alt"] + ")")

    doc.add_page_break()
    doc.add_heading("Collection notes", level=1)
    doc.add_paragraph(f"Successfully collected pages: {len(pages)}")
    doc.add_paragraph(f"Fetch failures: {len(errors)}")
    if errors:
        for url, err in errors:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(url + " — " + err)
    doc.add_paragraph(
        f"Scope: links on the exact configured host ({ALLOWED_HOST}) only. "
        "External domains and non-HTML assets are not included."
    )
    doc.save(OUTPUT_FILE)


def main():
    start_url = canonicalize(START_URL)
    if not start_url:
        raise SystemExit("START_URL must be an http(s) URL on a valid host.")

    robots_by_host: dict[str, RobotFileParser] = {}
    if not can_fetch(start_url, robots_by_host):
        raise SystemExit("The site's robots.txt disallows this crawler for the starting URL. Please request permission or review robots.txt.")

    queue = deque([start_url])
    seen = set()
    pages, errors = [], []
    while queue and len(seen) < MAX_PAGES:
        url = queue.popleft()
        url = canonicalize(url)
        if not url or url in seen:
            continue
        seen.add(url)
        print(f"[{len(seen)}] Fetching {url}")
        try:
            if not can_fetch(url, robots_by_host):
                print("  SKIPPED: disallowed by robots.txt")
                continue
            page = extract_page(url)
            pages.append(page)
            for link in page["links"]:
                if link not in seen:
                    queue.append(link)
        except Exception as exc:
            msg = f"{type(exc).__name__}: {exc}"
            print(f"  FAILED: {msg}")
            errors.append((url, msg))
        time.sleep(DELAY_SECONDS)

    # Stable ordering: main page first, then alphabetical title.
    pages.sort(key=lambda p: (p["url"] != start_url, p["title"].casefold(), p["url"]))
    build_docx(pages, errors)
    print(f"\nDone. Pages saved: {len(pages)}; failures: {len(errors)}")
    print(f"Word document: {OUTPUT_FILE}")
    if queue:
        print(f"Warning: MAX_PAGES={MAX_PAGES} reached; some links may remain unvisited.")

if __name__ == "__main__":
    main()
