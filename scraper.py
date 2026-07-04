#!/usr/bin/env python3
"""
IslamWeb Library Scraper
Scrapes all textbooks from islamweb.net/ar/library into structured JSON.
Uses concurrent workers with per-thread sessions for maximum throughput.
"""

import requests
from bs4 import BeautifulSoup
import json
import os
import re
import time
import sys
from urllib.parse import urljoin
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

BASE_URL = "https://islamweb.net/ar/library"
AJAX_BASE = "https://islamweb.net/ar/library/maktaba"
OUTPUT_DIR = Path("output")
DATA_DIR = Path("scraped_data")
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

# Per-thread sessions (thread-safe, avoids global lock)
thread_local = threading.local()

def get_session():
    if not hasattr(thread_local, "session"):
        s = requests.Session()
        s.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ar,en-US;q=0.7,en;q=0.3",
            "Connection": "close",
        })
        thread_local.session = s
    return thread_local.session

# Rate limiter: shared across all threads
class RateLimiter:
    def __init__(self, requests_per_sec=10):
        self.min_interval = 1.0 / requests_per_sec
        self.last_time = 0
        self.lock = threading.Lock()
    def wait(self):
        with self.lock:
            elapsed = time.time() - self.last_time
            if elapsed < self.min_interval:
                time.sleep(self.min_interval - elapsed)
            self.last_time = time.time()

limiter = RateLimiter(requests_per_sec=3)

def safe_get(url, **kwargs):
    for attempt in range(5):
        limiter.wait()
        try:
            session = get_session()
            resp = session.get(url, timeout=60, **kwargs)
            resp.encoding = "utf-8"
            if resp.status_code == 200:
                return resp
            print(f"  HTTP {resp.status_code} for {url[:80]}", flush=True, file=sys.stderr)
        except Exception as e:
            print(f"  Error fetching {url[:80]}: {e}", flush=True, file=sys.stderr)
        time.sleep(2 ** attempt)
    return None

def clean_text(html_content):
    soup = BeautifulSoup(html_content, "lxml")
    for tag in soup.find_all(["script", "style"]):
        tag.decompose()
    for span in soup.find_all("span", style=re.compile(r'display\s*:\s*none', re.I)):
        span.decompose()
    for a in soup.find_all("a", href=lambda h: h and h.startswith("#docu")):
        a.unwrap()
    for tag in soup.find_all(["span", "font"]):
        if tag.name == "font":
            tag.unwrap()
        elif tag.name == "span" and not tag.get("class"):
            tag.unwrap()
    text = soup.get_text(separator=" ")
    text = re.sub(r'\xa0+', ' ', text)
    text = re.sub(r'\u200c+', '', text)
    text = re.sub(r'\[\s*ص:\s*\d+\s*\]', '', text)
    text = re.sub(r'nindex\S+\s*', '', text)
    text = '\n'.join(line.strip() for line in text.split('\n') if line.strip())
    return text.strip()

def clean_html(html_content):
    soup = BeautifulSoup(html_content, "lxml")
    for tag in soup.find_all(["script", "style"]):
        tag.decompose()
    for span in soup.find_all("span", style=re.compile(r'display\s*:\s*none', re.I)):
        span.decompose()
    for a in soup.find_all("a"):
        a.unwrap()
    for tag in soup.find_all(attrs={"onmouseover": True}):
        for attr in list(tag.attrs):
            if attr in ("onmouseover", "onmouseout", "onclick"):
                del tag[attr]
    for span in soup.find_all("span"):
        classes = span.get("class", [])
        allowed = {"quran", "hadith", "names"}
        keep = [c for c in classes if c in allowed]
        if keep:
            span.attrs = {"class": " ".join(keep)}
        else:
            span.unwrap()
    for font in soup.find_all("font"):
        c = font.get("color", "")
        if c and re.match(r'^#[0-9a-fA-F]{6}$', c):
            font.name = "span"
            font.attrs = {"class": f"fc_{c}"}
        else:
            font.unwrap()
    for tag in soup.find_all(style=re.compile(r'display\s*:\s*none', re.I)):
        if tag.name != "span":
            del tag["style"]
    result = str(soup)
    result = re.sub(r'<html><body>', '', result)
    result = re.sub(r'</body></html>', '', result)
    result = re.sub(r'nindex\S+\s*', '', result)
    return result.strip()

# ============================================================
# STAGE 1: Subjects
# ============================================================
SUBJECTS_CACHE = DATA_DIR / "subjects.json"

def get_subjects():
    if SUBJECTS_CACHE.exists():
        with open(SUBJECTS_CACHE) as f:
            return json.load(f)
    url = f"{BASE_URL}/index.php?page=bookslist"
    resp = safe_get(url)
    if not resp:
        return []
    soup = BeautifulSoup(resp.text, "lxml")
    subjects = []
    for a in soup.select("div.leftblock.fatCatleft li a[href*='subject=']"):
        href = a.get("href", "")
        m = re.search(r'subject=(\d+)', href)
        if m:
            subjects.append({"id": int(m.group(1)), "name": a.get_text(strip=True)})
    with open(SUBJECTS_CACHE, "w", encoding="utf-8") as f:
        json.dump(subjects, f, ensure_ascii=False, indent=2)
    print(f"Found {len(subjects)} subjects", flush=True)
    return subjects

# ============================================================
# STAGE 2: Books in a subject
# ============================================================
def get_books_for_subject(subject_id, subject_name):
    books = []
    url = f"{BASE_URL}/index.php?page=bookslist&subject={subject_id}"
    resp = safe_get(url)
    if not resp:
        return books
    soup = BeautifulSoup(resp.text, "lxml")
    items = soup.select("li.answer[itemtype*='Book']")
    if not items:
        return books
    for item in items:
        a_tag = item.select_one("h2 a[href*='/ar/library/content/']")
        if not a_tag:
            continue
        href = a_tag.get("href", "")
        m = re.search(r'/content/(\d+)/', href)
        if not m:
            continue
        book_id = int(m.group(1))
        name_el = item.select_one("[itemprop='name']")
        author_el = item.select_one("[itemprop='author']")
        date_el = item.select_one("[itemprop='datePublished']")
        publisher_el = item.select_one("[itemprop='publisher']")
        idto_match = re.search(r'idto=(\d+)', href)
        total_pages = int(idto_match.group(1)) if idto_match else 0
        books.append({
            "book_id": book_id,
            "title": name_el.get_text(strip=True) if name_el else "",
            "author": author_el.get_text(strip=True) if author_el else "",
            "date": date_el.get("content", "") if date_el else "",
            "publisher": publisher_el.get_text(strip=True) if publisher_el else "",
            "total_pages": total_pages,
            "subject_id": subject_id,
            "subject_name": subject_name,
            "url": urljoin(BASE_URL, href),
        })
    return books

# ============================================================
# STAGE 3: Recursive tree expansion
# ============================================================
def expand_tree_node(book_id, node_id):
    items = []
    ajax_url = f"{AJAX_BASE}/nindex.php?id={node_id}&treeLevel=1&bookid={book_id}&page=bookssubtree"
    resp = safe_get(ajax_url)
    if not resp:
        return items
    soup = BeautifulSoup(resp.text, "lxml")
    for a in soup.select("a[href*='/ar/library/content/']"):
        href = a.get("href", "")
        title = a.get_text(strip=True)
        a_id = a.get("id", "")
        pagenum_m = re.search(r'/content/\d+/(\d+)/', href)
        pagenum = int(pagenum_m.group(1)) if pagenum_m else None
        items.append({
            "id": int(a_id) if a_id.isdigit() else a_id,
            "title": title,
            "pagenum": pagenum,
            "url": urljoin(BASE_URL, href),
        })
    for label in soup.select("label.tree_label[data-level]"):
        child_id = label.get("data-id", "")
        title = label.get_text(strip=True)
        if child_id:
            items.append({
                "id": int(child_id) if child_id.isdigit() else child_id,
                "title": title,
                "sections": expand_tree_node(book_id, child_id),
            })
    return items

def get_book_tree(book_id):
    url = f"{BASE_URL}/content/{book_id}/1"
    resp = safe_get(url)
    if not resp:
        return None, []
    soup = BeautifulSoup(resp.text, "lxml")
    book_title_el = soup.select_one("h3.txt-blue")
    book_title = book_title_el.get_text(strip=True) if book_title_el else ""
    author_el = soup.select_one("h4.txt-secondary")
    author = author_el.get_text(strip=True) if author_el else ""
    tree_items = soup.select("li.first-level")
    chapters = []

    def expand_chapter(item):
        label = item.select_one("label.tree_label")
        if not label:
            return None
        node_id = label.get("data-id", "")
        title = label.get_text(strip=True)
        chapter = {
            "id": int(node_id) if node_id.isdigit() else node_id,
            "title": title,
            "sections": expand_tree_node(book_id, node_id),
        }
        return chapter

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(expand_chapter, item) for item in tree_items}
        for f in as_completed(futures):
            ch = f.result()
            if ch:
                chapters.append(ch)
    chapters.sort(key=lambda c: c["id"] if isinstance(c["id"], int) else 0)
    return {"title": book_title, "author": author}, chapters

# ============================================================
# STAGE 4: Fetch content for a section
# ============================================================
def fetch_content(section_url, with_html=False):
    resp = safe_get(section_url)
    if not resp:
        return "", "", "", "", None
    soup = BeautifulSoup(resp.text, "lxml")
    pagebody = soup.select_one("#pagebody")
    pagebody_tashkeel = soup.select_one("#pagebody_thaskeel")
    text = clean_text(str(pagebody)) if pagebody else ""
    text_t = clean_text(str(pagebody_tashkeel)) if pagebody_tashkeel else ""
    html_text = clean_html(str(pagebody)) if pagebody and with_html else ""
    html_text_t = clean_html(str(pagebody_tashkeel)) if pagebody_tashkeel and with_html else ""
    part_el = soup.select_one(".partdropmenu .dropdown-toggle, .partdropmenu a")
    part = None
    if part_el:
        txt = part_el.get_text(strip=True)
        if txt.isdigit():
            part = int(txt)
    return text, text_t, html_text, html_text_t, part

# ============================================================
# STAGE 5: Full book scrape
# ============================================================
def scrape_book(book_info):
    book_id = book_info["book_id"]
    cache_file = OUTPUT_DIR / f"book_{book_id}.json"
    if cache_file.exists():
        with open(cache_file) as f:
            return json.load(f)

    print(f"\nScraping book {book_id}: {book_info['title']}", flush=True)

    meta, chapters = get_book_tree(book_id)
    if meta is None:
        print(f"  FAILED: book page", flush=True)
        return None

    def collect_leaves(nodes):
        leaves = []
        for n in nodes:
            if "url" in n:
                leaves.append(n)
            if "sections" in n:
                leaves.extend(collect_leaves(n["sections"]))
        return leaves
    all_sections = collect_leaves(chapters)
    leaf_count = len(all_sections)
    print(f"  Tree: {len(chapters)} chapters, {leaf_count} leaf sections", flush=True)

    results = []
    done_lock = threading.Lock()
    done_count = 0
    total = leaf_count
    t0 = time.time()

    def fetch_section(sec):
        nonlocal done_count
        text, text_t, html_text, html_text_t, part = fetch_content(sec["url"], with_html=True)
        sec["content"] = text
        sec["content_with_tashkeel"] = text_t
        sec["content_html"] = html_text
        sec["content_with_tashkeel_html"] = html_text_t
        if part is not None: sec["part"] = part
        with done_lock:
            done_count += 1
            n = done_count
        if n % 100 == 0 or n == 1:
            elapsed = time.time() - t0
            rate = n / elapsed if elapsed > 0 else 0
            eta = (total - n) / rate if rate > 0 else 0
            print(f"  [{n}/{total}] {rate:.1f}/s, ETA {eta:.0f}s - {sec['title'][:40]}", flush=True)

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(fetch_section, sec): sec for sec in all_sections}
        for f in as_completed(futures):
            f.result()

    result = {
        "book": book_info,
        "chapters": chapters,
    }

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)
    elapsed = time.time() - t0
    total_chars = sum(len(s.get("content","")) for s in all_sections)
    print(f"  Saved ({elapsed:.0f}s, {total_chars:,} chars)", flush=True)
    return result

# ============================================================
# REFRESH: re-fetch content to add HTML fields
# ============================================================
def refresh_book(book_id, checkpoint_every=500):
    cache_file = OUTPUT_DIR / f"book_{book_id}.json"
    if not cache_file.exists():
        print(f"No cache for book {book_id}")
        return None
    with open(cache_file) as f:
        data = json.load(f)

    def collect_leaves(nodes):
        leaves = []
        for n in nodes:
            if "url" in n:
                leaves.append(n)
            if "sections" in n:
                leaves.extend(collect_leaves(n["sections"]))
        return leaves
    all_sections = collect_leaves(data.get("chapters", []))
    print(f"Refreshing {len(all_sections)} sections for book {book_id}...", flush=True)

    # Skip sections already updated (have fc_ spans)
    need_refresh = []
    for s in all_sections:
        html = s.get("content_with_tashkeel_html", "")
        if "fc_" in html:
            continue
        need_refresh.append(s)
    skipped = len(all_sections) - len(need_refresh)
    if skipped:
        print(f"  {skipped} already up-to-date, {len(need_refresh)} to refresh", flush=True)

    done_count = 0
    last_save = 0
    done_lock = threading.Lock()
    total = len(need_refresh)
    t0 = time.time()

    def fetch_section(sec):
        nonlocal done_count, last_save
        sec.pop("content_html", None)
        sec.pop("content_with_tashkeel_html", None)
        text, text_t, html_text, html_text_t, part = fetch_content(sec["url"], with_html=True)
        sec["content"] = text
        sec["content_with_tashkeel"] = text_t
        sec["content_html"] = html_text
        sec["content_with_tashkeel_html"] = html_text_t
        if part is not None: sec["part"] = part
        with done_lock:
            done_count += 1
            n = done_count
        if n % 100 == 0 or n == 1:
            elapsed = time.time() - t0
            rate = n / elapsed if elapsed > 0 else 0
            eta = (total - n) / rate if rate > 0 else 0
            print(f"  [{n}/{total}] {rate:.1f}/s, ETA {eta:.0f}s", flush=True)
        if checkpoint_every and (n - last_save) >= checkpoint_every:
            with done_lock:
                with open(cache_file, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2, default=str)
                last_save = n
                print(f"  checkpoint saved ({n}/{total})", flush=True)

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(fetch_section, sec): sec for sec in need_refresh}
        for f in as_completed(futures):
            f.result()

    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    elapsed = time.time() - t0
    print(f"Refresh done ({elapsed:.0f}s)", flush=True)
    return data

# ============================================================
# MAIN
# ============================================================
def scrape_all():
    print("ISLAMWEB LIBRARY SCRAPER", flush=True)

    subjects = get_subjects()
    if not subjects:
        print("No subjects found!", flush=True)
        return

    all_books = []
    for subj in subjects:
        print(f"--- {subj['name']} (id={subj['id']}) ---", flush=True)
        books = get_books_for_subject(subj["id"], subj["name"])
        print(f"  {len(books)} books", flush=True)
        all_books.extend(books)

    all_books.sort(key=lambda b: b["book_id"])
    with open(OUTPUT_DIR / "all_books.json", "w", encoding="utf-8") as f:
        json.dump(all_books, f, ensure_ascii=False, indent=2)
    print(f"Total: {len(all_books)} books", flush=True)

    with open(OUTPUT_DIR / "book_ids.txt", "w") as f:
        for b in all_books:
            f.write(f"{b['book_id']}\t{b['title']}\t{b['author']}\t{b['subject_name']}\n")

    print("Books saved to output/book_ids.txt", flush=True)
    for book in all_books:
        result = scrape_book(book)
        if result is None:
            print(f"  SKIPPED {book['title']}", flush=True)

    gen_master_json()
    print("DONE", flush=True)

def scrape_single(book_id):
    # Fetch metadata directly from the book page instead of scanning all subjects
    url = f"{BASE_URL}/content/{book_id}/1"
    resp = safe_get(url)
    if not resp:
        print(f"Book {book_id} not found", flush=True)
        return
    soup = BeautifulSoup(resp.text, "lxml")
    title_el = soup.select_one("h3.txt-blue")
    title = title_el.get_text(strip=True) if title_el else ""
    author_el = soup.select_one("h4.txt-secondary")
    author = author_el.get_text(strip=True) if author_el else ""
    meta_author = soup.select_one("meta[itemprop='author']")
    meta_date = soup.select_one("meta[itemprop='datePublished']")
    meta_publisher = soup.select_one("meta[itemprop='name']")
    publisher = meta_publisher.get("content", "") if meta_publisher else ""
    date = meta_date.get("content", "") if meta_date else ""
    if not author and meta_author:
        author = meta_author.get("content", "")
    # Try to find subject from breadcrumbs or sidebar category
    subject_name = ""
    for a in soup.select("div.leftblock a[href*='subject=']"):
        sn = a.get_text(strip=True)
        if sn:
            subject_name = sn
            break
    if not subject_name:
        for a in soup.select("a[href*='subject=']"):
            sn = a.get_text(strip=True)
            if sn and sn not in ("", "قائمة الكتب"):
                subject_name = sn
                break
    # Fallback: scan subjects list to find this book
    if not subject_name:
        subs = get_subjects()
        for sub in subs:
            books = get_books_for_subject(sub["id"], sub["name"])
            for b in books:
                if b["book_id"] == book_id:
                    subject_name = sub["name"]
                    break
            if subject_name:
                break
            time.sleep(0.3)
    book_info = {
        "book_id": book_id,
        "title": title,
        "author": author,
        "publisher": publisher,
        "date": date,
        "total_pages": 0,
        "subject_id": 0,
        "subject_name": subject_name,
        "url": url,
    }
    if not title:
        print(f"Book {book_id} not found (no title)", flush=True)
        return
    result = scrape_book(book_info)
    if result:
        gen_master_json()
    return result

def gen_master_json():
    master = {"subjects": {}}
    for bf in sorted(OUTPUT_DIR.glob("book_*.json")):
        with open(bf) as f:
            data = json.load(f)
        sn = data["book"]["subject_name"]
        master["subjects"].setdefault(sn, {"name": sn, "books": []})
        master["subjects"][sn]["books"].append(data)
    with open(OUTPUT_DIR / "master.json", "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=2)
    print(f"Master JSON: {OUTPUT_DIR / 'master.json'}", flush=True)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "master":
            gen_master_json()
        elif sys.argv[1].isdigit():
            scrape_single(int(sys.argv[1]))
        else:
            print("Usage: python scraper.py [book_id|master]", flush=True)
    else:
        scrape_all()
