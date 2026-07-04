#!/usr/bin/env python3
"""
Convert scraped IslamWeb JSON to styled PDFs with TOC, PDF outlines, and colors.
Uses weasyprint for HTML→PDF rendering with Amiri Arabic font.
Preserves HTML with quran/hadith spans for proper coloring.
"""

import json
import sys
from pathlib import Path
import weasyprint

OUTPUT_DIR = Path("output")
FONT = "/usr/share/fonts/TTF/Amiri-Regular.ttf"
FONT_BOLD = "/usr/share/fonts/TTF/Amiri-Bold.ttf"

VOLUME_SIZE = 15  # chapters per volume

CSS = """
@page {
    size: A5;
    margin: 2cm 1.8cm;
    @bottom-center {
        content: counter(page);
        font-family: 'Amiri', serif;
        font-size: 10pt;
        color: #666;
    }
}
@page :first {
    @bottom-center { content: none; }
}
@page toc {
    @bottom-center { content: none; }
}

html {
    background: #faf7f0;
}
body {
    font-family: 'Amiri', serif;
    font-size: 12pt;
    line-height: 1.8;
    direction: rtl;
    text-align: right;
    color: #1a1a1a;
    margin: 0;
    padding: 0;
}

/* Full page background via overlay */
body::before {
    content: '';
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: #faf7f0;
    z-index: -1;
}

/* Colored elements */
.quran {
    color: #1a6b1a;
}
.hadith {
    color: #012D6B;
}

/* Cover */
.cover {
    page-break-after: always;
    text-align: center;
    padding-top: 30%;
}
.cover h1 {
    font-size: 26pt;
    color: #8B6914;
    margin-bottom: 1cm;
    line-height: 1.4;
}
.cover .author {
    font-size: 14pt;
    color: #1a3a5c;
    margin-bottom: 0.5cm;
}
.cover .publisher {
    font-size: 11pt;
    color: #888;
}
.cover .subject {
    font-size: 11pt;
    color: #888;
    margin-top: 1cm;
}
.cover .volume {
    font-size: 13pt;
    color: #8B6914;
    margin-top: 0.5cm;
}
.cover .chapters-range {
    font-size: 11pt;
    color: #666;
}

/* TOC */
.toc-page { page: toc; page-break-after: always; }
.toc-page h1 {
    text-align: center;
    color: #8B6914;
    font-size: 20pt;
    margin-bottom: 1cm;
}
.toc-page ul {
    list-style: none;
    padding: 0;
    margin: 0;
}
.toc-page li {
    font-size: 11pt;
    margin: 2px 0;
    padding: 2px 0;
}
.toc-page a {
    text-decoration: none;
    color: #1a1a1a;
    display: flex;
    justify-content: space-between;
}
.toc-page a::after {
    content: target-counter(attr(href), page);
    color: #8B6914;
    font-weight: bold;
}
.toc-page .toc-chapter {
    font-weight: bold;
    color: #1a3a5c;
    margin-top: 4px;
}

/* Chapters */
.chapter {
    page-break-before: always;
    font-size: 18pt;
    color: #1a3a5c;
    margin-bottom: 0.3cm;
    padding-bottom: 5px;
    border-bottom: 2px solid #8B6914;
    -weasy-bookmark-level: 1;
    -weasy-bookmark-label: content();
}

/* Sections */
.section-title {
    font-size: 13pt;
    color: #8B6914;
    margin: 0.4cm 0 0.1cm 0;
    padding-bottom: 2px;
    border-bottom: 1px solid #e0d5b8;
    -weasy-bookmark-level: 2;
    -weasy-bookmark-label: content();
}
.content {
    text-align: justify;
}
.content p {
    margin: 0.2em 0;
    text-indent: 1.5em;
}
.content p:first-of-type {
    text-indent: 0;
}
"""

def build_volume_html(data, chapters, volume_num, total_volumes):
    book = data["book"]
    parts = []

    ch_start = chapters[0]["title"] if chapters else ""
    ch_end = chapters[-1]["title"] if chapters else ""

    # Cover
    parts.append(f"""\
<div class="cover">
    <h1>{book['title']}</h1>
    <div class="volume">الجزء {volume_num} من {total_volumes}</div>
    <div class="chapters-range">{ch_start} → {ch_end}</div>
    <div class="author">{book['author']}</div>
    <div class="publisher">{book.get('publisher', '')}</div>
    <div class="subject">{book.get('subject_name', '')}</div>
</div>""")

    # TOC (chapters only; sections in PDF outlines)
    parts.append('<div class="toc-page"><h1>الفهرس</h1><ul>')
    for i, ch in enumerate(chapters, 1):
        aid = f'ch{i}' if volume_num == 1 else f'ch{volume_num}-{i}'
        parts.append(f'<li class="toc-chapter"><a href="#{aid}">{ch["title"]}</a></li>')
    parts.append('</ul></div>')

    # Chapters
    for i, ch in enumerate(chapters, 1):
        aid = f'ch{i}' if volume_num == 1 else f'ch{volume_num}-{i}'
        parts.append(f'<h2 id="{aid}" class="chapter">{ch["title"]}</h2>')
        for sec in ch["sections"]:
            title = sec["title"]
            content_html = sec.get("content_with_tashkeel_html", "") or sec.get("content_with_tashkeel", "")
            # Strip wrapper div attributes, keep only inner content
            content_html = strip_wrapper(content_html)
            parts.append(f'<div class="section-title">{title}</div>')
            parts.append(f'<div class="content">{content_html}</div>')

    body = '\n'.join(parts)

    return f"""\
<!DOCTYPE html>
<html dir="rtl" lang="ar">
<head>
<meta charset="utf-8">
<style>
@font-face {{
    font-family: 'Amiri';
    src: url('{FONT}') format('truetype');
    font-weight: normal;
}}
@font-face {{
    font-family: 'Amiri';
    src: url('{FONT_BOLD}') format('truetype');
    font-weight: bold;
}}
{CSS}
</style>
</head>
<body>
{body}
</body>
</html>"""

def strip_wrapper(html):
    """Remove the outer div wrapper (bookcontent-dic) from cleaned HTML."""
    import re
    html = re.sub(r'<div[^>]*id="pagebody[^"]*"[^>]*>\s*', '', html, count=1)
    html = re.sub(r'</div>\s*$', '', html)
    return html.strip()

def render_volume(data, chapters, volume_num, total_volumes, pdf_path):
    print(f"  Volume {volume_num}/{total_volumes} ({len(chapters)} chapters)...", flush=True)
    html_str = build_volume_html(data, chapters, volume_num, total_volumes)
    doc = weasyprint.HTML(string=html_str)
    doc.write_pdf(pdf_path)
    size_mb = pdf_path.stat().st_size / 1024 / 1024
    print(f"    Done: {pdf_path} ({size_mb:.1f} MB)", flush=True)

def json_to_pdf(json_path, pdf_stem=None, chapters_per_volume=VOLUME_SIZE):
    json_path = Path(json_path)
    if pdf_stem is None:
        pdf_stem = json_path.stem.replace("book_", "")

    print(f"Reading {json_path}...", flush=True)
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    chapters = data["chapters"]
    total_ch = len(chapters)

    if total_ch <= chapters_per_volume * 1.5:
        pdf_path = OUTPUT_DIR / f"book_{pdf_stem}.pdf"
        print(f"Single volume -> {pdf_path}", flush=True)
        render_volume(data, chapters, 1, 1, pdf_path)
        return [pdf_path]
    else:
        n_volumes = (total_ch + chapters_per_volume - 1) // chapters_per_volume
        pdfs = []
        for v in range(n_volumes):
            start = v * chapters_per_volume
            end = min((v + 1) * chapters_per_volume, total_ch)
            vol_chapters = chapters[start:end]
            pdf_path = OUTPUT_DIR / f"book_{pdf_stem}_v{v+1}.pdf"
            pdfs.append(pdf_path)
            render_volume(data, vol_chapters, v+1, n_volumes, pdf_path)
        return pdfs

def batch_convert(chapters_per_volume=VOLUME_SIZE):
    json_files = sorted(OUTPUT_DIR.glob("book_*.json"))
    if not json_files:
        print("No JSON files found in output/", flush=True)
        return
    for jf in json_files:
        pdf_stem = jf.stem.replace("book_", "")
        existing = list(OUTPUT_DIR.glob(f"book_{pdf_stem}_v*.pdf")) + list(OUTPUT_DIR.glob(f"book_{pdf_stem}.pdf"))
        if existing:
            print(f"Skipping {jf.name} (PDF exists)", flush=True)
            continue
        json_to_pdf(jf, pdf_stem=pdf_stem, chapters_per_volume=chapters_per_volume)

if __name__ == "__main__":
    cpv = VOLUME_SIZE

    args = sys.argv[1:]
    while args and args[0].startswith("--"):
        opt = args.pop(0)
        if opt == "--chapters-per-volume" and args:
            cpv = int(args.pop(0))
        else:
            print(f"Unknown option: {opt}", flush=True)
            sys.exit(1)

    if not args:
        print(f"Usage: {sys.argv[0]} [--chapters-per-volume N] <book.json|batch|all> [stem]", flush=True)
        print(f"  --chapters-per-volume N  Split into volumes of N chapters (default: {VOLUME_SIZE})", flush=True)
        sys.exit(1)

    if args[0] in ("batch", "all"):
        batch_convert(chapters_per_volume=cpv)
    else:
        json_path = args[0]
        pdf_stem = args[1] if len(args) > 1 else None
        json_to_pdf(json_path, pdf_stem=pdf_stem, chapters_per_volume=cpv)
