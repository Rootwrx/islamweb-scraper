# islamweb-scraper

Scrape and convert Islamic textbooks from [islamweb.net/ar/library](https://islamweb.net/ar/library) into styled DOCX.

## Pipeline

```
islamweb.net  ──scraper.py──►  JSON (with HTML)  ──to_docx.py──►  DOCX
```

## Scraper (`scraper.py`)

Fetches the entire book tree and content using concurrent HTTP requests.

### Usage

```bash
# Scrape a single book by ID
python3 scraper.py 209

# Scrape all books across all subjects
python3 scraper.py --all
```

### Features

- **Concurrent** — 4 worker threads, 3 req/s rate limit
- **Resilient** — 5 retries with exponential backoff, 60s timeout, `Connection: close` per request
- **Per-thread sessions** — avoids connection contention
- **Diacritics preserved** — full tashkeel (تشكيل) kept
- **Hidden tooltips stripped** — spans with `display:none` (`quranatt`, `hadithatt`, `mainsubjatt`, `hashiya_title`) removed
- **Output** — `output/book_<id>.json` with full chapter tree, metadata, and HTML content

### Commands

| Command | Description |
|---------|-------------|
| `python3 scraper.py <book_id>` | Scrape a single book |
| `python3 scraper.py <url>` | Scrape book from full URL |
| `python3 scraper.py --list` | List all subjects |
| `python3 scraper.py --subjects` | List books grouped by subject |
| `python3 scraper.py <book_id> --refresh` | Re-scrape cached book |
| `python3 scraper.py --all` | Scrape every book from all subjects |

## DOCX Converter (`to_docx.py`)

Converts scraped JSON into a professionally formatted DOCX with full RTL support, heading hierarchy, and colored Quran/Hadith.

### Usage

```bash
# Single volume
python3 to_docx.py 209 --full

# Multi-volume (10 chapters per volume by default)
python3 to_docx.py 209

# Custom volume size
python3 to_docx.py 209 --chapters-per-volume=5

# From URL or cached JSON
python3 to_docx.py output/book_209.json
```

### Formatting

| Element | Style |
|---------|-------|
| **Background** | `#FAF7F0` (warm cream) |
| **Font** | Scheherazade New (full complex-script support) |
| **Body text** | 14pt bold, 1.15× line spacing, RTL |
| **Quran verses** | Pure red `#FF0000` bold, wrapped in ﴿…﴾ |
| **Hadith** | Pure blue `#0000FF` bold, wrapped in «…» |
| **Subject headings** | Dark red bold `#B43232` |
| **Page markers** | `[ ص: N ]` stripped automatically |
| **Hidden tooltips** | `display:none` elements skipped |

### Heading Hierarchy

| Level | Size | Color |
|-------|------|-------|
| H1 | 26pt | Dark blue `#1A3A5C` |
| H2 | 24pt | Golden `#8B6914` |
| H3 | 22pt | Teal `#005A50` |
| H4 | 20pt | Dark red `#8C3232` |
| H5 | 18pt | Purple `#503278` |
| H6 | 17pt | Green `#3C6432` |
| H7 | 16pt | Brown `#965028` |
| H8 | 15pt | Steel blue `#283C64` |
| H9 | 14pt | Gray `#646464` |

### Cover Page

- Book title — 52pt golden bold
- Author — 32pt dark blue
- Publisher / subject — 24pt gray
- Volume info — 32pt golden (multi-volume only)

### Structure

- No TOC page — Word Navigation Pane serves as outline
- Hidden outline numbering (visible only in Navigation Pane)
- Consecutive `<br>` tags collapsed to single paragraph break
- Page break before each chapter section

### Google Docs Compatibility

The converter strips non-standard parts (`stylesWithEffects.xml`, `customXml/`) and avoids VML elements for reliable opening in Google Docs.

## Bookmarks

File `4- folders_collapsable.txt` documents the API format used for expanding the book tree.

## Requirements

```bash
pip install requests beautifulsoup4 lxml python-docx
```

Font: [Scheherazade New](https://scripts.sil.org/cms/scripts/page.php?site_id=nrsi&id=Scheherazade) must be installed on the system.
