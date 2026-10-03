Built by [Talha Pythoneer](https://www.talhapythoneer.com), web scraping and AI agents.

# Legal 500 Law Firm Scraper

A standalone Python scraper for [legal500.com](https://www.legal500.com) that
collects law firm profiles — practice areas, firm descriptions, office
contacts, and more — organized by the site's own geographic hierarchy
(Region → Sub-Region → Country → Law firms), for Europe, Middle East, Africa,
Americas, and Asia Pacific.

Built for a freelance data-extraction job; shared here as a reference
implementation / portfolio piece. A small, real sample of the output lives in
[`sample_output/`](sample_output/).

## What it collects

Per firm:

- **Practice areas**, each with its Legal 500 ranking tier (or "Firms to
  Watch" status where that applies instead of a numbered tier)
- **About** / firm description
- **Office address(es)**
- Listing (profile) URL
- Firm name, Country, Sub-Region, Region
- Website
- Phone number(s) and email(s) per office (Cloudflare-obfuscated emails are
  decoded automatically)
- Lawyers (name + position)
- Which extra editorial sections the firm has (News & developments, Client
  testimonials, Diversity, Teams, Interviews, Comparative guides)

Output is a single `.xlsx` workbook with one sheet per region. Inside each
sheet, rows are sorted by Sub-Region → Country → Firm name, so the geographic
hierarchy is immediately visible (e.g. Latin America → Argentina → firms),
plus a Summary sheet with per-region totals.

## Why no Selenium / Playwright

legal500.com is a Next.js site behind Cloudflare, but it's fully
server-rendered and Cloudflare's check here is TLS/fingerprint-based, not a
JS challenge. A real browser isn't needed to get the data — a plain HTTP
client that presents a convincing browser fingerprint is enough, and it's
far faster and lighter than driving a headless browser for ~4,000+ profile
pages.

This scraper uses [`curl_cffi`](https://github.com/lexiforest/curl_cffi)
(Chrome TLS/HTTP2 impersonation) + `lxml` for parsing + `openpyxl` for the
workbook, scraping at ~600–700 firms/minute with 10–12 worker threads — no
browser, no driver binaries.

## Installation

```bash
pip install -r requirements.txt
```

Requires Python 3.9+.

## Usage

```bash
# Full run — all 5 regions, every country, every firm (~4,000+ firms, ~6-8 min)
python legal500_scraper.py

# Just a couple of countries
python legal500_scraper.py --countries portugal,spain

# Just a couple of regions
python legal500_scraper.py --regions americas,europe

# Quick smoke test
python legal500_scraper.py --limit 20

# More/fewer concurrent workers (default 10)
python legal500_scraper.py --workers 16

# Rebuild the .xlsx from the existing cache only — no network calls
python legal500_scraper.py --build-only
```

Full options:

```
--regions TEXT       Comma-separated region filter, e.g. "europe,americas"
--countries TEXT      Comma-separated country slugs, e.g. "portugal,spain"
--limit INT           Cap the number of firms scraped (for testing)
--workers INT          Concurrent worker threads (default 10)
--cache PATH           Resumable JSONL cache path (default legal500_cache.jsonl)
--output PATH          Output .xlsx path (default legal500_law_firms.xlsx)
--build-only           Skip scraping; just rebuild the .xlsx from the cache
-v, --verbose          Debug logging
```

### Resumable by design

Every scraped firm is appended to a local JSONL cache file as soon as it's
scraped. If the run is interrupted (or crashes), re-running the same command
skips everything already cached and only fetches what's missing.

## Notes on data quality

- A small fraction of firms listed in Legal 500's own country directories
  (~4% in a full run) have profile links that 404 on legal500.com itself.
  Those rows still get a name/country/listing-URL, but the rest is blank —
  the data simply isn't available at the source, not a parsing failure.
- A handful of firms carry an already-corrupted character in the raw bytes
  legal500.com's own server sends — not something decoding on this end can
  recover.

## License

MIT — see [LICENSE](LICENSE).
