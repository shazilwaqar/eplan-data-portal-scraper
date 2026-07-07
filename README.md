# EPLAN Data Portal Scraper

A Playwright-based nested scraper for the [EPLAN Data Portal](https://dataportal.eplan.com) — a fully JS-rendered Angular SPA listing over 2.3 million industrial electrical components from manufacturers including Siemens, Rittal, and Schneider Electric.

## Why Playwright (Not the API)

The Data Portal exposes internal API endpoints, but the site is a fully client-rendered Angular application — `view-source` returns nothing but a loading shell with `<app-root>Loading...</app-root>`. Playwright was chosen deliberately to demonstrate real browser automation on a JS-rendered SPA, and to show the complete list→detail navigation pattern that clients typically need on product data collection tasks.

## What It Does

Performs true two-stage (list → detail) browser automation across paginated results:

1. Navigates the parts list page and extracts four fields from each component card
2. Follows each component's detail page link and extracts richer fields from a structured data table
3. Appends each row incrementally to a CSV file — crash-safe at any point in the run
4. Converts the final CSV to a professionally styled Excel workbook

## Project Structure

```
eplan-data-portal-scraper/
├── scraper.py           # Main scraper — pagination loop, list + detail extraction
├── playwright_utils.py  # Reusable browser engine, safe extractors, CSV/Excel helpers
├── eplan_data.csv       # Intermediate output (auto-generated)
├── eplan_data.xlsx      # Final styled Excel output (sample included)
├── requirements.txt
└── .gitignore
```

## Data Extracted

| Field | Source |
|---|---|
| Part Number | List page (`data-test` attribute on part link) |
| Manufacturer | List page (definition list row) |
| Designation | List page (definition list row) |
| Description | Detail page (structured data table) |
| Product Group | Detail page (structured data table) |
| Link | Constructed from detail page href |

Fields absent from a component's detail page are written as empty cells — some manufacturers leave Description blank.

## Technical Highlights

- Modular `playwright_utils.py` with reusable helpers: safe text/attribute extractors, CSV appender, styled Excel converter
- `wait_for_selector()` at both list and detail page level to handle Angular's async rendering
- Polite 3-second delay between detail page visits to avoid rate limiting
- Incremental CSV append after every row — data is safe even if the run crashes mid-way
- Final Excel output auto-styled with branded header, frozen panes, auto-fit column widths
- Graceful per-item error handling: a failed item logs the error and continues the loop

## Output

Professionally formatted Excel file with one component per row. Sample output included: `eplan_data.xlsx`

## Setup

```bash
pip install playwright openpyxl pandas
playwright install chromium
python scraper.py
```

## Configuration

In `scraper.py`, adjust the page range to control how many pages to scrape (200 parts per page):

```python
for page_no in range(1, 11):  # Change 11 to scrape more or fewer pages
```

## Legal

EPLAN Data Portal's `robots.txt` places no restrictions on general crawlers. No terms of service clause prohibiting automated access was identified. All data collected is publicly visible without authentication.
