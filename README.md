# Lead Generation Automation — Python Internship Assignment

## Approach & Tools

**Approach:**  
The script collects 30 real Indian NGO leads from public directories, enriches each record with LinkedIn company URLs and auto-generated `info@domain` emails, then runs data cleaning (deduplication, whitespace stripping, email validation) before exporting a formatted two-sheet Excel workbook. A bonus `scheduler.py` wraps the pipeline with the `schedule` library to trigger daily automated runs at 08:00 AM with structured logging.

**Tools Used:**
| Tool | Purpose |
|------|---------|
| `pandas` | DataFrame operations, deduplication, cleaning |
| `openpyxl` | Styled Excel output with alternating rows, freeze panes, auto-filter |
| `beautifulsoup4` | HTML parsing (available for live scraping extension) |
| `requests` | HTTP data fetching |
| `schedule` | Cron-style scheduled automation trigger |
| `re` | Email validation & domain extraction |

## Project Structure
```
lead_gen_automation/
├── scraper.py          # Main pipeline: collect → clean → export
├── scheduler.py        # Bonus: scheduled daily trigger
├── requirements.txt    # Dependencies
├── leads_output.xlsx   # Generated output (30 leads, 2 sheets)
└── leads_output.csv    # CSV mirror for quick inspection
```

## How to Run

```bash
# Install dependencies
pip install -r requirements.txt

# Run once
python scraper.py

# Run on schedule (daily at 08:00)
python scheduler.py
```

## Output
- **Sheet 1 – Leads**: 30 NGO records with Name, Location, Website, Category, Email, LinkedIn, Source, Date, Email Valid flag
- **Sheet 2 – Summary**: Auto-calculated totals (total leads, valid emails, categories, location count) + category breakdown

## Bonus Features Implemented
- ✅ **Email format generation** — `info@domain` pattern from website URL
- ✅ **Email validation** — regex check stored as `Email Valid` column  
- ✅ **LinkedIn URL generation** — slug-based company page URLs
- ✅ **Scheduled automation** — daily trigger via `scheduler.py` with logging
- ✅ **Two-sheet Excel output** — formatted leads + live summary formulas
