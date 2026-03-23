# Distributor One-Pager Report Generator

## Project Purpose
Generates JSW One TMT distributor one-pager reports in `.docx` format from a master Excel file containing sales, dealer, district, and invoice data.

## Architecture: WAT Framework
- `workflows/` - Markdown SOPs (see `generate_distributor_one_pager.md`)
- `tools/` - Python scripts for data extraction and report generation
- `data/` - Source Excel file (do not modify)
- `output/` - Generated .docx reports
- `.tmp/` - Temporary/intermediate files

## Key Commands
```bash
# Generate single report
python tools/generate_one_pager.py "DISTRIBUTOR NAME" "STATE"

# Batch generate by zone (loads workbook once, outputs to output/{date}/)
python tools/batch_generate.py North              # North region, today's date
python tools/batch_generate.py East 2026-03-09    # East region, specific date

# Test data extraction only
python tools/extract_distributor_data.py "DISTRIBUTOR NAME" "STATE"
```

## Important Conventions
- **Filter key**: Always use Distributor Name + State combo (some distributors operate across states)
- **FY26 definition**: April 2025 through March 2026
- **Invoice dates are strings** (YYYY-MM-DD format), not datetime objects - parse with `strptime`
- **Metric formulas are user-defined** - do not change without confirmation:
  - Active = transacted in last 3 rolling months
  - SoB% = Avg Monthly FY26 Sales / Counter Potential (**transacting dealers only**)
  - Reach% = Counter Potential (**transacting dealers only**) / Retail Demand
  - Market Share% = Reach% * SoB%
  - "Transacting" = dealers who lifted (sales > 0) in that month; CP summed per month then averaged

## Report Formatting
- **All numbers rounded to integers** — no decimals anywhere (sales, percentages)
- **Mar'26 excluded** — all monthly tables show 11 months (Apr'25–Feb'26)
- **Title**: `JSW ONE TMT - <Distributor Name> - Performance Feb'26`
- **SM/TM source**: `data/District wise SM_TM name.xlsx` (preferred), fallback to District master

## Section-Specific Notes
- **Sales Manpower**: 2-column table (Designation, Name). SM/TM from dedicated SM_TM file.
- **Sales by Location**: Parent rows (Retailer, Self-stocking) show integer sums. Child rows show **% only**.
- **Key Districts**: Only Very High, High, Medium categories shown. Low districts excluded.
- **New Dealers**: Counted from `Onboarded Dealer Months` column in FY 26 dealer sales (e.g. "Apr'25").
- **Top 10 Dealers**: Sorted by FY26 total sales descending. Wider dealer name/district columns.

## Dependencies
- Python 3.14+
- `openpyxl` - Excel reading
- `python-docx` - Word document generation
