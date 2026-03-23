# Blueprint: Data Extraction Pipeline

## Goal
Convert Excel market feedback data + JSW price list into a single JS data file for the dashboard.

## Inputs Required
- `data/Market Feedback Report.xlsx` — dealer survey data with brands, prices, states, districts
- `data/Price List.xlsx` — JSW ONE TMT distributor landed prices (optional, skipped if absent)

## Script
`scripts/extract_data.py` — Run via `python scripts/extract_data.py` or double-click `refresh.bat`

## Steps
1. **Find Excel** — Auto-detects latest `.xlsx` in `/data/` (excludes `Price List.xlsx`)
2. **Detect sheet** — Finds sheet with columns: Brand name, Amount, Auto state, Months
3. **Read data** — Loads full sheet into pandas DataFrame
4. **GST normalize** — Auto-detects Price Type column, converts inclusive prices to excl. GST (÷ 1.18)
5. **Clean & filter** — Removes nulls, filters to 30k-80k range, normalizes states/districts to uppercase, maps brand names
6. **Build metadata** — Extracts months, states, districts, top 20 brands
7. **Build filterable data** — Converts each row to compact JSON format (`b`, `a`, `s`, `d`, `m`, `q`, `t`, `c`, `w`, `wl`)
8. **Inject JSW prices** — Calls `parse_price_list()` to add JSW ONE TMT 550 and 550D records
9. **Write output** — Generates `src/data/dashboard-data.js` as a JavaScript constant

## Output
`src/data/dashboard-data.js` — contains `const DASHBOARD_DATA = { metadata: {...}, filterableData: [...] };`

## Key Constants
- `GST_RATE = 1.18` (18%)
- `DEALER_MARGIN = 3000` (added to JSW distributor prices)
- `PRICE_LIST_FILE = 'Price List.xlsx'`
- Price range filter: 30,000 to 80,000

## Edge Cases
- **No Excel file in /data/**: Script exits with error message
- **No matching sheet**: Script exits listing available sheets
- **No GST column**: Assumes all prices are already excl. GST
- **No Price List.xlsx**: Skips JSW injection, dashboard works without it
- **Brand "jsw neo"**: Auto-mapped to "JSW Neosteel" (line 130)

## Verification
After running, check:
1. Console output shows row counts and JSW record counts
2. `src/data/dashboard-data.js` file size is reasonable (~1MB)
3. Refresh browser — dashboard loads with updated data
