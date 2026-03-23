# Project Learnings

Track errors, solutions, and insights. System gets smarter with each entry.

## Format
Date | Component | Issue | Resolution | Insight

---

## Entries
(Add new entries at top)

### 2026-03-17 | extract_data.py | Price List.xlsx picked as main data file
`find_latest_excel()` auto-selects the most recently modified .xlsx. After adding Price List.xlsx, it was chosen instead of the Market Feedback Report.
**Fix:** Exclude `PRICE_LIST_FILE` from the file list in `find_latest_excel()`.
**Insight:** When adding new Excel files to `/data/`, always check if `find_latest_excel()` needs an exclusion filter.

### 2026-03-17 | extract_data.py | Tab name spacing inconsistency
Price List.xlsx tabs have inconsistent spacing: `"07th Feb -26"`, `"16th Feb - 26"`, `"16th Jan-26"`.
**Fix:** Use `re.sub(r'\s+', ' ', sheet_name.strip())` to normalize whitespace before parsing.
**Insight:** Always normalize whitespace when parsing user-created Excel tab/sheet names.

### 2026-03-17 | extract_data.py | Odisha composite price string
Odisha cell contains `"51100, Odisha West - 50400"` instead of a plain number.
**Fix:** `str(val).split(',')[0].strip()` extracts the first number. Helper function `_extract_price()` handles both numeric and string cells.
**Insight:** Never assume Excel cells are purely numeric — always handle string variants.

### 2026-03-17 | extract_data.py | State name mapping complexity
Price List uses mixed-case regional names: "Kashmir", "Jammu", "Uttar Pradesh (West)", "Uttar Pradesh (E+C)".
Dashboard uses uppercase state names: "JAMMU AND KASHMIR", "UTTAR PRADESH".
**Fix:** Explicit `STATE_MAP` dictionary with `.lower()` lookup. Multiple sub-regions map to the same state, producing multiple data points that are averaged by the dashboard.
**Insight:** When merging external price data, always build an explicit mapping table — don't rely on fuzzy matching.

### 2026-03-17 | dashboard.js | Benchmark brand was hardcoded null
Line 7 had `const BENCHMARK_BRAND = null;` — would need manual code change each time.
**Fix:** Changed to `DASHBOARD_DATA.metadata.benchmarkBrand` so it's data-driven. Python script sets it when JSW records are present.
**Insight:** Make configuration data-driven whenever possible — avoids manual code edits when data changes.

### 2026-03-17 | extract_data.py | Dealer margin calculation
JSW prices in Price List.xlsx are distributor landed (excl. GST). Dashboard shows dealer landed prices.
**Fix:** Add `DEALER_MARGIN = 3000` constant and apply it during record generation.
**Insight:** Document the price basis (distributor vs dealer, incl/excl GST) clearly in constants and comments.
