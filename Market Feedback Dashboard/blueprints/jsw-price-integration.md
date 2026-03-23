# Blueprint: JSW ONE TMT Price Integration

## Goal
Parse JSW ONE TMT prices from Price List.xlsx and inject as dealer-landed-price records into the dashboard data.

## Inputs Required
- `data/Price List.xlsx` — Excel file with date-wise tabs containing Fe-550 and Fe-550D prices

## Script
`scripts/extract_data.py` → function `parse_price_list()` (called automatically from `main()`)

## Price List Structure
- **Tabs**: Date-wise (e.g., "16th Jan-26", "05th Feb-26", "16th Mar-26")
- **Column B**: State name
- **Column C**: Major city
- **Column F** (index 5): Fe-550 price (Basic + Freight, 12-32mm)
- **Column J** (index 9): Fe-550D price
- **Prices are**: Distributor landed, excluding GST

## Two Brands (Treated Separately)
| Brand Name | Grade | Coverage |
|------------|-------|----------|
| JSW ONE TMT 550 | Fe-550 | All ~20 states |
| JSW ONE TMT 550D | Fe-550D | ~14 states (some eastern states have no Fe-550D data) |

## Price Conversion
`Dealer Landed Price = Distributor Price + Rs 3,000`

All prices in the Excel are already excl. GST — no GST conversion needed.

## State Name Mapping
| Price List | Dashboard |
|------------|-----------|
| Kashmir | JAMMU AND KASHMIR |
| Jammu | JAMMU AND KASHMIR |
| Chandigarh | CHANDIGARH |
| Himachal Pradesh | HIMACHAL PRADESH |
| Uttarakhand | UTTARAKHAND |
| Punjab | PUNJAB |
| Delhi | DELHI |
| Haryana | HARYANA |
| Rajasthan | RAJASTHAN |
| Uttar Pradesh (West) | UTTAR PRADESH |
| Uttar Pradesh (E+C) | UTTAR PRADESH |
| Chhattisgarh | CHHATTISGARH |
| Madhya Pradesh | MADHYA PRADESH |
| Odisha | ODISHA |
| Jharkhand | JHARKHAND |
| Bihar | BIHAR |
| West Bengal | WEST BENGAL |
| Gujarat | GUJARAT |
| Maharashtra | MAHARASHTRA |

States with multiple entries (J&K, UP, WB) produce multiple records — dashboard averages them.

## Tab Selection Logic
Multiple tabs exist per month. Script picks the **latest date** per month:
- Parses tab name with regex: `(\d+)\w*\s+(\w+)\s*-\s*(\d+)`
- Groups by month key (e.g., "Jan-26")
- Selects tab with highest day number

## Edge Cases
- **Tab name spacing varies**: "07th Feb -26" vs "16th Feb - 26" — whitespace normalized with `re.sub(r'\s+', ' ', ...)`
- **Odisha composite string**: Cell contains `"51100, Odisha West - 50400"` — `_extract_price()` takes first number before comma
- **Empty Fe-550D cells**: Some states (Odisha, Jharkhand, Bihar, parts of WB, Maharashtra) have no Fe-550D prices — silently skipped
- **New states**: Bihar and Gujarat don't exist in market feedback data — added to metadata automatically
- **Price List.xlsx missing**: Function returns empty list, dashboard works without JSW data

## Dashboard Behavior
- JSW brands inserted at front of `topBrands` → pre-selected in default 5-brand dropdown
- JSW ONE TMT 550 set as `benchmarkBrand` → thicker chart lines, accent colors
- District field is empty (`''`) for JSW records — correctly excluded when user filters by specific district
- Company field is "JSW ONE Distribution Ltd" for all records

## How to Update Prices
1. Add new tabs to `data/Price List.xlsx` with the same column structure
2. Run `python scripts/extract_data.py`
3. Script auto-detects new tabs and picks the latest per month
4. Refresh browser
