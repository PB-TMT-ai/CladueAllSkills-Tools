"""
Extract and aggregate market feedback data from Excel into a JS data file.

Usage:
  1. Place your latest Excel file in the /data folder
  2. Run: python scripts/extract_data.py
     (or double-click refresh.bat)
  3. Refresh the dashboard in your browser

The script auto-detects the most recently modified .xlsx file in /data.
"""

import pandas as pd
import json
import os
import glob
import re
from collections import defaultdict
from datetime import datetime

# Paths relative to project root
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FOLDER = os.path.join(PROJECT_ROOT, 'data')
OUTPUT_PATH = os.path.join(PROJECT_ROOT, 'src', 'data', 'dashboard-data.js')

# The report sheet name — the script also tries to auto-detect it
EXPECTED_COLUMNS = ['Brand name', 'Amount', 'Auto state', 'Months']

# Candidate names for the Price Type / GST column (case-insensitive matching also used)
GST_COLUMN_CANDIDATES = [
    'Price Type', 'Price type', 'price type', 'PRICE TYPE',
    'Pricing Type', 'GST Type', 'gst type',
]

GST_RATE = 1.18  # 18% GST for TMT steel

# JSW ONE TMT Price List configuration
PRICE_LIST_FILE = 'Price List.xlsx'
DEALER_MARGIN = 3000  # Add to distributor landed price to get dealer landed price
JSW_550_BRAND = 'JSW ONE TMT 550'
JSW_550D_BRAND = 'JSW ONE TMT 550D'


def find_latest_excel():
    """Find the most recently modified .xlsx file in /data folder,
    excluding the JSW Price List (handled separately)."""
    pattern = os.path.join(DATA_FOLDER, '*.xlsx')
    files = glob.glob(pattern)
    # Exclude the JSW price list file — it's parsed separately
    files = [f for f in files if os.path.basename(f) != PRICE_LIST_FILE]
    if not files:
        print(f"ERROR: No .xlsx files found in {DATA_FOLDER}")
        print(f"  Place your Market Feedback Report .xlsx file in that folder and try again.")
        raise SystemExit(1)

    # Sort by modification time, newest first
    files.sort(key=os.path.getmtime, reverse=True)
    chosen = files[0]
    print(f"  Found {len(files)} Excel file(s), using latest:")
    print(f"  -> {os.path.basename(chosen)}")
    if len(files) > 1:
        print(f"    (other files: {', '.join(os.path.basename(f) for f in files[1:])})")
    return chosen


def find_report_sheet(excel_path):
    """Auto-detect the sheet containing the market feedback report data."""
    xl = pd.ExcelFile(excel_path)
    for name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=name, nrows=5)
            if all(col in df.columns for col in EXPECTED_COLUMNS):
                print(f"  Report sheet found: \"{name}\"")
                return name
        except Exception:
            continue

    print("ERROR: Could not find a sheet with the expected columns:")
    print(f"  Expected: {EXPECTED_COLUMNS}")
    print(f"  Available sheets: {xl.sheet_names}")
    raise SystemExit(1)


def find_gst_column(df):
    """Auto-detect the Price Type / GST Type column."""
    lowered_candidates = {c.strip().lower() for c in GST_COLUMN_CANDIDATES}
    for col in df.columns:
        if col.strip().lower() in lowered_candidates:
            return col
    # Fallback: fuzzy match
    for col in df.columns:
        low = col.strip().lower()
        if 'price' in low and 'type' in low:
            return col
        if 'gst' in low and 'type' in low:
            return col
    return None


def parse_price_list():
    """Parse Price List.xlsx and return synthetic dealer-landed-price records
    for JSW ONE TMT 550 and 550D."""
    price_list_path = os.path.join(DATA_FOLDER, PRICE_LIST_FILE)
    if not os.path.exists(price_list_path):
        print(f"  Price List not found at {price_list_path}, skipping JSW ONE TMT.")
        return []

    xl = pd.ExcelFile(price_list_path)

    # Step 1: Group tabs by month, pick latest date per month
    tabs_by_month = defaultdict(list)
    for sheet_name in xl.sheet_names:
        clean = re.sub(r'\s+', ' ', sheet_name.strip())
        match = re.match(r'(\d+)\w*\s+(\w+)\s*-\s*(\d+)', clean)
        if match:
            day, mon, year = match.groups()
            month_key = f"{mon}-{year}"  # e.g., "Jan-26"
            tabs_by_month[month_key].append((int(day), sheet_name))

    latest_tabs = {}
    for month_key, tabs in tabs_by_month.items():
        tabs.sort(key=lambda x: x[0])
        latest_tabs[month_key] = tabs[-1][1]  # sheet name with highest day

    # Step 2: State name mapping (Price List → dashboard uppercase)
    STATE_MAP = {
        'kashmir': 'JAMMU AND KASHMIR',
        'jammu': 'JAMMU AND KASHMIR',
        'chandigarh': 'CHANDIGARH',
        'himachal pradesh': 'HIMACHAL PRADESH',
        'uttarakhand': 'UTTARAKHAND',
        'punjab': 'PUNJAB',
        'delhi': 'DELHI',
        'haryana': 'HARYANA',
        'rajasthan': 'RAJASTHAN',
        'uttar pradesh (west)': 'UTTAR PRADESH',
        'uttar pradesh (e+c)': 'UTTAR PRADESH',
        'chhattisgarh': 'CHHATTISGARH',
        'madhya pradesh': 'MADHYA PRADESH',
        'odisha': 'ODISHA',
        'jharkhand': 'JHARKHAND',
        'bihar': 'BIHAR',
        'west bengal': 'WEST BENGAL',
        'gujarat': 'GUJARAT',
        'maharashtra': 'MAHARASHTRA',
    }

    # Step 3: Parse each latest tab and generate records
    records = []
    for month_key, sheet_name in sorted(latest_tabs.items()):
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)

        # Derive week number from tab date
        clean = re.sub(r'\s+', ' ', sheet_name.strip())
        day_match = re.match(r'(\d+)', clean)
        day_of_month = int(day_match.group(1)) if day_match else 15
        week_num = min((day_of_month - 1) // 7 + 1, 5)
        week_label = f"{month_key}-W{week_num}"

        for i in range(len(df)):
            raw_state = df.iloc[i, 1]  # Column B = state
            if pd.isna(raw_state):
                continue
            state_str = str(raw_state).strip()
            state_lower = state_str.lower()

            if state_lower not in STATE_MAP:
                continue

            mapped_state = STATE_MAP[state_lower]

            # Fe-550 price (Column F, index 5)
            fe550_price = _extract_price(df, i, 5)

            # Fe-550D price (Column J, index 9)
            fe550d_price = _extract_price(df, i, 9)

            if fe550_price is not None:
                records.append({
                    'b': JSW_550_BRAND,
                    'a': fe550_price + DEALER_MARGIN,
                    's': mapped_state,
                    'd': '',
                    'm': month_key,
                    'q': 'BIS',
                    't': None,
                    'c': 'JSW ONE Distribution Ltd',
                    'w': week_num,
                    'wl': week_label,
                })

            if fe550d_price is not None:
                records.append({
                    'b': JSW_550D_BRAND,
                    'a': fe550d_price + DEALER_MARGIN,
                    's': mapped_state,
                    'd': '',
                    'm': month_key,
                    'q': 'BIS',
                    't': None,
                    'c': 'JSW ONE Distribution Ltd',
                    'w': week_num,
                    'wl': week_label,
                })

    return records


def _extract_price(df, row_idx, col_idx):
    """Extract a numeric price from a cell, handling composite strings like
    '51100, Odisha West - 50400' by taking the first number."""
    if col_idx >= df.shape[1]:
        return None
    val = df.iloc[row_idx, col_idx]
    if pd.isna(val):
        return None
    # If already numeric
    if isinstance(val, (int, float)):
        return int(val)
    # Parse first number from string
    price_str = str(val).split(',')[0].strip()
    try:
        return int(float(price_str))
    except (ValueError, TypeError):
        return None


def normalize_gst(df, gst_col):
    """Convert all prices to excluding-GST basis.

    Uses substring matching to handle descriptive values like
    'dealer landing price excluding gst (12-25 mm per mt)'.
    Rows that mention 'excluding gst' are kept as-is.
    Rows that mention 'including gst' are converted (÷ 1.18).
    Rows with no GST mention are assumed to include GST and converted.
    """
    normalized = df[gst_col].fillna('').str.strip().str.lower()

    is_exc = normalized.str.contains('excluding gst|excl gst|excl\\.? gst|ex-gst|without gst|gst exclusive', regex=True)
    is_inc = normalized.str.contains('including gst|incl gst|incl\\.? gst|inc-gst|with gst|gst inclusive', regex=True) & ~is_exc
    # Rows with a non-empty value that don't mention GST at all — assume inclusive
    has_value = normalized != ''
    is_no_gst_mention = has_value & ~is_exc & ~is_inc
    # Combine: convert both explicit inc-GST and no-GST-mention rows
    needs_conversion = is_inc | is_no_gst_mention

    summary = {
        'total': len(df),
        'inc_gst_converted': int(needs_conversion.sum()),
        'exc_gst_kept': int(is_exc.sum()),
        'unknown_kept': int((~has_value).sum()),
        'unique_values': sorted(normalized[has_value].unique().tolist()),
    }

    df.loc[needs_conversion, 'Amount'] = (df.loc[needs_conversion, 'Amount'] / GST_RATE).round(0)

    return df, summary


def clean_data(df):
    """Filter and normalize the raw data."""
    df = df[df['Brand name'].notna() & df['Amount'].notna()].copy()
    df = df[(df['Amount'] >= 30000) & (df['Amount'] <= 80000)].copy()

    df['Auto state'] = df['Auto state'].fillna('').str.strip().str.upper()
    df['Auto district'] = df['Auto district'].fillna('').str.strip().str.upper()

    df['Brand name'] = df['Brand name'].str.strip()
    df.loc[df['Brand name'].str.lower() == 'jsw neo', 'Brand name'] = 'JSW Neosteel'

    df['Quality_clean'] = df['Quality'].fillna('Unknown')
    df.loc[df['Quality_clean'] == 'BIS certified', 'Quality_clean'] = 'BIS'
    df.loc[df['Quality_clean'] == 'Non-BIS certified', 'Quality_clean'] = 'NonBIS'

    df['Delivery_clean'] = df['Timeliness of delivery'].fillna('')
    df['Months'] = df['Months'].str.strip()

    df['Last Modified Date'] = pd.to_datetime(df['Last Modified Date'], errors='coerce')
    df['WeekOfMonth'] = df['Last Modified Date'].apply(
        lambda d: min((d.day - 1) // 7 + 1, 5) if pd.notna(d) else 1
    )
    df['WeekLabel'] = df.apply(
        lambda r: f"{r['Months']}-W{int(r['WeekOfMonth'])}" if pd.notna(r['Months']) else '', axis=1
    )

    df = df[df['Auto state'] != ''].copy()
    return df


def get_top_brands(df, n=20):
    """Get top N brands by row count."""
    return df['Brand name'].value_counts().head(n).index.tolist()


def build_month_order(df):
    """Build chronological month order from the data itself."""
    # Parse "Mon-YY" into sortable dates
    month_strs = df['Months'].dropna().unique()
    parsed = []
    for m in month_strs:
        try:
            dt = datetime.strptime(m, '%b-%y')
            parsed.append((dt, m))
        except ValueError:
            continue
    parsed.sort(key=lambda x: x[0])
    return [m for _, m in parsed]


def build_metadata(df, top_brands, month_order):
    """Build metadata object."""
    months_present = [m for m in month_order if m in df['Months'].unique()]
    states = sorted(df['Auto state'].unique().tolist())

    districts = {}
    for state in states:
        state_districts = sorted(
            df[df['Auto state'] == state]['Auto district'].unique().tolist()
        )
        districts[state] = [d for d in state_districts if d]

    all_brands = df['Brand name'].value_counts().index.tolist()

    return {
        'totalRows': len(df),
        'generatedAt': datetime.now().isoformat(),
        'months': months_present,
        'states': states,
        'districts': districts,
        'topBrands': top_brands,
        'allBrands': all_brands,
        'benchmarkBrand': None
    }


def build_filterable_data(df):
    """Build compact per-row data for client-side filtering."""
    rows = []
    for _, r in df.iterrows():
        rows.append({
            'b': r['Brand name'],
            'a': round(float(r['Amount'])),
            's': r['Auto state'],
            'd': r['Auto district'],
            'm': r['Months'],
            'q': r['Quality_clean'],
            't': r['Delivery_clean'] if r['Delivery_clean'] else None,
            'c': str(r['Registered Company Name']) if pd.notna(r['Registered Company Name']) else '',
            'w': int(r['WeekOfMonth']),
            'wl': r['WeekLabel'],
        })
    return rows


def main():
    print("=" * 50)
    print("Market Feedback Dashboard — Data Refresh")
    print("=" * 50)

    print("\n1. Finding Excel file...")
    excel_path = find_latest_excel()

    print("\n2. Detecting report sheet...")
    sheet_name = find_report_sheet(excel_path)

    print("\n3. Reading data...")
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    print(f"  Raw rows: {len(df)}")

    print("\n3.5. GST normalization...")
    gst_col = find_gst_column(df)
    if gst_col:
        print(f"  Found GST column: \"{gst_col}\"")
        df, gst_summary = normalize_gst(df, gst_col)
        print(f"  Converted {gst_summary['inc_gst_converted']} rows to Ex-GST (÷ {GST_RATE})")
        print(f"  Already Ex-GST: {gst_summary['exc_gst_kept']} rows")
        if gst_summary['unknown_kept'] > 0:
            print(f"  Empty price type: {gst_summary['unknown_kept']} rows (kept as-is)")
        if gst_summary['unique_values']:
            print(f"  Price type values found: {gst_summary['unique_values']}")
    else:
        print("  No 'Price Type' column found — assuming all prices are Ex-GST")
        gst_summary = None

    print("\n4. Cleaning & filtering...")
    df = clean_data(df)
    print(f"  Clean rows: {len(df)}")

    top_brands = get_top_brands(df, n=20)
    print(f"  Top 20 brands: {top_brands[:5]}... (+{len(top_brands)-5} more)")

    print("\n5. Building dashboard data...")
    month_order = build_month_order(df)
    metadata = build_metadata(df, top_brands, month_order)
    filterable_data = build_filterable_data(df)

    metadata['priceNormalization'] = {
        'applied': gst_col is not None,
        'basis': 'Excl. GST',
        'gstRate': 18,
        'rowsConverted': gst_summary['inc_gst_converted'] if gst_summary else 0,
    }

    # Step 5.5: Parse JSW ONE TMT price list and inject
    print("\n5.5. Parsing JSW ONE TMT price list...")
    jsw_records = parse_price_list()
    if jsw_records:
        filterable_data.extend(jsw_records)
        fe550_count = sum(1 for r in jsw_records if r['b'] == JSW_550_BRAND)
        fe550d_count = sum(1 for r in jsw_records if r['b'] == JSW_550D_BRAND)
        print(f"  Added {len(jsw_records)} JSW ONE TMT records")
        print(f"    {JSW_550_BRAND}: {fe550_count} records")
        print(f"    {JSW_550D_BRAND}: {fe550d_count} records")

        # Add new states from price list
        jsw_states = set(r['s'] for r in jsw_records)
        new_states = jsw_states - set(metadata['states'])
        if new_states:
            metadata['states'] = sorted(set(metadata['states']) | new_states)
            print(f"    New states added: {new_states}")
            for ns in new_states:
                metadata['districts'][ns] = []

        # Insert JSW brands at front of topBrands and allBrands
        for brand_name in [JSW_550D_BRAND, JSW_550_BRAND]:
            if brand_name not in metadata['allBrands']:
                metadata['allBrands'].insert(0, brand_name)
            if brand_name not in metadata['topBrands']:
                metadata['topBrands'].insert(0, brand_name)

        # Ensure months from price list are present
        jsw_months = set(r['m'] for r in jsw_records)
        existing_months = set(metadata['months'])
        new_months = jsw_months - existing_months
        if new_months:
            all_months = list(existing_months | new_months)
            parsed_months = []
            for m in all_months:
                try:
                    d = datetime.strptime(m, '%b-%y')
                    parsed_months.append((d, m))
                except ValueError:
                    continue
            parsed_months.sort(key=lambda x: x[0])
            metadata['months'] = [m for _, m in parsed_months]

        metadata['totalRows'] = len(filterable_data)
        metadata['benchmarkBrand'] = JSW_550_BRAND
    else:
        print("  No JSW ONE TMT records found.")

    dashboard_data = {
        'metadata': metadata,
        'filterableData': filterable_data,
    }

    os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
    source_name = os.path.basename(excel_path)
    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(f'// Auto-generated by extract_data.py on {datetime.now().strftime("%Y-%m-%d %H:%M")}\n')
        f.write(f'// Source: {source_name}\n')
        f.write('// Filtered: non-null brand, non-null amount, amount in 30k-80k range, prices normalized to Excl. GST\n\n')
        f.write('const DASHBOARD_DATA = ')
        json.dump(dashboard_data, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')

    file_size = os.path.getsize(OUTPUT_PATH)
    print(f"\n6. Done!")
    print(f"  Output: {OUTPUT_PATH}")
    print(f"  Size: {file_size / 1024:.0f} KB | Rows: {len(filterable_data)}")
    print(f"\n  -> Refresh your browser to see the updated dashboard.")
    print("=" * 50)


if __name__ == '__main__':
    main()
