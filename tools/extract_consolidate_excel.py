"""
Tool: extract_consolidate_excel.py
Purpose: Extract and consolidate all Excel data from North-Marketing-Execution-Reports
         into categorized CSV summaries.
Input: Directory path containing Excel files
Output: Consolidated CSVs in .tmp/ directory + summary stats to stdout
"""

import sys
import os
import pandas as pd
import openpyxl
import json
from datetime import datetime

REPORT_DIR = sys.argv[1] if len(sys.argv) > 1 else "/home/user/North-Marketing-Execution-Reports"
OUTPUT_DIR = "/home/user/CladueAllSkills-Tools/.tmp"

os.makedirs(OUTPUT_DIR, exist_ok=True)


def dedup_columns(df):
    """Normalize and deduplicate column names."""
    raw_cols = [str(c).strip().lower().replace('\n', ' ') for c in df.columns]
    seen = {}
    deduped = []
    for col in raw_cols:
        if col in seen:
            seen[col] += 1
            deduped.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            deduped.append(col)
    df.columns = deduped
    return df


def safe_read_sheet(filepath, sheet_name=None):
    """Read an Excel sheet, handling messy headers by scanning for the real header row."""
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if not rows:
            return pd.DataFrame()

        # Find the row that looks like a header (contains 'S.No' or 'S.N' or 'SR')
        header_idx = 0
        for i, row in enumerate(rows):
            row_str = " ".join([str(c).lower() for c in row if c is not None])
            if any(k in row_str for k in ['s.no', 's.n.', 'sr.', 'region', 'dealer', 'city']):
                header_idx = i
                break

        # Deduplicate headers
        raw_headers = [str(c).strip() if c else f"col_{j}" for j, c in enumerate(rows[header_idx])]
        seen = {}
        headers = []
        for h in raw_headers:
            if h in seen:
                seen[h] += 1
                headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                headers.append(h)
        data = rows[header_idx + 1:]

        # Filter out completely empty rows
        data = [r for r in data if any(c is not None and str(c).strip() for c in r)]

        # Truncate or pad rows to match header length
        cleaned = []
        for r in data:
            r = list(r)
            if len(r) < len(headers):
                r.extend([None] * (len(headers) - len(r)))
            elif len(r) > len(headers):
                r = r[:len(headers)]
            cleaned.append(r)

        df = pd.DataFrame(cleaned, columns=headers)
        return df
    except Exception as e:
        print(f"  ⚠ Error reading {os.path.basename(filepath)} [{sheet_name}]: {e}")
        return pd.DataFrame()


def classify_file(filename):
    """Classify a file into a category based on its name."""
    fn = filename.lower()
    if any(k in fn for k in ['impact wall', 'impact wp', 'jsw impact']):
        return 'impact_wall'
    elif any(k in fn for k in ['gsb', 'nlb', 'board', 'd2r']):
        return 'gsb_nlb_boards'
    elif any(k in fn for k in ['inshop', 'in-shop']):
        return 'inshop_branding'
    elif any(k in fn for k in ['bill no', 'bill ']):
        return 'bills'
    elif any(k in fn for k in ['wp', 'wall painting', 'wall &', 'dealer wall', 'shop painting', 'dealer wp', 'dealer_final', 'complete work']):
        return 'dealer_wall_painting'
    else:
        return 'other'


def extract_dealer_wp():
    """Extract dealer wall/shop painting data."""
    files_map = {
        'Bill no 145 Dealer WP (1).xlsx': [('Sheet1', None)],
        'JSW Complete Work 21.6.2025 (1).xlsx': [('Sheet1', None)],
        'JSW_Dealer WP_Final (1) (1).xlsx': [('HR', 'Haryana'), ('DL', 'Delhi')],
        'JSW One Tmt Dealer Wall & Shop Painting 1 March 2026 (2) (1).xlsx': [('Sheet1', None)],
        'PB & Jammu Wall painting (1) (2).xlsx': [('WP', None)],
        'UP WP (3).xlsx': [('WP', None)],
        'UP WP (4).xlsx': [('WP', None)],
        'WP - 100625 (4).xlsx': [('HR', 'Haryana'), ('NCR', 'NCR'), ('RAj', 'Rajasthan')],
        'WP-200625 (5).xlsx': [('Haryana', 'Haryana'), ('Rajasthan', 'Rajasthan'), ('NCR - UP', 'NCR-UP'), ('NCR-Delhi', 'Delhi')],
        'Bill No 321 (1).xlsx': [('HR', 'Haryana')],
    }

    all_rows = []
    for fname, sheets in files_map.items():
        fpath = os.path.join(REPORT_DIR, fname)
        if not os.path.exists(fpath):
            continue
        for sheet_name, region_label in sheets:
            df = safe_read_sheet(fpath, sheet_name)
            if df.empty:
                continue

            df = dedup_columns(df)
            df['_source_file'] = fname
            df['_source_sheet'] = sheet_name
            if region_label:
                df['_region_label'] = region_label
            all_rows.append(df)

    if all_rows:
        combined = pd.concat(all_rows, ignore_index=True, sort=False)
        return combined
    return pd.DataFrame()


def extract_impact_wall():
    """Extract impact wall painting data."""
    files_map = {
        'Impact WP - 100625 (4).xlsx': [('HR', 'Haryana'), ('Ghaziabad', 'Ghaziabad')],
        'JSW IMPACT - 080725.xlsx': [('Haryana', 'Haryana')],
        'JSW Impact Wall Xls Jaipur Area (1).xlsx': [('Sheet1', 'Rajasthan')],
        'UP Impact Wall (3).xlsx': [('Sheet1', 'UP')],
        'JSW One Tmt Impact Wall 01 March 2026 (1) (1).xlsx': [('Sheet1', None)],
        'JSW One Tmt Impact Wall 31 December 2025 (3).xlsx': [('Sheet1', None)],
    }

    all_rows = []
    for fname, sheets in files_map.items():
        fpath = os.path.join(REPORT_DIR, fname)
        if not os.path.exists(fpath):
            continue
        for sheet_name, region_label in sheets:
            df = safe_read_sheet(fpath, sheet_name)
            if df.empty:
                continue
            df = dedup_columns(df)
            df['_source_file'] = fname
            df['_source_sheet'] = sheet_name
            if region_label:
                df['_region_label'] = region_label
            all_rows.append(df)

    if all_rows:
        return pd.concat(all_rows, ignore_index=True, sort=False)
    return pd.DataFrame()


def extract_gsb_nlb():
    """Extract GSB/NLB board installation data."""
    files = [
        'D2R NLB LIST UPDATED TILL JUNE 25 FINAL (1) (1) (1).xlsx',
        'Delhi HR Boards_5th Feb_Invoice Final (1).xlsx',
        'GSB -  June 2025_300625 (2).xlsx',
        'GSB+ NLB (1).xlsx',
        'JSW GSB NLB Board FINAL 22 jan 2026 (3).xlsx',
        'JSW One Tmt Dealer GSB ,NLB Installation 8 July 2025.xlsx',
    ]

    all_rows = []
    for fname in files:
        fpath = os.path.join(REPORT_DIR, fname)
        if not os.path.exists(fpath):
            continue
        df = safe_read_sheet(fpath)
        if df.empty:
            continue
        df = dedup_columns(df)
        df['_source_file'] = fname
        all_rows.append(df)

    if all_rows:
        return pd.concat(all_rows, ignore_index=True, sort=False)
    return pd.DataFrame()


def extract_inshop():
    """Extract in-shop branding data."""
    fpath = os.path.join(REPORT_DIR, 'inshop (PB, HR, RJ, Delhi ) (1).xlsx')
    if not os.path.exists(fpath):
        return pd.DataFrame()
    df = safe_read_sheet(fpath)
    df = dedup_columns(df)
    df['_source_file'] = os.path.basename(fpath)
    return df


def extract_bills():
    """Extract bill/invoice data."""
    files = [
        'Bill no 119 (1).xlsx',
        'Bill no 121 (1).xlsx',
    ]
    all_rows = []
    for fname in files:
        fpath = os.path.join(REPORT_DIR, fname)
        if not os.path.exists(fpath):
            continue
        df = safe_read_sheet(fpath)
        if df.empty:
            continue
        df = dedup_columns(df)
        df['_source_file'] = fname
        all_rows.append(df)

    if all_rows:
        return pd.concat(all_rows, ignore_index=True, sort=False)
    return pd.DataFrame()


def generate_summary(name, df):
    """Generate and print summary stats for a dataframe."""
    print(f"\n{'='*80}")
    print(f"  {name.upper()}")
    print(f"{'='*80}")
    print(f"  Total rows: {len(df)}")
    print(f"  Columns: {list(df.columns)[:15]}{'...' if len(df.columns) > 15 else ''}")

    # State breakdown if available
    state_col = None
    for c in df.columns:
        if 'state' in c.lower():
            state_col = c
            break
    if state_col:
        states = df[state_col].dropna().astype(str).str.strip()
        states = states[states != '']
        if len(states) > 0:
            print(f"\n  State breakdown:")
            for s, count in states.value_counts().head(10).items():
                print(f"    {s}: {count}")

    # Source files
    if '_source_file' in df.columns:
        print(f"\n  Source files ({df['_source_file'].nunique()}):")
        for f in df['_source_file'].unique():
            count = len(df[df['_source_file'] == f])
            print(f"    {f}: {count} rows")


def main():
    print("=" * 80)
    print("  NORTH MARKETING EXECUTION REPORTS - EXCEL DATA CONSOLIDATION")
    print(f"  Source: {REPORT_DIR}")
    print(f"  Output: {OUTPUT_DIR}")
    print(f"  Run at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

    # 1. Dealer Wall Painting
    print("\n[1/5] Extracting Dealer Wall/Shop Painting data...")
    wp_df = extract_dealer_wp()
    if not wp_df.empty:
        wp_df.to_csv(os.path.join(OUTPUT_DIR, "consolidated_dealer_wall_painting.csv"), index=False)
        generate_summary("Dealer Wall/Shop Painting", wp_df)

    # 2. Impact Wall
    print("\n[2/5] Extracting Impact Wall data...")
    impact_df = extract_impact_wall()
    if not impact_df.empty:
        impact_df.to_csv(os.path.join(OUTPUT_DIR, "consolidated_impact_wall.csv"), index=False)
        generate_summary("Impact Wall Painting", impact_df)

    # 3. GSB/NLB
    print("\n[3/5] Extracting GSB/NLB Board data...")
    gsb_df = extract_gsb_nlb()
    if not gsb_df.empty:
        gsb_df.to_csv(os.path.join(OUTPUT_DIR, "consolidated_gsb_nlb_boards.csv"), index=False)
        generate_summary("GSB/NLB Boards", gsb_df)

    # 4. In-shop
    print("\n[4/5] Extracting In-shop Branding data...")
    inshop_df = extract_inshop()
    if not inshop_df.empty:
        inshop_df.to_csv(os.path.join(OUTPUT_DIR, "consolidated_inshop_branding.csv"), index=False)
        generate_summary("In-shop Branding", inshop_df)

    # 5. Bills
    print("\n[5/5] Extracting Bills data...")
    bills_df = extract_bills()
    if not bills_df.empty:
        bills_df.to_csv(os.path.join(OUTPUT_DIR, "consolidated_bills.csv"), index=False)
        generate_summary("Bills/Invoices", bills_df)

    # Overall summary
    print("\n" + "=" * 80)
    print("  CONSOLIDATION COMPLETE")
    print("=" * 80)
    totals = {
        "Dealer Wall/Shop Painting": len(wp_df),
        "Impact Wall Painting": len(impact_df),
        "GSB/NLB Boards": len(gsb_df),
        "In-shop Branding": len(inshop_df),
        "Bills/Invoices": len(bills_df),
    }
    for name, count in totals.items():
        print(f"  {name}: {count} rows")
    print(f"  TOTAL: {sum(totals.values())} rows")
    print(f"\n  Output files:")
    for f in sorted(os.listdir(OUTPUT_DIR)):
        if f.endswith('.csv'):
            size = os.path.getsize(os.path.join(OUTPUT_DIR, f))
            print(f"    {f} ({size:,} bytes)")


if __name__ == "__main__":
    main()
