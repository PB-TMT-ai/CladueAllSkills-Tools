"""
Batch generate distributor one-pager reports for a given zone.
Loads the workbook ONCE, then iterates over all matching (distributor, state) pairs.

Usage:
    python tools/batch_generate.py              # North region, today's date
    python tools/batch_generate.py North         # North region, today's date
    python tools/batch_generate.py East 2026-03-09  # East region, specific date
"""

import sys
import os
import time
from datetime import date

# Add tools dir to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from extract_distributor_data import (
    load_workbook, load_district_master, load_achi_data,
    load_mou_targets, load_dealer_sales, load_invoice_data,
    load_smtm_data, extract_all_from_loaded, normalize
)
from generate_one_pager import generate_one_pager


def get_distributors_by_zone(district_rows, zone="North"):
    """Get unique (distributor_name, state) pairs for a given zone."""
    zone_norm = normalize(zone)
    pairs = set()
    for r in district_rows:
        if normalize(r['zone']) == zone_norm:
            pairs.add((r['distributor'], r['state']))
    return sorted(pairs)


def batch_generate(zone="North", output_date=None):
    """Generate one-pager reports for all distributors in a zone."""
    if output_date is None:
        output_date = date.today().strftime('%Y-%m-%d')

    output_dir = os.path.join(r"D:\Distributor One pager\output", output_date)
    os.makedirs(output_dir, exist_ok=True)

    # === LOAD WORKBOOK ONCE ===
    print("=" * 60)
    print(f"BATCH GENERATION: {zone} Region")
    print("=" * 60)

    t0 = time.time()
    print("\nLoading workbook and all sheets (one-time)...")
    wb = load_workbook()

    district_rows = load_district_master(wb)
    achi_rows = load_achi_data(wb)
    mou_rows = load_mou_targets(wb)
    dealer_rows = load_dealer_sales(wb)
    invoice_rows = load_invoice_data(wb)

    wb.close()

    smtm_rows = load_smtm_data()

    load_time = time.time() - t0
    print(f"Loaded in {load_time:.1f}s")

    # === GET DISTRIBUTOR LIST ===
    pairs = get_distributors_by_zone(district_rows, zone)
    total = len(pairs)
    print(f"\nFound {total} distributors in {zone} zone")
    print("-" * 60)

    # === GENERATE REPORTS ===
    success = 0
    errors = []

    for idx, (dist_name, state) in enumerate(pairs, 1):
        safe_name = dist_name.replace(' ', '_').replace('/', '_')
        safe_state = state.replace(' ', '_')
        filename = f"{safe_name}_{safe_state}.docx"
        output_path = os.path.join(output_dir, filename)

        try:
            t1 = time.time()
            data = extract_all_from_loaded(
                dist_name, state,
                district_rows, achi_rows, mou_rows,
                dealer_rows, invoice_rows, smtm_rows
            )
            generate_one_pager(data, output_path)
            elapsed = time.time() - t1
            print(f"  [{idx}/{total}] OK  {dist_name} ({state}) - {elapsed:.1f}s")
            success += 1
        except Exception as e:
            print(f"  [{idx}/{total}] ERR {dist_name} ({state}) - {e}")
            errors.append((dist_name, state, str(e)))

    # === SUMMARY ===
    total_time = time.time() - t0
    print("\n" + "=" * 60)
    print(f"BATCH COMPLETE")
    print(f"  Total:    {total}")
    print(f"  Success:  {success}")
    print(f"  Errors:   {len(errors)}")
    print(f"  Time:     {total_time:.1f}s")
    print(f"  Output:   {output_dir}")

    if errors:
        print(f"\nFailed distributors:")
        for dist, state, err in errors:
            print(f"  - {dist} ({state}): {err}")

    print("=" * 60)
    return success, errors


if __name__ == '__main__':
    zone = sys.argv[1] if len(sys.argv) > 1 else "North"
    output_date = sys.argv[2] if len(sys.argv) > 2 else None
    batch_generate(zone, output_date)
