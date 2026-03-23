"""
generate_consolidated_map_v14.py
V14 — bare choropleth only. No header, no sidebar, no legends, no colorbar.
"""

import json
import os

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import openpyxl

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data_3grade.json")
EXCEL_PATH = os.path.join(BASE_DIR, "Data", "FY 27_AOP_V2.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_v14.png")

# ── State name mapping (data → GeoJSON STNAME) ────────────────────────
STATE_NAME_MAP = {
    "UTTAR PRADESH": "UTTAR PRADESH",
    "RAJASTHAN": "RAJASTHAN",
    "DELHI": "DELHI",
    "HARYANA": "HARYANA",
    "PUNJAB": "PUNJAB",
    "UTTARAKHAND": "UTTARAKHAND",
    "JAMMU AND KASHMIR": "JAMMU & KASHMIR",
    "HIMACHAL PRADESH": "HIMACHAL PRADESH",
    "ODISHA": "ODISHA",
    "MADHYA PRADESH": "MADHYA PRADESH",
    "MAHARASHTRA": "MAHARASHTRA",
    "JHARKHAND": "JHARKHAND",
    "WEST BENGAL": "WEST BENGAL",
    "BIHAR": "BIHAR",
    "GUJARAT": "GUJARAT",
    "TELANGANA": "TELANGANA",
    "ANDHRA PRADESH": "ANDHRA PRADESH",
    "ASSAM": "ASSAM",
    "CHHATTISGARH": "CHHATTISGARH",
}


def _read_fy27_sales():
    """Read FY27 sales from 'Market share - comparison Y-o-Y' tab, col T."""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Market share - comparison Y-o-Y"]
    sales = {}
    for row in ws.iter_rows(min_row=5, max_row=31, values_only=True):
        state_raw = row[1]
        fy_sales = row[19]
        if state_raw is None or fy_sales is None:
            continue
        state = str(state_raw).strip().upper()
        if state in ("", "TOTAL"):
            continue
        try:
            val = float(fy_sales)
        except (ValueError, TypeError):
            continue
        if val <= 0:
            continue
        geo = STATE_NAME_MAP.get(state)
        if geo:
            sales[geo] = val
    wb.close()
    return sales


# ── Custom yellow colormap (light yellow → deep amber) ────────────────
_YELLOW_CMAP = mcolors.LinearSegmentedColormap.from_list(
    'yellow_gradient',
    ['#FFFDE7', '#FFF9C4', '#FFF176', '#FFEE58', '#FDD835', '#F9A825', '#F57F17'],
)


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    gdf = gpd.read_file(GEOJSON_PATH)
    fy27_sales = _read_fy27_sales()

    # ── Single-axis figure, transparent background ─────────────────────
    fig, ax_map = plt.subplots(figsize=(30, 30))
    fig.patch.set_alpha(0.0)

    # ── Choropleth — colored by FY27 sales (yellow gradient) ─────────
    sales_states = set(fy27_sales.keys())
    gdf['sales'] = gdf['STNAME'].map(fy27_sales).fillna(0)

    gdf[~gdf['STNAME'].isin(sales_states)].plot(
        ax=ax_map, color='#ECEFF1', edgecolor='#B0BEC5', linewidth=0.5)

    cmap = _YELLOW_CMAP
    gdf_yes = gdf[gdf['STNAME'].isin(sales_states)].copy()
    if not gdf_yes.empty:
        vals = gdf_yes['sales'].values
        vmin = max(vals[vals > 0].min(), 1) if (vals > 0).any() else 1
        norm = mcolors.LogNorm(vmin=vmin, vmax=vals.max())
        gdf_yes.plot(ax=ax_map,
                     color=[cmap(norm(v)) if v > 0 else '#ECEFF1' for v in gdf_yes['sales']],
                     edgecolor='#BFA000', linewidth=0.8)

    # ── Map styling ───────────────────────────────────────────────────
    ax_map.set_xlim(68, 97)
    ax_map.set_ylim(6, 37)
    ax_map.set_aspect('equal')
    ax_map.axis('off')
    ax_map.patch.set_alpha(0.0)

    # ── Colorbar at bottom ────────────────────────────────────────────
    if sales_states:
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig.colorbar(sm, ax=ax_map, location='bottom', shrink=0.5,
                            pad=0.03, aspect=30)
        cbar.set_label("FY 2026-27 Sales ('000 MT)", fontsize=12, fontweight='bold')
        cbar.ax.tick_params(labelsize=10)

    # ── Save ──────────────────────────────────────────────────────────
    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', transparent=True,
                pad_inches=0.1)
    plt.close()
    print(f"V14 map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
