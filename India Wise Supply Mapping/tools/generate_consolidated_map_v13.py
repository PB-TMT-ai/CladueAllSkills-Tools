"""
generate_consolidated_map_v13.py
V13 consolidated map — no plant icons or labels, choropleth only.
- Choropleth: states colored by FY27 sales (YELLOW gradient)
- Factory icons DOUBLED in size at plant locations with numbered labels
- Ambashakti Industries moved to Ghaziabad, UP
- Chhattisgarh plants spread for better representation
- Slim sidebar: color key (plant counts) + grade breakdown totals
"""

import json
import os

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.gridspec as gridspec
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import numpy as np
from PIL import Image, ImageDraw
import openpyxl

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data_3grade.json")
EXCEL_PATH = os.path.join(BASE_DIR, "Data", "FY 27_AOP_V2.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_v13.png")

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

# ── State label positions (lon, lat) ──────────────────────────────────
STATE_LABEL_POS = {
    "UTTAR PRADESH": (80.9, 27.0),
    "RAJASTHAN": (72.5, 26.0),
    "DELHI": (73.5, 28.5),
    "HARYANA": (73.2, 30.4),
    "PUNJAB": (72.0, 32.2),
    "UTTARAKHAND": (80.5, 31.2),
    "JAMMU & KASHMIR": (75.3, 34.5),
    "HIMACHAL PRADESH": (77.8, 33.2),
    "ODISHA": (84.0, 20.5),
    "MADHYA PRADESH": (78.6, 23.5),
    "MAHARASHTRA": (75.7, 19.3),
    "JHARKHAND": (85.5, 23.6),
    "WEST BENGAL": (88.5, 22.0),
    "BIHAR": (86.0, 25.6),
    "GUJARAT": (70.5, 22.3),
    "ANDHRA PRADESH": (79.7, 15.0),
    "TELANGANA": (79.0, 17.8),
    "ASSAM": (93.0, 26.1),
    "CHHATTISGARH": (82.2, 21.3),
}

STATE_ABBREVS = {
    "UTTAR PRADESH": "UP", "RAJASTHAN": "RAJ", "DELHI": "DEL",
    "HARYANA": "HR", "PUNJAB": "PB", "UTTARAKHAND": "UK",
    "JAMMU & KASHMIR": "J&K", "HIMACHAL PRADESH": "HP",
    "ODISHA": "OD", "MADHYA PRADESH": "MP", "MAHARASHTRA": "MH",
    "JHARKHAND": "JH", "WEST BENGAL": "WB", "BIHAR": "BR",
    "GUJARAT": "GJ", "ANDHRA PRADESH": "AP", "TELANGANA": "TG",
    "ASSAM": "AS", "CHHATTISGARH": "CG",
}

# ── Plant locations (lon, lat) ─────────────────────────────────────────
# V11: Ambashakti Industries → Ghaziabad; CG plants spread out
PLANT_LOCATIONS = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (78.18, 26.22),     # Gwalior, Madhya Pradesh
    "API Ispat And Powertech Private Limited": (81.63, 21.25),# Raipur, Chhattisgarh
    "Aditya Industries": (77.31, 30.52),                      # Kala Amb, Himachal Pradesh
    "Ambashakti Industries Limited": (77.44, 28.67),          # Ghaziabad, Uttar Pradesh
    "GERMAN GREEN STEEL AND POWER LIMITED": (69.86, 23.24),   # Kutch, Gujarat
    "Maharashtra - New Plant": (73.9, 18.5),                  # Maharashtra
    "Rashmi Steel": (87.33, 22.33),                           # Kharagpur, West Bengal
    "Real Ispat": (81.63, 21.25),                             # Raipur, Chhattisgarh
    "SKA Ispat Private Limited": (81.63, 21.25),              # Raipur, Chhattisgarh
    "Telangana - New Plant": (78.5, 17.4),                    # Telangana
}

PLANT_SHORT_NAMES = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": "Ambashakti Gwalior",
    "API Ispat And Powertech Private Limited": "API Ispat",
    "Aditya Industries": "Aditya Industries",
    "Ambashakti Industries Limited": "Ambashakti UP",
    "GERMAN GREEN STEEL AND POWER LIMITED": "German Green Steel",
    "Maharashtra - New Plant": "Maharashtra Plant",
    "Rashmi Steel": "Rashmi Steel",
    "Real Ispat": "Real Ispat",
    "SKA Ispat Private Limited": "SKA Ispat",
    "Telangana - New Plant": "Telangana Plant",
}

PLANT_STATES = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": "MP",
    "API Ispat And Powertech Private Limited": "CG",
    "Aditya Industries": "HP",
    "Ambashakti Industries Limited": "UP",
    "GERMAN GREEN STEEL AND POWER LIMITED": "GJ",
    "Maharashtra - New Plant": "MH",
    "Rashmi Steel": "WB",
    "Real Ispat": "CG",
    "SKA Ispat Private Limited": "CG",
    "Telangana - New Plant": "TG",
}

# ── Display offsets for overlapping CG plants ──────────────────────────
# All 3 CG plants share (81.63, 21.25) — spread them visually
PLANT_DISPLAY_OFFSETS = {
    "API Ispat And Powertech Private Limited": (-1.5, 0.8),   # upper-left
    "SKA Ispat Private Limited": (1.5, 0.8),                  # upper-right
    "Real Ispat": (0.0, -1.2),                                # below
}

# ── Colors ─────────────────────────────────────────────────────────────
COLOR_550 = '#D4A017'    # Gold — Fe 550
COLOR_OH = '#E65100'     # Deep Orange — One Helix
COLOR_550D = '#1565C0'   # Blue — Fe 550D
COLOR_MULTI = '#7B1FA2'  # Purple — multiple grades


def _read_fy27_sales():
    """Read FY27 sales from 'Market share - comparison Y-o-Y' tab, col T (state in col B)."""
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True, read_only=True)
    ws = wb["Market share - comparison Y-o-Y"]
    sales = {}
    for row in ws.iter_rows(min_row=5, max_row=31, values_only=True):
        state_raw = row[1]   # Column B (0-indexed = 1)
        fy_sales = row[19]   # Column T (0-indexed = 19)
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


def create_factory_icon(size=96):
    """Create a simple factory icon as a PIL image. V11: doubled size (96px)."""
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    s = size
    draw.rectangle([s*0.1, s*0.35, s*0.65, s*0.9], fill=(40, 40, 40, 255))
    draw.rectangle([s*0.15, s*0.05, s*0.3, s*0.35], fill=(60, 60, 60, 255))
    draw.rectangle([s*0.4, s*0.15, s*0.55, s*0.35], fill=(60, 60, 60, 255))
    draw.polygon([(s*0.65, s*0.35), (s*0.65, s*0.9), (s*0.9, s*0.9), (s*0.9, s*0.55)],
                 fill=(50, 50, 50, 255))
    draw.rectangle([s*0.2, s*0.5, s*0.35, s*0.65], fill=(255, 220, 50, 255))
    draw.rectangle([s*0.42, s*0.5, s*0.57, s*0.65], fill=(255, 220, 50, 255))
    for cx, cy, r in [(s*0.22, s*0.02, s*0.06), (s*0.17, s*0.0, s*0.04),
                       (s*0.47, s*0.1, s*0.05)]:
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(180, 180, 180, 200))
    return img


def _compute(plant_data):
    """Compute state & plant aggregations with 3 grades."""
    s550, s_oh, s550d = {}, {}, {}
    for pdata in plant_data.values():
        for state, vol in pdata.get("fe_550", {}).items():
            geo = STATE_NAME_MAP.get(state)
            if geo:
                s550[geo] = s550.get(geo, 0) + vol
        for state, vol in pdata.get("one_helix", {}).items():
            geo = STATE_NAME_MAP.get(state)
            if geo:
                s_oh[geo] = s_oh.get(geo, 0) + vol
        for state, vol in pdata.get("fe_550d", {}).items():
            geo = STATE_NAME_MAP.get(state)
            if geo:
                s550d[geo] = s550d.get(geo, 0) + vol

    all_states = set(s550) | set(s_oh) | set(s550d)
    state_totals = {s: s550.get(s, 0) + s_oh.get(s, 0) + s550d.get(s, 0) for s in all_states}

    pt, pg = {}, {}
    for pn, pd in plant_data.items():
        if pn not in PLANT_LOCATIONS:
            continue
        a = sum(pd.get("fe_550", {}).values())
        b = sum(pd.get("one_helix", {}).values())
        c = sum(pd.get("fe_550d", {}).values())
        pt[pn] = a + b + c
        grades = []
        if a > 0:
            grades.append('fe550')
        if b > 0:
            grades.append('oh')
        if c > 0:
            grades.append('fe550d')
        if len(grades) > 1:
            pg[pn] = 'multi'
        elif len(grades) == 1:
            pg[pn] = grades[0]
        else:
            pg[pn] = 'fe550'

    return s550, s_oh, s550d, state_totals, pt, pg


def _fmt(v):
    return f"{v/1000:.1f}K" if v >= 1000 else f"{v:.0f}"


# ── Custom yellow colormap (light yellow → deep amber) ────────────────
_YELLOW_CMAP = mcolors.LinearSegmentedColormap.from_list(
    'yellow_gradient',
    ['#FFFDE7', '#FFF9C4', '#FFF176', '#FFEE58', '#FDD835', '#F9A825', '#F57F17'],
)


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    gdf = gpd.read_file(GEOJSON_PATH)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    s550, s_oh, s550d, state_totals, pt_totals, pt_grades = _compute(plant_data)

    # ── Read FY27 sales for choropleth gradient ───────────────────────
    fy27_sales = _read_fy27_sales()

    # ── Layout ────────────────────────────────────────────────────────
    fig = plt.figure(figsize=(42, 30), facecolor='white')
    gs = gridspec.GridSpec(1, 2, width_ratios=[85, 15], wspace=0.01)
    ax_map = fig.add_subplot(gs[0])
    ax_info = fig.add_subplot(gs[1])
    ax_info.axis('off')

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

    if sales_states:
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig.colorbar(sm, ax=ax_map, location='bottom', shrink=0.5,
                            pad=0.03, aspect=30)
        cbar.set_label("FY 2026-27 Sales ('000 MT)", fontsize=12, fontweight='bold')
        cbar.ax.tick_params(labelsize=10)

    ax_map.set_title("JSW ONE TMT — Plant-wise Supply Network\nFY 2026-27",
                     fontsize=26, fontweight='bold', pad=24,
                     color='#1B2631', fontfamily='sans-serif')

    # ══════════════════════════════════════════════════════════════════
    # SIDEBAR — Color Key (plant counts) + Grade Breakdown
    # ══════════════════════════════════════════════════════════════════
    y = 0.70

    # ── Count plants per grade ────────────────────────────────────────
    n550, n_oh, n550d, n_multi = 0, 0, 0, 0
    for pn, pd in plant_data.items():
        if pn not in PLANT_LOCATIONS:
            continue
        has550 = sum(pd.get("fe_550", {}).values()) > 0
        has_oh = sum(pd.get("one_helix", {}).values()) > 0
        has550d = sum(pd.get("fe_550d", {}).values()) > 0
        cnt = int(has550) + int(has_oh) + int(has550d)
        if cnt > 1:
            n_multi += 1
        if has550:
            n550 += 1
        if has_oh:
            n_oh += 1
        if has550d:
            n550d += 1

    # ── Color Key ─────────────────────────────────────────────────────
    ax_info.text(0.05, y, "COLOR KEY", fontsize=16, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.035

    for color, label, count in [
        (COLOR_550,  "Fe 550",          n550),
        (COLOR_OH,   "OH (One Helix)",  n_oh),
        (COLOR_550D, "Fe 550D",         n550d),
        (COLOR_MULTI, "Multiple grades", n_multi),
    ]:
        ax_info.plot(0.08, y - 0.003, 's', markersize=16, color=color,
                     transform=ax_info.transAxes, clip_on=False)
        ax_info.text(0.18, y, f"{label}  ({count} plants)", fontsize=13,
                     color='#37474F', transform=ax_info.transAxes, va='top')
        y -= 0.035

    ax_info.text(0.08, y, "Darker shade = higher FY sales",
                 fontsize=11, color='#8D6E00', fontstyle='italic',
                 transform=ax_info.transAxes, va='top')

    # ── Grade Breakdown (3-grade) ─────────────────────────────────────
    y -= 0.055
    ax_info.plot([0.05, 0.95], [y, y], color='#B0BEC5', linewidth=0.8,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.035

    t550 = sum(s550.values())
    t_oh = sum(s_oh.values())
    t550d = sum(s550d.values())
    grand = t550 + t_oh + t550d

    ax_info.text(0.05, y, "GRADE BREAKDOWN", fontsize=16, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.038
    ax_info.text(0.05, y, f"Fe 550: {t550:>9,.0f} MT", fontsize=14,
                 color=COLOR_550, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"OH:     {t_oh:>9,.0f} MT", fontsize=14,
                 color=COLOR_OH, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"Fe 550D:{t550d:>9,.0f} MT", fontsize=14,
                 color=COLOR_550D, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.040
    ax_info.plot([0.05, 0.95], [y, y], color='#CFD8DC', linewidth=0.5,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.032
    ax_info.text(0.05, y, f"Total:  {grand:>9,.0f} MT", fontsize=15,
                 color='#1B2631', fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"Plants: {len(PLANT_LOCATIONS):>9}",
                 fontsize=13, color='#546E7A', fontfamily='monospace',
                 transform=ax_info.transAxes, va='top')

    # ── Save ──────────────────────────────────────────────────────────
    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', facecolor='white',
                pad_inches=0.3)
    plt.close()
    print(f"V13 map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
