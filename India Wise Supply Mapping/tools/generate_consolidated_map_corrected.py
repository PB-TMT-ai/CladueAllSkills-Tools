"""
generate_consolidated_map_corrected.py
Corrected consolidated map — clean map with 3-grade breakdown (Fe 550 / OH / Fe 550D).
- Choropleth: states colored by total supply volume (green gradient)
- Factory icons at plant locations with numbered labels
- NO state labels, NO plant directory, NO state table
- Slim sidebar: color key + grand totals (3-grade: Fe 550, OH, Fe 550D)
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

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data_3grade.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_corrected.png")

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

# ── Plant locations (lon, lat) ─────────────────────────────────────────
PLANT_LOCATIONS = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (78.18, 26.22),
    "API Ispat And Powertech Private Limited": (81.6, 21.25),
    "Aditya Industries": (75.6, 30.7),
    "Ambashakti Industries Limited": (77.70, 28.45),
    "GERMAN GREEN STEEL AND POWER LIMITED": (78.0, 23.2),
    "Maharashtra - New Plant": (73.9, 18.5),
    "Rashmi Steel": (88.3, 22.6),
    "Real Ispat": (83.0, 25.3),
    "SKA Ispat Private Limited": (81.5, 27.5),
    "Telangana - New Plant": (78.5, 17.4),
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
    "Aditya Industries": "PB",
    "Ambashakti Industries Limited": "UP",
    "GERMAN GREEN STEEL AND POWER LIMITED": "MP",
    "Maharashtra - New Plant": "MH",
    "Rashmi Steel": "WB",
    "Real Ispat": "UP",
    "SKA Ispat Private Limited": "UP",
    "Telangana - New Plant": "TG",
}

# ── Display offsets for overlapping plant markers ──────────────────────
PLANT_DISPLAY_OFFSETS = {
    "Ambashakti Industries Limited": (2.0, 1.5),
    "SKA Ispat Private Limited": (1.5, 1.5),
    "Real Ispat": (1.8, 0.0),
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (-1.8, -0.5),
}

# ── State centroids (lon, lat) — tuned for grade label placement ───────
STATE_CENTROIDS = {
    "UTTAR PRADESH": (80.9, 27.0),
    "RAJASTHAN": (73.2, 26.0),
    "DELHI": (75.5, 28.65),
    "HARYANA": (75.0, 29.8),
    "PUNJAB": (74.2, 31.2),
    "UTTARAKHAND": (79.8, 30.6),
    "JAMMU & KASHMIR": (75.3, 34.0),
    "HIMACHAL PRADESH": (77.2, 32.3),
    "ODISHA": (84.0, 20.5),
    "MADHYA PRADESH": (78.6, 23.5),
    "MAHARASHTRA": (75.7, 19.5),
    "JHARKHAND": (85.3, 23.6),
    "WEST BENGAL": (87.9, 23.0),
    "BIHAR": (85.9, 25.6),
    "GUJARAT": (71.6, 22.3),
    "ANDHRA PRADESH": (79.7, 15.5),
    "TELANGANA": (79.0, 18.0),
    "ASSAM": (92.9, 26.1),
    "CHHATTISGARH": (81.9, 21.3),
}

# ── Colors ─────────────────────────────────────────────────────────────
COLOR_550 = '#D4A017'    # Gold — Fe 550
COLOR_OH = '#E65100'     # Deep Orange — One Helix
COLOR_550D = '#1565C0'   # Blue — Fe 550D
COLOR_MULTI = '#7B1FA2'  # Purple — multiple grades


def create_factory_icon(size=48):
    """Create a simple factory icon as a PIL image."""
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


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    gdf = gpd.read_file(GEOJSON_PATH)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    s550, s_oh, s550d, state_totals, pt_totals, pt_grades = _compute(plant_data)

    # ── Layout: large map + minimal right sidebar ──────────────────────
    fig = plt.figure(figsize=(30, 22), facecolor='white')
    gs = gridspec.GridSpec(1, 2, width_ratios=[88, 12], wspace=0.01)
    ax_map = fig.add_subplot(gs[0])
    ax_info = fig.add_subplot(gs[1])
    ax_info.axis('off')

    # ── Choropleth ────────────────────────────────────────────────────
    supply_states = set(state_totals.keys())
    gdf['volume'] = gdf['STNAME'].map(state_totals).fillna(0)

    gdf[~gdf['STNAME'].isin(supply_states)].plot(
        ax=ax_map, color='#ECEFF1', edgecolor='#B0BEC5', linewidth=0.5)

    gdf_yes = gdf[gdf['STNAME'].isin(supply_states)].copy()
    if not gdf_yes.empty:
        vals = gdf_yes['volume'].values
        vmin = max(vals[vals > 0].min(), 1) if (vals > 0).any() else 1
        norm = mcolors.LogNorm(vmin=vmin, vmax=vals.max())
        cmap = plt.cm.Greens
        gdf_yes.plot(ax=ax_map,
                     color=[cmap(norm(v)) if v > 0 else '#ECEFF1' for v in gdf_yes['volume']],
                     edgecolor='#4CAF50', linewidth=0.8)

    # ── Per-state grade volume labels ──────────────────────────────────
    for state, vol in state_totals.items():
        if state not in STATE_CENTROIDS:
            continue
        cx, cy = STATE_CENTROIDS[state]
        fe550 = s550.get(state, 0)
        oh = s_oh.get(state, 0)
        fe550d = s550d.get(state, 0)

        # Line 1: total volume
        line1 = f"{_fmt(vol)}"
        ax_map.text(
            cx, cy + 0.45, line1,
            fontsize=8, ha='center', va='center', fontweight='bold',
            color='#1B2631',
            bbox=dict(facecolor='white', alpha=0.85, edgecolor='#B0BEC5',
                      linewidth=0.5, pad=2.0, boxstyle='round,pad=0.4'),
            zorder=8,
        )

        # Line 2: grade split (Fe 550 / OH / Fe 550D)
        parts = []
        if fe550 > 0:
            parts.append(f"550:{_fmt(fe550)}")
        if oh > 0:
            parts.append(f"OH:{_fmt(oh)}")
        if fe550d > 0:
            parts.append(f"550D:{_fmt(fe550d)}")
        line2 = " | ".join(parts) if parts else ""

        if line2:
            ax_map.text(
                cx, cy - 0.55, line2,
                fontsize=6.5, ha='center', va='center', fontweight='bold',
                color='#546E7A',
                bbox=dict(facecolor='white', alpha=0.85, edgecolor='none',
                          pad=1.0, boxstyle='round,pad=0.2'),
                zorder=8,
            )

    # ── Factory icons at plant locations ──────────────────────────────
    factory_img = create_factory_icon(size=48)
    factory_arr = np.array(factory_img)

    plant_order = list(PLANT_LOCATIONS.keys())
    for i, pname in enumerate(plant_order, 1):
        true_lon, true_lat = PLANT_LOCATIONS[pname]
        offset = PLANT_DISPLAY_OFFSETS.get(pname)
        if offset:
            disp_lon = true_lon + offset[0]
            disp_lat = true_lat + offset[1]
            ax_map.plot([disp_lon, true_lon], [disp_lat, true_lat],
                        color='#78909C', linewidth=0.8, linestyle='--',
                        alpha=0.6, zorder=9)
            ax_map.plot(true_lon, true_lat, 'o', markersize=4,
                        color='#546E7A', zorder=9)
        else:
            disp_lon, disp_lat = true_lon, true_lat

        imagebox = OffsetImage(factory_arr, zoom=0.4)
        ab = AnnotationBbox(imagebox, (disp_lon, disp_lat), frameon=False, zorder=10)
        ax_map.add_artist(ab)

        ax_map.text(disp_lon, disp_lat - 0.7, str(i),
                    fontsize=7, ha='center', va='top', fontweight='bold',
                    color='#1B2631',
                    bbox=dict(facecolor='white', alpha=0.85, edgecolor='#90A4AE',
                              linewidth=0.4, pad=1.0, boxstyle='round,pad=0.2'),
                    zorder=11)

    # ── Map styling ───────────────────────────────────────────────────
    ax_map.set_xlim(68, 97)
    ax_map.set_ylim(6, 37)
    ax_map.set_aspect('equal')
    ax_map.axis('off')

    if supply_states:
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig.colorbar(sm, ax=ax_map, location='bottom', shrink=0.5,
                            pad=0.03, aspect=30)
        cbar.set_label('Total Inbound Supply (MT)', fontsize=10, fontweight='bold')
        cbar.ax.tick_params(labelsize=8)

    ax_map.set_title("JSW ONE TMT — Plant-wise Supply Network\nFY 2026-27",
                     fontsize=22, fontweight='bold', pad=20,
                     color='#1B2631', fontfamily='sans-serif')

    # ══════════════════════════════════════════════════════════════════
    # SIDEBAR — Color Key + Grand Totals only (no plant list)
    # ══════════════════════════════════════════════════════════════════
    y = 0.70  # start mid-page since sidebar is now very light

    # ── Color Key ─────────────────────────────────────────────────────
    ax_info.text(0.05, y, "COLOR KEY", fontsize=14, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.035

    for color, label in [(COLOR_550, "Fe 550"),
                          (COLOR_OH, "OH (One Helix)"),
                          (COLOR_550D, "Fe 550D"),
                          (COLOR_MULTI, "Multiple grades")]:
        ax_info.plot(0.08, y - 0.003, 's', markersize=14, color=color,
                     transform=ax_info.transAxes, clip_on=False)
        ax_info.text(0.18, y, label, fontsize=11, color='#37474F',
                     transform=ax_info.transAxes, va='top')
        y -= 0.035

    ax_info.text(0.08, y, "Darker green = more supply",
                 fontsize=10, color='#2E7D32', fontstyle='italic',
                 transform=ax_info.transAxes, va='top')

    # ── Grand Totals (3-grade) ────────────────────────────────────────
    y -= 0.055
    ax_info.plot([0.05, 0.95], [y, y], color='#B0BEC5', linewidth=0.8,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.035

    t550 = sum(s550.values())
    t_oh = sum(s_oh.values())
    t550d = sum(s550d.values())
    grand = t550 + t_oh + t550d

    ax_info.text(0.05, y, "GRADE BREAKDOWN", fontsize=14, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.038
    ax_info.text(0.05, y, f"Fe 550: {t550:>9,.0f} MT", fontsize=12,
                 color=COLOR_550, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"OH:     {t_oh:>9,.0f} MT", fontsize=12,
                 color=COLOR_OH, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"Fe 550D:{t550d:>9,.0f} MT", fontsize=12,
                 color=COLOR_550D, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.040
    ax_info.plot([0.05, 0.95], [y, y], color='#CFD8DC', linewidth=0.5,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.032
    ax_info.text(0.05, y, f"Total:  {grand:>9,.0f} MT", fontsize=13,
                 color='#1B2631', fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"Plants: {len(PLANT_LOCATIONS):>9}",
                 fontsize=11, color='#546E7A', fontfamily='monospace',
                 transform=ax_info.transAxes, va='top')

    # ── Save ──────────────────────────────────────────────────────────
    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', facecolor='white',
                pad_inches=0.3)
    plt.close()
    print(f"Corrected map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
