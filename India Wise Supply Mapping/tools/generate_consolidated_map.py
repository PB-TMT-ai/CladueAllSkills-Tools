"""
generate_consolidated_map.py
Single consolidated India map — choropleth + factory-icon plant markers + compact legend.
- States colored by total inbound supply volume (green gradient)
- Factory icons at each plant location (with number label)
- Compact right legend: plant directory + color key + totals
"""

import json
import math
import os

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.colors as mcolors
import matplotlib.gridspec as gridspec
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import numpy as np
from PIL import Image, ImageDraw

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map.png")

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

# ── State centroids (lon, lat) — tuned to avoid label overlaps ────────
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

# ── State abbreviations for map labels ─────────────────────────────────
STATE_ABBREVS = {
    "UTTAR PRADESH": "UP",
    "RAJASTHAN": "RAJ",
    "DELHI": "DEL",
    "HARYANA": "HR",
    "PUNJAB": "PB",
    "UTTARAKHAND": "UK",
    "JAMMU & KASHMIR": "J&K",
    "HIMACHAL PRADESH": "HP",
    "ODISHA": "OD",
    "MADHYA PRADESH": "MP",
    "MAHARASHTRA": "MH",
    "JHARKHAND": "JH",
    "WEST BENGAL": "WB",
    "BIHAR": "BR",
    "GUJARAT": "GJ",
    "ANDHRA PRADESH": "AP",
    "TELANGANA": "TG",
    "ASSAM": "AS",
    "CHHATTISGARH": "CG",
}

# ── Plant locations (lon, lat) ─────────────────────────────────────────
PLANT_LOCATIONS = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (78.18, 26.22),
    "API Ispat And Powertech Private Limited": (81.6, 21.25),
    "Aditya Industries": (75.6, 30.7),
    "Ambashakti Industries Limited": (79.8, 27.2),
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

# ── Display offsets for overlapping plant markers (lon_offset, lat_offset) ──
PLANT_DISPLAY_OFFSETS = {
    "Ambashakti Industries Limited": (-1.5, 1.2),
    "SKA Ispat Private Limited": (1.5, 1.5),
    "Real Ispat": (1.8, 0.0),
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (-1.8, -0.5),
}

# ── Colors ─────────────────────────────────────────────────────────────
COLOR_550 = '#D4A017'
COLOR_550D = '#1565C0'
COLOR_BOTH = '#7B1FA2'


def create_factory_icon(size=48):
    """Create a simple factory icon as a PIL image."""
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    s = size
    # Main building body
    draw.rectangle([s*0.1, s*0.35, s*0.65, s*0.9], fill=(40, 40, 40, 255))
    # Chimney 1
    draw.rectangle([s*0.15, s*0.05, s*0.3, s*0.35], fill=(60, 60, 60, 255))
    # Chimney 2
    draw.rectangle([s*0.4, s*0.15, s*0.55, s*0.35], fill=(60, 60, 60, 255))
    # Roof / sawtooth
    draw.polygon([(s*0.65, s*0.35), (s*0.65, s*0.9), (s*0.9, s*0.9), (s*0.9, s*0.55)],
                 fill=(50, 50, 50, 255))
    # Windows
    draw.rectangle([s*0.2, s*0.5, s*0.35, s*0.65], fill=(255, 220, 50, 255))
    draw.rectangle([s*0.42, s*0.5, s*0.57, s*0.65], fill=(255, 220, 50, 255))
    # Smoke puffs
    for cx, cy, r in [(s*0.22, s*0.02, s*0.06), (s*0.17, s*0.0, s*0.04),
                       (s*0.47, s*0.1, s*0.05)]:
        draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(180, 180, 180, 200))
    return img


def _compute_aggregations(plant_data):
    """Compute state-level and plant-level volume aggregations."""
    state_fe550 = {}
    state_fe550d = {}
    for pdata in plant_data.values():
        for state, vol in pdata.get("fe_550", {}).items():
            geo = STATE_NAME_MAP.get(state)
            if geo:
                state_fe550[geo] = state_fe550.get(geo, 0) + vol
        for state, vol in pdata.get("fe_550d", {}).items():
            geo = STATE_NAME_MAP.get(state)
            if geo:
                state_fe550d[geo] = state_fe550d.get(geo, 0) + vol

    state_totals = {}
    all_states = set(state_fe550) | set(state_fe550d)
    for s in all_states:
        state_totals[s] = state_fe550.get(s, 0) + state_fe550d.get(s, 0)

    plant_totals = {}
    plant_grades = {}
    for pname, pdata in plant_data.items():
        if pname not in PLANT_LOCATIONS:
            continue
        t550 = sum(pdata.get("fe_550", {}).values())
        t550d = sum(pdata.get("fe_550d", {}).values())
        plant_totals[pname] = t550 + t550d
        if t550 > 0 and t550d > 0:
            plant_grades[pname] = 'both'
        elif t550d > 0:
            plant_grades[pname] = 'fe550d'
        else:
            plant_grades[pname] = 'fe550'

    return state_fe550, state_fe550d, state_totals, plant_totals, plant_grades


def _fmt_vol(vol):
    """Format volume: 142,843 → '142.8K', 7,958 → '8.0K'"""
    if vol >= 1000:
        return f"{vol/1000:.1f}K"
    return f"{vol:.0f}"


def generate_consolidated_map():
    """Generate the consolidated supply map."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    gdf = gpd.read_file(GEOJSON_PATH)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    (state_fe550, state_fe550d, state_totals,
     plant_totals, plant_grades) = _compute_aggregations(plant_data)

    # ── Figure layout: large map + slim right legend ─────────────────────
    fig = plt.figure(figsize=(30, 22), facecolor='white')
    gs = gridspec.GridSpec(1, 2, width_ratios=[78, 22], wspace=0.01)
    ax_map = fig.add_subplot(gs[0])
    ax_info = fig.add_subplot(gs[1])
    ax_info.axis('off')

    # ── Choropleth: color supply states by total volume ────────────────
    supply_states = set(state_totals.keys())
    gdf['volume'] = gdf['STNAME'].map(state_totals).fillna(0)

    gdf_no = gdf[~gdf['STNAME'].isin(supply_states)]
    gdf_no.plot(ax=ax_map, color='#ECEFF1', edgecolor='#B0BEC5', linewidth=0.5)

    gdf_yes = gdf[gdf['STNAME'].isin(supply_states)].copy()
    if not gdf_yes.empty:
        vol_values = gdf_yes['volume'].values
        vmin = max(vol_values[vol_values > 0].min(), 1) if (vol_values > 0).any() else 1
        vmax = vol_values.max()
        norm = mcolors.LogNorm(vmin=vmin, vmax=vmax)
        cmap = plt.cm.Greens
        colors = [cmap(norm(v)) if v > 0 else '#ECEFF1' for v in gdf_yes['volume']]
        gdf_yes.plot(ax=ax_map, color=colors, edgecolor='#4CAF50', linewidth=0.8)

    # ── State volume labels (with Fe 550 / Fe 550D split) ────────────────
    for state, vol in state_totals.items():
        if state not in STATE_CENTROIDS:
            continue
        cx, cy = STATE_CENTROIDS[state]
        abbrev = STATE_ABBREVS.get(state, state[:3])
        fe550 = state_fe550.get(state, 0)
        fe550d = state_fe550d.get(state, 0)

        line1 = f"{abbrev}  {_fmt_vol(vol)}"
        parts = []
        if fe550 > 0:
            parts.append(f"550: {_fmt_vol(fe550)}")
        if fe550d > 0:
            parts.append(f"550D: {_fmt_vol(fe550d)}")
        line2 = "  |  ".join(parts) if len(parts) == 2 else (parts[0] if parts else "")

        ax_map.text(
            cx, cy + 0.35, line1,
            fontsize=8, ha='center', va='center', fontweight='bold',
            color='#1B2631',
            bbox=dict(facecolor='white', alpha=0.85, edgecolor='#B0BEC5',
                      linewidth=0.5, pad=2.0, boxstyle='round,pad=0.4'),
            zorder=8,
        )
        if line2:
            ax_map.text(
                cx, cy - 0.55, line2,
                fontsize=6.5, ha='center', va='center', fontweight='bold',
                color='#546E7A',
                bbox=dict(facecolor='white', alpha=0.85, edgecolor='none',
                          pad=1.0, boxstyle='round,pad=0.2'),
                zorder=8,
            )

    # ── Factory icons at plant locations ─────────────────────────────────
    factory_img = create_factory_icon(size=48)
    factory_arr = np.array(factory_img)

    plant_order = list(PLANT_LOCATIONS.keys())
    for i, pname in enumerate(plant_order, 1):
        true_lon, true_lat = PLANT_LOCATIONS[pname]
        offset = PLANT_DISPLAY_OFFSETS.get(pname)
        if offset:
            disp_lon = true_lon + offset[0]
            disp_lat = true_lat + offset[1]
            # Leader line from display position to true location
            ax_map.plot(
                [disp_lon, true_lon], [disp_lat, true_lat],
                color='#78909C', linewidth=0.8, linestyle='--',
                alpha=0.6, zorder=9,
            )
            ax_map.plot(
                true_lon, true_lat, 'o', markersize=4,
                color='#546E7A', zorder=9,
            )
        else:
            disp_lon, disp_lat = true_lon, true_lat

        # Factory icon
        imagebox = OffsetImage(factory_arr, zoom=0.4)
        ab = AnnotationBbox(imagebox, (disp_lon, disp_lat), frameon=False, zorder=10)
        ax_map.add_artist(ab)

        # Number label below icon
        ax_map.text(
            disp_lon, disp_lat - 0.7, str(i),
            fontsize=7, ha='center', va='top', fontweight='bold',
            color='#1B2631',
            bbox=dict(facecolor='white', alpha=0.85, edgecolor='#90A4AE',
                      linewidth=0.4, pad=1.0, boxstyle='round,pad=0.2'),
            zorder=11,
        )

    # ── Map bounds & styling ───────────────────────────────────────────
    ax_map.set_xlim(68, 97)
    ax_map.set_ylim(6, 37)
    ax_map.set_aspect('equal')
    ax_map.axis('off')

    # ── Colorbar below map ─────────────────────────────────────────────
    if supply_states:
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig.colorbar(
            sm, ax=ax_map, location='bottom', shrink=0.5,
            pad=0.03, aspect=30,
        )
        cbar.set_label('Total Inbound Supply (MT)', fontsize=10, fontweight='bold')
        cbar.ax.tick_params(labelsize=8)

    # ── Title ──────────────────────────────────────────────────────────
    ax_map.set_title(
        "JSW ONE TMT — Plant-wise Supply Network\nFY 2026-27",
        fontsize=22, fontweight='bold', pad=20,
        color='#1B2631', fontfamily='sans-serif',
    )

    # ══════════════════════════════════════════════════════════════════════
    # SLIM RIGHT LEGEND (Plant Directory + Color Key + Totals only)
    # ══════════════════════════════════════════════════════════════════════

    y = 0.96

    # ── Plant Directory ─────────────────────────────────────────────────
    ax_info.text(0.05, y, "PLANT DIRECTORY", fontsize=13, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.03

    for i, pname in enumerate(plant_order, 1):
        grade = plant_grades.get(pname, 'fe550')
        if grade == 'both':
            marker_color = COLOR_BOTH
        elif grade == 'fe550d':
            marker_color = COLOR_550D
        else:
            marker_color = COLOR_550

        short = PLANT_SHORT_NAMES.get(pname, pname[:20])
        pstate = PLANT_STATES.get(pname, "")
        vol = plant_totals.get(pname, 0)

        # Number badge (colored by grade)
        ax_info.text(
            0.04, y - 0.005, str(i), fontsize=8, ha='center', va='center',
            fontweight='bold', color='white',
            bbox=dict(facecolor=marker_color, edgecolor='none',
                      pad=2.5, boxstyle='round,pad=0.3'),
            transform=ax_info.transAxes, zorder=6,
        )
        # Plant name + state + volume
        ax_info.text(
            0.10, y, f"{short} ({pstate})",
            fontsize=10, fontweight='bold', color='#1B2631',
            transform=ax_info.transAxes, va='top',
        )
        ax_info.text(
            0.10, y - 0.022, f"{vol:,.0f} MT",
            fontsize=8.5, color='#546E7A',
            transform=ax_info.transAxes, va='top',
        )
        y -= 0.052

    # ── Divider ────────────────────────────────────────────────────────
    y -= 0.01
    ax_info.plot([0.05, 0.95], [y, y], color='#B0BEC5', linewidth=0.8,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.025

    # ── Color Key ──────────────────────────────────────────────────────
    ax_info.text(0.05, y, "COLOR KEY", fontsize=12, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.028
    ax_info.plot(0.07, y - 0.003, 's', markersize=12, color=COLOR_550,
                 transform=ax_info.transAxes, clip_on=False)
    ax_info.text(0.13, y, "Fe 550 plant", fontsize=9, color='#37474F',
                 transform=ax_info.transAxes, va='top')
    y -= 0.028
    ax_info.plot(0.07, y - 0.003, 's', markersize=12, color=COLOR_550D,
                 transform=ax_info.transAxes, clip_on=False)
    ax_info.text(0.13, y, "Fe 550D plant", fontsize=9, color='#37474F',
                 transform=ax_info.transAxes, va='top')
    y -= 0.028
    ax_info.plot(0.07, y - 0.003, 's', markersize=12, color=COLOR_BOTH,
                 transform=ax_info.transAxes, clip_on=False)
    ax_info.text(0.13, y, "Both grades", fontsize=9, color='#37474F',
                 transform=ax_info.transAxes, va='top')
    y -= 0.028
    ax_info.text(0.07, y, "Darker green = more supply",
                 fontsize=9, color='#2E7D32', fontstyle='italic',
                 transform=ax_info.transAxes, va='top')

    # ── Grand Totals ───────────────────────────────────────────────────
    y -= 0.04
    ax_info.plot([0.05, 0.95], [y, y], color='#B0BEC5', linewidth=0.8,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.025
    t550 = sum(state_fe550.values())
    t550d = sum(state_fe550d.values())
    grand = t550 + t550d
    ax_info.text(0.05, y, "GRAND TOTAL", fontsize=12, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.028
    ax_info.text(0.05, y, f"Fe 550:  {t550:>10,.0f} MT", fontsize=10,
                 color=COLOR_550, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.025
    ax_info.text(0.05, y, f"Fe 550D: {t550d:>10,.0f} MT", fontsize=10,
                 color=COLOR_550D, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.025
    ax_info.text(0.05, y, f"Total:   {grand:>10,.0f} MT", fontsize=11,
                 color='#1B2631', fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.025
    ax_info.text(0.05, y, f"Plants:  {len(PLANT_LOCATIONS):>10}",
                 fontsize=10, color='#546E7A', fontfamily='monospace',
                 transform=ax_info.transAxes, va='top')

    # ── Save ───────────────────────────────────────────────────────────
    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', facecolor='white',
                pad_inches=0.3)
    plt.close()
    print(f"Map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate_consolidated_map()
