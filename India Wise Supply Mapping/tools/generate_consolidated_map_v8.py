"""
generate_consolidated_map_v8.py
V8 consolidated India map — based on V7 premium design.
- State grade labels increased 25% from V7
- Choropleth: green gradient by total supply volume
- Compact factory icons (no numbers) at plant locations
- Per-state grade labels: clean stacked cards, color-coded per grade
- NO state name labels, NO plant directory, NO state table
- Slim sidebar: color key + grade breakdown
"""

import json
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
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data_3grade.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_v8.png")

# ── State name mapping ────────────────────────────────────────────────
STATE_NAME_MAP = {
    "UTTAR PRADESH": "UTTAR PRADESH", "RAJASTHAN": "RAJASTHAN",
    "DELHI": "DELHI", "HARYANA": "HARYANA", "PUNJAB": "PUNJAB",
    "UTTARAKHAND": "UTTARAKHAND", "JAMMU AND KASHMIR": "JAMMU & KASHMIR",
    "HIMACHAL PRADESH": "HIMACHAL PRADESH", "ODISHA": "ODISHA",
    "MADHYA PRADESH": "MADHYA PRADESH", "MAHARASHTRA": "MAHARASHTRA",
    "JHARKHAND": "JHARKHAND", "WEST BENGAL": "WEST BENGAL",
    "BIHAR": "BIHAR", "GUJARAT": "GUJARAT", "TELANGANA": "TELANGANA",
    "ANDHRA PRADESH": "ANDHRA PRADESH", "ASSAM": "ASSAM",
    "CHHATTISGARH": "CHHATTISGARH",
}

# ── State label positions — carefully tuned to avoid overlaps ─────────
STATE_LABEL_POS = {
    "UTTAR PRADESH":    (80.9, 27.0),
    "RAJASTHAN":        (72.5, 26.0),
    "DELHI":            (73.5, 28.5),
    "HARYANA":          (73.2, 30.4),
    "PUNJAB":           (72.0, 32.2),
    "UTTARAKHAND":      (80.5, 31.2),
    "JAMMU & KASHMIR":  (75.3, 34.5),
    "HIMACHAL PRADESH": (77.8, 33.2),
    "ODISHA":           (84.0, 20.5),
    "MADHYA PRADESH":   (78.6, 23.5),
    "MAHARASHTRA":      (75.7, 19.3),
    "JHARKHAND":        (85.5, 23.6),
    "WEST BENGAL":      (88.5, 22.0),
    "BIHAR":            (86.0, 25.6),
    "GUJARAT":          (70.5, 22.3),
    "ANDHRA PRADESH":   (79.7, 15.0),
    "TELANGANA":        (79.0, 17.8),
    "ASSAM":            (93.0, 26.1),
    "CHHATTISGARH":     (82.2, 21.3),
}

# ── Plant locations ───────────────────────────────────────────────────
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

PLANT_DISPLAY_OFFSETS = {
    "Ambashakti Industries Limited": (2.0, 1.5),
    "SKA Ispat Private Limited": (1.5, 1.5),
    "Real Ispat": (1.8, 0.0),
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (-1.8, -0.5),
}

# ── Colors ────────────────────────────────────────────────────────────
C550  = '#B8860B'   # Dark gold — Fe 550
C_OH  = '#D84315'   # Deep orange — OH
C550D = '#1565C0'   # Blue — Fe 550D
C_MULTI = '#7B1FA2' # Purple


def create_factory_icon(size=64):
    """Crisp factory icon at higher resolution."""
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size
    # Chimneys
    d.rectangle([s*0.12, s*0.05, s*0.28, s*0.35], fill=(55, 71, 79, 255))
    d.rectangle([s*0.38, s*0.15, s*0.54, s*0.35], fill=(69, 90, 100, 255))
    # Smoke
    for cx, cy, r in [(s*0.20, s*0.03, s*0.05), (s*0.15, s*0.01, s*0.035),
                       (s*0.46, s*0.11, s*0.04)]:
        d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(176, 190, 197, 180))
    # Main body
    d.rounded_rectangle([s*0.05, s*0.35, s*0.65, s*0.92], radius=s*0.03,
                         fill=(55, 71, 79, 255))
    # Extension
    d.polygon([(s*0.65, s*0.35), (s*0.65, s*0.92), (s*0.92, s*0.92), (s*0.92, s*0.52)],
              fill=(69, 90, 100, 255))
    # Windows (warm glow)
    d.rounded_rectangle([s*0.14, s*0.48, s*0.32, s*0.66], radius=s*0.015,
                         fill=(253, 216, 53, 255))
    d.rounded_rectangle([s*0.40, s*0.48, s*0.58, s*0.66], radius=s*0.015,
                         fill=(253, 216, 53, 255))
    # Door
    d.rounded_rectangle([s*0.72, s*0.58, s*0.85, s*0.92], radius=s*0.015,
                         fill=(253, 216, 53, 200))
    # Base
    d.rectangle([s*0.02, s*0.90, s*0.95, s*0.95], fill=(38, 50, 56, 255))
    return img


def _compute(plant_data):
    s550, s_oh, s550d = {}, {}, {}
    for pdata in plant_data.values():
        for st, v in pdata.get("fe_550", {}).items():
            g = STATE_NAME_MAP.get(st)
            if g: s550[g] = s550.get(g, 0) + v
        for st, v in pdata.get("one_helix", {}).items():
            g = STATE_NAME_MAP.get(st)
            if g: s_oh[g] = s_oh.get(g, 0) + v
        for st, v in pdata.get("fe_550d", {}).items():
            g = STATE_NAME_MAP.get(st)
            if g: s550d[g] = s550d.get(g, 0) + v

    all_st = set(s550) | set(s_oh) | set(s550d)
    totals = {s: s550.get(s, 0) + s_oh.get(s, 0) + s550d.get(s, 0) for s in all_st}
    return s550, s_oh, s550d, totals


def _fmt(v):
    return f"{v/1000:.1f}K" if v >= 1000 else f"{v:.0f}"


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    gdf = gpd.read_file(GEOJSON_PATH)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    s550, s_oh, s550d, state_totals = _compute(plant_data)

    # ── Figure ────────────────────────────────────────────────────────
    fig = plt.figure(figsize=(36, 26), facecolor='white')
    gs = gridspec.GridSpec(1, 2, width_ratios=[87, 13], wspace=0.01)
    ax = fig.add_subplot(gs[0])
    ax_info = fig.add_subplot(gs[1])
    ax_info.axis('off')

    # ── Choropleth ────────────────────────────────────────────────────
    supply_st = set(state_totals.keys())
    gdf['volume'] = gdf['STNAME'].map(state_totals).fillna(0)
    gdf[~gdf['STNAME'].isin(supply_st)].plot(ax=ax, color='#ECEFF1',
                                              edgecolor='#B0BEC5', linewidth=0.5)
    gdf_yes = gdf[gdf['STNAME'].isin(supply_st)].copy()
    if not gdf_yes.empty:
        vals = gdf_yes['volume'].values
        vmin = max(vals[vals > 0].min(), 1) if (vals > 0).any() else 1
        norm = mcolors.LogNorm(vmin=vmin, vmax=vals.max())
        cmap = plt.cm.Greens
        gdf_yes.plot(ax=ax,
                     color=[cmap(norm(v)) if v > 0 else '#ECEFF1' for v in gdf_yes['volume']],
                     edgecolor='#4CAF50', linewidth=0.8)

    # ══════════════════════════════════════════════════════════════════
    # GRADE LABELS — 25% larger than V7
    # V7: LINE_H=0.75, FONT=9, card_w=3.8, text offsets ±1.4/1.5
    # V8: LINE_H=0.94, FONT=11.25, card_w=4.75, text offsets ±1.75/1.88
    # ══════════════════════════════════════════════════════════════════
    LINE_H = 0.94           # 0.75 * 1.25
    FONT_GRADE = 11.25      # 9 * 1.25
    FONT_GRADE_VAL = 11.25  # 9 * 1.25

    for state in state_totals:
        pos = STATE_LABEL_POS.get(state)
        if not pos:
            continue
        cx, cy = pos
        f550  = s550.get(state, 0)
        f_oh  = s_oh.get(state, 0)
        f550d = s550d.get(state, 0)

        lines = []
        if f550  > 0: lines.append(('550',  _fmt(f550),  C550))
        if f_oh  > 0: lines.append(('OH',   _fmt(f_oh),  C_OH))
        if f550d > 0: lines.append(('550D', _fmt(f550d), C550D))
        if not lines:
            continue

        n = len(lines)
        # Background card — 25% wider and taller
        card_w = 4.75          # 3.8 * 1.25
        card_h = n * LINE_H + 0.44   # padding 0.35 * 1.25 ≈ 0.44
        card_x = cx - card_w / 2
        card_y = cy - card_h / 2
        bg = mpatches.FancyBboxPatch(
            (card_x, card_y), card_w, card_h,
            boxstyle="round,pad=0.19",   # 0.15 * 1.25
            facecolor='white', edgecolor='#90A4AE',
            linewidth=0.7, alpha=0.92, zorder=7,
        )
        ax.add_patch(bg)

        # Draw each grade line
        top_y = cy + (n - 1) * LINE_H / 2
        for idx, (gname, gval, gcol) in enumerate(lines):
            ly = top_y - idx * LINE_H
            # Grade name (left-aligned within card)
            ax.text(cx - 1.75, ly, f"{gname}:", fontsize=FONT_GRADE,
                    ha='left', va='center', fontweight='bold', color=gcol,
                    fontfamily='sans-serif', zorder=8)
            # Value (right-aligned within card)
            ax.text(cx + 1.88, ly, gval, fontsize=FONT_GRADE_VAL,
                    ha='right', va='center', fontweight='bold', color='#263238',
                    fontfamily='sans-serif', zorder=8)

    # ── Factory icons (compact, no numbers) ──────────────────────────
    factory_img = create_factory_icon(size=64)
    factory_arr = np.array(factory_img)

    for pname in PLANT_LOCATIONS:
        true_lon, true_lat = PLANT_LOCATIONS[pname]
        off = PLANT_DISPLAY_OFFSETS.get(pname)
        if off:
            dx, dy = true_lon + off[0], true_lat + off[1]
            ax.plot([dx, true_lon], [dy, true_lat],
                    color='#78909C', linewidth=0.7, linestyle='--', alpha=0.5, zorder=9)
            ax.plot(true_lon, true_lat, 'o', markersize=3, color='#546E7A', zorder=9)
        else:
            dx, dy = true_lon, true_lat

        ib = OffsetImage(factory_arr, zoom=0.25)
        ab = AnnotationBbox(ib, (dx, dy), frameon=False, zorder=10)
        ax.add_artist(ab)

    # ── Map chrome ────────────────────────────────────────────────────
    ax.set_xlim(68, 97)
    ax.set_ylim(6, 37)
    ax.set_aspect('equal')
    ax.axis('off')

    if supply_st:
        sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
        sm.set_array([])
        cbar = fig.colorbar(sm, ax=ax, location='bottom', shrink=0.45, pad=0.03, aspect=30)
        cbar.set_label('Total Inbound Supply (MT)', fontsize=11, fontweight='bold')
        cbar.ax.tick_params(labelsize=9)

    ax.set_title("JSW ONE TMT — Plant-wise Supply Network\nFY 2026-27",
                 fontsize=24, fontweight='bold', pad=20,
                 color='#1B2631', fontfamily='sans-serif')

    # ══════════════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════════════
    y = 0.70
    ax_info.text(0.05, y, "COLOR KEY", fontsize=14, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.035
    for col, lab in [(C550, "Fe 550"), (C_OH, "OH (One Helix)"),
                     (C550D, "Fe 550D"), (C_MULTI, "Multiple grades")]:
        ax_info.plot(0.08, y - 0.003, 's', markersize=14, color=col,
                     transform=ax_info.transAxes, clip_on=False)
        ax_info.text(0.18, y, lab, fontsize=11, color='#37474F',
                     transform=ax_info.transAxes, va='top')
        y -= 0.035
    ax_info.text(0.08, y, "Darker green = more supply",
                 fontsize=10, color='#2E7D32', fontstyle='italic',
                 transform=ax_info.transAxes, va='top')

    y -= 0.055
    ax_info.plot([0.05, 0.95], [y, y], color='#B0BEC5', linewidth=0.8,
                 transform=ax_info.transAxes, clip_on=False)
    y -= 0.035

    t550  = sum(s550.values())
    t_oh  = sum(s_oh.values())
    t550d = sum(s550d.values())
    grand = t550 + t_oh + t550d

    ax_info.text(0.05, y, "GRADE BREAKDOWN", fontsize=14, fontweight='bold',
                 color='#1B2631', transform=ax_info.transAxes, va='top')
    y -= 0.038
    ax_info.text(0.05, y, f"Fe 550: {t550:>9,.0f} MT", fontsize=12,
                 color=C550, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"OH:     {t_oh:>9,.0f} MT", fontsize=12,
                 color=C_OH, fontfamily='monospace', fontweight='bold',
                 transform=ax_info.transAxes, va='top')
    y -= 0.032
    ax_info.text(0.05, y, f"Fe 550D:{t550d:>9,.0f} MT", fontsize=12,
                 color=C550D, fontfamily='monospace', fontweight='bold',
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

    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', facecolor='white', pad_inches=0.3)
    plt.close()
    print(f"V8 map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
