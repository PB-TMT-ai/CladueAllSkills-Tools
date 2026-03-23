"""
generate_consolidated_map_v3.py
Minimal, decluttered consolidated India supply map.
- Single-line state labels (abbreviation + total only)
- Small numbered plant circles, gently nudged to avoid overlap
- No colorbar, no arrows, no leader lines
- Clean sidebar: plant directory + state table + totals
"""

import json
import os

import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.gridspec as gridspec

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_v3.png")

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

# ── State label positions — spread out to prevent overlap ─────────────
STATE_LABEL_POS = {
    "UTTAR PRADESH": (80.5, 27.2),
    "RAJASTHAN": (72.8, 26.0),
    "DELHI": (74.5, 28.5),
    "HARYANA": (74.0, 30.1),
    "PUNJAB": (73.2, 31.6),
    "UTTARAKHAND": (80.2, 30.8),
    "JAMMU & KASHMIR": (75.3, 34.3),
    "HIMACHAL PRADESH": (77.6, 32.5),
    "ODISHA": (84.0, 20.5),
    "MADHYA PRADESH": (78.6, 23.5),
    "MAHARASHTRA": (75.7, 19.5),
    "JHARKHAND": (85.3, 23.6),
    "WEST BENGAL": (88.2, 22.5),
    "BIHAR": (85.9, 25.6),
    "GUJARAT": (71.0, 22.3),
    "ANDHRA PRADESH": (79.7, 15.3),
    "TELANGANA": (79.0, 18.0),
    "ASSAM": (93.0, 26.1),
    "CHHATTISGARH": (82.0, 21.3),
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

# ── Plant display positions — gently nudged to avoid overlap ──────────
# No leader lines; just slightly adjusted lat/lon so circles don't sit on top
# of each other. The sidebar legend maps number → plant name.
PLANT_POSITIONS = {
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (76.8, 25.9),      # nudged W from Gwalior
    "API Ispat And Powertech Private Limited": (81.6, 21.25),
    "Aditya Industries": (75.6, 30.7),
    "Ambashakti Industries Limited": (79.0, 28.0),           # nudged NW in UP
    "GERMAN GREEN STEEL AND POWER LIMITED": (77.5, 23.0),    # nudged slightly W in MP
    "Maharashtra - New Plant": (73.9, 18.5),
    "Rashmi Steel": (88.3, 22.6),
    "Real Ispat": (84.2, 25.3),                              # nudged E
    "SKA Ispat Private Limited": (82.0, 27.2),               # nudged E in UP
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

# ── Colors ─────────────────────────────────────────────────────────────
COLOR_550 = '#D4A017'
COLOR_550D = '#1565C0'
COLOR_BOTH = '#7B1FA2'


def _compute(plant_data):
    """Compute state & plant aggregations."""
    s550, s550d = {}, {}
    for pdata in plant_data.values():
        for st, vol in pdata.get("fe_550", {}).items():
            geo = STATE_NAME_MAP.get(st)
            if geo:
                s550[geo] = s550.get(geo, 0) + vol
        for st, vol in pdata.get("fe_550d", {}).items():
            geo = STATE_NAME_MAP.get(st)
            if geo:
                s550d[geo] = s550d.get(geo, 0) + vol
    totals = {s: s550.get(s, 0) + s550d.get(s, 0) for s in set(s550) | set(s550d)}

    pt, pg = {}, {}
    for pn, pd in plant_data.items():
        if pn not in PLANT_POSITIONS:
            continue
        a = sum(pd.get("fe_550", {}).values())
        b = sum(pd.get("fe_550d", {}).values())
        pt[pn] = a + b
        pg[pn] = 'both' if a > 0 and b > 0 else ('fe550d' if b > 0 else 'fe550')
    return s550, s550d, totals, pt, pg


def _fmt(v):
    return f"{v/1000:.1f}K" if v >= 1000 else f"{v:.0f}"


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    gdf = gpd.read_file(GEOJSON_PATH)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    s550, s550d, st_totals, pt_totals, pt_grades = _compute(plant_data)

    # ── Layout ────────────────────────────────────────────────────────
    fig = plt.figure(figsize=(36, 24), facecolor='white')
    gs = gridspec.GridSpec(1, 2, width_ratios=[62, 38], wspace=0.02)
    ax = fig.add_subplot(gs[0])
    ax_sb = fig.add_subplot(gs[1])
    ax_sb.axis('off')

    supply_set = set(st_totals.keys())
    gdf['vol'] = gdf['STNAME'].map(st_totals).fillna(0)

    # ── Base map ──────────────────────────────────────────────────────
    gdf[~gdf['STNAME'].isin(supply_set)].plot(
        ax=ax, color='#F5F5F5', edgecolor='#CFD8DC', linewidth=0.4)

    gdf_s = gdf[gdf['STNAME'].isin(supply_set)].copy()
    if not gdf_s.empty:
        vals = gdf_s['vol'].values
        vmin = max(vals[vals > 0].min(), 1) if (vals > 0).any() else 1
        norm = mcolors.LogNorm(vmin=vmin, vmax=vals.max())
        cmap = plt.cm.Greens
        gdf_s.plot(ax=ax,
                   color=[cmap(norm(v)) if v > 0 else '#F5F5F5' for v in gdf_s['vol']],
                   edgecolor='#66BB6A', linewidth=0.5)

    # ── State labels — one line: "UP 159K" ────────────────────────────
    for state, vol in st_totals.items():
        pos = STATE_LABEL_POS.get(state)
        if not pos:
            continue
        abbr = STATE_ABBREVS.get(state, state[:3])
        ax.text(
            pos[0], pos[1], f"{abbr} {_fmt(vol)}",
            fontsize=7, ha='center', va='center', fontweight='bold',
            color='#263238',
            bbox=dict(facecolor='white', alpha=0.80, edgecolor='none',
                      pad=1.0, boxstyle='round,pad=0.25'),
            zorder=8,
        )

    # ── Plant markers — small numbered circles ────────────────────────
    plant_order = list(PLANT_POSITIONS.keys())
    for i, pn in enumerate(plant_order, 1):
        lon, lat = PLANT_POSITIONS[pn]
        g = pt_grades.get(pn, 'fe550')
        ec = COLOR_BOTH if g == 'both' else (COLOR_550D if g == 'fe550d' else COLOR_550)
        ax.plot(lon, lat, 'o', markersize=11, color='white',
                markeredgecolor=ec, markeredgewidth=2.0, zorder=10)
        ax.text(lon, lat, str(i), fontsize=6.5, ha='center', va='center',
                fontweight='bold', color='#1B2631', zorder=11)

    # ── Map chrome ────────────────────────────────────────────────────
    ax.set_xlim(68, 97)
    ax.set_ylim(6, 37)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title("JSW ONE TMT — Plant-wise Supply Network\nFY 2026-27",
                 fontsize=20, fontweight='bold', pad=16, color='#1B2631')

    # ══════════════════════════════════════════════════════════════════
    # SIDEBAR
    # ══════════════════════════════════════════════════════════════════
    y = 0.97

    # ── Plant Directory ───────────────────────────────────────────────
    ax_sb.text(0.04, y, "PLANT DIRECTORY", fontsize=13, fontweight='bold',
               color='#1B2631', transform=ax_sb.transAxes, va='top')
    y -= 0.026

    for i, pn in enumerate(plant_order, 1):
        g = pt_grades.get(pn, 'fe550')
        mc = COLOR_BOTH if g == 'both' else (COLOR_550D if g == 'fe550d' else COLOR_550)
        short = PLANT_SHORT_NAMES[pn]
        pst = PLANT_STATES[pn]
        vol = pt_totals.get(pn, 0)

        ax_sb.plot(0.05, y - 0.002, 'o', markersize=10, color='white',
                   markeredgecolor=mc, markeredgewidth=1.8,
                   transform=ax_sb.transAxes, zorder=5, clip_on=False)
        ax_sb.text(0.05, y - 0.002, str(i), fontsize=6.5, ha='center', va='center',
                   fontweight='bold', color='#1B2631',
                   transform=ax_sb.transAxes, zorder=6)
        ax_sb.text(0.10, y, f"{short} ({pst})  —  {vol:,.0f} MT",
                   fontsize=9, fontweight='bold', color='#1B2631',
                   transform=ax_sb.transAxes, va='top')
        y -= 0.032

    # ── Divider ───────────────────────────────────────────────────────
    y -= 0.008
    ax_sb.plot([0.04, 0.96], [y, y], color='#CFD8DC', linewidth=0.6,
               transform=ax_sb.transAxes, clip_on=False)
    y -= 0.018

    # ── State Table ───────────────────────────────────────────────────
    ax_sb.text(0.04, y, "STATE-WISE SUPPLY", fontsize=13, fontweight='bold',
               color='#1B2631', transform=ax_sb.transAxes, va='top')
    y -= 0.022

    # Header row
    for label, x, ha, color in [
        ('State', 0.04, 'left', '#37474F'), ('Fe 550', 0.32, 'right', COLOR_550),
        ('Fe 550D', 0.52, 'right', COLOR_550D), ('Total', 0.74, 'right', '#1B2631'),
        ('#Plants', 0.94, 'right', '#546E7A')]:
        ax_sb.text(x, y, label, fontsize=8, fontweight='bold', color=color,
                   transform=ax_sb.transAxes, va='top', ha=ha)
    y -= 0.010
    ax_sb.plot([0.04, 0.96], [y, y], color='#E0E0E0', linewidth=0.4,
               transform=ax_sb.transAxes, clip_on=False)
    y -= 0.012

    # Plant count per state
    spc = {}
    for pn, pd in plant_data.items():
        if pn not in PLANT_POSITIONS:
            continue
        served = set()
        for st in pd.get("fe_550", {}):
            geo = STATE_NAME_MAP.get(st)
            if geo:
                served.add(geo)
        for st in pd.get("fe_550d", {}):
            geo = STATE_NAME_MAP.get(st)
            if geo:
                served.add(geo)
        for s in served:
            spc[s] = spc.get(s, 0) + 1

    for state, total in sorted(st_totals.items(), key=lambda x: x[1], reverse=True):
        abbr = STATE_ABBREVS.get(state, state[:3])
        f5 = s550.get(state, 0)
        f5d = s550d.get(state, 0)
        np_ = spc.get(state, 0)

        ax_sb.text(0.04, y, abbr, fontsize=8, color='#37474F', fontweight='bold',
                   transform=ax_sb.transAxes, va='top')
        ax_sb.text(0.32, y, _fmt(f5) if f5 > 0 else "—",
                   fontsize=8, color=COLOR_550 if f5 > 0 else '#D0D0D0',
                   transform=ax_sb.transAxes, va='top', ha='right')
        ax_sb.text(0.52, y, _fmt(f5d) if f5d > 0 else "—",
                   fontsize=8, color=COLOR_550D if f5d > 0 else '#D0D0D0',
                   transform=ax_sb.transAxes, va='top', ha='right')
        ax_sb.text(0.74, y, f"{total:,.0f}", fontsize=8, color='#1B2631',
                   fontweight='bold', transform=ax_sb.transAxes, va='top', ha='right')
        ax_sb.text(0.94, y, str(np_), fontsize=8, color='#546E7A',
                   transform=ax_sb.transAxes, va='top', ha='right')
        y -= 0.022

    # ── Divider ───────────────────────────────────────────────────────
    y -= 0.006
    ax_sb.plot([0.04, 0.96], [y, y], color='#CFD8DC', linewidth=0.6,
               transform=ax_sb.transAxes, clip_on=False)
    y -= 0.016

    # ── Color Key (inline) ────────────────────────────────────────────
    ax_sb.text(0.04, y, "COLOR KEY", fontsize=10, fontweight='bold',
               color='#1B2631', transform=ax_sb.transAxes, va='top')
    y -= 0.020
    for col, lbl in [(COLOR_550, "Fe 550 (Gold)"),
                      (COLOR_550D, "Fe 550D (Blue)"),
                      (COLOR_BOTH, "Both (Purple)")]:
        ax_sb.plot(0.06, y - 0.002, 's', markersize=8, color=col,
                   transform=ax_sb.transAxes, clip_on=False)
        ax_sb.text(0.10, y, lbl, fontsize=8, color='#455A64',
                   transform=ax_sb.transAxes, va='top')
        y -= 0.018
    ax_sb.text(0.06, y, "Darker green = higher volume",
               fontsize=7.5, color='#2E7D32', fontstyle='italic',
               transform=ax_sb.transAxes, va='top')

    # ── Grand Totals ──────────────────────────────────────────────────
    y -= 0.022
    ax_sb.plot([0.04, 0.96], [y, y], color='#CFD8DC', linewidth=0.6,
               transform=ax_sb.transAxes, clip_on=False)
    y -= 0.018
    t5 = sum(s550.values())
    t5d = sum(s550d.values())
    gt = t5 + t5d
    ax_sb.text(0.04, y, "GRAND TOTAL", fontsize=10, fontweight='bold',
               color='#1B2631', transform=ax_sb.transAxes, va='top')
    y -= 0.020
    ax_sb.text(0.04, y, f"Fe 550:  {t5:>10,.0f} MT", fontsize=9,
               color=COLOR_550, fontfamily='monospace', fontweight='bold',
               transform=ax_sb.transAxes, va='top')
    y -= 0.018
    ax_sb.text(0.04, y, f"Fe 550D: {t5d:>10,.0f} MT", fontsize=9,
               color=COLOR_550D, fontfamily='monospace', fontweight='bold',
               transform=ax_sb.transAxes, va='top')
    y -= 0.018
    ax_sb.text(0.04, y, f"Total:   {gt:>10,.0f} MT", fontsize=9.5,
               color='#1B2631', fontfamily='monospace', fontweight='bold',
               transform=ax_sb.transAxes, va='top')
    y -= 0.018
    ax_sb.text(0.04, y, f"Plants:  {len(PLANT_POSITIONS):>10}",
               fontsize=9, color='#546E7A', fontfamily='monospace',
               transform=ax_sb.transAxes, va='top')

    # ── Save ──────────────────────────────────────────────────────────
    plt.savefig(OUTPUT_FILE, dpi=300, bbox_inches='tight', facecolor='white',
                pad_inches=0.3)
    plt.close()
    print(f"V3 map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
