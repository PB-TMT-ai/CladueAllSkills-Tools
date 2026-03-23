"""
Generate an interactive India map showing JSW ONE TMT Conversion Mill Network.

Uses the same Folium approach as the UP distributor territory maps:
  - Python → Folium → Standalone HTML (~self-contained, open in browser)
  - High-detail GeoJSON for India (Natural Earth 10m for accurate, complete map)
  - Custom factory markers with proper connecting lines to stacked bar overlays
  - Dynamic per-plant blue intensity (darkest→medium→lightest by value rank)

Output: jsw_one_tmt_mill_network.html
"""

import json
import os
import folium
from folium import DivIcon, Marker, FeatureGroup, PolyLine
import geopandas as gpd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(BASE_DIR, "jsw_one_tmt_mill_network.html")
INDIA_GEOJSON_PATH = os.path.join(BASE_DIR, "india_outline_hd.geojson")

# ── India Bounding Box ──────────────────────────────────────────────────────
INDIA_CENTER = [22.5, 79.5]
INDIA_BOUNDS = [[6.0, 67.5], [36.5, 98.0]]

# ── Plant Data ──────────────────────────────────────────────────────────────
PLANTS = [
    {"id": 1, "name": "AmbaShakti Gwalior", "location": "Gwalior", "state": "MP",
     "lat": 26.22, "lon": 78.18, "lrf": 3400, "freight": 1524, "margin": -231},
    {"id": 2, "name": "Real Ispat", "location": "Raipur", "state": "CG",
     "lat": 21.25, "lon": 81.63, "lrf": 4500, "freight": 2453, "margin": 15},
    {"id": 3, "name": "Rashmi", "location": "Kharagpur", "state": "WB",
     "lat": 22.35, "lon": 87.32, "lrf": 3500, "freight": 1933, "margin": 646},
    {"id": 4, "name": "Jai Bharat", "location": "Kala Amb", "state": "HP",
     "lat": 30.52, "lon": 77.28, "lrf": 3400, "freight": 1379, "margin": -1775},
    {"id": 5, "name": "SKA Ispat", "location": "Raipur", "state": "CG",
     "lat": 21.75, "lon": 82.10, "lrf": 3500, "freight": 2453, "margin": 1015},
    {"id": 6, "name": "AmbaShakti Sikandrabad", "location": "Sikandrabad", "state": "UP",
     "lat": 28.35, "lon": 77.69, "lrf": 3400, "freight": 1339, "margin": -1188},
    {"id": 7, "name": "German Steel", "location": "Kutch", "state": "GJ",
     "lat": 23.73, "lon": 69.86, "lrf": 2500, "freight": 1635, "margin": -1870},
]

# ── Color Palette ───────────────────────────────────────────────────────────
BLUES = ["#08306B", "#2171B5", "#6BAED6"]  # dark, medium, light
RED = "#E05252"


def get_segment_colors(plant):
    """Assign blue shades by per-plant value ranking (highest→darkest).
    Negative margin overridden to red."""
    vals = [
        ("lrf", plant["lrf"]),
        ("freight", plant["freight"]),
        ("margin", abs(plant["margin"])),
    ]
    ranked = sorted(vals, key=lambda x: x[1], reverse=True)
    color_map = {}
    for i, (key, _) in enumerate(ranked):
        color_map[key] = BLUES[i]
    if plant["margin"] < 0:
        color_map["margin"] = RED
    return color_map


def fmt(val):
    """Format ₹ value with Indian numbering."""
    abs_val = abs(val)
    s_str = str(abs_val)
    if len(s_str) > 3:
        last3 = s_str[-3:]
        rest = s_str[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        parts.insert(0, rest)
        s = ",".join(parts) + "," + last3
    else:
        s = s_str
    prefix = "-" if val < 0 else ""
    return f"₹{prefix}{s}"


# ── GeoJSON Download ────────────────────────────────────────────────────────
def ensure_india_geojson():
    """Download India GeoJSON from Natural Earth 10m (high detail) if not cached.
    Uses moderate simplification (0.01°) to keep ~2000 points — accurate but not huge."""
    if os.path.exists(INDIA_GEOJSON_PATH):
        print(f"India GeoJSON exists: {INDIA_GEOJSON_PATH}")
        return

    print("Downloading India from Natural Earth (10m — high detail)...")
    world = gpd.read_file(
        "https://naciscdn.org/naturalearth/10m/cultural/ne_10m_admin_0_countries.zip"
    )
    india = world[world["NAME"] == "India"]
    india_out = india.copy()
    # Light simplification: 0.01° ≈ 1 km — preserves coastline, islands, borders
    india_out["geometry"] = india_out.geometry.simplify(
        tolerance=0.01, preserve_topology=True
    )
    india_out.to_file(INDIA_GEOJSON_PATH, driver="GeoJSON")
    size_kb = os.path.getsize(INDIA_GEOJSON_PATH) / 1024
    print(f"Saved to {INDIA_GEOJSON_PATH} ({size_kb:.0f} KB)")


# ── Bar / Label HTML builders ───────────────────────────────────────────────

MAX_TOTAL = max(p["lrf"] + p["freight"] + abs(p["margin"]) for p in PLANTS)
MAX_BAR_PX = 220  # wider bars for readability


def build_bar_marker_html(plant):
    """Full HTML for the floating label + bar next to each factory on the map.
    Includes plant name, location/state, and a wide stacked bar with values."""
    colors = get_segment_colors(plant)

    w_lrf = plant["lrf"] / MAX_TOTAL * MAX_BAR_PX
    w_freight = plant["freight"] / MAX_TOTAL * MAX_BAR_PX
    w_margin = abs(plant["margin"]) / MAX_TOTAL * MAX_BAR_PX
    total_w = w_lrf + w_freight + w_margin

    segments = [
        (w_lrf, colors["lrf"], fmt(plant["lrf"]), "LRF"),
        (w_freight, colors["freight"], fmt(plant["freight"]), "Freight"),
        (w_margin, colors["margin"], fmt(plant["margin"]), "Margin"),
    ]

    # ── Name + location label ──
    name_html = (
        f'<div style="font-family:Calibri,Segoe UI,sans-serif;font-size:13px;'
        f'font-weight:700;color:#1A202C;margin-bottom:1px;white-space:nowrap;'
        f'text-shadow:0 0 4px #fff, 0 0 4px #fff, 1px 1px 2px rgba(255,255,255,1);">'
        f'{plant["name"]}</div>'
    )
    loc_html = (
        f'<div style="font-family:Calibri,Segoe UI,sans-serif;font-size:11px;'
        f'color:#718096;margin-bottom:4px;white-space:nowrap;'
        f'text-shadow:0 0 3px #fff, 1px 1px 1px rgba(255,255,255,1);">'
        f'{plant["location"]}, {plant["state"]}</div>'
    )

    # ── Stacked bar ──
    bar_html = (
        '<div style="display:flex;height:22px;border-radius:4px;overflow:hidden;'
        'box-shadow:0 1px 5px rgba(0,0,0,0.18);">'
    )
    for w, color, label, seg_name in segments:
        # Show text inside if segment wide enough, else show above
        show_inside = w > 44
        bar_html += (
            f'<div style="width:{w:.1f}px;height:100%;background:{color};'
            f'display:flex;align-items:center;justify-content:center;'
            f'color:{"#fff" if show_inside else "transparent"};'
            f'font-size:{"10px" if w > 60 else "9px"};'
            f'font-family:Calibri,sans-serif;font-weight:600;white-space:nowrap;'
            f'letter-spacing:0.01em;">'
            f'{label if show_inside else ""}</div>'
        )
    bar_html += "</div>"

    # ── Value labels above bar for narrow segments ──
    overflow_labels = ""
    running_x = 0
    for w, color, label, seg_name in segments:
        if w <= 44:
            mid_x = running_x + w / 2
            overflow_labels += (
                f'<span style="position:absolute;left:{mid_x:.0f}px;top:-2px;'
                f'transform:translateX(-50%);font-size:9px;color:#4A5568;'
                f'font-family:Calibri,sans-serif;font-weight:600;white-space:nowrap;'
                f'text-shadow:0 0 3px #fff, 1px 1px 1px #fff;">{label}</span>'
            )
        running_x += w

    # ── Segment labels underneath bar ──
    seg_labels_html = '<div style="display:flex;margin-top:2px;">'
    for w, color, label, seg_name in segments:
        seg_labels_html += (
            f'<div style="width:{w:.1f}px;text-align:center;'
            f'font-size:9px;color:#8899A6;font-family:Calibri,sans-serif;'
            f'white-space:nowrap;overflow:visible;">'
            f'{seg_name}</div>'
        )
    seg_labels_html += "</div>"

    outer_w = max(total_w, 120)
    html = (
        f'<div style="pointer-events:auto;width:{outer_w + 8:.0f}px;position:relative;">'
        f'{name_html}{loc_html}'
        f'<div style="position:relative;">'
        f'{overflow_labels}'
        f'{bar_html}'
        f'</div>'
        f'{seg_labels_html}'
        f'</div>'
    )
    return html, outer_w


def build_popup_html(plant):
    """Rich popup HTML on click."""
    colors = get_segment_colors(plant)
    # Build a wider bar for popup
    pw = 240
    w_lrf = plant["lrf"] / MAX_TOTAL * pw
    w_freight = plant["freight"] / MAX_TOTAL * pw
    w_margin = abs(plant["margin"]) / MAX_TOTAL * pw

    segs = [
        (w_lrf, colors["lrf"], fmt(plant["lrf"]), "LRF Cost"),
        (w_freight, colors["freight"], fmt(plant["freight"]), "Freight from Plant"),
        (w_margin, colors["margin"], fmt(plant["margin"]), "JSW One Margin"),
    ]

    bar = '<div style="display:flex;height:20px;border-radius:4px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.12);margin:8px 0;">'
    for w, c, lbl, _ in segs:
        show = w > 50
        bar += f'<div style="width:{w:.0f}px;height:100%;background:{c};display:flex;align-items:center;justify-content:center;color:{"#fff" if show else "transparent"};font-size:10px;font-family:Calibri,sans-serif;font-weight:600;">{lbl if show else ""}</div>'
    bar += "</div>"

    rows = ""
    for _, c, lbl, seg_name in segs:
        rows += f"""
        <tr>
            <td style="padding:3px 10px 3px 0;">
                <span style="display:inline-block;width:12px;height:12px;border-radius:2px;background:{c};margin-right:6px;vertical-align:middle;"></span>
                {seg_name}
            </td>
            <td style="text-align:right;font-weight:700;padding:3px 0;">{lbl}</td>
        </tr>"""

    html = f"""
    <div style="font-family:Calibri,Segoe UI,sans-serif;min-width:260px;padding:6px;">
        <div style="font-size:16px;font-weight:700;color:#1A202C;margin-bottom:2px;">{plant["name"]}</div>
        <div style="font-size:12px;color:#718096;margin-bottom:4px;">{plant["location"]}, {plant["state"]}</div>
        {bar}
        <table style="font-size:12px;color:#4A5568;border-collapse:collapse;width:100%;">
            {rows}
        </table>
    </div>"""
    return html


# ── Factory Icon ────────────────────────────────────────────────────────────
def factory_icon_svg(size=32, color="#1A202C"):
    """Bold sawtooth-roof factory SVG icon."""
    s = size
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{s}" height="{s}" viewBox="0 0 {s} {s}">'
        f'<path d="M0,{s} L0,{s*0.42:.1f} L{s*0.18:.1f},{s*0.10:.1f} '
        f'L{s*0.18:.1f},{s*0.42:.1f} L{s*0.36:.1f},{s*0.10:.1f} '
        f'L{s*0.36:.1f},{s*0.42:.1f} L{s*0.54:.1f},{s*0.10:.1f} '
        f'L{s*0.54:.1f},{s*0.42:.1f} L{s*0.72:.1f},{s*0.42:.1f} '
        f'L{s*0.72:.1f},0 L{s*0.88:.1f},0 '
        f'L{s*0.88:.1f},{s*0.42:.1f} L{s},{s*0.42:.1f} '
        f'L{s},{s} Z" fill="{color}" />'
        f'</svg>'
    )
    return svg


# ── Create Map ──────────────────────────────────────────────────────────────
def create_map():
    print("Creating map...")
    ensure_india_geojson()

    with open(INDIA_GEOJSON_PATH, "r", encoding="utf-8") as f:
        india_geojson = json.load(f)

    m = folium.Map(
        location=INDIA_CENTER,
        tiles="cartodbpositron",
        zoom_start=5,
        max_bounds=True,
    )
    m.fit_bounds(INDIA_BOUNDS)
    m.options["minZoom"] = 4
    m.options["maxZoom"] = 12

    # ── India outline (high-detail GeoJSON) ────────────────────────────────
    folium.GeoJson(
        india_geojson,
        style_function=lambda f: {
            "fillColor": "#F0EEEB",
            "color": "#A0998E",
            "weight": 1.8,
            "fillOpacity": 0.85,
        },
        highlight_function=lambda f: {
            "weight": 2.5,
            "fillOpacity": 0.92,
        },
    ).add_to(m)

    # ── Plant markers + bars + connecting lines ────────────────────────────
    plant_group = FeatureGroup(name="Conversion Mills")
    lines_group = FeatureGroup(name="Connector Lines")

    # Offsets: (lat_offset, lon_offset) from factory icon to bar label anchor
    BAR_OFFSETS = {
        1: (0.4, 2.2),      # Gwalior → right
        2: (-1.0, 2.5),     # Real Ispat → right+up
        3: (1.0, -1.5),     # Rashmi → above-left
        4: (0.4, 2.2),      # Jai Bharat → right
        5: (1.0, 2.5),      # SKA Ispat → right+down
        6: (-0.6, 2.2),     # Sikandrabad → right
        7: (-1.2, 0.5),     # German Steel → below-right
    }

    for plant in PLANTS:
        lat, lon = plant["lat"], plant["lon"]
        lat_off, lon_off = BAR_OFFSETS[plant["id"]]
        bar_lat = lat + lat_off
        bar_lon = lon + lon_off

        # ── Connecting line (solid, subtle) ──
        PolyLine(
            locations=[[lat, lon], [bar_lat, bar_lon]],
            color="#94A3B8",
            weight=1.5,
            dash_array="6,4",
            opacity=0.7,
        ).add_to(lines_group)

        # ── Factory icon ──
        icon_html = factory_icon_svg(size=32, color="#1A202C")
        Marker(
            location=[lat, lon],
            icon=DivIcon(
                html=icon_html,
                icon_size=(32, 32),
                icon_anchor=(16, 16),
                class_name="factory-icon",
            ),
            tooltip=folium.Tooltip(
                f"<b style='font-size:13px'>{plant['name']}</b><br>"
                f"<span style='color:#718096'>{plant['location']}, {plant['state']}</span><br>"
                f"LRF: {fmt(plant['lrf'])} · Freight: {fmt(plant['freight'])}<br>"
                f"Margin: <b style='color:{get_segment_colors(plant)['margin']}'>"
                f"{fmt(plant['margin'])}</b>",
                sticky=True,
            ),
            popup=folium.Popup(build_popup_html(plant), max_width=320),
        ).add_to(plant_group)

        # ── Stacked bar label ──
        bar_html, bar_w = build_bar_marker_html(plant)
        Marker(
            location=[bar_lat, bar_lon],
            icon=DivIcon(
                html=bar_html,
                icon_size=(int(bar_w + 12), 70),
                icon_anchor=(0, 35),
                class_name="bar-label",
            ),
        ).add_to(plant_group)

    lines_group.add_to(m)
    plant_group.add_to(m)

    # ── Title ──────────────────────────────────────────────────────────────
    title_html = """
    <div style="
        position: fixed;
        top: 12px; left: 50%;
        transform: translateX(-50%);
        z-index: 9999;
        background: rgba(255,255,255,0.96);
        padding: 12px 32px;
        border-radius: 8px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.12);
        font-family: Georgia, serif;
        font-size: 20px;
        font-weight: 600;
        color: #1A202C;
        letter-spacing: 0.02em;
        pointer-events: none;
    ">
        JSW ONE TMT — Conversion Mill Network
    </div>
    """
    m.get_root().html.add_child(folium.Element(title_html))

    # ── Legend ──────────────────────────────────────────────────────────────
    legend_html = """
    <div style="
        position: fixed;
        bottom: 18px; left: 50%;
        transform: translateX(-50%);
        z-index: 9999;
        background: rgba(255,255,255,0.96);
        padding: 12px 24px;
        border-radius: 8px;
        box-shadow: 0 2px 12px rgba(0,0,0,0.10);
        font-family: 'Calibri', 'Segoe UI', sans-serif;
        font-size: 13px;
        color: #4A5568;
        display: flex;
        gap: 24px;
        align-items: center;
        pointer-events: none;
    ">
        <span style="font-weight:700; color:#1A202C; margin-right:2px;">Legend:</span>
        <span>
            <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#08306B;vertical-align:middle;margin-right:4px;"></span>
            Highest
        </span>
        <span>
            <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#2171B5;vertical-align:middle;margin-right:4px;"></span>
            Middle
        </span>
        <span>
            <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#6BAED6;vertical-align:middle;margin-right:4px;"></span>
            Lowest
        </span>
        <span style="border-left:1px solid #CBD5E0; padding-left:18px;">
            <span style="display:inline-block;width:14px;height:14px;border-radius:3px;background:#E05252;vertical-align:middle;margin-right:4px;"></span>
            Negative Margin
        </span>
        <span style="border-left:1px solid #CBD5E0; padding-left:18px; color:#718096;">
            Segments: LRF Cost &nbsp;|&nbsp; Freight &nbsp;|&nbsp; JSW One Margin
        </span>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    # ── Custom CSS ─────────────────────────────────────────────────────────
    custom_css = """
    <style>
        .factory-icon {
            filter: drop-shadow(1px 2px 3px rgba(0,0,0,0.3));
            transition: transform 0.15s ease;
        }
        .factory-icon:hover {
            transform: scale(1.15);
        }
        .bar-label {
            pointer-events: none !important;
        }
        .bar-label > div, .bar-label > div > div {
            pointer-events: none !important;
        }
        .leaflet-popup-content-wrapper {
            border-radius: 10px !important;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15) !important;
        }
        .leaflet-popup-content {
            margin: 12px 14px !important;
            line-height: 1.5 !important;
        }
        .leaflet-tooltip {
            font-family: 'Calibri', 'Segoe UI', sans-serif !important;
            font-size: 13px !important;
            line-height: 1.5 !important;
            padding: 8px 12px !important;
            border-radius: 6px !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15) !important;
            max-width: 280px !important;
        }
    </style>
    """
    m.get_root().html.add_child(folium.Element(custom_css))

    # ── Layer Control ──────────────────────────────────────────────────────
    folium.LayerControl(collapsed=False).add_to(m)

    # ── Save ───────────────────────────────────────────────────────────────
    m.save(OUTPUT_PATH)
    size_kb = os.path.getsize(OUTPUT_PATH) / 1024
    print(f"\nMap saved to: {OUTPUT_PATH}")
    print(f"File size: {size_kb:.0f} KB")
    print("Open in browser to view.")


if __name__ == "__main__":
    create_map()
