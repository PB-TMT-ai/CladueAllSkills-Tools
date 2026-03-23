"""
generate_supply_map.py
Generates India supply maps showing plant location and supply regions.
- Plant state gets a factory icon
- States receiving supply are colored
- States not receiving supply are light gray
"""

import json
import os
import sys
import geopandas as gpd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from shapely.geometry import Point
import numpy as np
from PIL import Image, ImageDraw

# ── Paths ──────────────────────────────────────────────────────────────
BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")

# ── State name mapping (data → GeoJSON) ───────────────────────────────
STATE_NAME_MAP = {
    "UTTAR PRADESH": "Uttar Pradesh",
    "RAJASTHAN": "Rajasthan",
    "DELHI": "Delhi",
    "HARYANA": "Haryana",
    "PUNJAB": "Punjab",
    "UTTARAKHAND": "Uttaranchal",
    "JAMMU AND KASHMIR": "Jammu and Kashmir",
    "HIMACHAL PRADESH": "Himachal Pradesh",
    "ODISHA": "Orissa",
    "MADHYA PRADESH": "Madhya Pradesh",
    "MAHARASHTRA": "Maharashtra",
    "JHARKHAND": "Jharkhand",
    "WEST BENGAL": "West Bengal",
    "BIHAR": "Bihar",
    "GUJARAT": "Gujarat",
    "TELANGANA": "Andhra Pradesh",  # Telangana was part of AP in old maps
    "ANDHRA PRADESH": "Andhra Pradesh",
    "ASSAM": "Assam",
    "CHHATTISGARH": "Chhattisgarh",
}

# Approximate centroids for plant state labels (lon, lat)
STATE_CENTROIDS = {
    "Uttar Pradesh": (80.9, 27.0),
    "Rajasthan": (73.7, 26.5),
    "Delhi": (77.1, 28.65),
    "Haryana": (76.1, 29.0),
    "Punjab": (75.3, 31.0),
    "Uttaranchal": (79.0, 30.1),
    "Jammu and Kashmir": (75.3, 34.0),
    "Himachal Pradesh": (77.2, 31.8),
    "Orissa": (84.0, 20.5),
    "Madhya Pradesh": (78.6, 23.5),
    "Maharashtra": (75.7, 19.5),
    "Jharkhand": (85.3, 23.6),
    "West Bengal": (87.9, 23.0),
    "Bihar": (85.9, 25.6),
    "Gujarat": (71.6, 22.3),
    "Andhra Pradesh": (79.7, 16.5),
    "Assam": (92.9, 26.1),
    "Chhattisgarh": (81.9, 21.3),
    "Karnataka": (75.7, 15.3),
    "Kerala": (76.3, 10.5),
    "Tamil Nadu": (78.7, 11.0),
    "Goa": (74.0, 15.4),
    "Sikkim": (88.5, 27.5),
    "Arunachal Pradesh": (94.7, 28.2),
    "Nagaland": (94.6, 26.1),
    "Manipur": (93.9, 24.8),
    "Mizoram": (92.9, 23.2),
    "Tripura": (91.7, 23.8),
    "Meghalaya": (91.4, 25.5),
    "Chandigarh": (76.8, 30.7),
    "Dadra and Nagar Haveli": (73.0, 20.3),
    "Daman and Diu": (72.8, 20.4),
    "Lakshadweep": (72.2, 10.6),
    "Puducherry": (79.8, 12.0),
    "Andaman and Nicobar": (92.7, 11.7),
}

# Plant locations (approximate city coordinates: lon, lat)
PLANT_LOCATIONS = {
    "AIC IRON INDUSTRIES PRIVATE LIMITED": (86.2, 22.8),        # Jamshedpur, Jharkhand
    "AMBASHAKTI UDYOG LIMITED- GWALIOR": (78.18, 26.22),        # Gwalior, MP
    "API Ispat And Powertech Private Limited": (81.6, 21.25),    # Raipur, Chhattisgarh
    "Aditya Industries": (76.0, 30.7),                           # Punjab (Mandi Gobindgarh area)
    "Ambashakti Industries Limited": (80.35, 26.45),             # Kanpur area, UP
    "GERMAN GREEN STEEL AND POWER LIMITED": (78.0, 23.2),        # MP (Bhopal area)
    "Maharashtra - New Plant": (73.9, 18.5),                     # Maharashtra (Pune area)
    "Rashmi Steel": (88.3, 22.6),                                # Kolkata area, WB
    "Real Ispat": (83.0, 25.3),                                  # UP (Varanasi area)
    "SKA Ispat Private Limited": (80.9, 26.85),                  # UP (Lucknow area)
    "Telangana - New Plant": (78.5, 17.4),                       # Hyderabad, Telangana
}


def create_factory_icon(size=40):
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


def generate_map(plant_name, product_grade, supply_states, plant_state_geo,
                 plant_loc, output_path):
    """
    Generate a single supply map.

    Args:
        plant_name: Name of the plant
        product_grade: "Fe 550" or "Fe 550D"
        supply_states: dict {geo_state_name: total_mt}
        plant_state_geo: GeoJSON state name where plant is located
        plant_loc: (lon, lat) of plant
        output_path: Path to save the map image
    """
    gdf = gpd.read_file(GEOJSON_PATH)

    fig, ax = plt.subplots(1, 1, figsize=(12, 14), facecolor='white')

    # Color scheme
    supply_color = "#2196F3" if product_grade == "Fe 550" else "#FF5722"
    plant_state_color = "#1565C0" if product_grade == "Fe 550" else "#BF360C"
    no_supply_color = "#E8E8E8"
    border_color = "#666666"

    # Classify states
    def get_color(row):
        name = row['NAME_1']
        if name in supply_states:
            if name == plant_state_geo:
                return plant_state_color  # Darker shade for plant's own state
            return supply_color
        return no_supply_color

    gdf['color'] = gdf.apply(get_color, axis=1)

    # Plot map
    gdf.plot(ax=ax, color=gdf['color'], edgecolor=border_color, linewidth=0.5)

    # Add supply quantity labels on colored states
    for geo_name, mt in supply_states.items():
        if geo_name in STATE_CENTROIDS:
            cx, cy = STATE_CENTROIDS[geo_name]
            mt_label = f"{mt:,.0f}" if mt >= 1 else ""
            if mt_label:
                ax.annotate(mt_label, xy=(cx, cy), fontsize=6, fontweight='bold',
                           ha='center', va='center', color='white',
                           bbox=dict(boxstyle='round,pad=0.2', facecolor='black',
                                     alpha=0.6, edgecolor='none'))

    # Add factory icon at plant location
    if plant_loc:
        factory_img = create_factory_icon(size=60)
        factory_arr = np.array(factory_img)
        imagebox = OffsetImage(factory_arr, zoom=0.5)
        ab = AnnotationBbox(imagebox, plant_loc, frameon=False, zorder=10)
        ax.add_artist(ab)

        # Plant name label near icon
        short_name = plant_name.split(' - ')[0] if ' - ' in plant_name else plant_name
        if len(short_name) > 30:
            short_name = short_name[:28] + "..."
        ax.annotate(short_name, xy=(plant_loc[0], plant_loc[1] - 0.8),
                    fontsize=6, ha='center', va='top', fontweight='bold',
                    color='#333333',
                    bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                              edgecolor='#333333', alpha=0.9, linewidth=0.5))

    # Set map bounds (focus on mainland India)
    ax.set_xlim(68, 98)
    ax.set_ylim(6, 38)
    ax.set_aspect('equal')
    ax.axis('off')

    # Title
    total_supply = sum(supply_states.values())
    title = f"{plant_name}\n{product_grade} Supply Map | FY 2026-27 | Total: {total_supply:,.0f} MT"
    ax.set_title(title, fontsize=13, fontweight='bold', pad=20, color='#333333')

    # Legend
    legend_elements = [
        mpatches.Patch(facecolor=plant_state_color, edgecolor=border_color,
                       label=f'Plant State ({plant_state_geo})'),
        mpatches.Patch(facecolor=supply_color, edgecolor=border_color,
                       label='Supply Region'),
        mpatches.Patch(facecolor=no_supply_color, edgecolor=border_color,
                       label='No Supply'),
    ]
    ax.legend(handles=legend_elements, loc='lower left', fontsize=9,
              framealpha=0.9, edgecolor='#CCCCCC')

    plt.tight_layout()
    plt.savefig(output_path, dpi=200, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"  Saved: {output_path}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(DATA_PATH, 'r') as f:
        data = json.load(f)

    for plant_name, plant_data in data.items():
        plant_state_data = plant_data.get("plant_state")
        plant_state_geo = STATE_NAME_MAP.get(plant_state_data, plant_state_data) if plant_state_data else None
        plant_loc = PLANT_LOCATIONS.get(plant_name)

        for grade_key, grade_label in [("fe_550", "Fe 550"), ("fe_550d", "Fe 550D")]:
            raw_supply = plant_data.get(grade_key, {})
            if not raw_supply:
                print(f"  Skipping {plant_name} / {grade_label} — no supply data")
                continue

            # Map data state names to GeoJSON names
            supply_states = {}
            for data_state, mt in raw_supply.items():
                geo_name = STATE_NAME_MAP.get(data_state)
                if geo_name:
                    # Aggregate (Telangana + AP both map to Andhra Pradesh in old GeoJSON)
                    supply_states[geo_name] = supply_states.get(geo_name, 0) + mt

            # Build filename
            safe_name = plant_name.replace(" ", "_").replace("-", "_").replace(".", "")
            safe_name = "".join(c for c in safe_name if c.isalnum() or c == "_")
            filename = f"{safe_name}_{grade_key}.png"
            output_path = os.path.join(OUTPUT_DIR, filename)

            print(f"Generating: {plant_name} / {grade_label}")
            generate_map(plant_name, grade_label, supply_states, plant_state_geo,
                         plant_loc, output_path)

    print("\nDone! All maps saved to:", OUTPUT_DIR)


if __name__ == "__main__":
    main()
