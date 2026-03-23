"""
Generate two interactive maps of West Bengal.
  1. District Potential Map (choropleth + stock yard pins + district labels)
  2. Distributor Territory Map (color-coded territories + dealer count circles + district labels)

GeoJSON auto-downloaded from datta07/INDIAN-SHAPEFILES on first run.
Outputs: west_bengal_district_potential_map.html, west_bengal_distributor_territory_map.html

Notes:
  - GeoJSON has 23 features; Jhargram and Kalimpong are newer districts (post-2017)
    not present in Excel — they render grey (no data). This is expected.
  - "KEDIA PIPES PRIVATE LMITED" preserves the typo from the Excel source exactly.
"""

import json
import os
import warnings
import requests
import folium
from folium.features import GeoJsonTooltip
import openpyxl

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# --- Paths -----------------------------------------------------------------
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH   = r"C:\Users\2750834\Downloads\Book 67.xlsx"
GEOJSON_URL  = "https://raw.githubusercontent.com/datta07/INDIAN-SHAPEFILES/master/STATES/WEST%20BENGAL/WEST%20BENGAL_DISTRICTS.geojson"
GEOJSON_PATH = os.path.join(BASE_DIR, "west_bengal_districts.geojson")
MAP1_PATH    = os.path.join(BASE_DIR, "west_bengal_district_potential_map.html")
MAP2_PATH    = os.path.join(BASE_DIR, "west_bengal_distributor_territory_map.html")

# --- State Identity ---------------------------------------------------------
STATE_NAME  = "WEST BENGAL"
STATE_LABEL = "West Bengal"

# --- State Bounding Box (with ~0.15 deg padding) ---------------------------
STATE_BOUNDS = [[21.33, 85.67], [27.37, 90.02]]
STATE_CENTER  = [24.35, 87.85]

# --- Name mapping: Excel -> GeoJSON ----------------------------------------
EXCEL_TO_GEOJSON = {
    "24 Paraganas North":  "North Twenty Four Pargan*",
    "24 Paraganas South":  "South Twenty Four Pargan*",
    "Medinipur East":      "Purba Medinipur",
    "Dinajpur Uttar":      "Uttar Dinajpur",
    "Dinajpur Dakshin":    "Dakshin Dinajpur",
    "DARJEELING":          "Darjiling",
    "Purulia":             "Puruliya",
}

GEOJSON_TO_EXCEL = {v: k for k, v in EXCEL_TO_GEOJSON.items()}

# --- Color schemes ----------------------------------------------------------
POTENTIAL_COLORS = {
    "Very High": "#006d2c",
    "High":      "#66c266",
    "Medium":    "#ffe066",
    "Low":       "#ef7070",
}

DISTRIBUTOR_COLORS = {
    "NERIUM MULTICOM LLP":         "#e41a1c",
    "KEDIA PIPES PRIVATE LMITED":  "#377eb8",   # note: typo preserved from Excel
    "CHANDRA TRADERS":             "#4daf4a",
    "Maheshwari Traders":          "#ff7f00",
    "SILIGURI BUILDERS STORES":    "#984ea3",
    "Jai Mata Di Steels":          "#a65628",
}

DISTRIBUTOR_SHORT = {
    "NERIUM MULTICOM LLP":         "Nerium Multicom",
    "KEDIA PIPES PRIVATE LMITED":  "Kedia Pipes",
    "CHANDRA TRADERS":             "Chandra Traders",
    "Maheshwari Traders":          "Maheshwari Traders",
    "SILIGURI BUILDERS STORES":    "Siliguri Builders",
    "Jai Mata Di Steels":          "Jai Mata Di Steels",
}


# --- Helpers ----------------------------------------------------------------

def normalize_name(excel_name):
    return EXCEL_TO_GEOJSON.get(excel_name, excel_name)


def geojson_name_to_excel(geo_name):
    return GEOJSON_TO_EXCEL.get(geo_name, geo_name)


# --- GeoJSON download -------------------------------------------------------

def ensure_geojson():
    if not os.path.exists(GEOJSON_PATH):
        print(f"Downloading GeoJSON from:\n  {GEOJSON_URL}")
        r = requests.get(GEOJSON_URL, verify=False, timeout=60)
        r.raise_for_status()
        with open(GEOJSON_PATH, "w", encoding="utf-8") as f:
            json.dump(r.json(), f)
        print(f"Saved: {GEOJSON_PATH}")
    else:
        print(f"Using cached GeoJSON: {GEOJSON_PATH}")


# --- Load Data --------------------------------------------------------------

def load_geojson():
    with open(GEOJSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    # datta07 GeoJSON uses 'dtname' instead of 'district' — add alias
    for feat in data["features"]:
        feat["properties"]["district"] = feat["properties"]["dtname"]
    return data


def load_excel():
    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()

    district_potential    = {}
    district_distributors = {}
    stockyard_districts   = set()
    district_dealers      = {}

    for row in rows:
        if row[1] != STATE_NAME:
            continue
        district    = row[0]
        distributor = row[3]
        category    = row[4]
        stockyard   = str(row[5]).strip().lower() if row[5] else ""
        dealers     = int(row[6]) if row[6] is not None else 0

        norm = normalize_name(district)

        if norm not in district_potential:
            district_potential[norm] = category

        district_distributors.setdefault(norm, []).append(distributor)

        if stockyard == "yes":
            stockyard_districts.add(norm)

        if norm not in district_dealers or dealers > district_dealers[norm]:
            district_dealers[norm] = dealers

    # Build reverse: distributor -> [districts]  (skip None distributor)
    distributor_districts = {}
    for dist, distribs in district_distributors.items():
        for d in distribs:
            if d is not None:
                distributor_districts.setdefault(d, []).append(dist)

    return district_potential, district_distributors, distributor_districts, stockyard_districts, district_dealers


def compute_centroids(geojson_data):
    centroids = {}
    areas     = {}
    for feature in geojson_data["features"]:
        name   = feature["properties"]["district"]
        coords = feature["geometry"]["coordinates"]

        flat = []
        def flatten(c):
            if isinstance(c[0], (int, float)):
                flat.append(c)
            else:
                for item in c:
                    flatten(item)
        flatten(coords)

        if flat:
            lngs = [p[0] for p in flat]
            lats = [p[1] for p in flat]
            centroids[name] = [sum(lats) / len(lats), sum(lngs) / len(lngs)]
            areas[name]     = (max(lngs) - min(lngs)) * (max(lats) - min(lats))

    return centroids, areas


# --- Base Map ---------------------------------------------------------------

def create_base_map():
    m = folium.Map(location=STATE_CENTER, tiles="cartodbpositron", max_bounds=True)
    m.fit_bounds(STATE_BOUNDS)
    m.options["minZoom"] = 7
    m.options["maxZoom"] = 12
    return m


# --- Map 1: District Potential ----------------------------------------------

def create_potential_map(geojson_data, district_potential, district_distributors, centroids, stockyard_districts):
    m = create_base_map()

    def style_fn(feature):
        geo_name   = feature["properties"]["district"]
        excel_name = geojson_name_to_excel(geo_name)
        cat = district_potential.get(geo_name) or district_potential.get(excel_name)
        return {
            "fillColor":   POTENTIAL_COLORS.get(cat, "#d9d9d9"),
            "color":       "#555555",
            "weight":      1.2,
            "fillOpacity": 0.75,
        }

    def highlight_fn(feature):
        return {"weight": 3, "fillOpacity": 0.9}

    for feature in geojson_data["features"]:
        geo_name   = feature["properties"]["district"]
        excel_name = geojson_name_to_excel(geo_name)
        cat      = district_potential.get(geo_name) or district_potential.get(excel_name) or "No Data"
        distribs = district_distributors.get(geo_name) or district_distributors.get(excel_name) or []
        color    = POTENTIAL_COLORS.get(cat, "#d9d9d9")

        popup_html = f"""
        <div style="min-width:200px; font-family:Arial,sans-serif;">
            <h4 style="margin:0 0 8px 0; color:#333;">{geo_name}</h4>
            <div style="background:{color}; color:white; padding:4px 8px; border-radius:3px;
                        display:inline-block; margin-bottom:8px; font-weight:bold;">
                {cat}
            </div>
            <br><b>Distributor(s):</b><br>
            {'<br>'.join(f'&bull; {d if d else "Unassigned"}' for d in distribs) if distribs else '<i>None assigned</i>'}
        </div>
        """
        feature["properties"]["_popup"] = popup_html

    folium.GeoJson(
        geojson_data,
        style_function=style_fn,
        highlight_function=highlight_fn,
        tooltip=folium.GeoJsonTooltip(
            fields=["district"],
            aliases=["District:"],
            style="font-size:13px; font-weight:bold; background-color:white; padding:6px;",
        ),
        popup=folium.GeoJsonPopup(fields=["_popup"], labels=False),
    ).add_to(m)

    # Stock yard pins
    stockyard_group = folium.FeatureGroup(name="Stock Yard Locations")
    for geo_name, centroid in centroids.items():
        excel_name = geojson_name_to_excel(geo_name)
        if geo_name in stockyard_districts or excel_name in stockyard_districts:
            folium.Marker(
                location=centroid,
                icon=folium.Icon(icon="home", prefix="fa", color="darkblue"),
                tooltip=f"Stock Yard: {geo_name}",
                popup=f"<b>Stock Yard</b><br>{geo_name}",
            ).add_to(stockyard_group)
    stockyard_group.add_to(m)

    # District labels
    label_group = folium.FeatureGroup(name="District Labels")
    for geo_name, centroid in centroids.items():
        folium.Marker(
            location=centroid,
            icon=folium.DivIcon(
                html=(f'<div style="font-size:11px; font-weight:bold; text-align:center; '
                      f'white-space:nowrap; color:#000; '
                      f'text-shadow: 0 0 5px white, 0 0 5px white, 0 0 5px white, 0 0 5px white; '
                      f'pointer-events:none;">{geo_name}</div>'),
                icon_size=(0, 0),
                icon_anchor=(30, 8),
            ),
        ).add_to(label_group)
    label_group.add_to(m)

    # Legend
    legend_html = """
    <div style="position:fixed; bottom:30px; left:30px; z-index:1000;
         background:white; padding:14px 18px; border:2px solid #888; border-radius:6px;
         font-family:Arial,sans-serif; font-size:13px; box-shadow:2px 2px 6px rgba(0,0,0,0.3);">
        <b style="font-size:14px;">District Potential</b><br><br>
        <div style="margin-bottom:4px;"><i style="background:#006d2c;width:18px;height:14px;display:inline-block;border:1px solid #333;"></i>&nbsp; Very High</div>
        <div style="margin-bottom:4px;"><i style="background:#66c266;width:18px;height:14px;display:inline-block;border:1px solid #333;"></i>&nbsp; High</div>
        <div style="margin-bottom:4px;"><i style="background:#ffe066;width:18px;height:14px;display:inline-block;border:1px solid #333;"></i>&nbsp; Medium</div>
        <div style="margin-bottom:4px;"><i style="background:#ef7070;width:18px;height:14px;display:inline-block;border:1px solid #333;"></i>&nbsp; Low</div>
        <div style="margin-bottom:8px;"><i style="background:#d9d9d9;width:18px;height:14px;display:inline-block;border:1px solid #333;"></i>&nbsp; No Data</div>
        <b style="font-size:13px;">Markers</b><br><br>
        <div style="margin-bottom:4px;"><i style="background:#00008b;width:14px;height:14px;display:inline-block;border:1px solid #333;border-radius:2px;"></i>&nbsp; Stock Yard</div>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    title_html = f"""
    <div style="position:fixed; top:10px; left:50%; transform:translateX(-50%); z-index:1000;
         background:white; padding:10px 24px; border:2px solid #888; border-radius:6px;
         font-family:Arial,sans-serif; font-size:18px; font-weight:bold; color:#333;
         box-shadow:2px 2px 6px rgba(0,0,0,0.3);">
        {STATE_LABEL} &mdash; District Potential Map
    </div>
    """
    m.get_root().html.add_child(folium.Element(title_html))

    folium.LayerControl().add_to(m)
    m.save(MAP1_PATH)
    print(f"Map 1 saved: {MAP1_PATH}")


# --- Map 2: Distributor Territories -----------------------------------------

def create_territory_map(geojson_data, district_potential, district_distributors, centroids, district_dealers, areas):
    m = create_base_map()

    def style_fn(feature):
        geo_name   = feature["properties"]["district"]
        excel_name = geojson_name_to_excel(geo_name)
        distribs   = district_distributors.get(geo_name) or district_distributors.get(excel_name) or []
        primary    = distribs[0] if distribs else None
        return {
            "fillColor":   DISTRIBUTOR_COLORS.get(primary, "#d9d9d9"),
            "color":       "#555555",
            "weight":      1.5,
            "fillOpacity": 0.55,
        }

    def highlight_fn(feature):
        return {"weight": 3, "fillOpacity": 0.8}

    for feature in geojson_data["features"]:
        geo_name   = feature["properties"]["district"]
        excel_name = geojson_name_to_excel(geo_name)
        cat      = district_potential.get(geo_name) or district_potential.get(excel_name) or "No Data"
        distribs = district_distributors.get(geo_name) or district_distributors.get(excel_name) or []

        distrib_lines = ""
        for d in distribs:
            col   = DISTRIBUTOR_COLORS.get(d, "#333")
            short = DISTRIBUTOR_SHORT.get(d, (d[:20] if d else "Unassigned"))
            distrib_lines += (f'<div style="margin:2px 0;">'
                              f'<span style="color:{col};font-size:16px;">&#9632;</span> {short}</div>')

        popup_html = f"""
        <div style="min-width:220px; font-family:Arial,sans-serif;">
            <h4 style="margin:0 0 6px 0; color:#333;">{geo_name}</h4>
            <div style="margin-bottom:6px;">Potential: <b>{cat}</b></div>
            <b>Distributor(s):</b>
            {distrib_lines if distrib_lines else '<div><i>None assigned</i></div>'}
        </div>
        """
        feature["properties"]["_popup"] = popup_html

    folium.GeoJson(
        geojson_data,
        style_function=style_fn,
        highlight_function=highlight_fn,
        tooltip=folium.GeoJsonTooltip(
            fields=["district"],
            aliases=["District:"],
            style="font-size:13px; font-weight:bold; background-color:white; padding:6px;",
        ),
        popup=folium.GeoJsonPopup(fields=["_popup"], labels=False),
    ).add_to(m)

    # District labels (adaptive font size)
    label_group = folium.FeatureGroup(name="District Labels")
    for geo_name, centroid in centroids.items():
        area = areas.get(geo_name, 0.5)
        if area >= 0.8:
            name_size = 10
        elif area >= 0.4:
            name_size = 9
        else:
            name_size = 8

        label_html = (
            f'<div style="font-family: Segoe UI, Calibri, sans-serif; '
            f'font-weight: 400; text-align:center; white-space:nowrap; color:#111; '
            f'text-shadow: 0 0 3px white, 0 0 3px white, 0 0 3px white; '
            f'pointer-events:none; line-height:1.3;">'
            f'<div style="font-size:{name_size}px;">{geo_name}</div>'
            f'</div>'
        )
        folium.Marker(
            location=centroid,
            icon=folium.DivIcon(html=label_html, icon_size=(0, 0), icon_anchor=(30, 8)),
        ).add_to(label_group)
    label_group.add_to(m)

    # Dealer count circles
    dealer_group = folium.FeatureGroup(name="Dealer Count (>100 MT)")
    for geo_name, centroid in centroids.items():
        excel_name = geojson_name_to_excel(geo_name)
        dealers = district_dealers.get(geo_name) or district_dealers.get(excel_name, 0)
        if dealers > 0:
            circle_html = (
                f'<div style="display:flex; align-items:center; justify-content:center; '
                f'width:24px; height:24px; border-radius:50%; '
                f'background:#d32f2f; color:white; '
                f'font-family: Segoe UI, Calibri, sans-serif; '
                f'font-size:12px; font-weight:700; '
                f'border:2px solid white; '
                f'box-shadow:0 1px 4px rgba(0,0,0,0.5); '
                f'pointer-events:auto; cursor:pointer;">'
                f'{dealers}</div>'
            )
            folium.Marker(
                location=[centroid[0] - 0.06, centroid[1]],
                icon=folium.DivIcon(html=circle_html, icon_size=(24, 24), icon_anchor=(12, 12)),
                tooltip=f"{geo_name}: {dealers} Dealers (>100 MT)",
            ).add_to(dealer_group)
    dealer_group.add_to(m)

    # Legend
    legend_items = ""
    for dist_name, color in DISTRIBUTOR_COLORS.items():
        short = DISTRIBUTOR_SHORT.get(dist_name, dist_name[:20])
        legend_items += (
            f'<div style="margin-bottom:4px;">'
            f'<i style="background:{color};width:18px;height:14px;display:inline-block;'
            f'border:1px solid #333;"></i>&nbsp; {short}</div>'
        )

    legend_html = f"""
    <div style="position:fixed; bottom:30px; left:30px; z-index:1000;
         background:white; padding:14px 18px; border:2px solid #888; border-radius:6px;
         font-family:Arial,sans-serif; font-size:13px; box-shadow:2px 2px 6px rgba(0,0,0,0.3);">
        <b style="font-size:14px;">Distributor Territories</b><br><br>
        {legend_items}
        <div><i style="background:#d9d9d9;width:18px;height:14px;display:inline-block;
                border:1px solid #333;"></i>&nbsp; Unassigned</div>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))

    title_html = f"""
    <div style="position:fixed; top:10px; left:50%; transform:translateX(-50%); z-index:1000;
         background:white; padding:10px 24px; border:2px solid #888; border-radius:6px;
         font-family:Arial,sans-serif; font-size:18px; font-weight:bold; color:#333;
         box-shadow:2px 2px 6px rgba(0,0,0,0.3);">
        {STATE_LABEL} &mdash; Distributor Territory Map
    </div>
    """
    m.get_root().html.add_child(folium.Element(title_html))

    folium.LayerControl().add_to(m)
    m.save(MAP2_PATH)
    print(f"Map 2 saved: {MAP2_PATH}")


# --- Main -------------------------------------------------------------------

def main():
    ensure_geojson()

    print("Loading GeoJSON...")
    geojson_data = load_geojson()
    geo_names = sorted(set(f["properties"]["district"] for f in geojson_data["features"]))
    print(f"  {len(geo_names)} districts in GeoJSON")

    print("Loading Excel data...")
    district_potential, district_distributors, distributor_districts, stockyard_districts, district_dealers = load_excel()
    print(f"  {len(district_potential)} districts with potential data")
    print(f"  {len(distributor_districts)} distributors")
    print(f"  Stock yards: {sorted(stockyard_districts)}")

    matched, unmatched_geo = [], []
    for gn in geo_names:
        en = geojson_name_to_excel(gn)
        if gn in district_potential or en in district_potential:
            matched.append(gn)
        else:
            unmatched_geo.append(gn)
    print(f"\n  Matched: {len(matched)}/{len(geo_names)} districts")
    if unmatched_geo:
        print(f"  Unmatched (will show grey): {unmatched_geo}")

    print("\nDistributor summary:")
    for dist_name, districts in sorted(distributor_districts.items()):
        short = DISTRIBUTOR_SHORT.get(dist_name, dist_name[:20])
        print(f"  {short}: {len(districts)} districts")

    from collections import Counter
    cats = Counter(district_potential.values())
    print("\nPotential breakdown:")
    for cat in ["Very High", "High", "Medium", "Low"]:
        print(f"  {cat}: {cats.get(cat, 0)}")

    centroids, areas = compute_centroids(geojson_data)
    print(f"\nComputed centroids for {len(centroids)} districts")

    print(f"\n--- Generating Map 1: {STATE_LABEL} District Potential ---")
    create_potential_map(geojson_data, district_potential, district_distributors, centroids, stockyard_districts)

    geojson_data = load_geojson()

    print(f"\n--- Generating Map 2: {STATE_LABEL} Distributor Territories ---")
    create_territory_map(geojson_data, district_potential, district_distributors, centroids, district_dealers, areas)

    print(f"\nDone! Open in browser:")
    print(f"  {MAP1_PATH}")
    print(f"  {MAP2_PATH}")


if __name__ == "__main__":
    main()
