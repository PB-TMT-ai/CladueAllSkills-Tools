"""
generate_html_map_v7.py  —  Complete rewrite.
Pure D3.js SVG map — screenshot-ready, high-quality data-viz design.
- India-only (white BG, no tiles, no other countries)
- Green choropleth by total supply volume
- SVG factory icons at plant locations
- Per-state grade labels: compact color-coded pills on hover,
  with a clean tooltip panel. On the static view, only small dots
  per state to keep the map clean — grade breakdown visible via
  a right-side panel listing ALL states.
- Right sidebar: color key + grade breakdown + state-by-state table
"""

import json
import os

BASE_DIR = r"D:\India Wise Supply Mapping"
GEOJSON_PATH = os.path.join(BASE_DIR, ".tmp", "india_states_v2.geojson")
DATA_PATH = os.path.join(BASE_DIR, ".tmp", "plant_supply_data_3grade.json")
OUTPUT_DIR = os.path.join(BASE_DIR, ".tmp", "maps")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "consolidated_supply_map_v7.html")

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

PLANTS = [
    {"name": "Ambashakti Gwalior",  "lat": 26.22, "lon": 78.18, "st": "MP"},
    {"name": "API Ispat",           "lat": 21.25, "lon": 81.60, "st": "CG"},
    {"name": "Aditya Industries",   "lat": 30.70, "lon": 75.60, "st": "PB"},
    {"name": "Ambashakti UP",       "lat": 28.45, "lon": 77.70, "st": "UP"},
    {"name": "German Green Steel",  "lat": 23.20, "lon": 78.00, "st": "MP"},
    {"name": "Maharashtra Plant",   "lat": 18.50, "lon": 73.90, "st": "MH"},
    {"name": "Rashmi Steel",        "lat": 22.60, "lon": 88.30, "st": "WB"},
    {"name": "Real Ispat",          "lat": 25.30, "lon": 83.00, "st": "UP"},
    {"name": "SKA Ispat",           "lat": 27.50, "lon": 81.50, "st": "UP"},
    {"name": "Telangana Plant",     "lat": 17.40, "lon": 78.50, "st": "TG"},
]

STATE_CENTROIDS = {
    "UTTAR PRADESH": [27.0, 80.9], "RAJASTHAN": [26.0, 73.2],
    "DELHI": [28.65, 77.2], "HARYANA": [29.2, 76.0],
    "PUNJAB": [31.0, 75.3], "UTTARAKHAND": [30.3, 79.0],
    "JAMMU & KASHMIR": [33.7, 75.3], "HIMACHAL PRADESH": [31.8, 77.2],
    "ODISHA": [20.5, 84.0], "MADHYA PRADESH": [23.5, 78.6],
    "MAHARASHTRA": [19.5, 76.0], "JHARKHAND": [23.6, 85.3],
    "WEST BENGAL": [23.0, 87.9], "BIHAR": [25.6, 85.9],
    "GUJARAT": [22.3, 71.6], "ANDHRA PRADESH": [15.9, 79.7],
    "TELANGANA": [17.8, 79.0], "ASSAM": [26.1, 92.9],
    "CHHATTISGARH": [21.3, 81.9],
}


def _fmt(v):
    if v >= 1000:
        return f"{v/1000:.1f}K"
    return f"{v:.0f}"


def generate():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(GEOJSON_PATH, 'r') as f:
        geojson = json.load(f)
    with open(DATA_PATH, 'r') as f:
        plant_data = json.load(f)

    # Aggregate
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

    all_states = set(s550) | set(s_oh) | set(s550d)
    state_totals = {s: s550.get(s,0)+s_oh.get(s,0)+s550d.get(s,0) for s in all_states}

    # Plant totals
    plant_keys = [p["name"] for p in PLANTS]
    plant_full = {
        "Ambashakti Gwalior": "AMBASHAKTI UDYOG LIMITED- GWALIOR",
        "API Ispat": "API Ispat And Powertech Private Limited",
        "Aditya Industries": "Aditya Industries",
        "Ambashakti UP": "Ambashakti Industries Limited",
        "German Green Steel": "GERMAN GREEN STEEL AND POWER LIMITED",
        "Maharashtra Plant": "Maharashtra - New Plant",
        "Rashmi Steel": "Rashmi Steel",
        "Real Ispat": "Real Ispat",
        "SKA Ispat": "SKA Ispat Private Limited",
        "Telangana Plant": "Telangana - New Plant",
    }
    for p in PLANTS:
        fn = plant_full[p["name"]]
        pd = plant_data.get(fn, {})
        p["fe550"] = sum(pd.get("fe_550", {}).values())
        p["oh"] = sum(pd.get("one_helix", {}).values())
        p["fe550d"] = sum(pd.get("fe_550d", {}).values())
        p["total"] = p["fe550"] + p["oh"] + p["fe550d"]

    # Inject into geojson
    for feat in geojson['features']:
        sn = feat['properties'].get('STNAME', '')
        feat['properties']['total_volume'] = state_totals.get(sn, 0)
        feat['properties']['fe550'] = s550.get(sn, 0)
        feat['properties']['oh'] = s_oh.get(sn, 0)
        feat['properties']['fe550d'] = s550d.get(sn, 0)

    # Build state label data
    state_labels = []
    for st in sorted(all_states):
        c = STATE_CENTROIDS.get(st)
        if not c: continue
        state_labels.append({
            "name": st,
            "lat": c[0], "lon": c[1],
            "fe550": s550.get(st, 0),
            "oh": s_oh.get(st, 0),
            "fe550d": s550d.get(st, 0),
            "total": state_totals[st],
        })

    t550 = sum(s550.values())
    t_oh = sum(s_oh.values())
    t550d = sum(s550d.values())
    grand = t550 + t_oh + t550d

    geojson_str = json.dumps(geojson)
    labels_str = json.dumps(state_labels)
    plants_str = json.dumps(PLANTS)

    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>JSW ONE TMT — Supply Map V7</title>
<script src="https://d3js.org/d3.v7.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Inter',system-ui,sans-serif; background:#fff; overflow:hidden; }}

#container {{ display:flex; width:100vw; height:100vh; }}

/* ── Map area ───────────────────────────────────── */
#map-area {{
  flex:1; position:relative; background:#fff;
  display:flex; align-items:center; justify-content:center;
}}

/* ── Title ──────────────────────────────────────── */
.title-bar {{
  position:absolute; top:16px; left:50%; transform:translateX(-50%);
  text-align:center; z-index:10;
}}
.title-bar h1 {{
  font-size:20px; font-weight:900; color:#1a1a2e;
  letter-spacing:-0.3px; margin:0;
}}
.title-bar h2 {{
  font-size:12px; font-weight:600; color:#64748b;
  margin-top:2px; letter-spacing:0.5px;
}}

/* ── Sidebar ────────────────────────────────────── */
#sidebar {{
  width:300px; min-width:300px; background:#f8fafc;
  border-left:1px solid #e2e8f0; padding:20px 18px;
  display:flex; flex-direction:column; gap:0;
  overflow-y:auto; font-size:13px;
}}

.sb-section {{ margin-bottom:18px; }}
.sb-title {{
  font-size:11px; font-weight:800; color:#94a3b8;
  text-transform:uppercase; letter-spacing:1.2px;
  margin-bottom:10px;
}}

/* Color key */
.ck-row {{ display:flex; align-items:center; gap:8px; margin-bottom:7px; }}
.ck-dot {{ width:14px; height:14px; border-radius:3px; flex-shrink:0; }}
.ck-label {{ font-size:12.5px; font-weight:600; color:#334155; }}

/* Grade breakdown */
.gb-row {{
  display:flex; justify-content:space-between; align-items:center;
  padding:5px 0; border-bottom:1px solid #f1f5f9;
}}
.gb-row:last-child {{ border-bottom:none; }}
.gb-name {{ font-weight:700; font-size:12.5px; }}
.gb-val {{ font-weight:700; font-size:12.5px; color:#1e293b; font-variant-numeric:tabular-nums; }}
.gb-total {{
  display:flex; justify-content:space-between;
  padding:8px 0 0; margin-top:4px;
  border-top:2px solid #cbd5e1;
  font-weight:900; font-size:14px; color:#0f172a;
}}

/* State table */
.st-table {{ width:100%; border-collapse:collapse; font-size:11px; }}
.st-table th {{
  text-align:left; font-weight:700; color:#64748b;
  padding:4px 3px; border-bottom:2px solid #e2e8f0;
  font-size:10px; text-transform:uppercase; letter-spacing:0.5px;
}}
.st-table td {{
  padding:3px 3px; border-bottom:1px solid #f1f5f9;
  font-variant-numeric:tabular-nums; font-weight:600;
}}
.st-table tr:hover {{ background:#f0fdf4; }}
.st-name {{ font-weight:700; color:#334155; font-size:10.5px; }}

/* Tooltip */
.map-tooltip {{
  position:absolute; pointer-events:none; z-index:100;
  background:rgba(15,23,42,0.92); color:#fff;
  padding:10px 14px; border-radius:8px;
  font-size:12px; line-height:1.5;
  box-shadow:0 4px 20px rgba(0,0,0,0.25);
  backdrop-filter:blur(8px);
  max-width:220px;
  opacity:0; transition:opacity 0.15s;
}}
.map-tooltip.visible {{ opacity:1; }}
.map-tooltip .tt-state {{ font-weight:800; font-size:13px; margin-bottom:4px; }}
.map-tooltip .tt-row {{ display:flex; justify-content:space-between; gap:12px; }}
.map-tooltip .tt-grade {{ font-weight:600; opacity:0.8; }}
.map-tooltip .tt-val {{ font-weight:700; font-variant-numeric:tabular-nums; }}

/* SVG styles */
.state-path {{
  stroke:#fff; stroke-width:0.5; cursor:pointer;
  transition: opacity 0.15s, stroke-width 0.15s;
}}
.state-path:hover {{ stroke-width:1.5; stroke:#1a1a2e; opacity:0.9; }}
.state-path.inactive {{ fill:#eceff1 !important; stroke:#cfd8dc; cursor:default; }}
.state-path.inactive:hover {{ stroke-width:0.5; stroke:#cfd8dc; opacity:1; }}

.factory-icon {{ cursor:pointer; }}
.factory-icon:hover {{ filter:brightness(1.2) drop-shadow(0 2px 6px rgba(0,0,0,0.3)); }}

/* Grade label on map */
.state-grade-label {{
  font-family:'Inter',sans-serif;
  pointer-events:none;
}}
</style>
</head>
<body>

<div id="container">
  <div id="map-area">
    <div class="title-bar">
      <h1>JSW ONE TMT — Plant-wise Supply Network</h1>
      <h2>FY 2026-27</h2>
    </div>
    <svg id="map-svg"></svg>
    <div class="map-tooltip" id="tooltip"></div>
  </div>

  <div id="sidebar">
    <div class="sb-section">
      <div class="sb-title">Color Key</div>
      <div class="ck-row"><div class="ck-dot" style="background:#b8860b"></div><span class="ck-label">Fe 550</span></div>
      <div class="ck-row"><div class="ck-dot" style="background:#d84315"></div><span class="ck-label">OH (One Helix)</span></div>
      <div class="ck-row"><div class="ck-dot" style="background:#1565c0"></div><span class="ck-label">Fe 550D</span></div>
      <div style="font-size:11px;color:#16a34a;font-style:italic;margin-top:4px;">Darker green = higher supply volume</div>
    </div>

    <div class="sb-section">
      <div class="sb-title">Grade Breakdown</div>
      <div class="gb-row"><span class="gb-name" style="color:#b8860b">Fe 550</span><span class="gb-val">{t550:,.0f} MT</span></div>
      <div class="gb-row"><span class="gb-name" style="color:#d84315">OH</span><span class="gb-val">{t_oh:,.0f} MT</span></div>
      <div class="gb-row"><span class="gb-name" style="color:#1565c0">Fe 550D</span><span class="gb-val">{t550d:,.0f} MT</span></div>
      <div class="gb-total"><span>Total</span><span>{grand:,.0f} MT</span></div>
      <div style="display:flex;justify-content:space-between;font-size:11px;color:#64748b;font-weight:600;margin-top:2px;">
        <span>Plants</span><span>{len(PLANTS)}</span>
      </div>
    </div>

    <div class="sb-section" style="flex:1">
      <div class="sb-title">State-wise Breakdown</div>
      <table class="st-table">
        <thead><tr>
          <th>State</th><th style="text-align:right;color:#b8860b">550</th>
          <th style="text-align:right;color:#d84315">OH</th>
          <th style="text-align:right;color:#1565c0">550D</th>
        </tr></thead>
        <tbody id="state-table-body"></tbody>
      </table>
    </div>
  </div>
</div>

<script>
const geo = {geojson_str};
const stateLabels = {labels_str};
const plants = {plants_str};

// ── Populate state table ────────────────────────────────────────────
const tbody = document.getElementById('state-table-body');
const sorted = stateLabels.slice().sort((a,b) => b.total - a.total);
sorted.forEach(s => {{
  const tr = document.createElement('tr');
  const fmt = v => v >= 1000 ? (v/1000).toFixed(1)+'K' : v > 0 ? v.toFixed(0) : '—';
  tr.innerHTML = `<td class="st-name">${{s.name.split(' ').map(w=>w[0]+w.slice(1).toLowerCase()).join(' ')}}</td>
    <td style="text-align:right;color:#b8860b;font-weight:600">${{fmt(s.fe550)}}</td>
    <td style="text-align:right;color:#d84315;font-weight:600">${{fmt(s.oh)}}</td>
    <td style="text-align:right;color:#1565c0;font-weight:600">${{fmt(s.fe550d)}}</td>`;
  tbody.appendChild(tr);
}});

// ── D3 Map ──────────────────────────────────────────────────────────
const mapArea = document.getElementById('map-area');
const W = mapArea.clientWidth;
const H = mapArea.clientHeight;

const svg = d3.select('#map-svg')
  .attr('width', W).attr('height', H)
  .attr('viewBox', `0 0 ${{W}} ${{H}}`);

// Projection — Mercator centered on India
const projection = d3.geoMercator()
  .center([82, 23])
  .scale(Math.min(W, H) * 1.55)
  .translate([W * 0.48, H * 0.52]);

const path = d3.geoPath().projection(projection);

// Color scale (log)
const volumes = stateLabels.map(s => s.total).filter(v => v > 0);
const vMin = d3.min(volumes) || 1;
const vMax = d3.max(volumes) || 1;
const colorScale = d3.scaleLog()
  .domain([vMin, vMax])
  .range([0, 1])
  .clamp(true);

function stateColor(vol) {{
  if (vol <= 0) return '#eceff1';
  const t = colorScale(vol);
  return d3.interpolate('#e8f5e9', '#1b5e20')(t);
}}

// Tooltip
const tooltip = document.getElementById('tooltip');
function showTip(evt, props) {{
  const fe = props.fe550 || 0, oh = props.oh || 0, fd = props.fe550d || 0;
  const fmt = v => v >= 1000 ? (v/1000).toFixed(1)+'K' : v > 0 ? v.toFixed(0) : '—';
  let html = `<div class="tt-state">${{props.STNAME}}</div>`;
  if (fe > 0) html += `<div class="tt-row"><span class="tt-grade" style="color:#fbbf24">550:</span><span class="tt-val">${{fmt(fe)}}</span></div>`;
  if (oh > 0) html += `<div class="tt-row"><span class="tt-grade" style="color:#fb923c">OH:</span><span class="tt-val">${{fmt(oh)}}</span></div>`;
  if (fd > 0) html += `<div class="tt-row"><span class="tt-grade" style="color:#60a5fa">550D:</span><span class="tt-val">${{fmt(fd)}}</span></div>`;
  tooltip.innerHTML = html;
  tooltip.classList.add('visible');
  tooltip.style.left = (evt.clientX + 14) + 'px';
  tooltip.style.top = (evt.clientY - 10) + 'px';
}}
function hideTip() {{ tooltip.classList.remove('visible'); }}

// Draw states
svg.selectAll('path.state-path')
  .data(geo.features)
  .join('path')
    .attr('class', d => {{
      const vol = d.properties.total_volume || 0;
      return 'state-path' + (vol <= 0 ? ' inactive' : '');
    }})
    .attr('d', path)
    .attr('fill', d => stateColor(d.properties.total_volume || 0))
    .on('mouseover', function(evt, d) {{
      if (d.properties.total_volume > 0) showTip(evt, d.properties);
    }})
    .on('mousemove', function(evt) {{
      tooltip.style.left = (evt.clientX + 14) + 'px';
      tooltip.style.top = (evt.clientY - 10) + 'px';
    }})
    .on('mouseout', hideTip);

// ── Grade labels on states ──────────────────────────────────────────
const fmt = v => v >= 1000 ? (v/1000).toFixed(1)+'K' : v > 0 ? v.toFixed(0) : '';

const labelG = svg.selectAll('g.state-grade-label')
  .data(stateLabels)
  .join('g')
    .attr('class', 'state-grade-label')
    .attr('transform', d => {{
      const [x, y] = projection([d.lon, d.lat]);
      return `translate(${{x}},${{y}})`;
    }});

labelG.each(function(d) {{
  const g = d3.select(this);
  const lines = [];
  if (d.fe550 > 0) lines.push({{ prefix: '550', val: fmt(d.fe550), color: '#b8860b' }});
  if (d.oh > 0)    lines.push({{ prefix: 'OH',  val: fmt(d.oh),    color: '#d84315' }});
  if (d.fe550d > 0) lines.push({{ prefix: '550D',val: fmt(d.fe550d),color: '#1565c0' }});
  if (lines.length === 0) return;

  const lineH = 14;
  const padX = 6, padY = 4;
  const totalH = lines.length * lineH;
  const cardH = totalH + padY * 2;
  const cardW = 68;
  const startY = -cardH / 2;

  // Card background
  g.append('rect')
    .attr('x', -cardW/2).attr('y', startY)
    .attr('width', cardW).attr('height', cardH)
    .attr('rx', 4).attr('ry', 4)
    .attr('fill', 'rgba(255,255,255,0.92)')
    .attr('stroke', '#94a3b8').attr('stroke-width', 0.6);

  // Grade lines
  lines.forEach((ln, i) => {{
    const ty = startY + padY + i * lineH + lineH * 0.75;
    g.append('text')
      .attr('x', -cardW/2 + padX).attr('y', ty)
      .attr('font-size', '9px').attr('font-weight', '800')
      .attr('fill', ln.color)
      .text(ln.prefix + ':');
    g.append('text')
      .attr('x', cardW/2 - padX).attr('y', ty)
      .attr('text-anchor', 'end')
      .attr('font-size', '9px').attr('font-weight', '700')
      .attr('fill', '#1e293b')
      .text(ln.val);
  }});
}});

// ── Factory icons ───────────────────────────────────────────────────
const factoryG = svg.selectAll('g.factory-icon')
  .data(plants)
  .join('g')
    .attr('class', 'factory-icon')
    .attr('transform', d => {{
      const [x, y] = projection([d.lon, d.lat]);
      return `translate(${{x}},${{y}})`;
    }});

factoryG.each(function(d) {{
  const g = d3.select(this);
  const s = 0.6; // scale
  g.attr('transform', g.attr('transform') + ` scale(${{s}})`);

  // Chimney 1
  g.append('rect').attr('x',-12).attr('y',-22).attr('width',5).attr('height',10)
    .attr('rx',1).attr('fill','#455a64');
  // Chimney 2
  g.append('rect').attr('x',-3).attr('y',-18).attr('width',5).attr('height',6)
    .attr('rx',1).attr('fill','#546e7a');
  // Smoke
  g.append('circle').attr('cx',-10).attr('cy',-23).attr('r',2).attr('fill','#b0bec5').attr('opacity',0.5);
  g.append('circle').attr('cx',-1).attr('cy',-19).attr('r',1.5).attr('fill','#b0bec5').attr('opacity',0.4);
  // Main body
  g.append('rect').attr('x',-14).attr('y',-12).attr('width',20).attr('height',14)
    .attr('rx',1.5).attr('fill','#37474f');
  // Extension
  g.append('polygon').attr('points','6,-12 6,2 14,2 14,-5').attr('fill','#455a64');
  // Windows
  g.append('rect').attr('x',-10).attr('y',-8).attr('width',4.5).attr('height',4.5)
    .attr('rx',0.5).attr('fill','#fdd835');
  g.append('rect').attr('x',-3).attr('y',-8).attr('width',4.5).attr('height',4.5)
    .attr('rx',0.5).attr('fill','#fdd835');
  // Door
  g.append('rect').attr('x',7).attr('y',-4).attr('width',4).attr('height',6)
    .attr('rx',0.5).attr('fill','#fdd835').attr('opacity',0.7);
  // Base
  g.append('rect').attr('x',-15).attr('y',1.5).attr('width',30).attr('height',1.5)
    .attr('fill','#263238');
}});

// Add plant tooltips
factoryG.on('mouseover', function(evt, d) {{
  const fmt = v => v >= 1000 ? (v/1000).toFixed(1)+'K' : v > 0 ? v.toFixed(0) : '—';
  let html = `<div class="tt-state">${{d.name}}</div>
    <div style="font-size:11px;color:#94a3b8;margin-bottom:4px">${{d.st}}</div>`;
  if (d.fe550 > 0) html += `<div class="tt-row"><span class="tt-grade" style="color:#fbbf24">550:</span><span class="tt-val">${{fmt(d.fe550)}}</span></div>`;
  if (d.oh > 0)    html += `<div class="tt-row"><span class="tt-grade" style="color:#fb923c">OH:</span><span class="tt-val">${{fmt(d.oh)}}</span></div>`;
  if (d.fe550d > 0) html += `<div class="tt-row"><span class="tt-grade" style="color:#60a5fa">550D:</span><span class="tt-val">${{fmt(d.fe550d)}}</span></div>`;
  html += `<div style="border-top:1px solid rgba(255,255,255,0.2);margin-top:4px;padding-top:4px;font-weight:800">${{d.total.toLocaleString()}} MT</div>`;
  tooltip.innerHTML = html;
  tooltip.classList.add('visible');
  tooltip.style.left = (evt.clientX + 14) + 'px';
  tooltip.style.top = (evt.clientY - 10) + 'px';
}})
.on('mousemove', function(evt) {{
  tooltip.style.left = (evt.clientX + 14) + 'px';
  tooltip.style.top = (evt.clientY - 10) + 'px';
}})
.on('mouseout', hideTip);

</script>
</body>
</html>'''

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"HTML map saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
