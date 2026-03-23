# Project: Market Feedback Dashboard

## Core Principle
You operate as the decision-maker in a modular system. Your job is NOT to do everything
yourself. Your job is to read instructions, pick the right tools, handle errors intelligently, and
improve the system over time.

Why? 90% accuracy across 5 steps = 59% total success. Push repeatable work into tested
scripts. You focus on decisions.

## System Architecture
**Blueprints (/blueprints)** - Step-by-step instructions in markdown. Goal, inputs, scripts to use,
output, edge cases. Check here FIRST.

**Scripts (/scripts)** - Tested, deterministic code. Call these instead of writing from scratch.

**Workspace (/.workspace)** - Temp files. Never commit. Delete anytime.

## How You Operate
1. Check blueprints first - If one exists, follow it exactly
2. Use existing scripts - Only create new if nothing exists
3. Fail forward - Error -> Fix -> Test -> Update blueprint -> Add to LEARNINGS.md -> System smarter
4. Ask before creating - Don't overwrite blueprints without asking

## Tech Stack
- **Frontend**: Vanilla JavaScript (ES6+), HTML5
- **Styling**: Tailwind CSS (CDN) + custom CSS (`src/styles/dashboard.css`)
- **Charts**: Chart.js 4.x + @sgratzl/chartjs-chart-boxplot plugin (CDN)
- **Data Pipeline**: Python 3 (pandas, openpyxl, json)
- **Build**: None — static files served directly (no bundler)
- **Server**: Python `http.server` for local dev (port 8080)

## Project Structure
```
/index.html                  — Main page (loads CDN libs + dashboard.js)
/refresh.bat                 — Runs extract_data.py to regenerate data
/src/dashboard.js            — Dashboard logic: filters, charts, tables (848 lines)
/src/styles/dashboard.css    — Custom dark-theme styling (456 lines)
/src/data/dashboard-data.js  — Auto-generated from extract_data.py (DO NOT edit)
/scripts/extract_data.py     — Data extraction + JSW price parsing
/data/                       — Source Excel files (Market Feedback Report + Price List)
/blueprints/                 — Task SOPs
/.workspace/                 — Temp files (gitignored)
```

## Data Pipeline
1. Place Excel files in `/data/`
2. Run `python scripts/extract_data.py` (or double-click `refresh.bat`)
3. Script reads Market Feedback Report + Price List.xlsx
4. Outputs `src/data/dashboard-data.js` with all data merged
5. Refresh browser to see updated dashboard

## Data Record Schema
Each data point uses compact keys for performance:
- `b` = brand name (string)
- `a` = amount/price in INR, excl. GST (number)
- `s` = state (uppercase string)
- `d` = district (uppercase string, empty for JSW)
- `m` = month ("Jan-26" format)
- `q` = quality ("BIS" or "NonBIS")
- `t` = delivery timeliness (string or null)
- `c` = company/dealer name (string)
- `w` = week of month (1-5)
- `wl` = week label ("Jan-26-W3" format)

## JSW ONE TMT Integration
- Two brands: "JSW ONE TMT 550" (Fe-550) and "JSW ONE TMT 550D" (Fe-550D)
- Source: `data/Price List.xlsx` (date-wise tabs)
- Prices are distributor landed — script adds Rs 3000 for dealer landed
- JSW ONE TMT 550 is set as `benchmarkBrand` (gets visual emphasis)
- See `blueprints/jsw-price-integration.md` for details

## Code Standards
- Vanilla JS — no frameworks, no transpilation
- Functions are procedural and self-contained in dashboard.js
- Chart.js instances stored in `State.charts` object
- Filter state managed via global `State` object
- Format prices with `formatINR()`, states with `titleCase()`
- Max 5 brands selectable via custom multi-select component

## Error Protocol
1. Stop and read the full error
2. Isolate - which component/script failed
3. Fix and test
4. Document in LEARNINGS.md
5. Update relevant blueprint

## What NOT To Do
- Don't skip blueprint check
- Don't ignore errors and retry blindly
- Don't create files outside structure
- Don't write from scratch when blueprint exists
- Don't manually edit `src/data/dashboard-data.js` — it's auto-generated
