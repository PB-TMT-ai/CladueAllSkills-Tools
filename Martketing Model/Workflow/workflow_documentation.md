# JSW Steel TMT Marketing Budget Model - Workflow Documentation

## Overview
This model covers FY 2026-27 (Apr 2026 - Mar 2027) marketing budget planning for:
- **Objective 1**: One Helix TMT national launch (60% of budget)
- **Objective 2A**: JSW ONE TMT East India growth (24% of budget)
- **Objective 2B**: JSW ONE TMT UP & Haryana 2x volume target (16% of budget)

## How to Regenerate the Model
1. Open terminal in `D:\Martketing Model`
2. Run: `python Tools/build_model.py`
3. Output: `Model/JSW_TMT_Marketing_Budget_Model.xlsx`

## How to Use the Excel Model

### Changing Budget Allocation
1. Open **Master Inputs** tab
2. Edit cell C5 (Total Budget - default Rs. 5 Cr)
3. Edit cell C6 (Helix share % - default 60%)
4. Edit cell C8 (Sub-2A East share of Obj2 - default 60%)
5. All other tabs auto-update

### Adjusting ATL/BTL Mix (Helix only)
1. In **Master Inputs**, find "HELIX ATL vs BTL SPLIT" section
2. Change ATL Share % (default 40%, meaning 60% BTL)

### Changing Quarterly Distribution
1. In **Master Inputs**, find "QUARTERLY ALLOCATION" section
2. Adjust Q1-Q4 percentages (must sum to 100%)
3. The "Quarterly Check" cell shows YES/NO validation

### Updating Activity Unit Costs
1. Go to **Activity Costs** tab
2. Edit yellow cells in column C (Unit Cost)
3. Changes cascade to all objective tabs automatically

### Tracking Actuals (Post-Quarter)
1. Go to **Quarterly Review** tab
2. Enter actual spend in yellow cells for the completed quarter
3. Variance and deviation % auto-calculate
4. Cells turn red if deviation exceeds 15%

## Verification Checks (Built-in)
- Master Inputs tab has YES/NO checks for budget allocation, sub-objective sums, quarterly splits
- Each objective tab shows Budget vs Spend variance (green = within budget, red = over)
- Consolidated tab shows overall variance by objective

## Color Legend
| Color | Meaning |
|-------|---------|
| Yellow | Editable input cell |
| Green | Auto-calculated value |
| Blue | ATL activity / output |
| Gray | Total / summary row |
| Red | Alert / over-budget |

## Tab Structure
| # | Tab Name | Purpose |
|---|----------|---------|
| 1 | Dashboard | Executive summary with charts and KPIs |
| 2 | Master Inputs | All editable parameters |
| 3 | Obj1 Helix | National launch plan (ATL + BTL) |
| 4 | Obj2A East | East India growth (BTL-focused) |
| 5 | Obj2B UPHR | UP & Haryana 2x volume (intensive BTL) |
| 6 | Consolidated | Combined view with hierarchical rollup |
| 7 | Vol Projections | Volume targets, Cost/MT, A/S ratio |
| 8 | Activity Costs | Master unit cost reference |
| 9 | ROI Metrics | Scenario analysis, efficiency metrics |
| 10 | Quarterly Review | Actual vs Plan tracker |

## Formula Flow
```
Master Inputs + Activity Costs
        |
        v
Objective Tabs (3, 4, 5)
        |
        v
Consolidated (6)
        |
        v
Vol Projections (7) + ROI Metrics (9) + Quarterly Review (10)
        |
        v
Dashboard (1)
```
