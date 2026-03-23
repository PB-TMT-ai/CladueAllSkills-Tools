# Learnings - JSW Steel TMT Marketing Budget Model

## Reference File Analysis

### Key Data Extracted
- **One Helix**: 42 distributors, 27K MT/month exit rate, 121K MT annual, Rs. 48,868/MT selling price
- **JSW ONE TMT**: 56 distributors, 31K MT/month exit rate, 268K MT annual, Rs. 50,000/MT selling price
- **Dealer counts** (from #Dealers sheet): UP=152, Haryana=81, WB=159, Bihar=70, Jharkhand=61, Odisha=48 (JSW ONE baseline)
- **Quarterly pattern**: Q1=25%, Q2=15%, Q3=35% (peak construction), Q4=25%
- **Reference budget**: ~Rs. 2.6 Cr total across both brands (our model uses Rs. 5 Cr as per brief)

### Activity Cost Benchmarks (from reference)
- Wall Painting: Rs. 14-15/sq ft
- FM Radio: Rs. 2.5 Lac/campaign
- Hoardings: Rs. 50K each
- Digital: Rs. 5-6 Lac/campaign
- Events: Rs. 5 Lac/event
- Van Campaign: Rs. 7,000/day
- NLB: Rs. 1,200 (40/sqft x 30 sqft)
- GSB: Rs. 6,510
- Mason Meet: Rs. 300/pax (30 pax/meet)
- Contractor Meet: Rs. 1,000/pax
- Architect Meet: Rs. 3,500/pax

### Seasonal Patterns
- Q3 (Oct-Dec) gets highest allocation due to peak construction season
- Q2 (Jul-Sep) gets lowest due to monsoon slowdown
- Helix launch should front-load in Q1 for brand awareness, then ramp in Q3

## Design Decisions Made
1. **Budget split 60:40** (Helix:JSW ONE) - prioritizes national launch investment
2. **ATL/BTL 40:60** for Helix - BTL-heavy approach for dealer/retailer activation despite being a launch
3. **Sub-2A:Sub-2B = 60:40** within Objective 2 - broader East gets more due to wider geography
4. **4 East states**: Bihar, WB, Jharkhand, Odisha (core East, excluding UP/HR which are separate)
5. **Helix volume ramp-up curve**: 10/15/30/45% across Q1-Q4 (new brand launch curve)

## Technical Learnings
- openpyxl handles Excel formulas as strings - they evaluate when opened in Excel
- Sheet names with spaces require single quotes in cross-tab formulas: `'Master Inputs'!B5`
- Conditional formatting via CellIsRule and FormulaRule works well for validation
- Data validation prevents invalid inputs at the cell level
- Chart creation in openpyxl is limited but functional for basic bar/pie charts
- Sheet reordering via `wb.move_sheet(name, offset=delta)` works with calculated deltas

## Potential Improvements
- Add state-wise breakdown within Helix (North/South/East/West zones)
- Include competitor benchmarking data for cost/MT comparison
- Add monthly actual volume tracking alongside spend actuals
- Include digital campaign KPIs (impressions, CTR, leads) as sub-metrics
- Consider adding a "What-If Simulator" tab with data tables for sensitivity analysis
