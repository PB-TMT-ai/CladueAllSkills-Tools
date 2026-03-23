# Workflow: Generate Distributor One-Pager Reports

## Objective
Generate `.docx` one-pager reports for JSW One TMT distributors from the master Excel file.

## Required Inputs
- Excel file: `data/JSW One TMT_distributor reports.xlsx`
- Distributor Name (exact match from `FY 26_Distributor-wise achi.` sheet)
- State (exact match)

## Tools Used
1. `tools/extract_distributor_data.py` - Data extraction
2. `tools/generate_one_pager.py` - Docx generation
3. `tools/batch_generate.py` - Batch generation by zone

## How to Run

### Single distributor:
```bash
python tools/generate_one_pager.py "DISTRIBUTOR NAME" "STATE"
```

### Batch by zone:
```bash
python tools/batch_generate.py North              # North region, today's date
python tools/batch_generate.py East 2026-03-09    # East region, specific date
```
Zones: North, East, West, Central. Loads workbook once, outputs to `output/{date}/`.

## Data Source Mapping

### Source Sheets:
| Sheet | Purpose |
|---|---|
| `FY 26_Distributor-wise achi.` | Monthly achieved sales, revised MoU (49 distributors) |
| `FY 26_Original MoU_targets` | Monthly target values (original MoU) |
| `FY 26 dealer sales` | Dealer-level sales, onboarding, counter potential (1848 rows) |
| `District master` | District info, manpower, district-level sales (460 rows) |
| `Invoice` | Invoice-level data with Business Unit & Address tagging (49K rows) |

### Key Filter: Distributor + State
- Some distributors operate in multiple states (e.g., PAL CEMENT AND STEEL PVT. LTD. in HARYANA, PUNJAB, HIMACHAL PRADESH)
- Each combo produces one separate report
- District master must be filtered by BOTH distributor AND state

### Report Sections and Data Sources:

**Section 1: Basic Details** - District master (filtered by distributor + state)
**Section 2: Sales Manpower** - SM/TM from `data/District wise SM_TM name.xlsx` (fallback: District master). DGO/DSR from District master. 2-column table.
**Section 3: Sales Performance** - Distributor-wise achi + Original MoU targets. 11 months (no Mar'26).
**Section 4: Sales by Location** - Invoice sheet (Business Unit + Address taging columns). Parent rows show integer sums, child rows show **% only**. 11 months.
**Section 5: Channel Performance** - FY 26 dealer sales (aggregated by distributor). Rows: Sec. sales, New dealers, Transacting dealers, Active dealers, Trans/Active ratio (shown as %). No FY26 total for trans/active ratio.
**Section 6: Key Districts** - District master + FY 26 dealer sales (VH/H/M only)
**Section 7: Top 10 Dealers** - FY 26 dealer sales (sorted by FY26 total)

## Key Metric Definitions

| Metric | Formula | Source |
|---|---|---|
| Active dealer | Transacted at least once in the last 3 months | FY 26 dealer sales |
| New dealer | Counted by Onboarded Dealer Month (e.g., "Apr'25") | FY 26 dealer sales |
| Transacting dealer | Sales > 0 in that specific month | FY 26 dealer sales |
| SoB% | FY26 Avg Monthly Sales / Counter Potential of transacting dealers only | FY 26 dealer sales |
| Reach% | Counter Potential of transacting dealers / Retail Demand per month | Dealer sales + District master |
| Market Share% | Reach% * SoB% | Calculated |

## Known Issues / Edge Cases
- Invoice dates are stored as strings (format: YYYY-MM-DD), not datetime objects - parsed via `strptime`
- Some dealers have `None` counter potential - treated as 0
- Dealer segmentation values: Super Star, Star, Friendly, Challenger, Minor, Alien, Potential Not updated
- FY26 = Apr 2025 through Mar 2026 (11 months elapsed as of Feb'26)
- Mar'26 data may be 0 (future month)
- SoB% and Reach% use counter potential of **only transacting dealers** (lifted that month), averaged across months with transactions
- Sales by Location child rows display **% only** (not value + %)
- All numbers rounded to nearest integer (no decimals)
- Mar'26 column excluded from all monthly tables (11 months shown)
- Title format: `JSW ONE TMT - <Name> - Performance Feb'26`

## Output
- Single mode: saved to `output/{DISTRIBUTOR_NAME}_{STATE}.docx`
- Batch mode: saved to `output/{YYYY-MM-DD}/{DISTRIBUTOR_NAME}_{STATE}.docx`
- Naming: spaces replaced with underscores
- Format: A4 landscape, Calibri font, blue header styling
