# Price Difference Calculations

## Run
Close Excel first.
```bash
python tools/run_all.py "<filename>.xlsx" "<sheet-name>"
```
```bash
python tools/run_all.py "Price difference calculations - Jan26.xlsx" "Jan-26"
python tools/run_all.py   # Dec-25 defaults
```
Sheet name format: `Mon-YY`. Pricing sheet auto-derived as `"<name> pricing"`.
Auto-detects column layout from sheet names.

## Structure
```
tools/run_all.py     # Single script — does everything
workflows/run_all.md # SOP
```

## Layout Detection
| Sheet exists | Layout | Lookup sheet | Lookup data starts | Outputs |
|---|---|---|---|---|
| `union book` | Dec-25 | union book | Row 2 | P-T, V, W |
| `Order` | Jan-26+ | Order | Row 3 | L, M, N, O |

Both lookup sheets share column positions: C=Order ID, S=Grade, T=Diameter, U=Form, AD=Proposed Price, BG=SFDC Comment.

## Main Sheet Columns

### Dec-25 (col A = Invoice date)
Inputs: A(Invoice date), C(Order ID), H(Grade), I(Diameter), J(Form), K(Pin code), L/M/N(Order Y/M/D).
Outputs: P/Q/R(Invoice Y/M/D), S(SF comment), T(Cluster), V(Proposed), W(Applicable).

### Jan-26+ (col A = Order ID)
Inputs: A(Order ID), F(Grade), G(Diameter), H(Form), I(Pin code), J(Order Day).
Outputs: L(SF comment), M(Cluster), N(Proposed), O(Applicable).
Date period: Order Day + month/year parsed from sheet name.

## Pricing Model
```
FE 550:  Delhi_FE550_base[period]  + Ladder[cluster] + diameter_extra + form_extra
FE 550D: Delhi_FE550D_base[period] + Ladder[cluster] + diameter_extra + form_extra
```
Extras (Instructions sheet): 8mm +3500, 10mm +2250, U-Bend/Fish-Bend +600.

## Results

### Dec-25 (2,426 rows)
P/Q/R 100%, S 100% match, T 99.6%, V 99.7%, W 96.9%.
Gaps: pincodes 122022/122050, 7 placeholder rows, 27 month=0, 31 GUJARAT.

### Jan-26 (2,175 rows)
L 41.1% (correct — 60% have no comment), M 99.7%, N 100%, O 99.7%.
Gaps: pincode 201308 (6 rows).

## Key Patterns
- openpyxl, close Excel before running.
- Two-pass: read-only (`data_only=True`) for lookups + pricing, then read-write for output.
- Headers row 3, data row 4. Normalize: `.strip().upper()` for grades/forms/clusters.
