# Run All Tasks

## Command
```bash
python tools/run_all.py "<filename>.xlsx" "<sheet-name>"
```
Close Excel first. ~2-10 min depending on file size.

## New Month Checklist
1. Get the new Excel file
2. Confirm sheets exist: main sheet, `"<sheet> pricing"`, lookup sheet (`union book` or `Order`), `Pincode`, `Instructions`
3. Close Excel
4. Run the command
5. Check output stats

## Layout Auto-Detection
- `"union book"` sheet → Dec-25 layout (outputs P-T, V, W)
- `"Order"` sheet → Jan-26+ layout (outputs L, M, N, O)

## Adding a New Layout
If a future file has neither `union book` nor `Order`:
1. In `run_all.py` → `detect_layout()`: add the new sheet name
2. In `get_layout_config()`: add column mappings (sentinel, order_id, grade, diameter, form, pincode, order_day, outputs)
3. Set `has_invoice_date`, `has_order_ymd`, `lookup_sheet`, `lookup_data_start`
