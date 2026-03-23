# P&T Validation Checklist

[Back to Overview](sop_excel_generation.md)

After running `python tools/generate_pt_sop_excel.py`, open `output/JSW_ONE_PT_SOPs_Master.xlsx` and verify:

## Structure
- [ ] Sheet tab name is "Pipes & Tubes"
- [ ] All 3 journey phases present (Pre-Production, Production, Post-Production)
- [ ] All 7 sections present (A through G)
- [ ] Activity numbering is sequential (1 through 27)
- [ ] Header row is frozen (scroll down to verify)

## Content
- [ ] All source documents represented (check console output for skipped files)
- [ ] No unexpected duplicates processed
- [ ] Each activity has a description (column D not blank)
- [ ] Each activity has an owner (column E not blank)
- [ ] Each activity has process steps (column H not blank)
- [ ] SOP Link column (J) is blank (expected - no links available yet)
- [ ] One_Helix document split into 8 activities across correct sections

## Formatting
- [ ] Section headers appear with green fill
- [ ] Journey column (A) shows rotated text with blue fill
- [ ] Merged cells render correctly (no overlapping text)
- [ ] Column widths match Private Brands master

## Console Output Checks
- Each document should show its journey/section classification
- Final summary should show:
  - Total journeys: 3
  - Total activities: 27 (or more if new docs added)
- One content duplicate skipped (`Traceability_14[1].docx`)
- No `ERROR` messages in the output

## Current Activity Count by Section

| Section | Expected Activities |
|---------|-------------------|
| A. RAW MATERIAL MANAGEMENT | 5 |
| B. PRODUCTION PLANNING | 1 |
| C. MANUFACTURING & FINISHING | 3 |
| D. QUALITY TESTING | 7 |
| E. QUALITY MANAGEMENT | 5 |
| F. DISPATCH & LOGISTICS | 4 |
| G. DOCUMENTATION & INVOICING | 2 |
| **Total** | **27** |
