# Adding New P&T Documents

[Back to Overview](sop_excel_generation.md)

## Steps

1. **Save the document** (`.docx`) into `Documents/SOP_Pipes & tubes/`

2. **Add a classification entry** in `tools/generate_pt_sop_excel.py` under `PT_DOCUMENT_MAPPING`:

   ```python
   "SOP for New Process*": {
       "journey": "Production",              # Pre-Production, Production, or Post-Production
       "section": "D. QUALITY TESTING",       # See section list below
       "activity_prefix": "New Process Name",
   },
   ```

3. **Re-run the generator**:
   ```bash
   python tools/generate_pt_sop_excel.py
   ```

## Available Sections

| Journey | Section |
|---------|---------|
| Pre-Production | A. RAW MATERIAL MANAGEMENT |
| Pre-Production | B. PRODUCTION PLANNING |
| Production | C. MANUFACTURING & FINISHING |
| Production | D. QUALITY TESTING |
| Production | E. QUALITY MANAGEMENT |
| Post-Production | F. DISPATCH & LOGISTICS |
| Post-Production | G. DOCUMENTATION & INVOICING |

New sections can be added to `PT_SECTION_ORDER` in the generator script if needed.

## For Comprehensive Documents (Like One_Helix)

If a document contains multiple activities spanning different sections/journeys, add `"is_comprehensive": True` to the mapping and write a custom parser function (see `parse_one_helix_docx()` as an example).

## Notes

- Only `.docx` files are supported (no `.doc` or `.pdf`)
- Filename patterns use `*` wildcards (fnmatch syntax)
- Duplicate files (same MD5 hash) are automatically skipped
- All P&T docs follow the JSW ONE template (header table, teams table, process paragraphs, approval table)
- Documents using numbered section headers ("1.0 Purpose:", "6.0 Test Procedure:") are handled automatically
