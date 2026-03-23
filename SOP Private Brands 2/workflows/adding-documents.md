# Adding New Documents

[Back to Overview](sop_excel_generation.md)

## Steps

1. **Save the document** (.doc, .docx, or .pdf) into the `Documents/` folder

2. **Add a classification entry** in `tools/mapping/document_classifier.py`:

   ```python
   "NewDocPattern*": {
       "journey": "Pre - Order",        # or "Order" or "Post Order"
       "section": "SECTION LABEL",       # e.g., "G. INFLUENCER MANAGEMENT"
       "doc_type": "ops_sop",            # see document-types.md for options
       "activity_prefix": "Display Name",
   },
   ```

3. **Add Confluence link** (if available) to `Documents/PB SOP Links Confluence.xlsx`:
   - Column A: SOP name
   - Column B: Confluence URL

4. **Re-run the generator**:
   ```bash
   python tools/generate_sop_excel.py
   ```

## For PDF Documents (Scanned)

PDF documents require Tesseract OCR and Poppler installed on the system. The first run will be slow (OCR processing), but results are cached for subsequent runs.

If the PDF is a quality/ops manual spanning multiple journey phases, use `multi_journey` in the config:

```python
"New Quality Doc*": {
    "journey": "Order",
    "section": "J. QUALITY ASSURANCE",
    "doc_type": "quality_manual",
    "activity_prefix": "Quality Manual Name",
    "multi_journey": {
        "Order": "J. QUALITY ASSURANCE (MANUFACTURING)",
        "Post Order": "J. QUALITY ASSURANCE (DISPATCH)",
    },
},
```

## Notes

- Filename patterns use `*` wildcards (fnmatch syntax)
- `+` in filenames is normalized to spaces during matching
- Duplicate files (same MD5 hash) are automatically skipped
