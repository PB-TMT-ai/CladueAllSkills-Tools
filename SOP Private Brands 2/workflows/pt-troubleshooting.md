# P&T Troubleshooting & Edge Cases

[Back to Overview](sop_excel_generation.md)

## Common Issues

| Issue | Solution |
|-------|----------|
| Document skipped: "No classification found" | Add entry to `PT_DOCUMENT_MAPPING` in `tools/generate_pt_sop_excel.py` |
| Empty steps for a document | Check if the document uses numbered section headers (e.g., "6.0 Test Procedure:") — these are handled, but custom formats may need parser updates |
| Empty purpose/description | Verify the document has a "Purpose:" section (with or without "1.0" prefix) |
| Content duplicate skipped | Expected for files with `[1]` suffix (same MD5 hash); not an error |
| Fewer activities than expected | Check console for `ERROR` messages; verify all files are in `Documents/SOP_Pipes & tubes/` |
| One_Helix activities missing | Check that the document has 9 tables in the expected format (see `pt-document-types.md`) |

## Known Edge Cases

- **Numbered section headers**: Documents like Bend Test, Tensile Test use "1.0 Purpose:", "6.0 Test Procedure:" instead of plain "Purpose:", "Process:". The parser strips leading number prefixes automatically.
- **"Record" false positive**: "Record production quantity." in Finishing Activities must not trigger the "Records" section break. The parser uses a regex that requires the word "Records" to stand alone (not followed by other content words).
- **Traceability duplicate**: `SOP_P&T_Traceability and identification products_14[1].docx` is a byte-for-byte duplicate of the non-`[1]` version. MD5 dedup handles this correctly.
- **One_Helix custom format**: This document does NOT follow the standard JSW ONE template. It has a custom responsibilities table and 6 activity-based tables. The `parse_one_helix_docx()` function handles it separately.
- **Special characters in documents**: Some docs contain characters like `\xa0` (non-breaking space), `\u2013` (en-dash), degree symbols. UTF-8 handling in python-docx manages these correctly.

## Adding a New Section

If a new P&T document doesn't fit existing sections:

1. Add the new section label to `PT_SECTION_ORDER` in `tools/generate_pt_sop_excel.py`
2. Use the new section label in the `PT_DOCUMENT_MAPPING` entry
3. Sections are ordered by their position in `PT_SECTION_ORDER` within each journey

## System Requirements

| Component | Required For |
|-----------|-------------|
| Python 3.10+ | All processing |
| openpyxl | Excel generation (`pip install openpyxl`) |
| python-docx | .docx parsing (`pip install python-docx`) |

Note: Tesseract OCR and Poppler are NOT needed for P&T (all documents are .docx).
