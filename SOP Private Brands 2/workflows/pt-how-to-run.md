# How to Run the P&T SOP Generator

[Back to Overview](sop_excel_generation.md)

## Basic Usage

```bash
cd "D:\SOP Private Brands 2"
python tools/generate_pt_sop_excel.py
```

No CLI arguments — the script reads from `Documents/SOP_Pipes & tubes/` and writes to `output/`.

## Output Files

- `output/JSW_ONE_PT_SOPs_Master.xlsx` - The master Excel file (sheet tab: "Pipes & Tubes")
- `output/pt_manifest.json` - Build manifest (file hashes, timestamps)

## What Happens When You Run

1. **Discovers** all `.docx` files in `Documents/SOP_Pipes & tubes/`
2. **Deduplicates** files by MD5 hash (e.g., `Traceability_14[1].docx` is skipped)
3. **Classifies** each document via `PT_DOCUMENT_MAPPING` (journey + section)
4. **Parses** each document:
   - Standard JSW ONE template: extracts header metadata, teams, process steps from paragraphs
   - One_Helix comprehensive doc: extracts 8 activities from 9 specialized tables
5. **Builds** Journey > Section > Activity hierarchy (Pre-Production / Production / Post-Production)
6. **Generates** formatted Excel with merged cells, borders, and section/journey formatting
7. **Saves** manifest for change tracking

## Differences from Private Brands Generator

| Feature | Private Brands | Pipes & Tubes |
|---------|---------------|---------------|
| Script | `generate_sop_excel.py` | `generate_pt_sop_excel.py` |
| File types | .doc, .docx, .pdf, .xlsx | .docx only |
| Classification | `document_classifier.py` | `PT_DOCUMENT_MAPPING` in generator |
| Confluence links | Yes (matched from Excel) | Not available yet |
| Journeys | Pre-Order, Order, Post Order | Pre-Production, Production, Post-Production |
| OCR needed | Yes (scanned PDFs) | No (all .docx) |
