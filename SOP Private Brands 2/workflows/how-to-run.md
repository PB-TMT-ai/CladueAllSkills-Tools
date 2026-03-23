# How to Run the SOP Generator

[Back to Overview](sop_excel_generation.md)

## Basic Usage

```bash
cd "D:\SOP Private Brands 2"
python tools/generate_sop_excel.py
```

## Optional Arguments

```bash
python tools/generate_sop_excel.py --docs-dir "path/to/docs" --output "path/to/output"
```

| Argument | Default | Description |
|----------|---------|-------------|
| `--docs-dir` | `Documents/` | Path to source documents folder |
| `--output` | `output/` | Output directory for generated files |

## Output Files

- `output/JSW_ONE_PB_SOPs_Master.xlsx` - The master Excel file
- `output/manifest.json` - Build manifest (file hashes for change detection)
- `output/.ocr_cache/` - Cached OCR text from scanned PDFs (speeds up re-runs)

## What Happens When You Run

1. **Discovers** all `.doc`, `.docx`, and `.pdf` files in `Documents/`
2. **Deduplicates** files by MD5 hash
3. **Parses** each document (MIME/HTML for .doc, python-docx for .docx, OCR for .pdf)
4. **Classifies** documents and maps to journey/section structure
5. **Matches** Confluence links from `PB SOP Links Confluence.xlsx`
6. **Generates** formatted Excel with merged cells, hyperlinks, and styling
7. **Saves** manifest for change detection on next run
