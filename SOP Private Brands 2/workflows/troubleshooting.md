# Troubleshooting & Edge Cases

[Back to Overview](sop_excel_generation.md)

## Common Issues

| Issue | Solution |
|-------|----------|
| Document skipped: "No classification found" | Add entry to `DOCUMENT_MAPPING` in `tools/mapping/document_classifier.py` |
| Confluence link not matched | Check title spelling in `PB SOP Links Confluence.xlsx`; add fallback in `confluence_links.py` |
| Empty steps for a document | Check parser output; may need custom extraction rule in `field_extractor.py` |
| Description same as Steps | Description MUST come from a different source than steps — see [Document Types](document-types.md) for per-type rules |
| OCR steps contain table fragments | `_extract_steps_from_block()` filters non-action phrases; check action verb list and min-length thresholds |
| Steps missing from scanned PDF SOPs | Check if SOP uses "Step N" multi-line format (content on next line) — handled by `step_header` regex |
| QAP Testing activity missing | Table may be OCR-garbled; `_extract_qap_from_text()` fallback uses keyword matching |
| Merged cells look wrong in Excel | Ensure openpyxl version >= 3.1; check `excel_writer/writer.py` merge logic |
| PDF OCR produces garbled text | Check Tesseract installation; try increasing DPI in `pdf_parser.py`; review image preprocessing |
| PDF parser fails: "tesseract not found" | Install Tesseract OCR; set `TESSERACT_CMD` env variable to correct path |
| PDF parser fails: "poppler not found" | Install Poppler; set `POPPLER_PATH` env variable to the `bin/` directory |
| OCR is very slow | First run caches results in `output/.ocr_cache/`; subsequent runs use cache |

## Known Edge Cases

- `.doc` files from Confluence are MIME/HTML format, not binary Word - parsed with Python `email` module
- Duplicate files (same MD5 hash) are automatically skipped
- `PB SOP Links Confluence.xlsx` has a typo ("nfluencer" missing "I") - handled by fuzzy matching
- Two `PB+Retailer_Influencer+in+App` files have same size but different hashes (minor metadata diff)
- "Order Logging Process" heading contains "ORDER" - parser checks for "ORDER PHASE" specifically to avoid false phase detection
- Scanned PDF pages with < 50 characters of OCR text are skipped (likely diagrams/images)
- OCR text contains `--- PAGE N ---` markers — step extractor skips these, don't treat as section breaks
- OCR artifacts (`@`, `_`, `e`) before step numbers are stripped before matching
- Uppercase boilerplate lines (`JSW ONE TMT`, `JODL/PB TMT/...`) are skipped, not treated as section headings
- Scanner watermarks (`© Scanned with OKEN Scanner`) are filtered from step extraction
- Step extraction starts from "Work Flow:" / "Procedure:" marker to skip Teams & Responsibilities tables
- Bullet-prefixed table entries (e.g., `-Heavy visual bend`) are filtered: must be >40 chars or start with action verb

## System Requirements

| Component | Required For | Install Guide |
|-----------|-------------|---------------|
| Tesseract OCR | PDF parsing (scanned docs) | [UB-Mannheim releases](https://github.com/UB-Mannheim/tesseract/wiki) |
| Poppler | PDF-to-image conversion | [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases) |
| Python 3.10+ | All processing | Standard Python install |
| openpyxl | Excel generation | `pip install openpyxl` |
| python-docx | .docx parsing | `pip install python-docx` |
| pytesseract | OCR | `pip install pytesseract` |
| pdf2image | PDF rendering | `pip install pdf2image` |
| pdfplumber | Native PDF text | `pip install pdfplumber` |
