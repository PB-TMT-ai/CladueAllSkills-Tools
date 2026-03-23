# CLAUDE.md - JSW ONE Private Brands SOP Generator

## Build & Run
- `python tools/generate_sop_excel.py` - Run from project root to generate master Excel
- Output: `output/JSW_ONE_PB_SOPs_Master.xlsx` (sheet tab name: "Private Brands")
- OCR cache: `output/.ocr_cache/` (delete to force re-OCR of PDFs)

## Architecture (WAT Framework)
- Workflows: `workflows/` (7 focused .md files, index at sop_excel_generation.md)
- Tools: `tools/` (Python scripts — parsers, mapping, excel_writer)
- Agent role: orchestration only; never do execution that tools handle

## Key Gotchas
- .doc files are MIME/HTML (Confluence exports), NOT binary Word — parse with Python `email` module
- Phase detection: "Order Logging Process" heading contains "ORDER" — must match "ORDER PHASE" specifically
- PB SOP Links Confluence.xlsx has typo "nfluencer" (missing "I") — fuzzy matching handles it
- Scanned PDFs need Tesseract + Poppler binaries (paths in pdf_parser.py, configurable via env vars)
- `pdf2image` requires explicit `poppler_path=` on Windows; don't rely on PATH
- Two PB+Retailer_Influencer+in+App files: same size, different hashes (metadata diff, not duplicate)

## Description vs Steps Design (Critical)
Description (Col D) and Steps (Col H) MUST contain different content — never overlap:
- **OrderLogging (docx_sop)**: Description = contextual metadata (phase, owner, interface); Steps = procedure text. Source has no separate description column — description is synthesized from other columns.
- **Approval Workflows (workflow_doc)**: Description = use case text; Steps = approval hierarchy (initiator → approvers → rejection). Parsed from 6-column table (name, use case, initiator, approver, rejection, info).
- **Quality Manual (quality_manual)**: Description = Purpose + Scope from SOP body; Steps = numbered procedure. Extracted by `_extract_purpose_from_block()` and `_extract_steps_from_block()`.

## Field Extractor Patterns (`tools/mapping/field_extractor.py`)
- OCR step extraction: `_extract_steps_from_block()` handles inline `N. text`, multi-line `Step N\ntext`, and bullet points
- Procedure marker skip: extraction starts from "Work Flow:" / "Procedure:" to avoid Teams table content
- Table fragment filter: steps must be action sentences (verb-starting or >50 chars), not short noun phrases
- Bullet point filter: dash-prefixed lines need >40 chars or action verb to be treated as steps
- QAP Testing fallback: `_extract_qap_from_text()` uses keyword matching when OCR table parsing fails
- Purpose extraction: `_extract_purpose_from_block()` extracts Purpose+Scope for SOP descriptions

## Adding Documents
1. Drop file into Documents/ (.doc, .docx, .pdf)
2. Add DOCUMENT_MAPPING entry in `tools/mapping/document_classifier.py`
3. Add Confluence link to `Documents/PB SOP Links Confluence.xlsx`
4. Run generator — see `workflows/adding-documents.md` for details

## Bash Tips (Windows/Git Bash)
- Use `<< 'PYEOF'` heredoc for multi-line Python, not `python -c "..."` (escaping breaks `!=`)
- Tesseract silent install goes to `AppData\Local\Programs\`, not `Program Files\`
- Use GitHub API to find release URLs; don't guess tag names
- `pip` may not be on PATH; use `python -m pip` instead

---

## Pipes & Tubes SOP Generator
- `python tools/generate_pt_sop_excel.py` - Run from project root
- Output: `output/JSW_ONE_PT_SOPs_Master.xlsx` (sheet tab name: "Pipes & Tubes")
- Source docs: `Documents/SOP_Pipes & tubes/` (21 .docx files, all JSW ONE template format)
- Manifest: `output/pt_manifest.json`

### P&T Journey Structure
- Pre-Production: A. Raw Material Management, B. Production Planning
- Production: C. Manufacturing & Finishing, D. Quality Testing, E. Quality Management
- Post-Production: F. Dispatch & Logistics, G. Documentation & Invoicing

### P&T Key Gotchas
- All P&T docs are .docx (no .doc or .pdf) — parsed with python-docx only
- One_Helix doc has custom format (9 tables, 6 activity sections) — special parser
- Section headers use numbered format ("1.0 Purpose:", "6.0 Test Procedure:") — parser strips number prefixes
- "Record production quantity." must NOT trigger "Records" section break — regex requires standalone "Records" keyword
- Traceability_14[1].docx is a content duplicate — auto-skipped by MD5 dedup
- No Confluence/SOP links available yet — SOP Link column is blank
- Document classification is in `tools/generate_pt_sop_excel.py` (PT_DOCUMENT_MAPPING), NOT in `document_classifier.py`

### Adding P&T Documents
1. Drop .docx into `Documents/SOP_Pipes & tubes/`
2. Add `PT_DOCUMENT_MAPPING` entry in `tools/generate_pt_sop_excel.py`
3. Run: `python tools/generate_pt_sop_excel.py`
4. See `workflows/pt-adding-documents.md` for details
