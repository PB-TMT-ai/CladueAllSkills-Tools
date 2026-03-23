# Document Types Reference

[Back to Overview](sop_excel_generation.md)

## Supported Types

| Type | Structure | Extraction | When to Use |
|------|-----------|------------|-------------|
| `ops_sop` | Purpose, Stakeholders, Steps, Sign-off | One activity with all steps | Standard operating procedure documents with numbered steps |
| `workflow_doc` | Approval tables (6 cols: name, use case, initiator, approver, rejection, info) | Each table row → activity; steps built from approval hierarchy (initiator → approvers → rejection) | Documents with approval/workflow tables (set `expand_table_rows: True`) |
| `technical_doc` | Headings + paragraphs | Sections become activities | Technical documentation with multiple sections |
| `field_spec` | Field specification tables | One activity with fields as steps | Data enrichment / field definition documents |
| `demo_doc` | Screenshots + captions | Single minimal entry with link | Screenshot-heavy or demo walkthrough documents |
| `docx_sop` | Tables with Activity/Steps/Team columns | Activities grouped by phase; description = contextual metadata (phase, owner, interface) — NOT from step text | Structured .docx with table-based activity data (auto-detects phases) |
| `quality_manual` | Quality policy, QAP tables, SOPs | Multiple activities split across journeys; description from Purpose+Scope; steps from Procedure; QAP fallback via keyword matching | Scanned quality manuals (OCR-parsed PDFs) |

## Description vs Steps Column Rules

Each document type generates Description (Col D) and Steps (Col H) from **different sources** to avoid overlap:

| Type | Description Source | Steps Source |
|------|-------------------|--------------|
| `ops_sop` | Purpose field from document | Numbered process steps |
| `workflow_doc` | Use case column | Approval hierarchy (initiator → approvers → rejection) |
| `docx_sop` | Phase + Owner + Interface metadata | Steps column from source table |
| `quality_manual` | Purpose + Scope text from SOP body | Numbered procedure from Work Flow/Procedure section |
| `field_spec` | Purpose or document title | Field specification entries |
| `technical_doc` | Section heading summary | Sub-headings as steps |
| `demo_doc` | Purpose or first paragraph | Empty (demo/screenshot docs) |

## File Format Support

| Format | Parser | Notes |
|--------|--------|-------|
| `.doc` | `confluence_doc_parser.py` | Confluence MIME/HTML exports (not binary Word) |
| `.docx` | `docx_parser.py` | Standard Word documents via python-docx |
| `.pdf` | `pdf_parser.py` | OCR via pytesseract + pdf2image; results cached |
