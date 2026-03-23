# SOP Excel Generation Workflow

## Objective
Generate and maintain the master SOP Excel file (`JSW_ONE_PB_SOPs_Master.xlsx`) from source documents in the `Documents/` folder, formatted to match the `Reference files/Construct.xlsx` template.

## When to Run
- **New document added** to `Documents/`
- **Existing document updated** (content changes)
- **Confluence links updated** in `PB SOP Links Confluence.xlsx`
- **Periodic review** (monthly or as needed)

## Required Inputs
1. Source documents in `Documents/` (.doc, .docx, .pdf files)
2. `Documents/PB SOP Links Confluence.xlsx` (Confluence URL mappings)
3. `Reference files/Construct.xlsx` (format reference - read-only)

## Quick Start
```bash
cd "D:\SOP Private Brands 2"
python tools/generate_sop_excel.py
```
See [How to Run](how-to-run.md) for full CLI options.

## Guides (Private Brands)
- [How to Run](how-to-run.md) - CLI commands, arguments, and output files
- [Adding New Documents](adding-documents.md) - Step-by-step for onboarding new SOPs
- [Adding Confluence Links](adding-confluence-links.md) - URL mapping updates
- [Document Types Reference](document-types.md) - Supported document types and extraction rules
- [Validation Checklist](validation-checklist.md) - Post-generation verification steps
- [Troubleshooting & Edge Cases](troubleshooting.md) - Common issues and fixes

---

## Pipes & Tubes SOP Generation

Generate and maintain the P&T master SOP Excel file (`JSW_ONE_PT_SOPs_Master.xlsx`) from `.docx` documents in `Documents/SOP_Pipes & tubes/`.

### Quick Start
```bash
cd "D:\SOP Private Brands 2"
python tools/generate_pt_sop_excel.py
```

### Guides (Pipes & Tubes)
- [P&T How to Run](pt-how-to-run.md) - CLI command and output files
- [P&T Adding Documents](pt-adding-documents.md) - Onboarding new P&T SOPs
- [P&T Document Types](pt-document-types.md) - JSW ONE template structure and parser behavior
- [P&T Validation Checklist](pt-validation-checklist.md) - Post-generation verification
- [P&T Troubleshooting](pt-troubleshooting.md) - Common issues and edge cases
