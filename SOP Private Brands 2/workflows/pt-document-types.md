# P&T Document Types Reference

[Back to Overview](sop_excel_generation.md)

## JSW ONE Standard Template (20 of 21 documents)

All standard P&T SOPs follow the same template structure:

### Table Structure
| Table | Content | Columns |
|-------|---------|---------|
| Table 0 | Header metadata | "A JSW ONE PRODUCT" / "STANDARD OPERATING PROCEDURE" / Document ID, Version, Date |
| Table 1 | Teams & Responsibilities | Sl No, Team, Responsibility |
| Table 2+ | Process-specific content (optional) | Varies by SOP (acceptance criteria, NC classification, etc.) |
| Last table | Approval/Sign-off | Created by / Approved by |

### Paragraph Structure
Sections are detected by keywords, with or without numbered prefixes:

| Section | Keywords | Extracted As |
|---------|----------|-------------|
| Purpose | `Purpose:`, `1.0 Purpose:` | `description` field |
| Scope | `Scope:`, `2.0 Scope:` | Not used in output (context only) |
| Teams | `Teams and Responsibilities:` | Parsed from Table 1 instead |
| Process/Procedure | `Process:`, `6.0 Test Procedure:` | `steps` field (each paragraph = one step) |
| Inspection/Acceptance | `Inspection and Acceptance Criteria:` | Appended to `steps` |
| Re-test | `Re-testing Procedure:` | Appended to `steps` |
| Records | `Records` (standalone only) | Stops step collection |

### Process-Specific Tables
Some SOPs include additional tables between Teams and Approval:

| SOP Type | Table Content | Extracted As |
|----------|---------------|-------------|
| Dimension Test | Acceptance criteria (characteristic, criterion, standard, remark) | `remarks` + `remarks_details` |
| NC Products | Nonconformity classification (type, criterion, disposition) | `remarks` + `remarks_details` |
| Other test SOPs | No extra tables | Steps from paragraphs only |

## One_Helix Comprehensive Document (1 of 21)

`One_Helix_Pipes_Tubes_SOP_REVISED_FINAL 1.2.docx` has a unique structure:

### Table Structure (9 tables)
| Table | Content | Extracted As |
|-------|---------|-------------|
| Table 0 | Responsibilities (Maker/Checker) | Skipped |
| Tables 1-6 | Activity tables (Activity / Steps / Team / Interface) | 6 separate activities |
| Table 7 | Yield Metrics & Formulas | 1 activity (steps = metric formulas) |
| Table 8 | Yield Status & Actions | 1 activity (steps = status thresholds) |

### Activity-to-Journey Mapping
| Table | Activity | Journey | Section |
|-------|----------|---------|---------|
| 1 | HR Coil Purchase & Receipt | Pre-Production | A. RAW MATERIAL MANAGEMENT |
| 2 | Data Entry & Recording | Pre-Production | A. RAW MATERIAL MANAGEMENT |
| 3 | Slitting Process | Production | C. MANUFACTURING & FINISHING |
| 4 | FG Conversion (Pipe Making) | Production | C. MANUFACTURING & FINISHING |
| 5 | Sales Order & Dispatch | Post-Production | F. DISPATCH & LOGISTICS |
| 6 | Yield & Loss Tracking | Production | E. QUALITY MANAGEMENT |
| 7 | Yield Metrics & Formulas | Production | E. QUALITY MANAGEMENT |
| 8 | Yield Status & Corrective Actions | Production | E. QUALITY MANAGEMENT |

## File Format

All P&T documents are `.docx` format (Microsoft Word via python-docx). No `.doc` (Confluence MIME/HTML) or `.pdf` (scanned/OCR) files are present in the P&T collection.
