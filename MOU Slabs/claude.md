# Optimized Prompt — TMT Incentive Extraction Agent

> **Target AI:** Claude (with computer tools / file creation enabled)
> **Mode:** DETAIL — comprehensive extraction with structured Excel output
> **Techniques Applied:** Role assignment, constraint-based extraction, chain-of-thought decomposition, output specification

---

## Your Optimized Prompt

```
You are a Steel Industry Incentive Policy Analyst. Your task is to read the attached JSW Steel Annual Incentive Scheme document (FY 2025-26) and extract ALL incentive-related details specifically for the TMT product category, then present them in a well-structured Excel file.

---

## STEP 1: IDENTIFY TMT CATEGORIES

Read Table-1 in the document to identify every TMT-related product type. TMT falls under "Long Products" and may appear across multiple product groups and distribution channels (Retail, OEM). Capture:
- Product Family code (e.g., G)
- Product Group name
- Distribution channel (Retail / OEM)
- Any sub-categories or channel-specific variants (e.g., TMT Retail vs TMT OEMs)

---

## STEP 2: EXTRACT INCENTIVE DETAILS FOR EACH TYPE

For EACH TMT category identified, extract the following three incentive types. If a particular incentive does NOT apply to TMT, explicitly note "Not Applicable" with the reason.

### A. Volume Incentive
Extract from the main policy body + Annexure-1 (Table-3):
- Volume slabs (TOD ranges in MT) and corresponding incentive rates (₹/mt)
- Payout frequency (Half Yearly / Yearly)
- Qualification criteria (minimum achievement %, prorate rules)
- Half-yearly payout conditions (e.g., 50% volume compliance, 80% of volume incentive rate)
- Yearly payout conditions (100% achievement requirement)
- Monthly/quarterly minimum lifting requirements and exceptions
- Differential incentive rules (if any)
- Maximum payout cap rules (e.g., 120% cap on total quantity lifted)
- Any special conditions specific to TMT

### B. Consistency Incentive
Extract from Table-4 and related qualification section:
- Applicable product families for TMT (check Table-9 for confirmation)
- Incentive rate (₹/mt)
- Minimum pro-rata quarterly achievement % required
- Monthly minimum achievement % (lean period vs normal)
- Lean period definition and rules (max consecutive months, max per year)
- Payout frequency
- Enhancement rules (original qty vs enhanced qty treatment)
- Any TMT-specific conditions or exceptions

### C. Loyalty Incentive
Extract from Table-5 and related qualification section:
- Qualification criteria (FY 25-26 signing vs FY 24-25 comparison)
- Incentive rate (₹/mt)
- Volume achievement requirement
- Payout frequency
- Any TMT-specific conditions or exceptions

### D. Other Applicable Incentives
Check Table-9 (Summary of Incentives and Applicable Product Family) and confirm if any of these also apply to TMT:
- MSME Incentive (Table-6)
- Super Dealer Incentive (Table-7)
- Process Compliance Incentive (Table-8)
For each that applies, extract the full details. For each that doesn't apply, note "Not Applicable."

---

## STEP 3: EXTRACT GENERAL TERMS & CONDITIONS

Extract ALL general terms from Section (c) that impact TMT incentives:
- Signing deadline
- Quantity multiples rule
- Single vs multiple product family signing
- Combined buyer group treatment
- Payment timeline (30 days)
- Prime / NCO applicability and exclusions
- S1 category exclusion for Electrical product family (note if this impacts TMT)
- JSW Group supply rights
- Distribution channel rules (OEM vs Retail, permanence)
- Coated product group separation rules (if relevant to TMT)
- Sold-to-party basis for eligibility
- Maximum payout cap (120% of signed quantity)
- Draft provision rules (FY24-25 to FY25-26 transition)
- Enhancement rules (Q2, post Oct 15 window)
- Unethical practices clause
- Termination and withdrawal conditions
- Recovery and set-off provisions
- Interest rate on delayed payments (18% p.a.)
- Governing law

---

## STEP 4: CREATE EXCEL FILE

Create a professional, well-formatted .xlsx file with exactly 2 sheets:

### Sheet 1: "TMT Incentives — Consolidated"

Build ONE consolidated master table that captures ALL incentive types for ALL TMT sub-categories. The table must have the following columns:

| Column | Description |
|--------|-------------|
| TMT Category | e.g., TMT Retail, TMT OEM |
| Product Family Code | e.g., G |
| Distribution Channel | Retail / OEM |
| Incentive Type | Volume / Consistency / Loyalty / MSME / Super Dealer / Process Compliance |
| Incentive Rate (₹/mt) | The applicable rate |
| Volume Slab / Threshold | TOD range (for volume incentive) or qualifying threshold (for others). E.g., ">=1200 < 1800 MT" or ">=22.5% quarterly pro-rata" or ">=110% of FY24-25" |
| Payout Frequency | Half Yearly / Yearly / Annually |
| Qualification Criteria | Full qualification rule in one cell — e.g., "Min 20% quarterly prorate signed volume; exception of 1 quarter (first 3 Qs) where min 18% allowed; no exception in Q4" |
| Half-Yearly Payout Rule | Applicable rule for H1 payout (if any), else "N/A" |
| Yearly Payout Rule | Applicable rule for annual payout |
| Monthly Minimum Requirement | Monthly lifting minimum %, lean period rules (if any), else "N/A" |
| Maximum Payout Cap | Cap rule (e.g., "120% of signed quantity") |
| Special Conditions / Exceptions | Any TMT-specific conditions, exceptions, or notes. Use "[Verify]" tag if ambiguous. |
| Applicable? (Y/N) | Whether this specific incentive applies to this TMT category. If "N", state reason in Special Conditions column. |

**Row structure:**
- For Volume Incentive: Create ONE ROW PER SLAB per TMT category. E.g., if TMT Retail has 12 volume slabs, that's 12 rows — each with the same TMT Category, Incentive Type = "Volume", but different Volume Slab and Rate values. The Qualification Criteria, Payout Rules, and other common fields should be filled identically across all slab rows for that category (do NOT leave them blank after the first row — every row must be self-contained).
- For Consistency, Loyalty, MSME, Super Dealer, Process Compliance: Create ONE ROW per TMT category per incentive type.
- If an incentive type does NOT apply to a TMT category, still include ONE row with Applicable? = "N" and the reason.

This means the table will have many rows — that is expected and correct. Completeness over compactness.

### Sheet 2: "Terms & Conditions"
- Two-column layout: Clause # | Description
- All general T&Cs extracted in Step 3
- Each clause as its own row with the full text of the condition

---

## FORMATTING REQUIREMENTS
- Use Arial font throughout
- Header rows: Bold, dark blue background (#003366), white text
- Sub-headers: Bold, light blue background (#B8CCE4)
- Data cells: Left-aligned for text, center-aligned for rates/numbers
- Currency values formatted as ₹ #,##0
- Percentage values formatted as 0.0%
- Column widths auto-fitted to content
- Freeze top row (header) on each sheet
- Add light gray borders to all data cells
- Include a "Source: JSW Steel Annual Incentive Scheme FY 2025-26, dated 5th May 2025" footer note on each sheet

---

## IMPORTANT INSTRUCTIONS
- Read the ENTIRE document carefully before extracting. TMT information is spread across multiple sections, tables, and the annexure.
- Cross-reference Table-9 (Summary) with detailed sections to ensure nothing is missed.
- If any information is ambiguous or potentially applicable, include it with a note "[Verify]".
- Do NOT skip any slab, rate, condition, or exception — completeness is critical.
- Preserve exact figures and percentages from the document — do not round or approximate.
```

---

## Key Improvements

**Structured extraction flow** — The prompt follows a 4-step chain-of-thought (Identify → Extract → Terms → Output) so the agent doesn't miss scattered TMT references across different sections and annexures.

**Explicit table/section references** — Instead of vaguely asking "find TMT info," the prompt points to Table-1, Table-3 through Table-9, Annexure-1, and Section (c), ensuring the agent knows exactly where to look.

**Single consolidated master table** — All 6 incentive types for all TMT sub-categories live in one filterable table. One row per slab/threshold per category, so you can filter by Incentive Type, TMT Category, or Channel instantly. Every row is self-contained (no blank cells referencing rows above).

**2-tab clean layout** — Tab 1 = all incentive data in one place; Tab 2 = full T&Cs. No jumping between sheets to piece together the picture.

**Expanded scope** — Goes beyond Volume, Consistency, and Loyalty to also check MSME, Super Dealer, and Process Compliance via Table-9 cross-reference, since some may apply to TMT OEMs.

**Anti-hallucination guardrails** — "[Verify]" tagging for ambiguous items, "Not Applicable with reason" rows, and the instruction to read the entire document before extracting all prevent partial or incorrect extraction.

## Techniques Applied
Role assignment (Steel Industry Policy Analyst), chain-of-thought decomposition (4-step flow), constraint-based extraction (exact table references), output specification (6-sheet structure with formatting), and completeness verification (Table-9 cross-reference).

## Pro Tip
Attach the PDF directly when running this prompt. If the agent struggles with the annexure tables (which are image-heavy), consider also providing a text-extracted version of the annexure separately for higher accuracy on slab rates. You can also add `"If you cannot read a specific table clearly, flag it with [UNREADABLE — manual verification needed]"` to avoid silent failures.
