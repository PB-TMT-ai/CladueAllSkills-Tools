# Validation Checklist

[Back to Overview](sop_excel_generation.md)

After running the generator, open `output/JSW_ONE_PB_SOPs_Master.xlsx` and verify:

- [ ] Sheet tab name is "Private Brands"
- [ ] All 3 journey phases present (Pre-Order, Order, Post Order)
- [ ] All source documents represented (check console output for skipped files)
- [ ] No duplicate files processed
- [ ] Confluence links are clickable hyperlinks in SOP Link column
- [ ] Section headers appear with green fill
- [ ] Merged cells render correctly
- [ ] Activity numbering is sequential
- [ ] Description (Col D) ≠ Steps (Col H) — no overlap or substring match
- [ ] OrderLogging descriptions show phase/owner/interface metadata, NOT step text
- [ ] Approval Workflow steps show initiator → approver → rejection hierarchy
- [ ] Quality Manual descriptions come from Purpose+Scope, steps from Procedure section
- [ ] Quality Manual activities appear under "J. QUALITY ASSURANCE" sections
- [ ] OCR-extracted steps contain only action items (no table cell fragments or watermarks)
- [ ] OCR-extracted content is readable and accurate (may need manual review)

## Console Output Checks

- Each document should show `Type:` and either `Confluence link matched:` or `No Confluence link matched`
- Final summary should show total journeys (3) and total activities count
- No `ERROR` messages in the output
