# Workflow 05: Quality Validation

## Overview
Comprehensive QA checklist before SOP delivery.

**Time**: 3-5 minutes  
**Critical**: ALL checks must pass before delivery

---

## QA Checklist

### 1. Structure Validation

```
[ ] Hierarchical numbering correct (1., 1.a., 1.a.i.)
[ ] No gaps in sequence (no 1., 3. without 2.)
[ ] No duplicate numbers
[ ] All main activities have descriptive titles
[ ] Sub-activities properly nested under parents
[ ] Numbering follows pattern consistently
```

**AUTO-CHECK CODE**:
```python
def validate_numbering(activities):
    # Check for gaps, duplicates, proper hierarchy
    # Return list of issues
```

---

### 2. Table Structure

```
[ ] All rows have 5 columns
[ ] Column headers present: Activity | Steps | Team | Interface/Utilities | Sign off
[ ] All cells have content (except Sign off)
[ ] No merged cells (except intentional)
[ ] Table borders visible
[ ] Column widths appropriate (15%, 35%, 10%, 35%, 5%)
```

---

### 3. Content Completeness

```
[ ] Every activity has steps filled
[ ] Every activity has team assigned
[ ] Every system interaction has interface specified
[ ] All decision points have separate outcome rows
[ ] All conditional flows documented (If/Then)
[ ] Cross-references are accurate
[ ] No "TBD" or "Unknown" values
```

---

### 4. Team Assignment

```
[ ] All teams from approved list only:
    • Sales
    • Planning
    • Biz-ops
    • JOTS
    • Plant Operations
[ ] No typos in team names
[ ] Consistent naming across document
[ ] User confirmed all ambiguous assignments
[ ] Flowchart mappings applied where available
```

---

### 5. Image Quality

```
[ ] All source images extracted
[ ] All images embedded (not linked)
[ ] Image widths 400-450px
[ ] Images center-aligned in cells
[ ] Images contextually placed
[ ] No broken image references
[ ] No placeholder text like "[Image here]"
[ ] Images render when document opened
```

**IMAGE CHECK CODE**:
```python
def validate_images(docx_file):
    # Check all images embedded
    # Verify dimensions
    # Return list of issues
```

---

### 6. Conditional Logic

```
[ ] All decision points identified
[ ] Each outcome has separate row
[ ] Outcomes clearly labeled:
    • "If Approved" / "If Rejected"
    • "If Yes" / "If No"
    • "If Condition Met" / "If Condition Not Met"
[ ] Each outcome has explicit next steps
[ ] No orphaned conditions
[ ] Exception paths documented
```

---

### 7. Formatting

```
[ ] Font: Calibri 11pt (body), 12pt Bold (headers)
[ ] Line spacing: 1.15
[ ] Margins: 1 inch all sides
[ ] Table borders: 1pt weight
[ ] Header row: Bold + background color
[ ] No extra blank rows
[ ] Consistent spacing between sections
```

---

### 8. Version Control

```
[ ] Version number in filename ([ProjectName]_v[N].docx)
[ ] Version number in document header
[ ] No version conflicts with existing files
[ ] Saved to correct directory (/mnt/user-data/outputs/)
```

---

### 9. Cross-References

```
[ ] All "see Activity X" references valid
[ ] All "proceed to Section Y" references exist
[ ] No broken internal links
[ ] Consistent activity numbering in references
```

---

### 10. Special Cases

```
[ ] RRP breach approvals documented
[ ] Channel finance workflows included if applicable
[ ] Delivery instruction changes documented
[ ] Exception handling paths clear
[ ] Special notes highlighted appropriately
```

---

## Automated Checks

### Script 1: Numbering Validation
```python
def validate_hierarchical_numbering(activities):
    issues = []
    prev_main = 0
    
    for activity in activities:
        num = activity.number
        
        # Check main activities sequential
        if '.' not in num:
            if int(num) != prev_main + 1:
                issues.append(f"Gap: {prev_main} → {num}")
            prev_main = int(num)
    
    return issues
```

### Script 2: Image Verification
```python
from docx import Document

def check_images(docx_path):
    doc = Document(docx_path)
    issues = []
    
    for table in doc.tables:
        for row in table.rows:
            cell = row.cells[3]  # Interface/Utilities column
            if not has_image(cell) and requires_image(row):
                issues.append(f"Missing image: {row.cells[0].text}")
    
    return issues
```

### Script 3: Team Validation
```python
APPROVED_TEAMS = ['Sales', 'Planning', 'Biz-ops', 'JOTS', 'Plant Operations']

def validate_teams(activities):
    issues = []
    
    for activity in activities:
        if activity.team not in APPROVED_TEAMS:
            issues.append(f"Invalid team '{activity.team}' in {activity.number}")
    
    return issues
```

---

## Manual Review Areas

### Critical Sections to Review
1. **First 3 activities** - Set structure pattern
2. **All decision points** - Ensure outcomes complete
3. **Team handoff points** - Verify smooth transitions
4. **Image placements** - Contextual accuracy
5. **Last activity** - Proper conclusion

---

## Issue Resolution

### Common Issues & Fixes

**Issue 1: Numbering Gap**
```
SYMPTOM: Activities jump from 1.a to 1.c
FIX: Insert missing 1.b OR renumber 1.c → 1.b
```

**Issue 2: Missing Team**
```
SYMPTOM: Team column empty or "TBD"
FIX: Review activity → Apply inference → Ask user if needed
```

**Issue 3: Image Not Rendering**
```
SYMPTOM: Image shows placeholder icon
FIX: Re-embed image with correct format (PNG/JPG)
```

**Issue 4: Broken Cross-Reference**
```
SYMPTOM: "See Activity 5.3" but activity doesn't exist
FIX: Update reference OR add missing activity
```

---

## Pre-Delivery Report

Generate before presenting to user:

```
✅ Quality Validation Report

PASSED CHECKS:
✓ Structure (10/10)
✓ Content (8/8)
✓ Teams (5/5)
✓ Images (12/12)
✓ Formatting (7/7)

FLAGGED ITEMS:
⚠️ Activity 2.b: Image placement needs confirmation
⚠️ Activity 5.d: Team assignment user-confirmed as Biz-ops

STATISTICS:
• Total Activities: [X]
• Total Steps: [Y]
• Images Embedded: [Z]
• Decision Points: [N]
• Teams Involved: [List]

READY FOR DELIVERY: YES
```

---

## Delivery Checklist

Final verification before presenting:

```
[ ] All QA checks passed
[ ] All flagged items addressed
[ ] File saved to /mnt/user-data/outputs/
[ ] Filename has correct version number
[ ] Document opens without errors
[ ] User-facing delivery message prepared
[ ] Change log prepared (if update)
```

---

## Success Criteria

Document passes QA when:
1. ✅ All automated checks pass
2. ✅ All manual review areas checked
3. ✅ All issues resolved or flagged for user
4. ✅ Pre-delivery report generated
5. ✅ Delivery checklist complete
