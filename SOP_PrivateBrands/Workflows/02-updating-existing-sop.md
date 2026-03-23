# Workflow 02: Updating Existing SOP

## Overview
Modify existing SOP documents while preserving structure and version control.

**Estimated Time**: 15-20 minutes  
**Output**: `[ProjectName]_v[N+1].docx`

---

## Step 1: Receive Update Request (1 min)

### 1.1 Identify Source Version
```
REQUIRED INFO:
• Current version number: v[N]
• Source file location
• Requested changes

CONFIRM:
"Updating [ProjectName]_v[N].docx
Changes requested: [List]
Proceed? (Yes/No)"
```

### 1.2 Load Existing SOP
```
ACTION:
1. Open current version
2. Extract existing structure
3. Create inventory of current activities
4. Note existing images

FILE: /mnt/project/[ProjectName]_v[N].docx
```

---

## Step 2: Analyze Change Request (3-5 min)

### 2.1 Classify Change Type
```
CHANGE TYPES:
A. Add new activities
B. Modify existing steps
C. Update team assignments
D. Add/replace images
E. Reorganize structure
F. Add conditional flows

ACTION: Identify type(s) from request
```

### 2.2 Impact Assessment
```
CHECK:
• Affects numbering? (Yes/No)
• Requires new images? (Yes/No)
• Changes team assignments? (Yes/No)
• Adds conditional flows? (Yes/No)

IF numbering affected:
  → Plan full renumbering
IF new images needed:
  → Request image sources
```

---

## Step 3: Implement Changes (10-12 min)

### 3.1 Make Modifications
```
FOR EACH change:
  1. Locate target section
  2. Apply modification
  3. Renumber if necessary
  4. Update cross-references
  5. Verify consistency

PRESERVE:
• Existing formatting
• Image quality
• Table structure
```

### 3.2 Add New Content
```
IF adding activities:
  1. Insert rows at correct position
  2. Assign hierarchical numbers
  3. Fill all 5 columns
  4. Place images if applicable
  5. Update sequence numbers below
```

### 3.3 Update Teams
```
IF team changes:
  1. Identify affected activities
  2. Update Team column
  3. Verify consistency across document
  4. Note changes in delivery message
```

### 3.4 Add/Replace Images
```
IF image changes:
  1. Extract new images from source
  2. Remove old images (if replacing)
  3. Embed at 400-450px width
  4. Center-align in cell
  5. Verify rendering
```

---

## Step 4: Increment Version (2 min)

### 4.1 Save New Version
```
ACTIONS:
1. Save as [ProjectName]_v[N+1].docx
2. Update version number in document header
3. Add modification date
4. Move to /mnt/user-data/outputs/

VERIFY: New version number in filename
```

### 4.2 Document Changes
```
CREATE: Change summary

Template:
"📝 Changes in v[N+1]:
• Added: [List new activities]
• Modified: [List changed steps]
• Updated: [List team changes]
• Images: [List image updates]
"
```

---

## Step 5: Quality Validation (2 min)

### 5.1 Run Update-Specific Checks
```
VERIFY:
[ ] Version number incremented correctly
[ ] All modifications applied
[ ] Numbering sequence intact
[ ] New images embedded correctly
[ ] Cross-references updated
[ ] No duplicate rows
[ ] Formatting preserved
```

### 5.2 Deliver Updated SOP
```
Template:
"✅ SOP Update Complete

Version: v[N] → v[N+1]

Changes Applied:
• [Summary of modifications]

Files:
📄 [ProjectName]_v[N+1].docx

Verification:
✓ All changes applied
✓ Numbering updated
✓ Images embedded
✓ Formatting preserved

Review changes and confirm."
```

---

## Decision Points

### DP1: Renumbering Required
```
IF new activities inserted:
  → Renumber ALL subsequent activities
  → Update cross-references
  → Verify no gaps

NOTIFY USER:
"Activities renumbered due to insertion at [Position]. All subsequent numbers updated."
```

### DP2: Conflicting Modifications
```
IF modification conflicts with existing:
  → Present conflict
  → Request clarification
  → Do NOT make assumptions

Template:
"⚠️ Modification Conflict

Requested: [Change A]
Existing: [Content B]

How to resolve?
1. Replace B with A
2. Merge A and B
3. Keep B, skip A
"
```

---

## Common Update Scenarios

### Scenario A: Add Special Case Section
```
STEPS:
1. Identify insertion point
2. Create new subsection with heading
3. Add activities with sub-numbering (8.a., 8.a.i., etc.)
4. Insert conditional flow rows
5. Place relevant images
6. Update TOC if present
```

### Scenario B: Split Large Activity
```
STEPS:
1. Identify activity to split
2. Create sub-activities (1.a., 1.b.)
3. Distribute content across sub-activities
4. Maintain original numbering for others
5. Add detail steps if needed (1.a.i., 1.a.ii.)
```

### Scenario C: Merge Multiple Activities
```
STEPS:
1. Identify activities to merge
2. Combine steps in logical order
3. Remove redundancy
4. Assign single activity number
5. Renumber subsequent activities
6. Update cross-references
```

---

## Success Criteria

Update is complete when:
1. ✅ All requested changes applied
2. ✅ Version incremented correctly
3. ✅ QA checks pass
4. ✅ File saved to outputs directory
5. ✅ Change summary documented
