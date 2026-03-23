# Workflow 01: New SOP Creation

## Overview
Create a complete SOP from scratch using source materials (DOCX, PPTX, flowchart).

**Estimated Time**: 25-35 minutes  
**Output**: `[ProjectName]_v1.docx` in `/mnt/user-data/outputs/`

---

## Phase 1: Intake & Discovery (2-3 min)

### Step 1.1: Receive Materials
```
ACTION: User uploads files
CONFIRM: List all files received

Template:
"I have received:
 • [File1.docx]
 • [File2.pptx]
 • [Flowchart.png]
 
Ready to proceed? (Yes/No)"
```

### Step 1.2: Ask Clarifying Questions
```
REQUIRED QUESTIONS:
1. Which phase does this cover?
   → Pre-order / Order / Post-order / All

2. Are there priority activities to focus on?
   → List if any

3. Any known team assignment exceptions?
   → Note exceptions

Template:
"Quick questions before analysis:
 1. Phase coverage: [Pre-order/Order/Post-order/All]?
 2. Priority activities: [Any specific focus]?
 3. Team exceptions: [Any deviations from flowchart]?
"
```

### Step 1.3: Confirm Master Flowchart
```
CHECK: Is master flowchart present?
  → YES: Note availability
  → NO: Flag high uncertainty for team assignments

IF NO FLOWCHART:
"⚠️ No master flowchart detected. Team assignments will have higher uncertainty. Proceed with text-only analysis? (Yes/No)"
```

---

## Phase 2: Analysis & Extraction (8-12 min)

### Step 2.1: Parse All Documents
```
FOR EACH document:
  1. Extract text content
  2. Identify existing structure (tables, bullets)
  3. Extract all images with metadata
  4. Note document type (SOP / Partial / Notes)

TOOLS:
  → view tool for document inspection
  → Document parser logic (see tools/document-parser.md)
```

### Step 2.2: Create Content Inventory
```
OUTPUT CHECKLIST:
[ ] Activity count: [X] identified
[ ] Image count: [Y] extracted
[ ] Decision points: [Z] detected
[ ] Team indicators: [Present/Absent]

Template:
"📊 Content Extraction Complete

Activities: [X]
Images: [Y]
Decision Points: [Z]
Team Indicators: [Present/Absent per activity]
"
```

### Step 2.3: Map Teams Using Flowchart
```
FOR EACH activity:
  1. Locate in flowchart by text match
  2. Identify team by color:
     • Blue = Sales
     • Green = Planning
     • Pink = Biz-ops
     • Yellow = JOTS
     • Purple = Plant Operations
  3. If unclear → Add to clarification list

OUTPUT: Team mapping table
```

### Step 2.4: Identify Ambiguous Team Assignments
```
CRITERIA FOR FLAGGING:
• No flowchart box match
• Multiple candidate teams
• Conflicting indicators

Template:
"🔍 Team Clarification Needed

Activities requiring confirmation:
1. [Activity X]: Candidates [Team A / Team B]
   Context: [Brief description]
   
2. [Activity Y]: No clear indicator
   Context: [Brief description]

Please specify teams for each."
```

### Step 2.5: Present Image Inventory
```
LIST ALL EXTRACTED IMAGES:
• Image ID
• Source document
• Proposed placement
• Context

Template:
"📸 Images Extracted: [Y]

1. [Image1.png] - Source: [Doc1], Page 3
   Proposed: Activity 1.a (Shows SFDC interface)
   
2. [Image2.png] - Source: [Doc2], Slide 5
   Proposed: Activity 2.b (Flowchart diagram)

Confirm placements? (Yes/Modify)"
```

---

## Phase 3: Structure Construction (10-15 min)

### Step 3.1: Build Activity Hierarchy
```
STRUCTURE RULES:
• Main activities: 1., 2., 3.
• Sub-activities: 1.a., 1.b., 2.a.
• Detail steps: 1.a.i., 1.a.ii.

PROCESS:
1. Group related activities
2. Assign hierarchical numbers
3. Verify no gaps in sequence
```

### Step 3.2: Map to Business Phases
```
PHASE CLASSIFICATION:
• PRE-ORDER: Opportunity creation, approvals
• ORDER: Planning, Biz-ops execution, JOTS coordination
• POST-ORDER: Plant dispatch, GRN, invoicing

ACTION: Tag each activity with phase
```

### Step 3.3: Assign Teams (with User Confirmation)
```
PROCESS:
1. Apply flowchart mappings
2. Infer from keywords/systems
3. Request user confirmation on flagged items
4. Update activity table with teams

VERIFICATION:
"✅ Team Assignment Complete

Confirmed Assignments:
• Activities 1-5: Sales
• Activities 6-10: Planning
• Activities 11-15: Biz-ops

Flagged for Confirmation:
• Activity 16: [Options presented]

Proceed? (Yes/Revise)"
```

### Step 3.4: Place Images in Table Cells
```
PLACEMENT LOGIC:
• UI screenshot → Step describing that UI
• Flowchart → First activity in section
• Form template → Form completion step

ACTION:
1. Resize to 400-450px width
2. Embed in "Interface / Utilities" column
3. Center-align in cell
```

### Step 3.5: Create Conditional Flow Branches
```
IDENTIFY DECISION POINTS:
• Approval/Rejection scenarios
• Yes/No decisions
• Multiple outcome options

CREATE SEPARATE ROWS:
| 1.f. Submit for Approval | [Steps] | Sales | SFDC | |
| 1.f.i. If Approved | [Outcome steps] | Sales | SFDC | |
| 1.f.ii. If Rejected | [Rejection steps] | Sales | SFDC | |

VERIFY: Each outcome has explicit next steps
```

---

## Phase 4: DOCX Generation (5 min)

### Step 4.1: Initialize DOCX Document
```
ACTIONS:
1. Create blank DOCX
2. Add title: "JSW ONE TMT - STANDARD OPERATING PROCEDURE"
3. Add subtitle: [Process Name]
4. Add version: "Version [N]"

FILE PATH: /home/claude/[ProjectName]_v1.docx
```

### Step 4.2: Build 5-Column Table
```
COLUMN STRUCTURE:
| Activity | Steps | Team | Interface/Utilities | Sign off |
|----------|-------|------|---------------------|----------|

FORMATTING:
• Header row: Bold, background color
• Border: All cells
• Width ratios: 15% | 35% | 10% | 35% | 5%
```

### Step 4.3: Populate Table Rows
```
FOR EACH activity:
  1. Insert row
  2. Fill Activity (hierarchical number + title)
  3. Fill Steps (detailed instructions)
  4. Fill Team (responsible team name)
  5. Fill Interface (system + embedded image)
  6. Leave Sign off empty

IMAGE EMBEDDING:
• Use python-docx library
• Resize: 400-450px width
• Center-align in cell
```

### Step 4.4: Apply Formatting
```
APPLY:
• Font: Calibri 11pt for body
• Font: Calibri 12pt Bold for headers
• Line spacing: 1.15
• Margins: 1 inch all sides
• Table: Border weight 1pt
```

### Step 4.5: Save to Output Directory
```
ACTIONS:
1. Save file
2. Copy to /mnt/user-data/outputs/
3. Verify file accessible

FILENAME: [ProjectName]_v1.docx
```

---

## Phase 5: Quality Validation (3-5 min)

### Step 5.1: Run QA Checklist
```
STRUCTURE VALIDATION:
[ ] All activities have hierarchical numbering
[ ] No gaps in sequence (1 → 2 → 3, not 1 → 3)
[ ] Each row has Activity, Steps, Team filled
[ ] Sign off column empty

CONTENT COMPLETENESS:
[ ] Every decision point has separate outcome rows
[ ] All source images extracted and placed
[ ] Interface/Utilities filled for system interactions
[ ] Cross-references verified

TEAM ASSIGNMENT:
[ ] No "TBD" or "Unknown" values
[ ] Teams match approved list exactly
[ ] Flowchart mappings applied

IMAGE QUALITY:
[ ] All images render correctly
[ ] Widths 400-450px
[ ] Contextually placed
[ ] No broken references

CONDITIONAL LOGIC:
[ ] All If/Then scenarios documented
[ ] No orphaned conditions
[ ] Exception paths included
```

### Step 5.2: Present Deliverable
```
Template:
"✅ SOP Generation Complete

Summary:
• Phase: [Pre-order/Order/Post-order]
• Activities: [X]
• Steps: [Y]
• Images: [Z] embedded
• Conditional flows: [N]

Files Ready:
📄 [ProjectName]_v1.docx

Verification Checklist:
✓ All teams assigned
✓ All images render correctly
✓ Conditional logic complete
✓ Hierarchical numbering correct

Next Steps:
→ Review document
→ Confirm any flagged items
→ Request modifications if needed
"
```

### Step 5.3: Offer Post-Delivery Support
```
Template:
"Document ready for review.

Need modifications?
• Update specific sections
• Add missing activities
• Adjust team assignments
• Re-place images

Or ready to process next phase/section?
"
```

---

## Decision Points

### DP1: No Flowchart Available
```
IF flowchart missing:
  → Proceed with text-only analysis
  → Flag ALL team assignments for user review
  → Request verification before final output
  
NOTIFY USER:
"⚠️ Processing without flowchart. All team assignments will require your confirmation."
```

### DP2: Conflicting Team Indicators
```
IF multiple team candidates:
  → Present conflict evidence
  → Request user decision
  → Do NOT guess
  
Template:
"🔍 Conflict Detected

Activity X shows:
• Flowchart: Blue box (Sales)
• Text: 'Planning team analyzes...'

Please clarify: Sales or Planning?"
```

### DP3: Image Placement Unclear
```
IF image context ambiguous:
  → Default to earliest relevant activity
  → Flag for user verification
  → Add note in document
  
Template:
"📸 Image Placement Verification

Image: [Filename]
Context: [Description]
Proposed: Activity [X]

Confirm or specify alternative?"
```

---

## Common Issues & Solutions

### Issue 1: Missing Images
```
SYMPTOM: Source document has images but extraction fails
SOLUTION:
  1. Check file format (DOCX/PPTX/PDF)
  2. Try alternate extraction method
  3. Note failed extractions
  4. Request user to provide images separately
```

### Issue 2: Ambiguous Activity Numbering
```
SYMPTOM: Source has non-standard numbering
SOLUTION:
  1. Renumber hierarchically from scratch
  2. Map old → new numbering
  3. Verify no activities lost
  4. Note renumbering in delivery message
```

### Issue 3: Conditional Flows Not Explicit
```
SYMPTOM: Decision points implied but not documented
SOLUTION:
  1. Identify implied decisions
  2. Request explicit paths from user
  3. Add separate rows for each outcome
  4. Document assumptions made
```

---

## Success Criteria

Document is complete when:
1. ✅ All QA checkpoints pass
2. ✅ User confirms flagged items
3. ✅ File saved to `/mnt/user-data/outputs/`
4. ✅ Version number in filename
5. ✅ No "TBD" or "Unknown" values

---

## Next Steps After Delivery

→ Await user feedback  
→ Prepare for modifications if needed  
→ Ready to process next SOP section  
→ Archive working files in `/home/claude/`
