# Workflow 03: Image Extraction & Placement

## Overview
Extract images from source documents and embed them correctly in SOP tables.

**Time**: 5-8 minutes per document  
**Critical**: ALL images must be extracted

---

## Step 1: Identify Image Sources (1 min)

### 1.1 List All Documents
```
CHECK:
• DOCX files: Inline images, screenshots
• PPTX files: Slide images, diagrams
• PDF files: Embedded images, flowcharts
• PNG/JPG files: Standalone images

ACTION: Create source inventory
```

---

## Step 2: Extract Images (3-5 min)

### 2.1 Extract from DOCX
```
TOOLS: python-docx library

CODE PATTERN:
from docx import Document
from docx.oxml import parse_xml

doc = Document('file.docx')
for rel in doc.part.rels.values():
    if "image" in rel.target_ref:
        # Extract image binary
        # Save to /home/claude/images/
```

### 2.2 Extract from PPTX
```
TOOLS: python-pptx library

CODE PATTERN:
from pptx import Presentation

prs = Presentation('file.pptx')
for slide in prs.slides:
    for shape in slide.shapes:
        if hasattr(shape, "image"):
            # Extract image
```

### 2.3 Extract from PDF
```
TOOLS: PyMuPDF (fitz) library

CODE PATTERN:
import fitz

pdf = fitz.open('file.pdf')
for page in pdf:
    for img in page.get_images():
        # Extract image
```

---

## Step 3: Catalog Extracted Images (1 min)

### 3.1 Create Image Inventory
```
FORMAT:
| ID | Source | Page/Slide | Context | Proposed Placement |
|----|--------|------------|---------|-------------------|

ACTIONS:
1. Assign unique ID to each image
2. Note source document + location
3. Extract nearby text for context
4. Propose activity placement
```

### 3.2 Present to User
```
Template:
"📸 Images Extracted: [N]

1. img_001.png
   Source: OrderLogging_v3.docx, Page 2
   Context: Shows SFDC interface
   Proposed: Activity 1.a (Navigate to Portal)

2. img_002.png
   Source: Flowchart.pptx, Slide 1
   Context: Master process flow
   Proposed: Section header

Confirm placements? (Yes/Modify)"
```

---

## Step 4: Determine Placement Logic (2-3 min)

### 4.1 Apply Placement Rules
```
RULE 1: UI Screenshot
→ Place in step describing that UI interaction

RULE 2: Flowchart/Diagram
→ Place at first activity in that section

RULE 3: Form Template
→ Place in form completion step

RULE 4: Before/After Comparison
→ Stack vertically in same cell

RULE 5: Ambiguous Context
→ Ask user for placement
```

### 4.2 Associate with Activities
```
FOR EACH image:
  1. Match context to activity description
  2. Find nearest text match in SOP
  3. Assign to specific activity number
  4. Note if multiple images per activity

OUTPUT: Image-to-Activity mapping
```

---

## Step 5: Embed in DOCX (2-3 min)

### 5.1 Prepare Images
```
ACTIONS:
1. Resize to 400-450px width
2. Maintain aspect ratio
3. Convert to format compatible with DOCX
4. Optimize file size if >500KB

TOOLS: PIL/Pillow library
```

### 5.2 Insert in Table Cells
```
CODE PATTERN:
from docx.shared import Inches

cell = table.cell(row_idx, 3)  # Column 4: Interface/Utilities
paragraph = cell.paragraphs[0]
run = paragraph.add_run()
run.add_picture('image.png', width=Inches(4.0))

# Center-align
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

### 5.3 Handle Multiple Images
```
IF multiple images for same activity:
  1. Insert first image
  2. Add line break
  3. Insert second image
  4. Repeat as needed
  
ALTERNATIVE: Create sub-rows if >3 images
```

---

## Step 6: Quality Checks (1 min)

### 6.1 Verify All Images
```
CHECKLIST:
[ ] All source images extracted
[ ] No broken image references
[ ] All images render in document
[ ] Widths between 400-450px
[ ] Images center-aligned
[ ] Contextually placed
[ ] No placeholder text remaining
```

### 6.2 Document Failed Extractions
```
IF extraction fails:
  → Note in cell: "[Image extraction failed: filename]"
  → Add to issues list
  → Request from user

Template:
"⚠️ Image Extraction Issues

Failed:
• [filename1]: Corrupted format
• [filename2]: Access denied

Please provide these images separately."
```

---

## Common Issues & Solutions

### Issue 1: Image Quality Loss
```
SYMPTOM: Extracted image is blurry
SOLUTION:
  1. Extract at higher resolution
  2. Check source image quality
  3. Use lossless format (PNG)
  4. Avoid resizing up
```

### Issue 2: Wrong Image Association
```
SYMPTOM: Image doesn't match activity
SOLUTION:
  1. Re-read context from source
  2. Check surrounding text
  3. Ask user for clarification
  4. Move to correct activity
```

### Issue 3: PDF Extraction Fails
```
SYMPTOM: PDF images not extracting
SOLUTION:
  1. Check if PDF is image-based (scanned)
  2. Try OCR if needed
  3. Request original source
  4. Manual screenshot as fallback
```

---

## Decision Points

### DP1: Unclear Image Context
```
IF cannot determine placement:
  → Default to earliest relevant activity
  → Flag for user verification
  
ASK USER:
"Image [ID] context unclear. Best placement:
A. Activity [X] - [Reason]
B. Activity [Y] - [Reason]
Please specify."
```

### DP2: Image Too Large
```
IF image >1MB:
  → Compress without quality loss
  → Warn user if compression insufficient
  
IF image >5MB:
  → Request smaller version
  → Cannot embed efficiently
```

---

## Success Criteria

Image extraction complete when:
1. ✅ All images extracted from sources
2. ✅ All images cataloged with context
3. ✅ All placements confirmed
4. ✅ All images embedded in DOCX
5. ✅ All images render correctly
6. ✅ Quality checks pass
