# Tool: Document Parser

## Overview
Parse DOCX, PPTX, and PDF files to extract text content, structure, and images.

---

## DOCX Parsing

```python
from docx import Document

def parse_docx(file_path):
    """Extract content from DOCX file."""
    doc = Document(file_path)
    
    content = {
        'text': [],
        'tables': [],
        'images': [],
        'structure': []
    }
    
    # Extract paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            content['text'].append({
                'text': para.text,
                'style': para.style.name,
                'level': get_heading_level(para)
            })
    
    # Extract tables
    for table_idx, table in enumerate(doc.tables):
        table_data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        content['tables'].append({
            'index': table_idx,
            'data': table_data,
            'rows': len(table.rows),
            'cols': len(table.columns)
        })
    
    # Extract images (see image-embedding.md for details)
    content['images'] = extract_docx_images(doc, file_path)
    
    return content

def get_heading_level(paragraph):
    """Determine if paragraph is a heading and its level."""
    style = paragraph.style.name.lower()
    if 'heading' in style:
        try:
            return int(style.split('heading')[1].strip())
        except:
            return 0
    return 0
```

---

## PPTX Parsing

```python
from pptx import Presentation

def parse_pptx(file_path):
    """Extract content from PowerPoint file."""
    prs = Presentation(file_path)
    
    content = {
        'slides': [],
        'images': []
    }
    
    for slide_idx, slide in enumerate(prs.slides):
        slide_data = {
            'index': slide_idx,
            'title': '',
            'text': [],
            'images': []
        }
        
        for shape in slide.shapes:
            # Extract text
            if hasattr(shape, "text"):
                if shape.is_placeholder and shape.placeholder_format.type == 1:
                    slide_data['title'] = shape.text
                else:
                    slide_data['text'].append(shape.text)
            
            # Extract images
            if hasattr(shape, "image"):
                image_data = extract_pptx_image(shape, slide_idx, file_path)
                slide_data['images'].append(image_data)
                content['images'].append(image_data)
        
        content['slides'].append(slide_data)
    
    return content
```

---

## PDF Parsing

```python
import fitz  # PyMuPDF

def parse_pdf(file_path):
    """Extract content from PDF file."""
    doc = fitz.open(file_path)
    
    content = {
        'pages': [],
        'images': []
    }
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        page_data = {
            'number': page_num + 1,
            'text': page.get_text(),
            'images': []
        }
        
        # Extract images
        image_list = page.get_images()
        for img_idx, img in enumerate(image_list):
            image_data = extract_pdf_image(doc, img, page_num, img_idx)
            page_data['images'].append(image_data)
            content['images'].append(image_data)
        
        content['pages'].append(page_data)
    
    return content
```

---

## Activity Detection

```python
import re

def detect_activities(text_content):
    """Identify activity patterns in text."""
    
    activities = []
    
    # Pattern 1: Numbered lists (1., 2., 3.)
    pattern1 = r'^\d+\.\s+(.*?)$'
    
    # Pattern 2: Lettered sub-items (a., b., c.)
    pattern2 = r'^[a-z]\.\s+(.*?)$'
    
    # Pattern 3: Roman numerals (i., ii., iii.)
    pattern3 = r'^(i{1,3}|iv|v|vi{1,3}|ix|x)\.\s+(.*?)$'
    
    for para in text_content:
        text = para['text']
        
        match1 = re.match(pattern1, text, re.MULTILINE)
        match2 = re.match(pattern2, text, re.MULTILINE)
        match3 = re.match(pattern3, text, re.MULTILINE)
        
        if match1:
            activities.append({
                'level': 1,
                'number': text.split('.')[0] + '.',
                'description': match1.group(1),
                'type': 'main'
            })
        elif match2:
            activities.append({
                'level': 2,
                'number': text.split('.')[0] + '.',
                'description': match2.group(1),
                'type': 'sub'
            })
        elif match3:
            activities.append({
                'level': 3,
                'number': text.split('.')[0] + '.',
                'description': match3.group(1),
                'type': 'detail'
            })
    
    return activities
```

---

## Conditional Flow Detection

```python
def detect_conditional_flows(activities):
    """Identify decision points and conditional branches."""
    
    conditionals = []
    
    keywords = [
        'if approved', 'if rejected',
        'if yes', 'if no',
        'upon approval', 'upon rejection',
        'in case of', 'otherwise'
    ]
    
    for activity in activities:
        desc = activity['description'].lower()
        
        for keyword in keywords:
            if keyword in desc:
                conditionals.append({
                    'activity': activity['number'],
                    'type': 'conditional',
                    'trigger': keyword,
                    'description': activity['description']
                })
                break
    
    return conditionals
```

---

## Team Inference

```python
def infer_team_from_text(activity_text):
    """Attempt to infer team from activity description."""
    
    text_lower = activity_text.lower()
    
    # Keyword patterns
    patterns = {
        'Sales': ['sales', 'opportunity', 'customer', 'distributor portal'],
        'Planning': ['planning', 'inventory', 'analyze', 'coordinate with pricing'],
        'Biz-ops': ['biz-ops', 'order confirmation', 'sfdc operation', 'system order'],
        'JOTS': ['jots', 'vehicle', 'transportation', 'freight coordination'],
        'Plant Operations': ['plant', 'dispatch', 'invoice', 'grn', 'e-way bill']
    }
    
    scores = {team: 0 for team in patterns.keys()}
    
    for team, keywords in patterns.items():
        for keyword in keywords:
            if keyword in text_lower:
                scores[team] += 1
    
    max_score = max(scores.values())
    if max_score > 0:
        candidates = [team for team, score in scores.items() if score == max_score]
        if len(candidates) == 1:
            return candidates[0], 'medium'
        else:
            return candidates, 'low'  # Multiple candidates
    
    return None, 'none'  # No match
```

---

## Usage Example

```python
# Parse multiple source documents
docx_content = parse_docx('/mnt/project/OrderLogging_v3.docx')
pptx_content = parse_pptx('/mnt/project/flowchart.pptx')
pdf_content = parse_pdf('/mnt/project/VendorMasterSOPs.pdf')

# Detect activities
activities = detect_activities(docx_content['text'])

# Detect conditionals
conditionals = detect_conditional_flows(activities)

# Infer teams
for activity in activities:
    team, confidence = infer_team_from_text(activity['description'])
    activity['team'] = team
    activity['confidence'] = confidence

print(f"Extracted {len(activities)} activities")
print(f"Found {len(conditionals)} conditional flows")
print(f"Images: {len(docx_content['images']) + len(pptx_content['images'])}")
```
