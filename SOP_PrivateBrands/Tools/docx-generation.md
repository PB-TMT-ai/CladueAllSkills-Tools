# Tool: DOCX Generation

## Overview
Technical specifications for generating JSW SOP Word documents using python-docx.

---

## Core Library

```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
```

---

## Document Initialization

```python
def create_sop_document(title, process_name, version):
    """Initialize new SOP document with JSW branding."""
    
    doc = Document()
    
    # Set margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Add title
    title_para = doc.add_paragraph()
    title_run = title_para.add_run(title)
    title_run.font.name = 'Calibri'
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add subtitle
    subtitle_para = doc.add_paragraph()
    subtitle_run = subtitle_para.add_run(f"\n{process_name}\n")
    subtitle_run.font.name = 'Calibri'
    subtitle_run.font.size = Pt(14)
    subtitle_run.font.bold = True
    subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add version
    version_para = doc.add_paragraph()
    version_run = version_para.add_run(f"Version {version}")
    version_run.font.name = 'Calibri'
    version_run.font.size = Pt(11)
    version_run.font.italic = True
    version_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph()  # Spacer
    
    return doc
```

---

## Table Creation

```python
def create_sop_table(doc, num_rows=1):
    """Create 5-column SOP table with proper formatting."""
    
    # Create table (rows + 1 for header)
    table = doc.add_table(rows=num_rows + 1, cols=5)
    table.style = 'Table Grid'
    
    # Set column widths (percentages of page width)
    widths = [Inches(1.0), Inches(2.3), Inches(0.7), Inches(2.3), Inches(0.3)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
    
    # Header row
    header_cells = table.rows[0].cells
    headers = ['Activity', 'Steps', 'Team', 'Interface / Utilities used', 'Sign off']
    
    for idx, header_text in enumerate(headers):
        cell = header_cells[idx]
        cell.text = header_text
        
        # Format header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.name = 'Calibri'
        
        # Header background color (light gray)
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), 'D9D9D9')
        cell._element.get_or_add_tcPr().append(shading_elm)
    
    return table
```

---

## Adding Activities

```python
def add_activity_row(table, activity_num, steps, team, interface_text=None, image_path=None):
    """Add single activity row to table."""
    
    row = table.add_row()
    cells = row.cells
    
    # Activity number
    cells[0].text = activity_num
    format_cell_text(cells[0], bold=True)
    
    # Steps
    cells[1].text = steps
    format_cell_text(cells[1])
    
    # Team
    cells[2].text = team
    format_cell_text(cells[2])
    
    # Interface/Utilities
    if image_path:
        add_image_to_cell(cells[3], image_path)
    if interface_text:
        cells[3].text = interface_text
        format_cell_text(cells[3])
    
    # Sign off (leave empty)
    cells[4].text = ""
    
    return row

def format_cell_text(cell, bold=False):
    """Apply standard formatting to cell text."""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
            run.font.bold = bold
```

---

## Image Embedding

```python
def add_image_to_cell(cell, image_path, width_inches=4.0):
    """Embed image in table cell with proper sizing."""
    
    # Clear existing content
    cell.text = ""
    
    # Add image
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run()
    
    try:
        run.add_picture(image_path, width=Inches(width_inches))
        
        # Center-align
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        return True
    except Exception as e:
        # Handle failed image embedding
        cell.text = f"[Image error: {e}]"
        return False
```

---

## Table Borders

```python
def set_table_borders(table):
    """Ensure all table borders are visible."""
    
    from docx.oxml import parse_xml
    
    borders_xml = """
    <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tcBorders>
    """
    
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._element.get_or_add_tcPr()
            tcBorders = parse_xml(borders_xml)
            tcPr.append(tcBorders)
```

---

## Complete Workflow

```python
def generate_sop_docx(activities, output_path, title, process_name, version):
    """
    Complete workflow to generate SOP DOCX file.
    
    Args:
        activities: List of activity dictionaries with keys:
                   - number: "1.a.i"
                   - steps: "Description of steps"
                   - team: "Sales"
                   - interface: "SFDC"
                   - image_path: "/path/to/image.png"
        output_path: Save location
        title: Document title
        process_name: Process name
        version: Version number
    """
    
    # 1. Initialize document
    doc = create_sop_document(title, process_name, version)
    
    # 2. Create table
    table = create_sop_table(doc, num_rows=len(activities))
    
    # 3. Add activities
    for idx, activity in enumerate(activities):
        add_activity_row(
            table=table,
            activity_num=activity['number'],
            steps=activity['steps'],
            team=activity['team'],
            interface_text=activity.get('interface'),
            image_path=activity.get('image_path')
        )
    
    # 4. Apply borders
    set_table_borders(table)
    
    # 5. Save document
    doc.save(output_path)
    
    return output_path
```

---

## Usage Example

```python
# Define activities
activities = [
    {
        'number': '1.',
        'steps': 'Login to JSW ONE TMT Distributor Portal. Click on the "Opportunities" tab.',
        'team': 'Sales',
        'interface': 'Distributor Portal',
        'image_path': '/home/claude/images/login_portal.png'
    },
    {
        'number': '1.a.',
        'steps': 'On the Opportunities page, click on "Create new opportunity" button.',
        'team': 'Sales',
        'interface': 'Distributor Portal',
        'image_path': '/home/claude/images/create_opportunity.png'
    }
]

# Generate document
output_file = generate_sop_docx(
    activities=activities,
    output_path='/mnt/user-data/outputs/OrderLogging_v1.docx',
    title='JSW ONE TMT STANDARD OPERATING PROCEDURE',
    process_name='Order Logging Process',
    version='1'
)

print(f"SOP generated: {output_file}")
```

---

## Error Handling

```python
def safe_docx_generation(activities, output_path, **kwargs):
    """Wrapper with comprehensive error handling."""
    
    try:
        # Validate inputs
        if not activities:
            raise ValueError("No activities provided")
        
        # Check image paths exist
        for activity in activities:
            if 'image_path' in activity:
                if not os.path.exists(activity['image_path']):
                    print(f"⚠️ Image not found: {activity['image_path']}")
        
        # Generate document
        result = generate_sop_docx(activities, output_path, **kwargs)
        
        # Verify output
        if not os.path.exists(result):
            raise IOError(f"Failed to create {result}")
        
        return result, None
        
    except Exception as e:
        return None, str(e)
```

---

## Performance Optimization

```python
# For large documents (>50 activities):
# 1. Generate in batches
# 2. Compress images before embedding
# 3. Use lower resolution for screenshots (72 DPI sufficient)

from PIL import Image

def optimize_image(image_path, max_width=800):
    """Optimize image for DOCX embedding."""
    img = Image.open(image_path)
    
    # Resize if too large
    if img.width > max_width:
        ratio = max_width / img.width
        new_height = int(img.height * ratio)
        img = img.resize((max_width, new_height), Image.LANCZOS)
    
    # Save optimized
    optimized_path = image_path.replace('.png', '_opt.png')
    img.save(optimized_path, optimize=True, quality=85)
    
    return optimized_path
```
