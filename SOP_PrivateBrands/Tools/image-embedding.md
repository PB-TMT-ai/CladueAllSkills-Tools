# Tool: Image Embedding

## Overview
Extract images from source documents and embed in DOCX files.

---

## Extract from DOCX

```python
from docx import Document
import os

def extract_docx_images(doc_or_path, output_dir='/home/claude/images/'):
    """Extract all images from DOCX file."""
    
    if isinstance(doc_or_path, str):
        doc = Document(doc_or_path)
    else:
        doc = doc_or_path
    
    os.makedirs(output_dir, exist_ok=True)
    images = []
    
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.target_ref:
            image_data = rel.target_part.blob
            image_ext = rel.target_ref.split('.')[-1]
            image_name = f"docx_img_{len(images) + 1}.{image_ext}"
            image_path = os.path.join(output_dir, image_name)
            
            with open(image_path, 'wb') as img_file:
                img_file.write(image_data)
            
            images.append({
                'path': image_path,
                'name': image_name,
                'source': 'docx',
                'rel_id': rel_id
            })
    
    return images
```

---

## Extract from PPTX

```python
from pptx import Presentation

def extract_pptx_images(pptx_path, output_dir='/home/claude/images/'):
    """Extract all images from PowerPoint file."""
    
    prs = Presentation(pptx_path)
    os.makedirs(output_dir, exist_ok=True)
    images = []
    
    for slide_idx, slide in enumerate(prs.slides):
        for shape_idx, shape in enumerate(slide.shapes):
            if hasattr(shape, "image"):
                image = shape.image
                image_bytes = image.blob
                image_ext = image.ext
                image_name = f"pptx_slide{slide_idx + 1}_img{shape_idx + 1}.{image_ext}"
                image_path = os.path.join(output_dir, image_name)
                
                with open(image_path, 'wb') as img_file:
                    img_file.write(image_bytes)
                
                images.append({
                    'path': image_path,
                    'name': image_name,
                    'source': 'pptx',
                    'slide': slide_idx + 1,
                    'shape': shape_idx + 1
                })
    
    return images
```

---

## Extract from PDF

```python
import fitz  # PyMuPDF

def extract_pdf_images(pdf_path, output_dir='/home/claude/images/'):
    """Extract all images from PDF file."""
    
    doc = fitz.open(pdf_path)
    os.makedirs(output_dir, exist_ok=True)
    images = []
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        image_list = page.get_images()
        
        for img_idx, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            image_name = f"pdf_page{page_num + 1}_img{img_idx + 1}.{image_ext}"
            image_path = os.path.join(output_dir, image_name)
            
            with open(image_path, 'wb') as img_file:
                img_file.write(image_bytes)
            
            images.append({
                'path': image_path,
                'name': image_name,
                'source': 'pdf',
                'page': page_num + 1
            })
    
    return images
```

---

## Image Optimization

```python
from PIL import Image

def optimize_for_docx(image_path, target_width=800, quality=85):
    """Optimize image for DOCX embedding."""
    
    img = Image.open(image_path)
    
    # Convert RGBA to RGB if needed
    if img.mode == 'RGBA':
        rgb_img = Image.new('RGB', img.size, (255, 255, 255))
        rgb_img.paste(img, mask=img.split()[3])
        img = rgb_img
    
    # Resize if too large
    if img.width > target_width:
        ratio = target_width / img.width
        new_height = int(img.height * ratio)
        img = img.resize((target_width, new_height), Image.LANCZOS)
    
    # Save optimized
    optimized_path = image_path.replace(os.path.splitext(image_path)[1], '_opt.png')
    img.save(optimized_path, 'PNG', optimize=True)
    
    return optimized_path
```

---

## Embed in DOCX

```python
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

def embed_image_in_cell(cell, image_path, width=4.0):
    """Embed image in table cell."""
    
    # Optimize first
    opt_path = optimize_for_docx(image_path)
    
    # Clear cell
    cell.text = ""
    
    # Add image
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run()
    
    try:
        run.add_picture(opt_path, width=Inches(width))
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    except Exception as e:
        cell.text = f"[Image error: {str(e)}]"
        return False
```

---

## Complete Workflow

```python
def extract_all_images(source_files):
    """Extract images from all source files."""
    
    all_images = []
    
    for file_path in source_files:
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.docx':
            images = extract_docx_images(file_path)
        elif ext == '.pptx':
            images = extract_pptx_images(file_path)
        elif ext == '.pdf':
            images = extract_pdf_images(file_path)
        else:
            continue
        
        all_images.extend(images)
    
    return all_images
```
