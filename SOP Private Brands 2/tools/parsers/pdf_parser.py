"""
Parser for PDF documents using OCR (pytesseract + pdf2image) with pdfplumber fallback.
Handles scanned PDFs like the JSW One TMT Quality Manual.

Requires system binaries:
  - Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki
  - Poppler (pdftoppm): https://github.com/oschwartz10612/poppler-windows/releases
"""

import hashlib
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from PIL import ImageOps

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from models.sop_data import SOPDocument

# --- Configurable binary paths (override via environment variables) ---
TESSERACT_CMD = os.environ.get(
    'TESSERACT_CMD',
    r'C:\Users\2750834\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
)
POPPLER_PATH = os.environ.get(
    'POPPLER_PATH',
    r'C:\Users\2750834\AppData\Local\poppler\poppler-24.08.0\Library\bin'
)

# Set tesseract path
pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

# OCR cache directory
CACHE_DIR = os.path.join(os.path.dirname(__file__), '..', '..', 'output', '.ocr_cache')


def parse_pdf_sop(filepath: str) -> SOPDocument:
    """Parse a PDF document (scanned or native) and return a SOPDocument."""
    filename = os.path.basename(filepath)
    file_hash = hashlib.md5(Path(filepath).read_bytes()).hexdigest()

    # Step 1: Check OCR cache
    cached_text = _load_cached_text(file_hash)
    if cached_text:
        print(f"    Using cached OCR text ({len(cached_text)} chars)")
        raw_text = cached_text
    else:
        # Step 2: Try native text extraction first (fast)
        raw_text = _try_native_extraction(filepath)

        if not raw_text or len(raw_text.strip()) < 100:
            # Step 3: OCR extraction for scanned pages
            print(f"    Scanned PDF detected, running OCR...")
            raw_text = _ocr_extract(filepath)

        # Step 4: Cache the extracted text
        _save_cached_text(file_hash, raw_text, filename)

    # Step 5: Parse the raw text into SOPDocument fields
    title = _extract_title(raw_text, filename)
    headings = _extract_headings(raw_text)
    tables = _extract_tables(raw_text)
    purpose = _extract_purpose(raw_text)
    stakeholders = _extract_stakeholders(raw_text)
    steps = _extract_steps(raw_text)
    sign_off = _extract_sign_off(raw_text)

    return SOPDocument(
        filename=filename,
        title=title,
        doc_type="",  # Set by document_classifier
        purpose=purpose,
        stakeholders=stakeholders,
        steps=steps,
        tables=tables,
        headings=headings,
        raw_text=raw_text,
        sign_off_info=sign_off,
        escalation_info=None,
    )


def _try_native_extraction(filepath: str) -> str:
    """Try extracting text natively with pdfplumber (for non-scanned PDFs)."""
    try:
        all_text = []
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                all_text.append(text)
        return '\n'.join(all_text)
    except Exception as e:
        print(f"    pdfplumber failed: {e}")
        return ""


def _ocr_extract(filepath: str) -> str:
    """Extract text from scanned PDF pages using OCR."""
    all_text = []

    try:
        # Get page count first
        with pdfplumber.open(filepath) as pdf:
            total_pages = len(pdf.pages)
    except Exception:
        total_pages = 50  # Fallback estimate

    for page_num in range(1, total_pages + 1):
        try:
            # Convert single page to image at 300 DPI
            images = convert_from_path(
                filepath,
                dpi=300,
                first_page=page_num,
                last_page=page_num,
                poppler_path=POPPLER_PATH,
            )

            if not images:
                continue

            img = images[0]

            # Preprocess: grayscale + auto-contrast
            img = img.convert('L')
            img = ImageOps.autocontrast(img)

            # OCR with page-segmentation mode 6 (assume uniform block of text)
            text = pytesseract.image_to_string(img, lang='eng', config='--psm 6')
            text = _clean_ocr_text(text)

            if len(text.strip()) > 50:
                all_text.append(f"--- PAGE {page_num} ---\n{text}")
                if page_num % 5 == 0 or page_num == 1:
                    print(f"    OCR page {page_num}/{total_pages}...")
            # Skip pages with very little text (likely diagrams/images)

        except Exception as e:
            print(f"    OCR error on page {page_num}: {e}")
            continue

    print(f"    OCR complete: {len(all_text)} pages with text")
    return '\n\n'.join(all_text)


def _clean_ocr_text(text: str) -> str:
    """Clean up common OCR noise."""
    # Collapse multiple blank lines
    text = re.sub(r'\n{3,}', '\n\n', text)
    # Fix common OCR artifacts
    text = re.sub(r'[|]{2,}', '|', text)
    # Remove isolated single characters that are likely noise
    text = re.sub(r'\n[^\w\n]{1,2}\n', '\n', text)
    # Normalize whitespace within lines
    text = re.sub(r'[ \t]{3,}', '  ', text)
    return text.strip()


# --- Cache functions ---

def _load_cached_text(file_hash: str) -> str | None:
    """Load cached OCR text if it exists."""
    cache_file = os.path.join(CACHE_DIR, f'{file_hash}.txt')
    if os.path.exists(cache_file):
        with open(cache_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # Skip the metadata header (first 3 lines)
            lines = content.split('\n', 3)
            if len(lines) > 3:
                return lines[3]
    return None


def _save_cached_text(file_hash: str, text: str, filename: str):
    """Save extracted text to cache."""
    os.makedirs(CACHE_DIR, exist_ok=True)
    cache_file = os.path.join(CACHE_DIR, f'{file_hash}.txt')
    with open(cache_file, 'w', encoding='utf-8') as f:
        f.write(f"# OCR Cache: {filename}\n")
        f.write(f"# Extracted: {datetime.now().isoformat()}\n")
        f.write(f"# Hash: {file_hash}\n")
        f.write(text)


# --- Text parsing functions ---

def _extract_title(raw_text: str, filename: str) -> str:
    """Extract document title from OCR text."""
    # Look for prominent text near the top
    lines = raw_text.split('\n')[:30]
    for line in lines:
        line = line.strip()
        # Skip page markers and short lines
        if line.startswith('---') or len(line) < 5:
            continue
        # All-caps lines are likely titles
        if line.isupper() and len(line) > 10:
            return line.title()
        # Lines containing "Quality" or "Manual" or "QAP"
        if any(kw in line.upper() for kw in ['QUALITY', 'MANUAL', 'QAP']):
            return line.strip()

    # Fallback: use filename
    return filename.replace('.pdf', '').replace('_', ' ')


def _extract_headings(raw_text: str) -> list:
    """Extract headings from OCR text (ALL CAPS lines, numbered sections)."""
    headings = []
    for line in raw_text.split('\n'):
        line = line.strip()
        if not line or line.startswith('---'):
            continue

        # Numbered section headings: "1.0", "2.0", "1.", etc.
        if re.match(r'^\d+[\.\d]*\s+[A-Z]', line) and len(line) < 100:
            headings.append((2, line))
        # ALL CAPS headings (at least 3 words)
        elif line.isupper() and len(line.split()) >= 2 and len(line) < 80:
            headings.append((1, line.title()))
        # Known section keywords
        elif any(kw in line.upper() for kw in [
            'QUALITY POLICY', 'QUALITY ASSURANCE PLAN', 'QAP',
            'PROCESS FLOW', 'SOP', 'VISUAL INSPECTION',
            'PRE-DISPATCH', 'JOB SETUP', 'STANDARD OPERATING',
            'CORRECTIVE ACTION', 'APPROVAL'
        ]) and len(line) < 100:
            headings.append((2, line))

    return headings


def _extract_tables(raw_text: str) -> list:
    """Extract table-like structures from OCR text."""
    tables = []
    current_table = []
    in_table = False

    for line in raw_text.split('\n'):
        line = line.strip()

        # Detect table rows by pipe delimiters
        if '|' in line and line.count('|') >= 2:
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if cells:
                current_table.append(cells)
                in_table = True
        # Detect table rows by consistent spacing (tab-separated)
        elif '\t' in line and line.count('\t') >= 2:
            cells = [c.strip() for c in line.split('\t') if c.strip()]
            if len(cells) >= 3:
                current_table.append(cells)
                in_table = True
        else:
            if in_table and current_table and len(current_table) >= 2:
                tables.append(current_table)
            current_table = []
            in_table = False

    # Don't forget last table
    if current_table and len(current_table) >= 2:
        tables.append(current_table)

    return tables


def _extract_purpose(raw_text: str) -> str | None:
    """Extract purpose or quality policy text."""
    text_upper = raw_text.upper()

    # Look for Quality Policy section
    for keyword in ['QUALITY POLICY', 'PURPOSE', 'OBJECTIVE', 'SCOPE']:
        idx = text_upper.find(keyword)
        if idx >= 0:
            # Get text after the keyword, up to next heading or 500 chars
            after = raw_text[idx + len(keyword):idx + len(keyword) + 800]
            # Clean up
            after = re.sub(r'^[:\s\n]+', '', after)
            # Stop at next likely heading
            lines = after.split('\n')
            collected = []
            for line in lines:
                if line.strip().isupper() and len(line.strip()) > 10:
                    break
                if line.strip().startswith('---'):
                    break
                if line.strip():
                    collected.append(line.strip())
                if len(' '.join(collected)) > 500:
                    break
            if collected:
                return ' '.join(collected)[:500]

    return None


def _extract_stakeholders(raw_text: str) -> dict:
    """Extract stakeholder/owner information from OCR text."""
    stakeholders = {}
    known_roles = [
        'CEO', 'Quality Manager', 'Plant Operations', 'QA Head',
        'Regional Quality', 'Technical Head', 'Inspection',
    ]

    for role in known_roles:
        pattern = re.compile(rf'{re.escape(role)}[:\s]*([^\n]+)', re.IGNORECASE)
        match = pattern.search(raw_text)
        if match:
            stakeholders[role] = match.group(1).strip()[:100]

    # Also look for signature/sign-off blocks with names
    sign_pattern = re.compile(r'(?:Prepared|Approved|Reviewed|Authorized)\s+(?:by|By)[:\s]*([^\n]+)', re.IGNORECASE)
    for match in sign_pattern.finditer(raw_text):
        name = match.group(1).strip()[:100]
        if name:
            stakeholders[match.group(0).split()[0]] = name

    return stakeholders


def _extract_steps(raw_text: str) -> list:
    """Extract numbered steps from OCR text."""
    steps = []
    # Match numbered steps: "1.", "1)", "Step 1:", etc.
    step_pattern = re.compile(r'(?:^|\n)\s*(?:Step\s+)?(\d+)[.\)]\s+(.+)', re.IGNORECASE)
    for match in step_pattern.finditer(raw_text):
        step_text = match.group(2).strip()
        if len(step_text) > 10 and len(step_text) < 500:
            steps.append(step_text)

    # Deduplicate while preserving order
    seen = set()
    unique_steps = []
    for s in steps:
        key = s.lower()[:50]
        if key not in seen:
            seen.add(key)
            unique_steps.append(s)

    return unique_steps[:30]  # Cap at 30 steps


def _extract_sign_off(raw_text: str) -> str | None:
    """Extract sign-off information."""
    sign_keywords = ['sign-off', 'sign off', 'approved by', 'authorized by', 'reviewed by']
    for keyword in sign_keywords:
        idx = raw_text.lower().find(keyword)
        if idx >= 0:
            after = raw_text[idx:idx + 200]
            return after.strip()[:200]
    return None
