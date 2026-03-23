"""
Parser for .doc files exported from Confluence (MIME/HTML format).
Extracts headings, tables, and body text from the HTML content.
"""

import email
import email.policy
import re
from html.parser import HTMLParser
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from models.sop_data import SOPDocument


class _HTMLContentExtractor(HTMLParser):
    """Extracts structured content (headings, tables, paragraphs) from HTML."""

    def __init__(self):
        super().__init__()
        self.headings = []          # [(level, text)]
        self.tables = []            # [[[cell, cell, ...], ...]]  (list of tables, each is list of rows)
        self.paragraphs = []        # [text]
        self.lists = []             # [text] - list items

        self._current_tag = None
        self._tag_stack = []
        self._text_buf = []

        # Table parsing state
        self._in_table = False
        self._current_table = []
        self._current_row = []
        self._current_cell = []
        self._table_depth = 0

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        self._tag_stack.append(tag)
        self._current_tag = tag

        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            self._text_buf = []
        elif tag == 'table':
            self._table_depth += 1
            if self._table_depth == 1:
                self._in_table = True
                self._current_table = []
        elif tag == 'tr' and self._table_depth == 1:
            self._current_row = []
        elif tag in ('td', 'th') and self._table_depth == 1:
            self._current_cell = []
        elif tag == 'p' and not self._in_table:
            self._text_buf = []
        elif tag in ('li',):
            self._text_buf = []

    def handle_endtag(self, tag):
        tag = tag.lower()

        if tag in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6'):
            level = int(tag[1])
            text = ' '.join(''.join(self._text_buf).split()).strip()
            if text:
                self.headings.append((level, text))
            self._text_buf = []

        elif tag == 'table':
            if self._table_depth == 1 and self._current_table:
                self.tables.append(self._current_table)
                self._current_table = []
                self._in_table = False
            self._table_depth = max(0, self._table_depth - 1)

        elif tag == 'tr' and self._table_depth == 1:
            if self._current_row:
                self._current_table.append(self._current_row)
            self._current_row = []

        elif tag in ('td', 'th') and self._table_depth == 1:
            cell_text = ' '.join(''.join(self._current_cell).split()).strip()
            self._current_row.append(cell_text)
            self._current_cell = []

        elif tag == 'p' and not self._in_table:
            # If we're inside a <li>, DON'T flush to paragraphs — let <li>
            # end-tag collect the text.  (Confluence exports <li><p>text</p></li>)
            if 'li' not in self._tag_stack:
                text = ' '.join(''.join(self._text_buf).split()).strip()
                if text:
                    self.paragraphs.append(text)
                self._text_buf = []
            # else: keep _text_buf intact for the enclosing <li> to consume

        elif tag == 'li':
            text = ' '.join(''.join(self._text_buf).split()).strip()
            if text:
                self.lists.append(text)
            self._text_buf = []

        if self._tag_stack and self._tag_stack[-1] == tag:
            self._tag_stack.pop()
        self._current_tag = self._tag_stack[-1] if self._tag_stack else None

    def handle_data(self, data):
        if self._table_depth == 1 and any(t in ('td', 'th') for t in self._tag_stack):
            self._current_cell.append(data)
        elif any(t in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6') for t in self._tag_stack):
            self._text_buf.append(data)
        elif any(t == 'li' for t in self._tag_stack):
            self._text_buf.append(data)
        elif any(t == 'p' for t in self._tag_stack) and not self._in_table:
            self._text_buf.append(data)

    def handle_entityref(self, name):
        char = {'amp': '&', 'lt': '<', 'gt': '>', 'nbsp': ' ', 'quot': '"'}.get(name, f'&{name};')
        self.handle_data(char)

    def handle_charref(self, name):
        try:
            if name.startswith('x'):
                char = chr(int(name[1:], 16))
            else:
                char = chr(int(name))
            self.handle_data(char)
        except (ValueError, OverflowError):
            self.handle_data(f'&#{name};')


def _extract_html_from_mime(filepath: str) -> str:
    """Extract the HTML content from a Confluence MIME-exported .doc file."""
    with open(filepath, 'rb') as f:
        raw = f.read()

    msg = email.message_from_bytes(raw, policy=email.policy.default)

    for part in msg.walk():
        content_type = part.get_content_type()
        if content_type == 'text/html':
            payload = part.get_payload(decode=True)
            if payload:
                return payload.decode('utf-8', errors='replace')

    # Fallback: if not multipart, the entire content might be HTML
    payload = msg.get_payload(decode=True)
    if payload:
        return payload.decode('utf-8', errors='replace')

    raise ValueError(f"No HTML content found in {filepath}")


def _extract_title_from_html(html: str) -> str:
    """Extract the <title> tag content."""
    match = re.search(r'<title[^>]*>(.*?)</title>', html, re.DOTALL | re.IGNORECASE)
    if match:
        title = re.sub(r'<[^>]+>', '', match.group(1))
        # Decode HTML entities
        title = title.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        title = title.replace('&#39;', "'").replace('&quot;', '"')
        return ' '.join(title.split()).strip()
    return ""


def parse_confluence_doc(filepath: str) -> SOPDocument:
    """Parse a Confluence-exported .doc file and return a structured SOPDocument."""
    html = _extract_html_from_mime(filepath)

    # Extract title
    title = _extract_title_from_html(html)

    # Parse HTML structure
    extractor = _HTMLContentExtractor()
    try:
        extractor.feed(html)
    except Exception:
        pass  # Best-effort parsing

    # Build raw text from paragraphs and lists
    raw_text = '\n'.join(extractor.paragraphs + extractor.lists)

    # Extract purpose/scope
    purpose = _extract_section_text(extractor, 'purpose')

    # Extract stakeholders
    stakeholders = _extract_stakeholders(extractor)

    # Extract steps
    steps = _extract_numbered_steps(extractor)

    # Extract sign-off info
    sign_off = _extract_sign_off(extractor)

    # Extract escalation info
    escalation = _extract_escalation(extractor)

    filename = os.path.basename(filepath)

    return SOPDocument(
        filename=filename,
        title=title,
        doc_type="",  # Will be set by document_classifier
        purpose=purpose,
        stakeholders=stakeholders,
        steps=steps,
        tables=extractor.tables,
        headings=extractor.headings,
        raw_text=raw_text,
        sign_off_info=sign_off,
        escalation_info=escalation,
    )


def _extract_section_text(extractor: _HTMLContentExtractor, section_keyword: str) -> str | None:
    """Find paragraphs following a heading that contains the keyword."""
    target_level = None
    capture = False
    collected = []

    for heading_level, heading_text in extractor.headings:
        if section_keyword.lower() in heading_text.lower():
            target_level = heading_level
            break

    if target_level is None:
        # Try paragraphs that start with the keyword
        for para in extractor.paragraphs:
            if section_keyword.lower() in para.lower()[:50]:
                return para
        return None

    # Walk paragraphs and headings in document order to find text after the heading
    # Since we don't have ordering between paragraphs and headings, use raw_text
    pattern = re.compile(
        rf'(?:purpose|scope)[:\s]*(.*?)(?=\n(?:stakeholder|process|step|sign[\s-]off|escalation)|\Z)',
        re.IGNORECASE | re.DOTALL
    )
    match = pattern.search(extractor.paragraphs[0] if extractor.paragraphs else '')
    if not match:
        # Return all paragraphs up to the next heading-like paragraph
        for para in extractor.paragraphs:
            if section_keyword.lower() in para.lower():
                continue
            if any(kw in para.lower() for kw in ['stakeholder', 'process steps', 'sign-off', 'escalation']):
                break
            collected.append(para)
            if len(collected) >= 3:
                break
        return ' '.join(collected) if collected else None

    return match.group(1).strip() if match else None


def _extract_stakeholders(extractor: _HTMLContentExtractor) -> dict:
    """Extract stakeholder information from the document."""
    stakeholders = {}

    # Look for a stakeholders table
    for table in extractor.tables:
        if not table:
            continue
        header_row = ' '.join(table[0]).lower()
        if 'stakeholder' in header_row or 'role' in header_row or 'team' in header_row:
            for row in table[1:]:
                if len(row) >= 2:
                    stakeholders[row[0].strip()] = row[1].strip()
            return stakeholders

    # Look for stakeholder info in paragraphs
    capture = False
    for para in extractor.paragraphs:
        if 'stakeholder' in para.lower():
            capture = True
            continue
        if capture:
            if any(kw in para.lower() for kw in ['process', 'step', 'sign-off', 'scope']):
                break
            # Try to parse "Role - Name" format
            if '-' in para or ':' in para:
                parts = re.split(r'[-:]', para, maxsplit=1)
                if len(parts) == 2:
                    stakeholders[parts[0].strip()] = parts[1].strip()
            elif para.strip():
                stakeholders[para.strip()] = ""

    return stakeholders


def _extract_numbered_steps(extractor: _HTMLContentExtractor) -> list:
    """Extract numbered process steps from the document."""
    steps = []

    # First check list items (primary source for Confluence docs)
    if extractor.lists:
        for item in extractor.lists:
            cleaned = re.sub(r'^\d+[\.\)]\s*', '', item).strip()
            if cleaned:
                steps.append(cleaned)
        if steps:
            return steps

    # Check paragraphs for numbered steps after trigger keywords
    capture = False
    # Also check headings for "steps"/"process" triggers
    step_heading_found = any(
        any(kw in h_text.lower() for kw in ['step', 'process'])
        for _, h_text in extractor.headings
    )
    for para in extractor.paragraphs:
        if any(kw in para.lower() for kw in [
            'process steps', 'steps:', 'process:', 'workflow steps',
            'steps', 'process'
        ]):
            capture = True
            continue
        if step_heading_found and not capture:
            # If heading had 'steps'/'process', start capturing from first paragraph
            capture = True
        if capture:
            if any(kw in para.lower() for kw in ['sign-off', 'escalation', 'timelines', 'tat']):
                break
            # Capture numbered items
            if re.match(r'^\d+[\.\)]\s', para):
                cleaned = re.sub(r'^\d+[\.\)]\s*', '', para).strip()
                if cleaned:
                    steps.append(cleaned)
            elif para.strip() and not para.startswith(('Note:', 'Remark:')):
                steps.append(para.strip())

    # Fallback: check tables for step-like content
    if not steps:
        for table in extractor.tables:
            if not table:
                continue
            header = ' '.join(table[0]).lower()
            if any(kw in header for kw in ['step', 'activity', 'action', 'task']):
                step_col = None
                for i, cell in enumerate(table[0]):
                    if any(kw in cell.lower() for kw in ['step', 'activity', 'action', 'task', 'description']):
                        step_col = i
                        break
                if step_col is not None:
                    for row in table[1:]:
                        if step_col < len(row) and row[step_col].strip():
                            steps.append(row[step_col].strip())
                break

    return steps


def _extract_sign_off(extractor: _HTMLContentExtractor) -> str | None:
    """Extract sign-off information."""
    for table in extractor.tables:
        if not table:
            continue
        header = ' '.join(table[0]).lower()
        if 'sign' in header and 'off' in header:
            parts = []
            for row in table[1:]:
                parts.append(' | '.join(row))
            return '; '.join(parts) if parts else None

    # Check paragraphs
    for para in extractor.paragraphs:
        if 'sign' in para.lower() and 'off' in para.lower():
            return para
    return None


def _extract_escalation(extractor: _HTMLContentExtractor) -> str | None:
    """Extract escalation matrix information."""
    for table in extractor.tables:
        if not table:
            continue
        header = ' '.join(table[0]).lower()
        if 'escalation' in header or 'level' in header:
            parts = []
            for row in table[1:]:
                parts.append(' | '.join(row))
            return '; '.join(parts) if parts else None

    for para in extractor.paragraphs:
        if 'escalation' in para.lower():
            return para
    return None
