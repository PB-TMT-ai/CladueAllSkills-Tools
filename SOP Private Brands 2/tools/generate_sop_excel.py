"""
JSW ONE Private Brands SOP Master Excel Generator

Main orchestrator script. Reads all SOP documents from the Documents folder,
parses them, maps Confluence links, and generates a formatted master Excel file.

Usage:
    python tools/generate_sop_excel.py
    python tools/generate_sop_excel.py --docs-dir Documents --output output
"""

import argparse
import hashlib
import json
import os
import sys
from pathlib import Path
from datetime import datetime
from collections import defaultdict

# Add tools directory to path
TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, TOOLS_DIR)

from models.sop_data import Activity, Section, Journey
from parsers.confluence_doc_parser import parse_confluence_doc
from parsers.docx_parser import parse_docx_sop
from parsers.pdf_parser import parse_pdf_sop
from parsers.confluence_links import parse_confluence_links, match_confluence_link
from mapping.document_classifier import (
    classify_document, get_section_order, get_all_journeys_ordered
)
from mapping.field_extractor import extract_activities_from_doc, extract_activities_from_docx
from excel_writer.writer import write_master_excel


def main():
    parser = argparse.ArgumentParser(description='Generate PB SOP Master Excel')
    parser.add_argument('--docs-dir', default=None, help='Path to source documents folder')
    parser.add_argument('--output', default=None, help='Output directory')
    args = parser.parse_args()

    # Resolve paths relative to project root
    project_root = os.path.dirname(TOOLS_DIR)
    docs_dir = args.docs_dir or os.path.join(project_root, 'Documents')
    output_dir = args.output or os.path.join(project_root, 'output')

    print(f"[1/6] Discovering documents in: {docs_dir}")
    doc_files = discover_documents(docs_dir)
    print(f"  Found {len(doc_files)} unique documents")

    print(f"\n[2/6] Loading Confluence links...")
    # Search recursively for the links file (may be in a subfolder)
    links_file = None
    for candidate in Path(docs_dir).rglob('PB SOP Links Confluence.xlsx'):
        links_file = str(candidate)
        break
    if links_file is None:
        links_file = os.path.join(docs_dir, 'PB SOP Links Confluence.xlsx')
    confluence_links = {}
    if os.path.exists(links_file):
        confluence_links = parse_confluence_links(links_file)
        print(f"  Loaded {len(confluence_links)} Confluence links")
    else:
        print(f"  WARNING: Links file not found at {links_file}")

    print(f"\n[3/6] Parsing and classifying documents...")
    all_activities = []  # List of (journey, section, [Activity])
    seen_titles = set()  # Content-based dedup (same title + doc_type = skip)

    for filepath in doc_files:
        filename = os.path.basename(filepath)
        print(f"  Processing: {filename}")

        # Classify document
        config = classify_document(filename)
        if config is None:
            print(f"    WARNING: No classification found, skipping")
            continue

        doc_type = config.get("doc_type", "")
        print(f"    Type: {doc_type}")

        # Parse document
        try:
            if filepath.endswith('.docx'):
                parsed_doc = parse_docx_sop(filepath)
            elif filepath.endswith('.doc'):
                parsed_doc = parse_confluence_doc(filepath)
            elif filepath.endswith('.pdf'):
                parsed_doc = parse_pdf_sop(filepath)
            elif filepath.endswith('.xlsx'):
                from parsers.xlsx_sop_parser import parse_xlsx_sop
                parsed_doc = parse_xlsx_sop(filepath)
            else:
                print(f"    WARNING: Unsupported file type, skipping")
                continue
        except Exception as e:
            print(f"    ERROR parsing: {e}")
            continue

        # Set doc_type from classifier
        parsed_doc.doc_type = doc_type

        # Content-based dedup: skip if same title + doc_type already processed
        dedup_key = (parsed_doc.title.strip().lower(), doc_type)
        if dedup_key in seen_titles:
            print(f"    Skipping content duplicate (same title as earlier doc)")
            continue
        seen_titles.add(dedup_key)

        # Match Confluence link
        link = match_confluence_link(parsed_doc.title, confluence_links)
        if not link:
            link = match_confluence_link(filename, confluence_links)
        if link:
            print(f"    Confluence link matched: {link[:60]}...")
        else:
            print(f"    No Confluence link matched")

        # Extract activities
        try:
            if doc_type == "docx_sop":
                activities = extract_activities_from_docx(filepath, config, link)
                # Group by phase
                phase_groups = defaultdict(list)
                for act in activities:
                    phase = getattr(act, '_phase', config.get('journey', 'Pre - Order'))
                    phase_groups[phase].append(act)

                for phase, acts in phase_groups.items():
                    section_label = _get_section_for_docx_phase(phase)
                    all_activities.append((phase, section_label, acts))

            elif doc_type == "quality_manual":
                activities = extract_activities_from_doc(parsed_doc, config, link)
                multi_journey = config.get("multi_journey", {})
                if multi_journey:
                    for act in activities:
                        hint = getattr(act, '_journey_hint', None)
                        if hint and hint in multi_journey:
                            all_activities.append((hint, multi_journey[hint], [act]))
                        else:
                            journey = config.get('journey', 'Order')
                            section = multi_journey.get(journey, config.get('section', ''))
                            all_activities.append((journey, section, [act]))
                else:
                    journey = config.get('journey', 'Order')
                    section = config.get('section', 'UNCATEGORIZED')
                    all_activities.append((journey, section, activities))

            else:
                activities = extract_activities_from_doc(parsed_doc, config, link)
                journey = config.get('journey', 'Pre - Order')
                section = config.get('section', 'UNCATEGORIZED')
                all_activities.append((journey, section, activities))

            act_count = sum(1 for _, _, acts in all_activities) if doc_type == "docx_sop" else len(activities)
            print(f"    Extracted activities: {len(activities)}")

        except Exception as e:
            print(f"    ERROR extracting activities: {e}")
            import traceback
            traceback.print_exc()
            continue

    print(f"\n[4/6] Building journey structure...")
    journeys = build_journey_structure(all_activities)

    total_activities = sum(
        len(section.activities)
        for journey in journeys
        for section in journey.sections
    )
    print(f"  Total journeys: {len(journeys)}")
    print(f"  Total activities: {total_activities}")

    print(f"\n[5/6] Generating Excel file...")
    output_path = os.path.join(output_dir, 'JSW_ONE_PB_SOPs_Master.xlsx')
    write_master_excel(journeys, output_path)
    print(f"  Generated: {output_path}")

    print(f"\n[6/6] Saving manifest...")
    manifest_path = os.path.join(output_dir, 'manifest.json')
    save_manifest(manifest_path, doc_files, confluence_links)
    print(f"  Manifest saved: {manifest_path}")

    print(f"\nDone! Master Excel generated at: {output_path}")
    return output_path


def discover_documents(docs_dir: str) -> list:
    """Find all SOP source files recursively, deduplicate by content hash."""
    files = []
    seen_hashes = set()

    for f in sorted(Path(docs_dir).rglob('*')):
        if not f.is_file():
            continue
        if f.suffix not in ('.doc', '.docx', '.pdf', '.xlsx'):
            continue
        # Skip non-SOP xlsx files (Confluence links spreadsheet)
        if 'PB SOP Links' in f.name:
            continue
        h = hashlib.md5(f.read_bytes()).hexdigest()
        if h not in seen_hashes:
            seen_hashes.add(h)
            files.append(str(f))
        else:
            print(f"  Skipping duplicate: {f.name}")
    return files


def build_journey_structure(all_activities: list) -> list:
    """Build the final Journey > Section > Activity hierarchy.

    Args:
        all_activities: List of (journey_name, section_label, [Activity]) tuples

    Returns:
        List of Journey objects in display order
    """
    # Collect all activities grouped by journey and section
    structure = defaultdict(lambda: defaultdict(list))

    for journey_name, section_label, activities in all_activities:
        structure[journey_name][section_label].extend(activities)

    # Build Journey objects in order
    journeys = []
    sr_no_counter = 1

    for journey_name in get_all_journeys_ordered():
        if journey_name not in structure:
            continue

        sections = []
        section_data = structure[journey_name]

        # Order sections: known sections first (from SECTION_ORDER), then unknown
        ordered_section_labels = get_section_order(journey_name)
        seen_labels = set()

        for label in ordered_section_labels:
            if label in section_data:
                seen_labels.add(label)
                activities = section_data[label]
                # Assign sequential sr_no
                for act in activities:
                    act.sr_no = sr_no_counter
                    sr_no_counter += 1
                sections.append(Section(label=label, activities=activities))

        # Add any remaining sections not in the predefined order
        for label, activities in section_data.items():
            if label not in seen_labels:
                for act in activities:
                    act.sr_no = sr_no_counter
                    sr_no_counter += 1
                sections.append(Section(label=label, activities=activities))

        if sections:
            journeys.append(Journey(name=journey_name, sections=sections))

    return journeys


def _get_section_for_docx_phase(phase: str) -> str:
    """Map a docx phase to an appropriate section label."""
    phase_section_map = {
        "Pre - Order": "E. ORDER LOGGING (PRE-ORDER)",
        "Order": "E. ORDER LOGGING (ORDER)",
        "Post Order": "E. ORDER LOGGING (POST-ORDER)",
    }
    return phase_section_map.get(phase, "E. ORDER LOGGING")


def save_manifest(manifest_path: str, doc_files: list, confluence_links: dict):
    """Save a build manifest for change detection on re-runs."""
    manifest = {
        "generated_at": datetime.now().isoformat(),
        "generator": "tools/generate_sop_excel.py",
        "files": {},
        "confluence_links_count": len(confluence_links),
    }

    for filepath in doc_files:
        h = hashlib.md5(Path(filepath).read_bytes()).hexdigest()
        manifest["files"][os.path.basename(filepath)] = {
            "hash": h,
            "size": os.path.getsize(filepath),
        }

    os.makedirs(os.path.dirname(manifest_path) or '.', exist_ok=True)
    with open(manifest_path, 'w') as f:
        json.dump(manifest, f, indent=2)


if __name__ == '__main__':
    main()
