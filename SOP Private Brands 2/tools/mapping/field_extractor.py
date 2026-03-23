"""
Field extractor: converts parsed SOPDocument objects into Activity objects
based on document type. Each doc_type has its own extraction logic.
"""

import os
import sys
import re

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from models.sop_data import Activity, Section, Journey
from parsers.docx_parser import extract_docx_activities


def extract_activities_from_doc(parsed_doc, doc_config: dict, confluence_link: str | None) -> list:
    """Extract Activity objects from a parsed SOPDocument based on its type.

    Args:
        parsed_doc: SOPDocument from parser
        doc_config: classification config from document_classifier
        confluence_link: matched Confluence URL or None

    Returns:
        List of Activity objects
    """
    doc_type = doc_config.get("doc_type", "")
    prefix = doc_config.get("activity_prefix", parsed_doc.title)

    if doc_type == "ops_sop":
        return _extract_ops_sop(parsed_doc, prefix, confluence_link, doc_config)
    elif doc_type == "workflow_doc":
        return _extract_workflow_doc(parsed_doc, prefix, confluence_link, doc_config)
    elif doc_type == "technical_doc":
        return _extract_technical_doc(parsed_doc, prefix, confluence_link)
    elif doc_type == "field_spec":
        return _extract_field_spec(parsed_doc, prefix, confluence_link)
    elif doc_type == "demo_doc":
        return _extract_demo_doc(parsed_doc, prefix, confluence_link)
    elif doc_type == "quality_manual":
        return _extract_quality_manual(parsed_doc, prefix, confluence_link)
    elif doc_type == "marketing_visual":
        return _extract_marketing_visual(parsed_doc, prefix, confluence_link)
    elif doc_type == "marketing_activities_xlsx":
        return _extract_marketing_activities_xlsx(parsed_doc, prefix, confluence_link)
    else:
        # Generic fallback
        return _extract_generic(parsed_doc, prefix, confluence_link)


def extract_activities_from_docx(filepath: str, doc_config: dict, confluence_link: str | None) -> list:
    """Extract Activity objects from JSWOrderLogging_V18.docx.

    This document maps most directly to the Construct.xlsx format.
    Each table row becomes activity step data, grouped by phase/section.
    """
    raw_activities = extract_docx_activities(filepath)
    filename = os.path.basename(filepath)

    # Group by phase and activity number
    grouped = {}
    for act in raw_activities:
        phase = act['phase']
        act_num = act['activity_num']
        key = (phase, act_num)

        if key not in grouped:
            grouped[key] = {
                'phase': phase,
                'activity_num': act_num,
                'steps': [],
                'teams': set(),
                'interfaces': set(),
                'sign_offs': set(),
            }

        if act['steps']:
            grouped[key]['steps'].append(act['steps'])
        # Filter out single-char garbage values (e.g. "s" from source data entry errors)
        if act['team'] and len(act['team'].strip()) > 1:
            grouped[key]['teams'].add(act['team'])
        if act['interface'] and len(act['interface'].strip()) > 1:
            grouped[key]['interfaces'].add(act['interface'])
        if act['sign_off'] and len(act['sign_off'].strip()) > 1:
            grouped[key]['sign_offs'].add(act['sign_off'])

    activities = []
    for (phase, act_num), data in grouped.items():
        act_label = f"Activity {act_num}" if act_num else "Order Logging"
        # Build description from contextual metadata (phase, team, interface)
        # so it is completely independent of the step text in Col H
        team_str = ', '.join(data['teams']) if data['teams'] else ""
        intf_str = ', '.join(data['interfaces']) if data['interfaces'] else ""
        desc_parts = [f"{phase} phase"]
        if team_str:
            desc_parts.append(f"Owner: {team_str}")
        if intf_str:
            desc_parts.append(f"via {intf_str}")
        desc = ' | '.join(desc_parts)

        activity = Activity(
            activity_name=act_label,
            description=desc,
            description_details=[],
            owner=team_str,
            interface=intf_str,
            sign_off=', '.join(data['sign_offs']) if data['sign_offs'] else "",
            steps=data['steps'],
            flow_type="Positive",
            sop_link=confluence_link,
            remarks=None,
            source_document=filename,
        )
        activity._phase = phase  # Temporary attribute for phase grouping
        activities.append(activity)

    return activities


def _extract_ops_sop(parsed_doc, prefix: str, confluence_link: str | None,
                     doc_config: dict | None = None) -> list:
    """Extract from ops_sop documents (Rebates, Influencer meets, etc.).

    Creates one Activity per document with steps from the process section.
    """
    doc_config = doc_config or {}

    # Build owner from stakeholders
    owner_parts = []
    for role, name in parsed_doc.stakeholders.items():
        if name:
            owner_parts.append(f"{role}: {name}")
        else:
            owner_parts.append(role)
    owner = '; '.join(owner_parts) if owner_parts else ""

    # Use default_owner from config if no stakeholders found
    if not owner and doc_config.get("default_owner"):
        owner = doc_config["default_owner"]

    # Interface detection from stakeholders or raw text
    interface = _detect_interface(parsed_doc)

    # Build description from purpose or title
    description = parsed_doc.purpose or parsed_doc.title
    # If purpose looks truncated/garbled (starts with lowercase fragment), prefer title
    if description and not description[0].isupper() and parsed_doc.title:
        description = parsed_doc.title

    # Steps from parsed steps
    steps = parsed_doc.steps if parsed_doc.steps else []

    # If no numbered steps found, try bullet points from raw text
    if not steps and parsed_doc.raw_text:
        for line in parsed_doc.raw_text.split('\n'):
            line = line.strip()
            if line.startswith(('-', '\u2022', '*')) and len(line) > 10:
                steps.append(line.lstrip('-*\u2022 ').strip())

    # If still no steps, try to get from tables
    if not steps and parsed_doc.tables:
        for table in parsed_doc.tables:
            if not table or len(table) < 2:
                continue
            header_text = ' '.join(table[0]).lower()
            if any(kw in header_text for kw in ['step', 'activity', 'process', 'action']):
                for row in table[1:]:
                    # Take the most descriptive cell (longest text)
                    best_cell = max(row, key=len) if row else ''
                    if best_cell.strip():
                        steps.append(best_cell.strip())
                break

    # Build remarks from TAT/escalation info
    remarks_parts = []
    if parsed_doc.escalation_info:
        remarks_parts.append(f"Escalation: {parsed_doc.escalation_info}")
    remarks = '; '.join(remarks_parts) if remarks_parts else None

    activity = Activity(
        activity_name=prefix,
        description=description,
        description_details=[],
        owner=owner,
        interface=interface,
        sign_off=parsed_doc.sign_off_info or "",
        steps=steps,
        flow_type="Positive",
        sop_link=confluence_link,
        remarks=remarks,
        source_document=parsed_doc.filename,
    )

    return [activity]


def _extract_workflow_doc(parsed_doc, prefix: str, confluence_link: str | None, config: dict) -> list:
    """Extract from workflow documents (Approval Workflows).

    If expand_table_rows is True, each table row becomes a separate Activity.
    """
    activities = []

    if config.get("expand_table_rows"):
        for table in parsed_doc.tables:
            if not table or len(table) < 2:
                continue

            header = table[0]
            header_lower = [h.lower() for h in header]

            # Identify columns
            name_col = _find_col(header_lower, ['workflow', 'approval', 'name', 'type'])
            use_case_col = _find_col(header_lower, ['use case', 'description', 'scenario'])
            approver_col = _find_col(header_lower, ['approver', 'approve', 'who all'])
            initiator_col = _find_col(header_lower, ['initiat', 'who initiates'])
            rejection_col = _find_col(header_lower, ['rejection', 'what happens'])
            info_col = _find_col(header_lower, ['additional', 'remark', 'note', 'information'])

            for row in table[1:]:
                if all(not cell.strip() for cell in row):
                    continue

                act_name = _safe_get(row, name_col) if name_col is not None else prefix
                description = _safe_get(row, use_case_col) if use_case_col is not None else ""
                sign_off = _safe_get(row, approver_col) if approver_col is not None else ""
                initiator = _safe_get(row, initiator_col) if initiator_col is not None else ""
                rejection = _safe_get(row, rejection_col) if rejection_col is not None else ""
                remarks = _safe_get(row, info_col) if info_col is not None else None

                if not act_name.strip():
                    continue

                # Build workflow steps from approval hierarchy (distinct from description)
                wf_steps = []
                if initiator:
                    wf_steps.append(f"{initiator} initiates the approval request")
                if sign_off:
                    # Parse approval levels (e.g. "L1: Regional sales manager L2: Sales head")
                    for level in re.split(r'(?=L\d:)', sign_off):
                        level = level.strip()
                        if level:
                            wf_steps.append(f"{level} reviews and approves")
                if rejection:
                    wf_steps.append(f"On rejection: {rejection}")

                activity = Activity(
                    activity_name=act_name.strip(),
                    description=description,
                    owner=initiator or "Sales / BU Head",
                    interface="Salesforce",
                    sign_off=sign_off,
                    steps=wf_steps if wf_steps else [f"Approval workflow for: {act_name.strip()}"],
                    flow_type="Positive",
                    sop_link=confluence_link,
                    remarks=remarks if remarks and remarks.strip() else None,
                    source_document=parsed_doc.filename,
                )
                activities.append(activity)

    if not activities:
        # Fallback: single activity
        activities.append(Activity(
            activity_name=prefix,
            description=parsed_doc.purpose or parsed_doc.title,
            owner="",
            interface="Salesforce",
            sign_off="",
            steps=parsed_doc.steps,
            flow_type="Positive",
            sop_link=confluence_link,
            source_document=parsed_doc.filename,
        ))

    return activities


def _extract_technical_doc(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract from technical documents (Opportunity Workflow, etc.).

    Creates activities from major sections/headings.
    """
    activities = []

    # Group content by major headings (level 1-2)
    sections_found = []
    current_section = None

    for level, text in parsed_doc.headings:
        if level <= 2:
            if current_section:
                sections_found.append(current_section)
            current_section = {'heading': text, 'sub_headings': [], 'level': level}
        elif current_section:
            current_section['sub_headings'].append(text)

    if current_section:
        sections_found.append(current_section)

    if sections_found:
        for section in sections_found:
            steps = section['sub_headings'][:10]  # Limit sub-headings as steps
            activity = Activity(
                activity_name=section['heading'],
                description=f"Technical documentation: {section['heading']}",
                owner="Product / Tech",
                interface=_detect_interface(parsed_doc),
                sign_off="",
                steps=steps,
                flow_type="Positive",
                sop_link=confluence_link,
                source_document=parsed_doc.filename,
            )
            activities.append(activity)
    else:
        # Single activity fallback
        activities.append(Activity(
            activity_name=prefix,
            description=parsed_doc.purpose or parsed_doc.title,
            owner="Product / Tech",
            interface=_detect_interface(parsed_doc),
            sign_off="",
            steps=parsed_doc.steps[:10],
            flow_type="Positive",
            sop_link=confluence_link,
            source_document=parsed_doc.filename,
        ))

    return activities


def _extract_field_spec(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract from field specification documents (Data Enrichment).

    Creates one activity with field names as steps.
    """
    steps = []

    # Look for the field specification table
    for table in parsed_doc.tables:
        if not table or len(table) < 2:
            continue

        header = table[0]
        header_lower = [h.lower() for h in header]

        # Look for tables with field/attribute columns
        field_col = _find_col(header_lower, ['field', 'attribute', 'data', 'parameter', 'name'])
        if field_col is not None:
            for row in table[1:]:
                field_val = _safe_get(row, field_col)
                if field_val:
                    # Include additional context from other columns
                    extra = [_safe_get(row, i) for i in range(len(row)) if i != field_col and _safe_get(row, i)]
                    if extra:
                        steps.append(f"{field_val} ({', '.join(extra[:2])})")
                    else:
                        steps.append(field_val)

    if not steps:
        steps = parsed_doc.steps[:20]

    # Owner from approver table
    owner = ""
    for table in parsed_doc.tables:
        if not table:
            continue
        header_text = ' '.join(table[0]).lower()
        if 'approver' in header_text or 'owner' in header_text:
            for row in table[1:]:
                if row:
                    owner = ' | '.join(cell for cell in row if cell.strip())
                    break
            break

    activity = Activity(
        activity_name=prefix,
        description=parsed_doc.purpose or f"Field specification for {prefix}",
        owner=owner or "Product / Sales",
        interface="Salesforce",
        sign_off="",
        steps=steps,
        flow_type="Positive",
        sop_link=confluence_link,
        remarks="See original document for full field specifications",
        source_document=parsed_doc.filename,
    )

    return [activity]


def _extract_demo_doc(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract minimal entry from demo/screenshot-heavy documents.

    Creates a single row with title, link, and brief description.
    """
    description = parsed_doc.purpose or parsed_doc.title
    if not description or description == prefix:
        # Use first meaningful paragraph
        for para in (parsed_doc.raw_text or "").split('\n'):
            if len(para.strip()) > 20:
                description = para.strip()[:200]
                break

    activity = Activity(
        activity_name=prefix,
        description=description,
        owner="Product",
        interface="Salesforce / App",
        sign_off="",
        steps=[],
        flow_type="Positive",
        sop_link=confluence_link,
        remarks="See original document for visual walkthrough",
        source_document=parsed_doc.filename,
    )

    return [activity]


def _extract_quality_manual(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract activities from a quality manual PDF (OCR-parsed).

    Creates multiple activities:
    - Quality Policy statement
    - QAP Testing parameters (from tables)
    - SOP procedures (job setup, pre-dispatch, visual inspection)
    """
    activities = []

    # 1. Quality Policy activity
    policy_text = _find_section_text(parsed_doc.raw_text,
                                     ['quality policy', 'quality objective', 'committed to'])
    if policy_text:
        act = Activity(
            activity_name=f"{prefix} - Quality Policy",
            description=policy_text[:500],
            owner="CEO; Regional Quality Manager",
            interface=_detect_interface(parsed_doc),
            sign_off="QA Head",
            steps=[],
            flow_type="Positive",
            sop_link=confluence_link,
            source_document=parsed_doc.filename,
        )
        act._journey_hint = "Order"
        activities.append(act)

    # 2. QAP Testing activities (from tables with test/frequency columns)
    qap_found = False
    for table in parsed_doc.tables:
        if not table or len(table) < 2:
            continue
        header_lower = [h.lower() for h in table[0]]
        test_col = _find_col(header_lower, ['test', 'parameter', 'check', 'inspection'])
        if test_col is not None:
            steps = []
            for row in table[1:]:
                step_text = _safe_get(row, test_col)
                if step_text and len(step_text) > 3:
                    # Include context from other columns
                    extras = [_safe_get(row, i) for i in range(len(row))
                              if i != test_col and _safe_get(row, i)]
                    if extras:
                        steps.append(f"{step_text} ({', '.join(extras[:2])})")
                    else:
                        steps.append(step_text)
            if steps:
                act = Activity(
                    activity_name=f"{prefix} - QAP Testing",
                    description="Quality Assurance Plan - testing parameters and acceptance criteria",
                    owner="QA Head; Plant Operations Manager",
                    interface=_detect_interface(parsed_doc),
                    sign_off="Regional Quality Manager",
                    steps=steps[:20],
                    flow_type="Positive",
                    sop_link=confluence_link,
                    source_document=parsed_doc.filename,
                )
                act._journey_hint = "Order"
                activities.append(act)
                qap_found = True

    # 2b. QAP Testing fallback: extract from OCR text when table parsing fails
    if not qap_found and parsed_doc.raw_text:
        qap_steps = _extract_qap_from_text(parsed_doc.raw_text)
        if qap_steps:
            act = Activity(
                activity_name=f"{prefix} - QAP Testing",
                description="Quality Assurance Plan - testing parameters, sampling frequency and acceptance criteria per IS 1786",
                owner="QA Head; Plant Operations Manager",
                interface=_detect_interface(parsed_doc),
                sign_off="Regional Quality Manager",
                steps=qap_steps,
                flow_type="Positive",
                sop_link=confluence_link,
                source_document=parsed_doc.filename,
            )
            act._journey_hint = "Order"
            activities.append(act)

    # 3. SOP Procedure activities (detected from headings/text)
    sop_sections = _find_sop_sections(parsed_doc.raw_text, parsed_doc.headings)
    seen_sop_names = set()  # Deduplicate SOPs with same descriptive name
    for section_name, section_steps, descriptive_name, body_text in sop_sections:
        # Dedup: skip if we already have an SOP with the same descriptive name
        dedup_key = descriptive_name.lower().strip() if descriptive_name else section_name.lower().strip()
        if dedup_key in seen_sop_names:
            continue
        seen_sop_names.add(dedup_key)

        # Use descriptive name for the activity if available
        display_name = descriptive_name or section_name
        # Check both the section header AND the descriptive title for dispatch keywords
        combined_text = f"{section_name} {descriptive_name}".lower()
        is_dispatch = any(kw in combined_text
                          for kw in ['dispatch', 'pre-dispatch', 'storage', 'loading',
                                     'nonconform', 'cooling bed', 'mixing', 'yard'])
        # Extract Purpose + Scope from the SOP body for the description
        purpose_desc = _extract_purpose_from_block(body_text)
        act = Activity(
            activity_name=f"{prefix} - {display_name}",
            description=purpose_desc or f"SOP: {display_name}",
            owner="Plant Operations Manager; QA Head",
            interface=_detect_interface(parsed_doc),
            sign_off="Regional Quality Manager",
            steps=section_steps[:15],
            flow_type="Positive",
            sop_link=confluence_link,
            source_document=parsed_doc.filename,
        )
        act._journey_hint = "Post Order" if is_dispatch else "Order"
        activities.append(act)

    # 4. Fallback: if no structured content found, create single generic activity
    if not activities:
        act = Activity(
            activity_name=prefix,
            description=parsed_doc.purpose or "TMT Quality Assurance Manual",
            owner="QA Head",
            interface=_detect_interface(parsed_doc),
            sign_off="Regional Quality Manager",
            steps=parsed_doc.steps[:15],
            flow_type="Positive",
            sop_link=confluence_link,
            remarks="OCR-extracted from scanned PDF; verify content accuracy",
            source_document=parsed_doc.filename,
        )
        act._journey_hint = "Order"
        activities.append(act)

    return activities


def _extract_marketing_visual(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract from marketing visual specification documents (GSB, Wall Painting).

    The .docx contains a 5-column Approval/Execution table matching
    OrderLogging_V18 format: Activity, Steps, Team, Interface, Sign off.
    Each table row becomes a step in a single Activity.
    """
    steps = []
    owners = set()
    interfaces = set()
    sign_offs = set()

    for table in parsed_doc.tables:
        if not table or len(table) < 2:
            continue
        header_lower = [h.lower() for h in table[0]]
        # Look for the 5-column execution table (not the 2-col spec table)
        step_col = _find_col(header_lower, ['step'])
        if step_col is None:
            continue
        team_col = _find_col(header_lower, ['team'])
        iface_col = _find_col(header_lower, ['interface', 'utilities'])
        signoff_col = _find_col(header_lower, ['sign off', 'signoff', 'sign_off'])

        for row in table[1:]:
            step_text = _safe_get(row, step_col)
            if step_text.strip():
                steps.append(step_text.strip())
            team = _safe_get(row, team_col)
            if team and team.strip() and team.strip() != '-':
                owners.add(team.strip())
            iface = _safe_get(row, iface_col)
            if iface and iface.strip() and iface.strip() != '-':
                interfaces.add(iface.strip())
            so = _safe_get(row, signoff_col)
            if so and so.strip() and so.strip() != '-':
                sign_offs.add(so.strip())

    # Build specification details from spec table (2-column)
    spec_details = []
    for table in parsed_doc.tables:
        if not table or len(table) < 2:
            continue
        header_lower = [h.lower() for h in table[0]]
        if _find_col(header_lower, ['step']) is not None:
            continue  # skip execution table
        if _find_col(header_lower, ['item', 'detail']) is not None:
            for row in table[1:]:
                item = _safe_get(row, 0)
                detail = _safe_get(row, 1)
                if item and detail:
                    spec_details.append(f"{item}: {detail}")

    activity = Activity(
        activity_name=prefix,
        description=parsed_doc.purpose or parsed_doc.title,
        description_details=spec_details,
        owner='; '.join(sorted(owners)) if owners else "Marketing; Channel Team; Vendor",
        interface='; '.join(sorted(interfaces)) if interfaces else "WhatsApp; Field Execution",
        sign_off='; '.join(sorted(sign_offs)) if sign_offs else "Marketing",
        steps=steps,
        flow_type="Positive",
        sop_link=confluence_link,
        remarks="Converted from specification image",
        source_document=parsed_doc.filename,
    )
    return [activity]


def _extract_marketing_activities_xlsx(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Extract from SOP Marketing Activities.xlsx.

    Each data row becomes an Activity. The 5 process phase columns
    (Requirement gathering, Budget approval, Vendor selection,
    Output verification, Invoice processing) become the steps.
    """
    activities = []

    if not parsed_doc.tables:
        return [Activity(
            activity_name=prefix,
            description=parsed_doc.purpose or parsed_doc.title,
            owner="Marketing; Business Team",
            interface="Confluence; WhatsApp",
            sign_off="Marketing; Finance",
            steps=[],
            flow_type="Positive",
            sop_link=confluence_link,
            source_document=parsed_doc.filename,
        )]

    for table in parsed_doc.tables:
        if not table or len(table) < 2:
            continue

        header = table[0]
        header_lower = [h.lower() for h in header]

        # Find columns by keyword
        activity_col = _find_col(header_lower, ['activity', 'header'])
        phase_cols = [
            (_find_col(header_lower, ['requirement', 'gathering']), "Requirement Gathering"),
            (_find_col(header_lower, ['budget', 'approval']), "Budget Approval"),
            (_find_col(header_lower, ['vendor', 'selection']), "Vendor Selection"),
            (_find_col(header_lower, ['output', 'verification']), "Output Verification"),
            (_find_col(header_lower, ['invoice', 'processing']), "Invoice Processing"),
        ]

        for row in table[1:]:
            act_name = _safe_get(row, activity_col) if activity_col is not None else ""
            if not act_name.strip():
                continue

            # Build steps from phase columns
            steps = []
            for col_idx, phase_label in phase_cols:
                cell_val = _safe_get(row, col_idx) if col_idx is not None else ""
                if cell_val.strip():
                    steps.append(f"[{phase_label}] {cell_val.strip()}")

            activity = Activity(
                activity_name=act_name.strip(),
                description=f"Marketing activity: {act_name.strip()}",
                owner="Marketing; Business Team; Vendor",
                interface="Confluence; WhatsApp; Finance Portal",
                sign_off="Marketing Head; Finance",
                steps=steps,
                flow_type="Positive",
                sop_link=confluence_link,
                source_document=parsed_doc.filename,
            )
            activities.append(activity)

    return activities


def _find_section_text(raw_text: str, keywords: list) -> str | None:
    """Find and extract text from a section identified by keywords."""
    if not raw_text:
        return None
    text_lower = raw_text.lower()
    for keyword in keywords:
        idx = text_lower.find(keyword)
        if idx >= 0:
            # Get text after the keyword
            after = raw_text[idx:idx + 600]
            lines = after.split('\n')
            collected = []
            for i, line in enumerate(lines):
                if i == 0:
                    continue  # Skip the keyword line itself
                line = line.strip()
                if not line or line.startswith('---'):
                    if collected:
                        break
                    continue
                # Stop at next heading-like line
                if line.isupper() and len(line) > 10 and i > 2:
                    break
                collected.append(line)
                if len(' '.join(collected)) > 500:
                    break
            if collected:
                return ' '.join(collected)[:500]
    return None


def _find_sop_sections(raw_text: str, headings: list) -> list:
    """Find SOP procedure sections in the document.

    Returns list of (section_name, [steps], descriptive_name, body_text) tuples.
    The descriptive_name is the human-readable title extracted from the
    lines following the SOP header (e.g., "Joint Inspection of Chemistry").
    The body_text is the raw text of the SOP block for purpose extraction.
    """
    sop_keywords = [
        'job setup', 'job-set up', 'pre-dispatch', 'pre dispatch',
        'visual inspection', 'corrective action', 'standard operating',
        'incoming inspection', 'in-process', 'final inspection',
    ]

    sections = []

    # For 'standard operating procedure' sections, use regex to split
    # the raw text into SOP blocks and extract descriptive titles
    sop_pattern = re.compile(
        r'(STANDARD\s+OPERATING\s+PROCEDURE[^\n]*)\n(.*?)(?=STANDARD\s+OPERATING\s+PROCEDURE|\Z)',
        re.DOTALL | re.IGNORECASE
    )
    sop_matches = list(sop_pattern.finditer(raw_text))

    if sop_matches:
        for m in sop_matches:
            header_line = m.group(1).strip()
            body = m.group(2)

            # Extract descriptive name from first few lines after header
            descriptive_name = _extract_sop_descriptive_name(body)

            # Extract steps from the body
            steps = _extract_steps_from_block(body)
            if steps or descriptive_name:
                sections.append((header_line, steps, descriptive_name or header_line, body))
        return sections

    # Fallback: try headings for other SOP keywords
    for level, heading_text in headings:
        heading_lower = heading_text.lower()
        for kw in sop_keywords:
            if kw in heading_lower:
                idx = raw_text.lower().find(heading_lower[:30])
                if idx >= 0:
                    after = raw_text[idx + len(heading_text):idx + len(heading_text) + 1000]
                    steps = _extract_steps_from_block(after)
                    if steps:
                        sections.append((heading_text.strip(), steps, heading_text.strip(), after))
                break

    # If no headings matched, search raw text
    if not sections:
        for kw in sop_keywords:
            idx = raw_text.lower().find(kw)
            if idx >= 0:
                start = max(0, idx - 50)
                block_start = raw_text.rfind('\n', start, idx)
                if block_start < 0:
                    block_start = start
                block = raw_text[block_start:block_start + 1000]
                lines = block.strip().split('\n')
                section_name = lines[0].strip() if lines else kw.title()
                body = '\n'.join(lines[1:])
                steps = _extract_steps_from_block(body)
                if steps:
                    sections.append((section_name, steps, section_name, body))

    return sections


def _extract_sop_descriptive_name(body_text: str) -> str | None:
    """Extract the descriptive title from an SOP section body.

    Looks for lines like 'Joint Inspection of Chemistry', 'Pre-Dispatch
    Inspection -Guidelines', 'Handling of nonconforming products', etc.
    in the first few lines after the SOP header.
    """
    lines = [l.strip() for l in body_text.split('\n') if l.strip()][:8]

    # Known descriptive patterns in QA manual SOP bodies
    title_keywords = [
        'joint inspection', 'pre-dispatch', 'job set up', 'job setup',
        'handling of', 'avoiding mixing', 'rib area', 'bendability',
        'martensite', 'ring', 'mechanical properties', 'chemistry',
        'nonconform', 'cooling bed',
    ]

    # First pass: extract title from lines containing "Date:" (title + date
    # on the same line is the most reliable indicator of the SOP title).
    # Skip boilerplate JSW/Version lines that may also contain "Date:".
    for line in lines[:5]:
        line_lower = line.lower()
        if 'date:' not in line_lower or len(line) < 20:
            continue
        if any(skip in line_lower for skip in ['jsw one', 'version:']):
            continue
        parts = re.split(r'\s*Date:', line, flags=re.IGNORECASE)
        if parts[0].strip() and len(parts[0].strip()) > 5:
            clean = re.sub(r'[|_]+', '', parts[0]).strip()
            clean = re.sub(r'\s+', ' ', clean)
            if len(clean) > 5:
                return clean

    # Second pass: keyword matching on non-boilerplate lines
    for line in lines:
        line_lower = line.lower()
        # Skip boilerplate lines
        if any(skip in line_lower for skip in [
            'jsw one', 'version:', 'purpose:', 'scope:', 'contract manufacturer',
            'sgs qc', 'operation manager', 'date:', 'rev ', 'to check', 'to ensure',
            'to make', 'to avoid',
        ]):
            continue
        # Skip very short or noisy OCR lines
        if len(line) < 8 or line.startswith(('|', '_', '-')):
            continue
        # Match against known title patterns
        if any(kw in line_lower for kw in title_keywords):
            # Clean up OCR artifacts
            clean = re.sub(r'[|_]+', '', line).strip()
            clean = re.sub(r'\s+', ' ', clean)
            if len(clean) > 5:
                return clean

    return None


def _extract_steps_from_block(text: str) -> list:
    """Extract numbered steps from a text block.

    Handles two common formats:
    - Inline: ``1. Do something`` or ``Step 1. Do something`` (single line)
    - Multi-line: ``Step 1\\nDo something`` (heading on one line, content on next)
    """
    # Skip to the actual procedure/workflow section if present, to avoid
    # extracting numbered rows from Teams/Responsibilities tables.
    procedure_markers = [
        'work flow:', 'workflow:', 'procedure:', 'process:',
        'standard operating process:', 'operating process:',
    ]
    text_lower = text.lower()
    best_idx = -1
    for marker in procedure_markers:
        idx = text_lower.find(marker)
        if idx >= 0 and (best_idx < 0 or idx < best_idx):
            best_idx = idx
    if best_idx >= 0:
        # Start extraction from the procedure marker line
        text = text[best_idx:]

    steps = []
    lines = text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        i += 1
        if not line:
            continue
        # Skip OCR page markers (e.g., "--- PAGE 18 ---")
        if re.match(r'^---\s*PAGE\s+\d+\s*---$', line):
            continue
        # Skip scanner watermarks
        if 'scanned with' in line.lower():
            continue
        # Stop at next section/heading (but skip known SOP boilerplate
        # lines like "JSW ONE TMT", "STANDARD OPERATING PROCEDURE", etc.)
        if line.isupper() and len(line) > 10:
            line_upper = line.upper()
            if any(bp in line_upper for bp in ['JSW', 'TMT', 'JODL', 'VERSION']):
                continue
            break
        if line.startswith('---'):
            break
        # Strip leading OCR artifacts (@, _, e, etc.) before checking for steps
        clean_line = re.sub(r'^[^a-zA-Z0-9(]*', '', line).strip()
        # Inline numbered steps: "1. Do something" or "Step 1. Do something"
        match = re.match(r'(?:Step\s+)?(\d+)[.\),]\s+(.+)', clean_line, re.IGNORECASE)
        if match:
            step_num = int(match.group(1))
            step_text = match.group(2).strip()
            # Skip table cell fragments vs real workflow steps.
            # Real steps are action sentences; table rows are short noun phrases.
            expected_next = len(steps) + 1
            if step_num > expected_next + 3:
                continue  # Number too far ahead — likely a table row number
            # Action verbs that indicate a real workflow step
            _action_starts = (
                'check', 'ensure', 'verify', 'collect', 'perform', 'create',
                'load', 'navigate', 'submit', 'approve', 'monitor', 'follow',
                'all ', 'random ', 'rejected', 'once ', 'after ', 'if ',
                'push', 'generate', 'fill', 'complete', 'receive', 'conduct',
                'do ', 'prepare', 'confirm', 'attach', 'search', 'copy',
                'change', 'mark ', 'send', 'update', 'determine', 'process',
                'the ', 'materials', 'jsw one', 'sgs ',
            )
            text_lower = step_text.lower()
            is_action = any(text_lower.startswith(v) for v in _action_starts)
            is_sentence = step_text.endswith('.') or len(step_text) > 50
            if not is_action and not is_sentence:
                continue  # Likely a table cell fragment (noun phrase, not action)
            if len(step_text) > 3:
                steps.append(step_text)
            continue
        # Multi-line "Step N" format: heading line has just "Step N" (with optional
        # trailing punctuation), content follows on subsequent non-empty lines.
        step_header = re.match(r'Step\s+(\d+)\s*[.,;:]?\s*$', clean_line, re.IGNORECASE)
        if step_header:
            # Collect content from lines that follow until next Step/section
            content_parts = []
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line:
                    i += 1
                    continue
                # Skip page markers inside step content
                if re.match(r'^---\s*PAGE\s+\d+\s*---$', next_line):
                    i += 1
                    continue
                if 'scanned with' in next_line.lower():
                    i += 1
                    continue
                # Stop if we hit another Step header or section boundary
                if re.match(r'Step\s+\d+', next_line, re.IGNORECASE):
                    break
                if next_line.isupper() and len(next_line) > 10:
                    break
                if next_line.startswith('---'):
                    break
                # Stop at boilerplate endings
                if any(kw in next_line.lower() for kw in [
                    'prepared by', 'reviewed by', 'approved by',
                    'scanned with', 'records:', 'images'
                ]):
                    break
                content_parts.append(next_line)
                i += 1
                # Limit per-step content collection
                if len(content_parts) >= 4:
                    break
            if content_parts:
                steps.append(' '.join(content_parts))
            continue
        # Bullet points — only accept if they look like action/workflow items,
        # not table cell fragments (e.g., "-Heavy visual bend")
        if line.startswith(('-', '*', '\u2022')) and len(line) > 10:
            bullet_text = line.lstrip('-*\u2022 ').strip()
            # Require either a sentence-like length or an action verb
            if len(bullet_text) > 40 or bullet_text.endswith('.') or any(
                bullet_text.lower().startswith(v) for v in (
                    'check', 'ensure', 'verify', 'all ', 'if ', 'once',
                    'after', 'the ', 'follow', 'complete', 'perform',
                    'prepare', 'conduct', 'submit', 'generate', 'process',
                    'materials', 'jsw one', 'sgs ', 'bundles',
                )
            ):
                steps.append(bullet_text)

    return steps[:15]


def _extract_purpose_from_block(body_text: str) -> str | None:
    """Extract Purpose and Scope text from an SOP body block.

    Returns a combined string like:
        "To check and report the chemistry... Scope: Rolled product"
    """
    purpose = None
    scope = None

    for keyword, target in [('purpose:', 'purpose'), ('scope:', 'scope')]:
        idx = body_text.lower().find(keyword)
        if idx < 0:
            continue
        after = body_text[idx + len(keyword):]
        # Collect lines until next section keyword
        collected = []
        for line in after.split('\n'):
            line = line.strip()
            if not line:
                continue
            line_lower = line.lower()
            # Stop at next known section header
            if any(line_lower.startswith(s) for s in [
                'scope:', 'applicable', 'teams', 'work flow:', 'step ',
                'sampling', 'pre-testing', 'joint testing',
                'standard operating', 'records:',
            ]):
                break
            # Skip noisy/short OCR lines
            if len(line) < 5 or line.startswith(('|', '_', '---')):
                continue
            collected.append(line)
            if len(collected) >= 3:
                break
        text = ' '.join(collected).strip()
        # Clean OCR artifacts
        text = re.sub(r'[|_]+', '', text).strip()
        text = re.sub(r'\s+', ' ', text)
        if target == 'purpose' and text:
            purpose = text
        elif target == 'scope' and text:
            scope = text

    if purpose and scope:
        return f"{purpose}. Scope: {scope}"
    return purpose


def _extract_qap_from_text(raw_text: str) -> list:
    """Fallback QAP extraction from OCR text when table parsing fails.

    Searches for the "Testing and Sampling Plan" section and extracts
    test parameter names with their sampling frequency.
    """
    # Known test parameters from the TMT Quality Manual QAP table
    qap_tests = [
        ('weighing', 'Weight/m measurement of rolled product'),
        ('surface', 'Surface and dimension check of rolled product'),
        ('dispatch', 'Pre-dispatch inspection of each bundle during loading'),
        ('tensile', 'Tensile Test - YS, UTS, %E and %Total Elongation'),
        ('bendability', 'Bendability test of rolled product'),
        ('rib', 'Rib Geometry / AR Value check of rolled product'),
        ('martensi', 'Martensitic Ring Test - ring and core structure verification'),
        ('chemistry', 'Chemistry analysis - C, Mn, Si, S, P, Cr, Ni'),
    ]

    # Find the Testing and Sampling Plan section
    idx = raw_text.lower().find('testing and sampling plan')
    if idx < 0:
        return []

    # Extract a block of text around this section
    block = raw_text[idx:idx + 2000].lower()

    steps = []
    for keyword, description in qap_tests:
        if keyword in block:
            steps.append(description)

    return steps


def _extract_generic(parsed_doc, prefix: str, confluence_link: str | None) -> list:
    """Generic fallback extraction."""
    activity = Activity(
        activity_name=prefix,
        description=parsed_doc.purpose or parsed_doc.title,
        owner="",
        interface="",
        sign_off="",
        steps=parsed_doc.steps[:10],
        flow_type="Positive",
        sop_link=confluence_link,
        source_document=parsed_doc.filename,
    )
    return [activity]


def _detect_interface(parsed_doc) -> str:
    """Detect the primary interface/system from document content."""
    text = (parsed_doc.raw_text or "").lower()
    interfaces = []

    if 'salesforce' in text or 'sfdc' in text or 'sf ' in text:
        interfaces.append("Salesforce")
    if 'excel' in text or 'spreadsheet' in text or 'google sheet' in text:
        interfaces.append("Excel")
    if 'jira' in text:
        interfaces.append("JIRA")
    if 'sap' in text or 'erp' in text:
        interfaces.append("SAP/ERP")
    if 'email' in text:
        interfaces.append("Email")
    if 'app' in text or 'portal' in text or 'website' in text:
        interfaces.append("App/Portal")
    if 'oms' in text:
        interfaces.append("OMS")
    if 'zoho' in text:
        interfaces.append("Zoho Books")

    return ', '.join(interfaces[:3]) if interfaces else ""


def _find_col(header_lower: list, keywords: list) -> int | None:
    """Find the column index whose header contains any of the keywords."""
    for i, h in enumerate(header_lower):
        for kw in keywords:
            if kw in h:
                return i
    return None


def _safe_get(lst: list, index: int | None, default: str = '') -> str:
    """Safely get a string from a list by index."""
    if index is None or index < 0 or index >= len(lst):
        return default
    return lst[index].strip() if lst[index] else default
