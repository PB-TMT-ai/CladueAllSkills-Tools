"""
Data models for SOP document parsing and Excel generation.
Matches the hierarchical structure of Construct.xlsx:
  Journey > Section > Activity (with multi-row steps/descriptions)
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Activity:
    """A single activity row (may span multiple Excel rows for steps/sub-descriptions)."""
    sr_no: Optional[int] = None
    activity_name: str = ""
    description: str = ""
    description_details: list = field(default_factory=list)
    owner: str = ""
    interface: str = ""
    sign_off: str = ""
    steps: list = field(default_factory=list)
    flow_type: Optional[str] = None
    sop_link: Optional[str] = None
    remarks: Optional[str] = None
    remarks_details: list = field(default_factory=list)
    assignee_notes: Optional[str] = None
    source_document: str = ""


@dataclass
class Section:
    """A labeled group of activities within a journey phase (e.g., 'B. Sales Planning')."""
    label: str = ""
    activities: list = field(default_factory=list)


@dataclass
class Journey:
    """A top-level journey phase: Pre-Order, Order, or Post Order."""
    name: str = ""
    sections: list = field(default_factory=list)


@dataclass
class SOPDocument:
    """Intermediate representation from parsing a single source document."""
    filename: str = ""
    title: str = ""
    doc_type: str = ""  # ops_sop, workflow_doc, technical_doc, field_spec, demo_doc
    purpose: Optional[str] = None
    stakeholders: dict = field(default_factory=dict)
    steps: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    headings: list = field(default_factory=list)
    raw_text: str = ""
    sign_off_info: Optional[str] = None
    escalation_info: Optional[str] = None
