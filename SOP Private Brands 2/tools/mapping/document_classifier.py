"""
Document classification and journey/section mapping.
Maps each source document to its journey phase, section label, and document type.
"""

import os
import re
import fnmatch

# Static configuration: maps filename patterns to journey/section/type
DOCUMENT_MAPPING = {
    "*JSW*Order*Logging*": {
        "journey": None,  # Auto-detect from document headings (PRE-ORDER / ORDER / POST-ORDER)
        "section": None,  # Auto-detect from table context
        "doc_type": "docx_sop",
        "activity_prefix": "Order Logging",
        "auto_phase": True,
    },
    "*Rebates*Schemes*Price*": {
        "journey": "Post Order",
        "section": "F. FINANCE & RECONCILIATION",
        "doc_type": "ops_sop",
        "activity_prefix": "PB Rebates, Schemes & Price Difference",
    },
    "*Influencer*Bill*Clearance*": {
        "journey": "Pre - Order",
        "section": "G. INFLUENCER MANAGEMENT",
        "doc_type": "ops_sop",
        "activity_prefix": "Influencer Meets Bill Clearance",
    },
    "*Influencer*Data*Management*": {
        "journey": "Pre - Order",
        "section": "G. INFLUENCER MANAGEMENT",
        "doc_type": "ops_sop",
        "activity_prefix": "Influencer Data Management",
    },
    "*Influencer*Scheme*Disbursal*": {
        "journey": "Pre - Order",
        "section": "G. INFLUENCER MANAGEMENT",
        "doc_type": "ops_sop",
        "activity_prefix": "Influencer Scheme Disbursal & Compliance",
    },
    "*Approval*Workflows*": {
        "journey": "Order",
        "section": "H. APPROVAL WORKFLOWS",
        "doc_type": "workflow_doc",
        "activity_prefix": "Approval Workflow",
        "expand_table_rows": True,
    },
    "*Opportunity*Workflow*": {
        "journey": "Order",
        "section": "C. OPPORTUNITY MANAGEMENT",
        "doc_type": "technical_doc",
        "activity_prefix": "PB Opportunity Workflow",
    },
    "*Data*Enrichment*Specification*": {
        "journey": "Pre - Order",
        "section": "A. DATA MANAGEMENT",
        "doc_type": "field_spec",
        "activity_prefix": "Data Enrichment Specification",
    },
    "PB*Retailer*Influencer*Data*Enrichment*Demo*": {
        "journey": "Pre - Order",
        "section": "A. DATA MANAGEMENT",
        "doc_type": "demo_doc",
        "activity_prefix": "PB Retailer & Influencer Data Enrichment Demo",
    },
    "PB*Retailer*Influencer*in*App*": {
        "journey": "Order",
        "section": "I. DIGITAL CHANNELS",
        "doc_type": "demo_doc",
        "activity_prefix": "PB Retailer/Influencer in App",
    },
    "*Category*Component*Updates*": {
        "journey": "Pre - Order",
        "section": "A. DATA MANAGEMENT",
        "doc_type": "demo_doc",
        "activity_prefix": "Category Component Updates",
    },
    "*SOP*Influencer*Meets*": {
        "journey": "Pre - Order",
        "section": "G. INFLUENCER MANAGEMENT",
        "doc_type": "ops_sop",
        "activity_prefix": "Influencer Meeting SOP",
        "default_owner": "Sales",
    },
    "JSW One TMT Quality Manual*": {
        "journey": "Order",
        "section": "J. QUALITY ASSURANCE",
        "doc_type": "quality_manual",
        "activity_prefix": "TMT Quality Manual",
        "multi_journey": {
            "Order": "J. QUALITY ASSURANCE (MANUFACTURING)",
            "Post Order": "J. QUALITY ASSURANCE (DISPATCH)",
        },
    },
    "GSB*Dealer*Sign*Board*": {
        "journey": "Pre - Order",
        "section": "K. MARKETING ACTIVITIES",
        "doc_type": "marketing_visual",
        "activity_prefix": "Dealer Sign Board (GSB)",
    },
    "Wall*Painting*": {
        "journey": "Pre - Order",
        "section": "K. MARKETING ACTIVITIES",
        "doc_type": "marketing_visual",
        "activity_prefix": "Wall Painting",
    },
    "*SOP*Marketing*Activities*": {
        "journey": "Pre - Order",
        "section": "K. MARKETING ACTIVITIES",
        "doc_type": "marketing_activities_xlsx",
        "activity_prefix": "Marketing Activity",
    },
}

# Ordered list of sections within each journey (controls output ordering)
SECTION_ORDER = {
    "Pre - Order": [
        "A. DATA MANAGEMENT",
        "B. SALES PLANNING",
        "B. GTM & SAP CODE INITIATION",
        "B. CREDIT REQUEST",
        "C. SAP CODE CREATION & PRODUCT MAPPING",
        "E. ORDER LOGGING (PRE-ORDER)",
        "G. INFLUENCER MANAGEMENT",
        "K. MARKETING ACTIVITIES",
    ],
    "Order": [
        "C. OPPORTUNITY MANAGEMENT",
        "D. PRICING, PO & APPROVALS",
        "E. ORDER CREATION & PROCESSING",
        "E. ORDER LOGGING (ORDER)",
        "H. APPROVAL WORKFLOWS",
        "I. DIGITAL CHANNELS",
        "J. QUALITY ASSURANCE (MANUFACTURING)",
    ],
    "Post Order": [
        "E. ORDER LOGGING (POST-ORDER)",
        "F. FINANCE & RECONCILIATION",
        "J. QUALITY ASSURANCE (DISPATCH)",
    ],
}


def classify_document(filename: str) -> dict | None:
    """Classify a document by matching its filename to DOCUMENT_MAPPING.

    Returns the mapping config dict or None if no match found.
    """
    # Normalize filename: replace + with spaces for matching
    normalized = filename.replace('+', ' ').replace('  ', ' ')

    for pattern, config in DOCUMENT_MAPPING.items():
        # Try matching against both original and normalized names
        pattern_normalized = pattern.replace('+', ' ')
        if fnmatch.fnmatch(normalized, pattern_normalized) or fnmatch.fnmatch(filename, pattern):
            return config

    # Fallback: try keyword matching
    name_lower = normalized.lower()
    for pattern, config in DOCUMENT_MAPPING.items():
        # Extract key words from pattern
        keywords = [w.strip('*').lower() for w in pattern.split('*') if w.strip('*')]
        if all(kw in name_lower for kw in keywords if len(kw) > 2):
            return config

    return None


def get_section_order(journey: str) -> list:
    """Get the ordered list of sections for a journey phase."""
    return SECTION_ORDER.get(journey, [])


def get_all_journeys_ordered() -> list:
    """Return journey phase names in display order."""
    return ["Pre - Order", "Order", "Post Order"]
