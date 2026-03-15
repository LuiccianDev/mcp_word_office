"""
Document properties and page layout operations for Word Document Server.

This module provides functionality for reading/writing core document properties
(author, title, etc.) and modifying section page layouts (orientation, size).
"""

from typing import Any

from docx.document import Document
from docx.enum.section import WD_ORIENTATION
from docx.shared import Inches


def get_core_properties(doc: Document) -> dict[str, str | None]:
    """Get the core properties of the document.

    Args:
        doc: The Word document object.

    Returns:
        A dictionary containing author, title, subject, keywords, etc.
    """
    props = doc.core_properties
    return {
        "author": props.author,
        "title": props.title,
        "subject": props.subject,
        "keywords": props.keywords,
        "comments": props.comments,
        "category": props.category,
        "last_modified_by": props.last_modified_by,
        "created": props.created.isoformat() if props.created else None,
        "modified": props.modified.isoformat() if props.modified else None,
    }


def set_core_properties(doc: Document, **kwargs: str) -> None:
    """Set the core properties of the document.

    Args:
        doc: The Word document object.
        **kwargs: The properties to set (author, title, subject, keywords, etc.)
    """
    props = doc.core_properties
    # Valid core_properties text attributes in python-docx
    valid_attrs = ["author", "category", "comments", "content_status", "identifier",
                   "keywords", "language", "last_modified_by", "subject", "title", "version"]
    
    for key, value in kwargs.items():
        if key in valid_attrs and value is not None:
            setattr(props, key, value)


def set_page_layout(
    doc: Document, section_idx: int = 0, orientation: str = "portrait"
) -> None:
    """Set the page layout orientation for a specific section.

    This adjusts both the orientation enum and the page dimensions.
    Assuming standard Letter size for simplicity (8.5 x 11 inches).

    Args:
        doc: The Word document object.
        section_idx: The index of the section to modify.
        orientation: 'portrait' or 'landscape'

    Raises:
        IndexError: If the section index is out of bounds.
        ValueError: If orientation is not valid.
    """
    sections = doc.sections
    if section_idx < 0 or section_idx >= len(sections):
        raise IndexError(f"Section index {section_idx} is out of bounds. Document has {len(sections)} sections.")

    section = sections[section_idx]
    
    orientation = orientation.lower()
    if orientation == "landscape":
        section.orientation = WD_ORIENTATION.LANDSCAPE
        # Standard Letter Landscape
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)
    elif orientation == "portrait":
        section.orientation = WD_ORIENTATION.PORTRAIT
        # Standard Letter Portrait
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
    else:
        raise ValueError("Orientation must be 'portrait' or 'landscape'")
