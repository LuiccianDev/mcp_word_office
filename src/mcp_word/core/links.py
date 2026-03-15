"""
Link and bookmark operations for Word Document Server.

This module provides functionality for adding hyperlinks (internal and external)
and bookmarks to document paragraphs.
"""

from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def add_bookmark(paragraph: Paragraph, name: str) -> None:
    """Add a bookmark to the end of a paragraph.

    Args:
        paragraph: The paragraph to append the bookmark to.
        name: The name of the bookmark. Must be unique in the document.
              Should not contain spaces.
    """
    # Clean the name to be a valid xml id (no spaces)
    name = name.replace(" ", "_")
    
    # Create the xml elements
    bookmark_start = OxmlElement("w:bookmarkStart")
    bookmark_start.set(qn("w:id"), "0")
    bookmark_start.set(qn("w:name"), name)

    bookmark_end = OxmlElement("w:bookmarkEnd")
    bookmark_end.set(qn("w:id"), "0")

    # Append to the paragraph element
    paragraph._p.append(bookmark_start)
    paragraph._p.append(bookmark_end)


def add_hyperlink(
    paragraph: Paragraph, text: str, url: str | None = None, bookmark: str | None = None
) -> None:
    """Add a hyperlink into a paragraph.

    Args:
        paragraph: The paragraph where the hyperlink will be added.
        text: The clickable text of the hyperlink.
        url: The external URL. Provide this OR bookmark, not both.
        bookmark: The internal bookmark name to link to.

    Raises:
        ValueError: If neither url nor bookmark is provided, or both are.
    """
    if bool(url) == bool(bookmark):
        raise ValueError("Must provide exactly one of 'url' or 'bookmark'")

    # Create the w:hyperlink tag
    hyperlink = OxmlElement("w:hyperlink")

    if url:
        # Create a relationship for external link
        part = paragraph.part
        r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
        hyperlink.set(qn("r:id"), r_id)
    elif bookmark:
        # Link to internal bookmark
        bookmark_cleaned = bookmark.replace(" ", "_")
        hyperlink.set(qn("w:anchor"), bookmark_cleaned)

    # Create a new run with the text
    new_run = Run(OxmlElement("w:r"), paragraph)
    new_run.text = text
    
    # Optional: styling
    new_run.font.color.rgb = None
    new_run.font.color.theme_color = 10 # Hyperlink theme color typically
    new_run.font.underline = True

    # Append the run xml to the hyperlink xml
    hyperlink.append(new_run._r)
    
    # Append the hyperlink to the paragraph xml
    paragraph._p.append(hyperlink)

