"""
Utility functions for the Word Document Server.
This package contains utility modules for file operations and document handling.
"""

from mcp_word.utils.document_utils import (
    create_document_copy,
    ensure_docx_extension,
    extract_document_text,
    find_and_replace_text,
    find_paragraph_by_text,
    get_document_properties,
    get_document_structure,
)
from mcp_word.utils.extended_document_utils import find_text, get_paragraph_text


__all__ = [
    # document_utils
    "create_document_copy",
    "ensure_docx_extension",
    "extract_document_text",
    "find_and_replace_text",
    "find_paragraph_by_text",
    "get_document_properties",
    "get_document_structure",
    # extended_document_utils
    "get_paragraph_text",
    "find_text",
]
