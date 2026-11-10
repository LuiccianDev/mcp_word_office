"""
Utility functions for the Word Document Server.
This package contains utility modules for file operations and document handling.
"""

from word_mcp.utils.document_utils import (
    create_document_copy,
    ensure_docx_extension,
    extract_document_text,
    get_document_structure,
)

__all__ = [
    "create_document_copy",
    "ensure_docx_extension",
    "extract_document_text",
    "get_document_structure",
]
