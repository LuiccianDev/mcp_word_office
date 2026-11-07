"""Extended document tools shim for backward compatibility.

All functions have been consolidated into word_mcp.tools.document_tools.
"""

from word_mcp.tools.document_tools import (
    convert_to_pdf,
    find_text_in_document,
    get_paragraph_text_from_document,
)

__all__ = [
    "convert_to_pdf",
    "find_text_in_document",
    "get_paragraph_text_from_document",
]
