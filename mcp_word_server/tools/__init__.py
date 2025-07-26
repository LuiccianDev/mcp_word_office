"""
MCP tool implementations for the Word Document Server.

This package contains the MCP tool implementations that expose functionality
to clients through the Model Context Protocol.
"""
from mcp_word_server.tools.register_tools import register_all_tools

__all__ = [
    "register_all_tools",
]

# Document tools
# Content tools
""" from mcp_word_server.tools.content_tools import (
    add_heading,
    add_page_break,
    add_paragraph,
    add_picture,
    add_table,
    add_table_of_contents,
    delete_paragraph,
    search_and_replace,
)
from mcp_word_server.tools.document_tools import (
    copy_document,
    create_document,
    get_document_info,
    get_document_outline,
    get_document_text,
    list_available_documents,
    merge_documents,
)

# Footnote tools
from mcp_word_server.tools.footnote_tools import (
    add_endnote_to_document,
    add_footnote_to_document,
    convert_footnotes_to_endnotes_in_document,
    customize_footnote_style,
)

# Format tools
from mcp_word_server.tools.format_tools import (
    create_custom_style,
    format_table,
    format_text,
)

# Protection tools
from mcp_word_server.tools.protection_tools import (
    add_digital_signature,
    add_restricted_editing,
    protect_document,
    verify_document,
) """
