"""
Register all tools with the MCP server.
"""

from collections.abc import Callable
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from word_mcp.tools import (
    batch_tools,
    content_tools,
    document_tools,
    footnote_tools,
    format_tools,
    property_tools,
    protection_tools,
)


def register_all_tools(mcp: FastMCP) -> None:
    """Register all tools with the MCP server."""

    read_only = ToolAnnotations(readOnlyHint=True)
    write = ToolAnnotations(readOnlyHint=False)
    destructive = ToolAnnotations(readOnlyHint=False, destructiveHint=True)

    def _register(tool_fn: Callable[..., Any], annotations: ToolAnnotations) -> None:
        mcp.tool(annotations=annotations)(tool_fn)

    def register_read_only(tool_fn: Callable[..., Any]) -> None:
        _register(tool_fn, read_only)

    def register_write(tool_fn: Callable[..., Any]) -> None:
        _register(tool_fn, write)

    def register_destructive(tool_fn: Callable[..., Any]) -> None:
        _register(tool_fn, destructive)

    # BATCH & COMPOSITE TOOLS (High-efficiency, token-saving single-pass tools)
    # build_document - Build complete document with typography, spacing and elements
    register_write(batch_tools.build_document)
    # create_from_markdown - Create styled document directly from Markdown
    register_write(batch_tools.create_from_markdown)
    # batch_add_content - Append multiple elements to existing document in one pass
    register_write(batch_tools.batch_add_content)

    # DOCUMENT TOOLS
    # create_document - Create a new Word document
    register_write(document_tools.create_document)
    # copy_document - Create a copy of a Word document
    register_write(document_tools.copy_document)
    # get_document_info - Get information about a Word document
    register_read_only(document_tools.get_document_info)
    # get_document_text - Extract all text from a Word document
    register_read_only(document_tools.get_document_text)
    # get_document_outline - Get the structure of a Word document
    register_read_only(document_tools.get_document_outline)
    # list_available_documents - List all Word documents (paginated)
    register_read_only(document_tools.list_available_documents)
    # merge_documents - Merge multiple Word documents
    register_write(document_tools.merge_documents)

    # CONTENT TOOLS
    # add_heading - Add a heading to a Word document
    register_write(content_tools.add_heading)
    # add_paragraph - Add a paragraph to a Word document
    register_write(content_tools.add_paragraph)
    # add_picture - Add an image to a Word document
    register_write(content_tools.add_picture)
    # add_table - Add a table to a Word document
    register_write(content_tools.add_table)
    # add_page_break - Add a page break to a document
    register_write(content_tools.add_page_break)
    # delete_paragraph - Delete a paragraph from a document
    register_destructive(content_tools.delete_paragraph)
    # search_and_replace - Search and replace text
    register_destructive(content_tools.search_and_replace)
    # add_table_of_contents - Add a table of contents to a document
    register_write(content_tools.add_table_of_contents)

    # FORMAT TOOLS
    # format_text - Format specific text within a paragraph
    register_write(format_tools.format_text)
    # create_custom_style - Create a custom style
    register_write(format_tools.create_custom_style)
    # format_table - Format a table with borders and shading
    register_write(format_tools.format_table)
    # add_header - Add or update a document header
    register_write(format_tools.add_header)
    # add_footer - Add or update a document footer
    register_write(format_tools.add_footer)

    # PROTECTION TOOLS
    # protect_document - Add password protection to a document
    register_write(protection_tools.protect_document)
    # unprotect_document - Remove password protection
    register_write(protection_tools.unprotect_document)
    # add_restricted_editing - Add restricted editing
    register_write(protection_tools.add_restricted_editing)
    # add_digital_signature - Add a digital signature
    register_write(protection_tools.add_digital_signature)
    # verify_document - Verify document protection/signature
    register_read_only(protection_tools.verify_document)

    # FOOTNOTE TOOLS
    # add_footnote_to_document - Add a footnote to a document
    register_write(footnote_tools.add_footnote_to_document)
    # add_endnote_to_document - Add an endnote to a document
    register_write(footnote_tools.add_endnote_to_document)
    # convert_footnotes_to_endnotes_in_document - Convert footnotes to endnotes
    register_write(footnote_tools.convert_footnotes_to_endnotes_in_document)
    # customize_footnote_style - Customize footnote styling
    register_write(footnote_tools.customize_footnote_style)

    # EXTENDED DOCUMENT TOOLS
    # get_paragraph_text_from_document - Get text from a specific paragraph
    register_read_only(document_tools.get_paragraph_text_from_document)
    # find_text_in_document - Find text in a document
    register_read_only(document_tools.find_text_in_document)
    # convert_to_pdf - Convert Word document to PDF
    register_write(document_tools.convert_to_pdf)

    # LINK TOOLS
    # add_hyperlink - Insert a hyperlink into a document
    register_write(content_tools.add_hyperlink)
    # add_bookmark - Insert a bookmark anchor in a document
    register_write(content_tools.add_bookmark)

    # PROPERTY AND LAYOUT TOOLS
    register_read_only(property_tools.get_core_properties)
    register_write(property_tools.set_core_properties)
    register_write(property_tools.set_page_layout)
