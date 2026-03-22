"""
Register all tools with the MCP server.
"""

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from mcp_word.tools import (
    content_tools,
    document_tools,
    extended_document_tools,
    footnote_tools,
    format_tools,
    header_footer_tools,
    link_tools,
    property_tools,
    protection_tools,
)


def register_all_tools(mcp: FastMCP) -> None:
    """Register all tools with the MCP server."""

    read_only = ToolAnnotations(readOnlyHint=True)
    write = ToolAnnotations(readOnlyHint=False)
    destructive = ToolAnnotations(readOnlyHint=False, destructiveHint=True)

    def register_read_only(tool_fn: object) -> None:
        mcp.tool(annotations=read_only)(tool_fn)

    def register_write(tool_fn: object) -> None:
        mcp.tool(annotations=write)(tool_fn)

    def register_destructive(tool_fn: object) -> None:
        mcp.tool(annotations=destructive)(tool_fn)

    # DOCUMENT TOOLS
    # word_create_document - Create a new Word document
    register_write(document_tools.create_document)
    # word_copy_document - Create a copy of a Word document
    register_write(document_tools.copy_document)
    # word_get_document_info - Get information about a Word document
    register_read_only(document_tools.get_document_info)
    # word_get_document_text - Extract all text from a Word document
    register_read_only(document_tools.get_document_text)
    # word_get_document_outline - Get the structure of a Word document
    register_read_only(document_tools.get_document_outline)
    # word_list_documents - List all Word documents (paginated)
    register_read_only(document_tools.list_available_documents)
    # word_merge_documents - Merge multiple Word documents
    register_write(document_tools.merge_documents)

    # CONTENT TOOLS
    # word_add_heading - Add a heading to a Word document
    register_write(content_tools.add_heading)
    # word_add_paragraph - Add a paragraph to a Word document
    register_write(content_tools.add_paragraph)
    # word_add_picture - Add an image to a Word document
    register_write(content_tools.add_picture)
    # word_add_table - Add a table to a Word document
    register_write(content_tools.add_table)
    # word_add_page_break - Add a page break to a document
    register_write(content_tools.add_page_break)
    # word_delete_paragraph - Delete a paragraph from a document
    register_destructive(content_tools.delete_paragraph)
    # word_search_and_replace - Search and replace text
    register_destructive(content_tools.search_and_replace)
    # word_add_table_of_contents - Add a table of contents to a document
    register_write(content_tools.add_table_of_contents)

    # FORMAT TOOLS
    # word_format_text - Format specific text within a paragraph
    register_write(format_tools.format_text)
    # word_create_custom_style - Create a custom style
    register_write(format_tools.create_custom_style)
    # word_format_table - Format a table with borders and shading
    register_write(format_tools.format_table)


    # PROTECTION TOOLS
    # word_protect_document - Add password protection to a document
    register_write(protection_tools.protect_document)
    # word_unprotect_document - Remove password protection
    register_write(protection_tools.unprotect_document)
    # word_add_restricted_editing - Add restricted editing
    register_write(protection_tools.add_restricted_editing)
    # word_add_digital_signature - Add a digital signature
    register_write(protection_tools.add_digital_signature)
    # word_verify_document - Verify document protection/signature
    register_read_only(protection_tools.verify_document)


    # FOOTNOTE TOOLS
    # word_add_footnote - Add a footnote to a document
    register_write(footnote_tools.add_footnote_to_document)
    # word_add_endnote - Add an endnote to a document
    register_write(footnote_tools.add_endnote_to_document)
    # word_convert_footnotes - Convert footnotes to endnotes
    register_write(footnote_tools.convert_footnotes_to_endnotes_in_document)
    # word_customize_footnote_style - Customize footnote styling
    register_write(footnote_tools.customize_footnote_style)


    # EXTENDED DOCUMENT TOOLS
    # word_get_paragraph_text - Get text from a specific paragraph
    register_read_only(extended_document_tools.get_paragraph_text_from_document)
    # word_find_text - Find text in a document
    register_read_only(extended_document_tools.find_text_in_document)
    # word_convert_to_pdf - Convert Word document to PDF
    register_write(extended_document_tools.convert_to_pdf)

    # HEADER FOOTER TOOL 
    register_write(header_footer_tools.add_header)
    register_write(header_footer_tools.add_footer)

    # LINK TOOLS
    register_write(link_tools.add_hyperlink)
    register_write(link_tools.add_bookmark)

    # PROPERTY AND LAYOUT TOOLS
    register_read_only(property_tools.get_core_properties)
    register_write(property_tools.set_core_properties)
    register_write(property_tools.set_page_layout)
