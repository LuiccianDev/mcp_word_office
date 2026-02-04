"""
Register all tools with the MCP server.
"""

from mcp.server.fastmcp import FastMCP

from mcp_word.tools import (
    content_tools,
    document_tools,
    extended_document_tools,
    footnote_tools,
    format_tools,
    protection_tools,
)


def register_all_tools(mcp: FastMCP) -> None:
    """Register all tools with the MCP server."""

    # =====================
    # DOCUMENT TOOLS
    # =====================

    # word_create_document - Create a new Word document
    mcp.tool()(document_tools.create_document)

    # word_copy_document - Create a copy of a Word document
    mcp.tool()(document_tools.copy_document)

    # word_get_document_info - Get information about a Word document
    mcp.tool()(document_tools.get_document_info)

    # word_get_document_text - Extract all text from a Word document
    mcp.tool()(document_tools.get_document_text)

    # word_get_document_outline - Get the structure of a Word document
    mcp.tool()(document_tools.get_document_outline)

    # word_list_documents - List all Word documents (paginated)
    mcp.tool()(document_tools.list_available_documents)

    # word_merge_documents - Merge multiple Word documents
    mcp.tool()(document_tools.merge_documents)

    # =====================
    # CONTENT TOOLS
    # =====================

    # word_add_heading - Add a heading to a Word document
    mcp.tool()(content_tools.add_heading)

    # word_add_paragraph - Add a paragraph to a Word document
    mcp.tool()(content_tools.add_paragraph)

    # word_add_picture - Add an image to a Word document
    mcp.tool()(content_tools.add_picture)

    # word_add_table - Add a table to a Word document
    mcp.tool()(content_tools.add_table)

    # word_add_page_break - Add a page break to a document
    mcp.tool()(content_tools.add_page_break)

    # word_delete_paragraph - Delete a paragraph from a document
    mcp.tool()(content_tools.delete_paragraph)

    # word_search_and_replace - Search and replace text
    mcp.tool()(content_tools.search_and_replace)

    # =====================
    # FORMAT TOOLS
    # =====================

    # word_format_text - Format specific text within a paragraph
    mcp.tool()(format_tools.format_text)

    # word_create_custom_style - Create a custom style
    mcp.tool()(format_tools.create_custom_style)

    # word_format_table - Format a table with borders and shading
    mcp.tool()(format_tools.format_table)

    # =====================
    # PROTECTION TOOLS
    # =====================

    # word_protect_document - Add password protection to a document
    mcp.tool()(protection_tools.protect_document)

    # word_unprotect_document - Remove password protection
    mcp.tool()(protection_tools.unprotect_document)

    # word_add_restricted_editing - Add restricted editing
    mcp.tool()(protection_tools.add_restricted_editing)

    # word_add_digital_signature - Add a digital signature
    mcp.tool()(protection_tools.add_digital_signature)

    # word_verify_document - Verify document protection/signature
    mcp.tool()(protection_tools.verify_document)

    # =====================
    # FOOTNOTE TOOLS
    # =====================

    # word_add_footnote - Add a footnote to a document
    mcp.tool()(footnote_tools.add_footnote_to_document)

    # word_add_endnote - Add an endnote to a document
    mcp.tool()(footnote_tools.add_endnote_to_document)

    # word_convert_footnotes - Convert footnotes to endnotes
    mcp.tool()(footnote_tools.convert_footnotes_to_endnotes_in_document)

    # word_customize_footnote_style - Customize footnote styling
    mcp.tool()(footnote_tools.customize_footnote_style)

    # =====================
    # EXTENDED DOCUMENT TOOLS
    # =====================

    # word_get_paragraph_text - Get text from a specific paragraph
    mcp.tool()(extended_document_tools.get_paragraph_text_from_document)

    # word_find_text - Find text in a document
    mcp.tool()(extended_document_tools.find_text_in_document)

    # word_convert_to_pdf - Convert Word document to PDF
    mcp.tool()(extended_document_tools.convert_to_pdf)
