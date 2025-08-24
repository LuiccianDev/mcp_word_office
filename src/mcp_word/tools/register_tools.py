"""
Register all tools with the MCP server.
"""

# modulos de terceros
from mcp.server.fastmcp import FastMCP

# modulos propios
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
    # Document tools (create, copy, info, etc.)
    mcp.tool()(document_tools.create_document)
    mcp.tool()(document_tools.copy_document)
    mcp.tool()(document_tools.get_document_info)
    mcp.tool()(document_tools.get_document_text)
    mcp.tool()(document_tools.get_document_outline)
    mcp.tool()(document_tools.list_available_documents)

    # Content tools (paragraphs, headings, tables, etc.)
    mcp.tool()(content_tools.add_paragraph)
    mcp.tool()(content_tools.add_heading)
    mcp.tool()(content_tools.add_picture)
    mcp.tool()(content_tools.add_table)
    mcp.tool()(content_tools.add_page_break)
    mcp.tool()(content_tools.delete_paragraph)
    mcp.tool()(content_tools.search_and_replace)

    # Format tools (styling, text formatting, etc.)
    mcp.tool()(format_tools.create_custom_style)
    mcp.tool()(format_tools.format_text)
    mcp.tool()(format_tools.format_table)

    # Protection tools
    mcp.tool()(protection_tools.protect_document)
    mcp.tool()(protection_tools.unprotect_document)

    # Footnote tools
    mcp.tool()(footnote_tools.add_footnote_to_document)
    mcp.tool()(footnote_tools.add_endnote_to_document)
    mcp.tool()(footnote_tools.convert_footnotes_to_endnotes_in_document)
    mcp.tool()(footnote_tools.customize_footnote_style)

    # Extended document tools
    mcp.tool()(extended_document_tools.get_paragraph_text_from_document)
    mcp.tool()(extended_document_tools.find_text_in_document)
    mcp.tool()(extended_document_tools.convert_to_pdf)
