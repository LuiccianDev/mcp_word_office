"""
Core functionality for the Word Document Server.

This package contains the core functionality modules used by the Word Document Server.
"""

from office_mcp.core.batch import (
    core_build_document,
    core_create_from_markdown,
)
from office_mcp.core.content import (
    core_add_heading,
    core_add_paragraph,
    core_add_picture,
    core_add_page_break,
    core_delete_paragraph,
    core_add_table_of_contents,
    core_find_and_replace_text,
)
from office_mcp.core.document import (
    core_create_document,
    core_copy_document,
    core_merge_documents,
    core_convert_to_pdf,
    core_find_text,
    core_get_paragraph_text,
)
from office_mcp.core.footnotes import (
    core_add_endnote,
    core_add_footnote,
    core_convert_footnotes_to_endnotes,
    core_customize_footnote_formatting,
    find_footnote_references,
    get_format_symbols,
)
from office_mcp.core.formatting import core_format_text
from office_mcp.core.headers_footers import (
    core_set_section_header,
    core_set_section_footer,
)
from office_mcp.core.links import core_add_hyperlink, core_add_bookmark
from office_mcp.core.properties import (
    core_get_core_properties,
    core_set_core_properties,
    core_set_page_layout,
)
from office_mcp.core.protection import (
    core_add_protection_info,
    core_create_signature_info,
    core_is_section_editable,
    core_verify_document_protection,
    core_verify_signature,
    core_protect_document,
    core_unprotect_document,
)
from office_mcp.core.styles import create_style, ensure_heading_style, ensure_table_style
from office_mcp.core.tables import (
    apply_table_style,
    copy_table,
    set_cell_border,
    core_add_table,
)


__all__ = [
    # batch / composite
    "core_build_document",
    "core_create_from_markdown",
    # content
    "core_add_heading",
    "core_add_paragraph",
    "core_add_picture",
    "core_add_page_break",
    "core_delete_paragraph",
    "core_add_table_of_contents",
    "core_find_and_replace_text",
    # documents
    "core_create_document",
    "core_copy_document",
    "core_merge_documents",
    "core_convert_to_pdf",
    "core_find_text",
    "core_get_paragraph_text",
    # footnotes
    "core_add_endnote",
    "core_add_footnote",
    "core_convert_footnotes_to_endnotes",
    "core_customize_footnote_formatting",
    "find_footnote_references",
    "get_format_symbols",
    # formatting
    "core_format_text",
    # headers_footers
    "core_set_section_header",
    "core_set_section_footer",
    # links
    "core_add_hyperlink",
    "core_add_bookmark",
    # properties
    "core_get_core_properties",
    "core_set_core_properties",
    "core_set_page_layout",
    # protection
    "core_add_protection_info",
    "core_create_signature_info",
    "core_is_section_editable",
    "core_verify_document_protection",
    "core_verify_signature",
    "core_protect_document",
    "core_unprotect_document",
    # styles
    "create_style",
    "ensure_heading_style",
    "ensure_table_style",
    # tables
    "apply_table_style",
    "copy_table",
    "set_cell_border",
    "core_add_table",
]
