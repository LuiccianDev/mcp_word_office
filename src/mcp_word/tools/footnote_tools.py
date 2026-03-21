"""Footnote and endnote tools for Word Document Server.

These tools handle footnote and endnote functionality,
including adding, customizing, and converting between them.
"""

from typing import Any

from mcp_word.core import (
    core_add_footnote,
    core_add_endnote,
    core_convert_footnotes_to_endnotes,
    find_footnote_references,
    get_format_symbols,
    core_customize_footnote_formatting,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.models.response_models import (
    FootnoteResult,
    EndnoteResult,
    OperationResult,
)
from mcp_word.validation.document_validators import validate_docx_write


@validate_docx_write("filename")
async def add_footnote_to_document(
    filename: str, paragraph_index: int, footnote_text: str
) -> dict[str, Any]:
    """Add a footnote to a specific paragraph in a Word document."""
    try:
        p_index = int(paragraph_index)
        with document_context(filename, mode="write") as doc:
            result = core_add_footnote(doc, p_index, footnote_text)

        return FootnoteResult(
            status="success",
            filename=filename,
            footnote_id=result["footnote_id"],
            footnote_text=footnote_text,
            reference_position=f"paragraph {p_index}",
            message=f"Footnote added to paragraph {p_index} in {filename}",
        ).model_dump()

    except (ValueError, TypeError, IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add footnote: {str(error)}"),
            filename=filename,
            operation="add footnote",
        )


@validate_docx_write("filename")
async def add_endnote_to_document(
    filename: str, paragraph_index: int, endnote_text: str
) -> dict[str, Any]:
    """Add an endnote to a specific paragraph in a Word document."""
    try:
        p_index = int(paragraph_index)
        with document_context(filename, mode="write") as doc:
            result = core_add_endnote(doc, p_index, endnote_text)

        return EndnoteResult(
            status="success",
            filename=filename,
            endnote_id=result["endnote_id"],
            endnote_text=endnote_text,
            reference_position=f"paragraph {p_index}",
            message=f"Endnote added to paragraph {p_index} in {filename}",
        ).model_dump()

    except (ValueError, TypeError, IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add endnote: {str(error)}"),
            filename=filename,
            operation="add endnote",
        )


@validate_docx_write("filename")
async def convert_footnotes_to_endnotes_in_document(filename: str) -> dict[str, Any]:
    """Convert all footnotes to endnotes in a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            count = core_convert_footnotes_to_endnotes(doc)

        return OperationResult(
            status="success",
            message=f"Converted {count} footnotes to endnotes in {filename}",
            details={"count": count}
        ).model_dump()

    except (DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to convert footnotes: {str(error)}"),
            filename=filename,
            operation="convert footnotes to endnotes",
        )


@validate_docx_write("filename")
async def customize_footnote_style(
    filename: str,
    numbering_format: str = "1, 2, 3",
    start_number: int = 1,
    font_name: str | None = None,
    font_size: int | None = None,
) -> dict[str, Any]:
    """Customize footnote numbering and formatting in a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            # Find all existing footnote references
            footnote_refs = find_footnote_references(doc)

            # Generate format symbols
            format_symbols = get_format_symbols(
                numbering_format, len(footnote_refs) + start_number
            )

            # Apply custom formatting
            count = core_customize_footnote_formatting(
                doc, footnote_refs, format_symbols, start_number
            )

        return OperationResult(
            status="success",
            message=f"Footnote style and numbering customized in {filename}",
            details={"count": count}
        ).model_dump()

    except (DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to customize footnote style: {str(error)}"),
            filename=filename,
            operation="customize footnote style",
        )
