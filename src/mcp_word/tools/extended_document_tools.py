"""
Extended document tools for Word Document Server.

These tools provide enhanced document content extraction and search capabilities.
"""

from typing import Any

from mcp_word.core import (
    core_convert_to_pdf,
    core_find_text,
    core_get_paragraph_text,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.validation.document_validators import (
    validate_docx_read,
    validate_docx_write,
)


@validate_docx_read("filename")
async def get_paragraph_text_from_document(
    filename: str, paragraph_index: int
) -> dict[str, Any]:
    """Get text from a specific paragraph in a Word document."""
    try:
        with document_context(filename, mode="read") as doc:
            result = core_get_paragraph_text(doc, paragraph_index)
        
        return {
            "status": "success",
            "paragraph": result
        }
    except (IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to get paragraph text: {str(error)}"),
            filename=filename,
            operation="get paragraph text",
        )


@validate_docx_read("filename")
async def find_text_in_document(
    filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> dict[str, Any]:
    """Find occurrences of specific text in a Word document."""
    try:
        with document_context(filename, mode="read") as doc:
            result = core_find_text(doc, text_to_find, match_case, whole_word)
        
        return {
            "status": "success",
            "search_results": result
        }
    except (ValueError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to search for text: {str(error)}"),
            filename=filename,
            operation="find text",
        )


@validate_docx_write("filename")
async def convert_to_pdf(
    filename: str, output_filename: str | None = None
) -> dict[str, Any]:
    """Convert a Word document to PDF format."""
    try:
        # Conversion doesn't use document_context because it's usually 
        # an external process (Word or LibreOffice)
        pdf_path = core_convert_to_pdf(filename, output_filename)
        return {
            "status": "success",
            "message": "Document successfully converted to PDF",
            "pdf_path": pdf_path,
        }
    except Exception as e:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to convert document to PDF: {str(e)}"),
            filename=filename,
            operation="convert to PDF",
        )
