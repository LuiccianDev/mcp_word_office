"""FastMCP tools for document headers and footers manipulation."""

from typing import Any

from mcp_word.core import (
    core_set_section_header,
    core_set_section_footer,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.validation.document_validators import validate_docx_write


@validate_docx_write("file_path")
async def add_header(file_path: str, text: str, section_index: int = 0) -> dict[str, Any]:
    """Add or update the primary header of a specific section in the document."""
    try:
        with document_context(file_path, mode="write") as doc:
            core_set_section_header(doc, text, section_idx=section_index)
        
        return {
            "status": "success",
            "message": f"Header successfully updated in section {section_index} of {file_path}",
            "file_path": str(file_path)
        }
    except (IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add header: {str(error)}"),
            filename=file_path,
            operation="add header",
        )


@validate_docx_write("file_path")
async def add_footer(file_path: str, text: str, section_index: int = 0) -> dict[str, Any]:
    """Add or update the primary footer of a specific section in the document."""
    try:
        with document_context(file_path, mode="write") as doc:
            core_set_section_footer(doc, text, section_idx=section_index)

        return {
            "status": "success",
            "message": f"Footer successfully updated in section {section_index} of {file_path}",
            "file_path": str(file_path)
        }
    except (IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add footer: {str(error)}"),
            filename=file_path,
            operation="add footer",
        )
