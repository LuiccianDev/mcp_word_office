"""FastMCP tools for document links and bookmarks manipulation."""

from typing import Any

from mcp_word.core import (
    core_add_hyperlink,
    core_add_bookmark,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.validation.document_validators import validate_docx_write


@validate_docx_write("file_path")
async def add_hyperlink(
    file_path: str, paragraph_index: int, text: str, url: str | None = None, bookmark: str | None = None
) -> dict[str, Any]:
    """Add a hyperlink to a specific paragraph."""
    try:
        with document_context(file_path, mode="write") as doc:
            core_add_hyperlink(doc, paragraph_index, text, url=url, bookmark=bookmark)
        
        target = url if url else f"bookmark: {bookmark}"
        return {
            "status": "success",
            "message": f"Hyperlink to {target} added to paragraph {paragraph_index} in {file_path}",
            "file_path": str(file_path)
        }
    except (ValueError, IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add hyperlink: {str(error)}"),
            filename=file_path,
            operation="add hyperlink",
        )


@validate_docx_write("file_path")
async def add_bookmark(file_path: str, paragraph_index: int, name: str) -> dict[str, Any]:
    """Add a bookmark to a specific paragraph to enable internal linking."""
    try:
        with document_context(file_path, mode="write") as doc:
            core_add_bookmark(doc, paragraph_index, name)
        
        return {
            "status": "success",
            "message": f"Bookmark '{name}' added to paragraph {paragraph_index} in {file_path}",
            "file_path": str(file_path)
        }
    except (IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add bookmark: {str(error)}"),
            filename=file_path,
            operation="add bookmark",
        )
