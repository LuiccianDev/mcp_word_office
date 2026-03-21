"""FastMCP tools for document properties and page layout."""

from typing import Any

from mcp_word.core import (
    core_get_core_properties,
    core_set_core_properties,
    core_set_page_layout,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.validation.document_validators import validate_docx_read, validate_docx_write


@validate_docx_read("file_path")
async def get_core_properties(file_path: str) -> dict[str, Any]:
    """Get the core metadata properties of the document (author, title, etc.)."""
    try:
        with document_context(file_path, mode="read") as doc:
            props = core_get_core_properties(doc)
        
        return {
            "status": "success",
            "properties": props,
        }
    except (DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to get properties: {str(error)}"),
            filename=file_path,
            operation="get core properties",
        )


@validate_docx_write("file_path")
async def set_core_properties(
    file_path: str,
    author: str | None = None,
    title: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
    category: str | None = None
) -> dict[str, Any]:
    """Set the core metadata properties of the document."""
    try:
        # Build kwargs dictionary excluding None values
        kwargs = {}
        if author is not None: kwargs["author"] = author
        if title is not None: kwargs["title"] = title
        if subject is not None: kwargs["subject"] = subject
        if keywords is not None: kwargs["keywords"] = keywords
        if comments is not None: kwargs["comments"] = comments
        if category is not None: kwargs["category"] = category
        
        if not kwargs:
            return {"status": "error", "message": "No properties provided to update."}
            
        with document_context(file_path, mode="write") as doc:
            core_set_core_properties(doc, **kwargs)
        
        return {
            "status": "success",
            "message": f"Successfully updated properties: {list(kwargs.keys())}",
            "file_path": str(file_path)
        }
    except (DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to set properties: {str(error)}"),
            filename=file_path,
            operation="set core properties",
        )


@validate_docx_write("file_path")
async def set_page_layout(file_path: str, section_index: int = 0, orientation: str = "portrait") -> dict[str, Any]:
    """Set the page layout orientation (portrait or landscape) for a specific section."""
    try:
        with document_context(file_path, mode="write") as doc:
            core_set_page_layout(doc, section_idx=section_index, orientation=orientation)
        
        return {
            "status": "success",
            "message": f"Successfully set section {section_index} layout to {orientation}",
            "file_path": str(file_path)
        }
    except (ValueError, IndexError, DocumentProcessingError, OSError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to set page layout: {str(error)}"),
            filename=file_path,
            operation="set page layout",
        )
