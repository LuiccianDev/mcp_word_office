"""Document creation and manipulation tools for Word Document Server.

This module provides functions for creating, reading, and manipulating
Word documents with security checks for file system access.
"""

import os
from typing import Any

from mcp_word.core import (
    core_create_document,
    core_copy_document,
    core_merge_documents,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
    FileOperationError,
)
from mcp_word.utils.document_utils import (
    get_document_properties,
    get_document_structure,
    extract_document_text,
)
from mcp_word.utils.security import (
    get_allowed_directories as _get_allowed_directories,
    is_path_in_allowed_directories as _is_path_in_allowed_directories,
)
from mcp_word.validation.document_validators import (
    validate_docx_read,
    validate_docx_write,
)


async def create_document(
    filename: str, title: str | None = None, author: str | None = None
) -> dict[str, Any]:
    """Create a new Word document with optional metadata."""
    is_allowed, error_message = _is_path_in_allowed_directories(filename)
    if not is_allowed:
        return {
            "status": "error",
            "message": f"Cannot create document: {error_message}",
        }

    directory = os.path.dirname(filename)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as e:
            return ExceptionTool.handle_error(
                FileOperationError(f"Cannot create directory '{directory}': {str(e)}"),
                filename=filename,
                operation="create directory",
            )

    try:
        path = core_create_document(filename, title=title, author=author)
        return {
            "status": "success",
            "message": f"Document {path} created successfully",
        }
    except (OSError, ValueError) as e:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to create document: {str(e)}"),
            filename=filename,
            operation="create document",
        )


@validate_docx_read("filename")
async def get_document_info(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Get information about a Word document."""
    try:
        # Properties usually doesn't need full doc open via DocumentContext 
        # but for consistency we could, however document_utils.get_document_properties 
        # might be optimized. Let's stick to core if we had one for properties.
        # We have core_get_core_properties but it takes doc.
        with document_context(filename, mode="read") as doc:
            from mcp_word.core import core_get_core_properties
            properties = core_get_core_properties(doc)
            # Add some extra info
            properties["filename"] = os.path.basename(filename)
            properties["path"] = os.path.abspath(filename)
            properties["size_kb"] = round(os.path.getsize(filename) / 1024, 2)
            
        return {
            "status": "success",
            "properties": properties,
            "response_format": response_format
        }
    except (OSError, ValueError, DocumentProcessingError) as e:
        return ExceptionTool.handle_error(
            e if isinstance(e, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to get document info: {str(e)}"),
            filename=filename,
            operation="get document info",
        )


@validate_docx_read("filename")
async def get_document_text(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Extract all text from a Word document."""
    try:
        # extract_document_text currently takes path.
        result = extract_document_text(filename)
        return result
    except (OSError, ValueError) as e:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to get document text: {str(e)}"),
            filename=filename,
            operation="get document text",
        )


@validate_docx_read("filename")
async def get_document_outline(
    filename: str, response_format: str = "markdown"
) -> dict[str, Any]:
    """Get the structure of a Word document."""
    try:
        structure = get_document_structure(filename)
        return structure
    except (OSError, ValueError) as e:
        return ExceptionTool.handle_error(
            DocumentProcessingError(f"Failed to get document outline: {str(e)}"),
            filename=filename,
            operation="get document outline",
        )


async def list_available_documents(
    directory: str | None = None,
    page: int = 1,
    page_size: int = 20,
    response_format: str = "markdown",
) -> dict[str, Any]:
    """List all .docx files in the allowed directories with pagination support."""
    try:
        # Pagination and listing is more about filesystem than doc structure
        if page < 1:
            page = 1
        if page_size < 1:
            page_size = 20
        if page_size > 100:
            page_size = 100

        search_directories = _get_allowed_directories()
        if not search_directories:
            return {
                "status": "error",
                "message": "No accessible directories found in MCP configuration.",
                "documents": [],
                "pagination": {"page": page, "page_size": page_size, "total": 0, "has_more": False}
            }

        all_documents = []
        for search_dir in search_directories:
            if not os.path.exists(search_dir):
                continue

            docx_files = [f for f in os.listdir(search_dir) if f.lower().endswith(".docx") and not f.startswith("~$")]
            for file in sorted(docx_files):
                file_path = os.path.join(search_dir, file)
                size_kb = os.path.getsize(file_path) / 1024
                all_documents.append({
                    "name": file,
                    "path": os.path.abspath(file_path),
                    "size_kb": round(size_kb, 2),
                    "source_directory": search_dir,
                })

        total_documents = len(all_documents)
        offset = (page - 1) * page_size
        paginated_documents = all_documents[offset : offset + page_size]
        has_more = offset + page_size < total_documents

        return {
            "status": "success",
            "message": f"Found {total_documents} Word document(s). Showing page {page}.",
            "documents": paginated_documents,
            "pagination": {
                "page": page,
                "page_size": page_size,
                "total": total_documents,
                "has_more": has_more,
            },
            "response_format": response_format,
        }
    except Exception as e:
        return ExceptionTool.handle_error(
            FileOperationError(f"Failed to list documents: {str(e)}"),
            operation="list documents",
        )


@validate_docx_read("source_filename")
async def copy_document(
    source_filename: str, destination_filename: str | None = None
) -> dict[str, Any]:
    """Create a copy of a Word document."""
    try:
        success, message, dest_path = core_copy_document(source_filename, destination_filename)
        if success:
            return {
                "status": "success",
                "message": message,
                "destination_path": dest_path
            }
        raise FileOperationError(message)
    except (OSError, FileOperationError) as e:
        return ExceptionTool.handle_error(
            e if isinstance(e, FileOperationError)
            else FileOperationError(f"Failed to copy document: {str(e)}"),
            filename=source_filename,
            operation="copy document",
        )


@validate_docx_write("target_filename")
async def merge_documents(
    target_filename: str,
    source_filenames: list[str],
    add_page_breaks: bool = True,
    response_format: str = "markdown",
) -> dict[str, Any]:
    """Merge multiple Word documents into a single document."""
    try:
        path = core_merge_documents(target_filename, source_filenames, add_page_breaks)
        return {
            "status": "success",
            "message": f"Successfully merged {len(source_filenames)} documents into {path}",
            "file_path": path
        }
    except (OSError, DocumentProcessingError) as e:
        return ExceptionTool.handle_error(
            e if isinstance(e, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to merge documents: {str(e)}"),
            filename=target_filename,
            operation="merge documents",
        )
