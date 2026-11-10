"""Document creation and manipulation tools for Word Document Server.

This module provides functions for creating, reading, and manipulating
Word documents with security checks for file system access.
"""

import os
from typing import Any

from word_mcp.core import (
    core_convert_to_pdf,
    core_copy_document,
    core_create_document,
    core_find_text,
    core_get_core_properties,
    core_get_paragraph_text,
    core_merge_documents,
)
from word_mcp.core.document_context import document_context
from word_mcp.exception import (
    DocumentProcessingError,
    ExceptionTool,
    FileOperationError,
)
from word_mcp.utils.document_utils import (
    extract_document_text,
    get_document_structure,
)
from word_mcp.utils.security import (
    get_allowed_directories as _get_allowed_directories,
    is_path_in_allowed_directories as _is_path_in_allowed_directories,
)
from word_mcp.validation.document_validators import (
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
        # Core properties are read from the open document context.
        with document_context(filename, mode="read") as doc:
            properties: dict[str, Any] = core_get_core_properties(doc)
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

