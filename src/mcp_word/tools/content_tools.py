"""Content tools for Word Document Server.

This module provides functions to add and manipulate various types of content
in Word documents, including headings, paragraphs, tables, images, and page breaks.
"""

import os
from typing import Any

from mcp_word.core import (
    core_add_heading,
    core_add_paragraph,
    core_add_picture,
    core_add_page_break,
    core_delete_paragraph,
    core_add_table_of_contents,
    core_add_table,
    core_find_and_replace_text,
)
from mcp_word.core.document_context import document_context
from mcp_word.exception import (
    DocumentProcessingError,
    ExceptionTool,
)
from mcp_word.models.response_models import (
    HeadingResult,
    ParagraphResult,
    PictureResult,
    OperationResult,
    TableResult,
    TableOfContentsResult,
)
from mcp_word.validation.document_validators import validate_docx_write


@validate_docx_write("filename")
async def add_heading(filename: str, text: str, level: int = 1) -> dict[str, Any]:
    """Add a heading to a Word document."""
    try:
        heading_level: int = _validate_heading_level(level)
        with document_context(filename, mode="write") as doc:
            result = core_add_heading(doc, text, level=heading_level)
            
        return HeadingResult(
            status="success",
            filename=filename,
            heading_text=text,
            heading_level=heading_level,
            message=f"Heading '{text}' (level {heading_level}) added to {filename}"
        ).model_dump()
        
    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError) 
            else DocumentProcessingError(f"Failed to add heading: {str(error)}"),
            filename=filename,
            operation="add heading",
        )


@validate_docx_write("filename")
async def add_paragraph(
    filename: str, text: str, style: str | None = None
) -> dict[str, Any]:
    """Add a paragraph to a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            result = core_add_paragraph(doc, text, style=style)
            
        return ParagraphResult(
            status="success",
            filename=filename,
            paragraph_text=text,
            style_applied=result["style_applied"],
            message=f"Paragraph added to {filename}"
        ).model_dump()

    except (OSError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add paragraph: {str(error)}"),
            filename=filename,
            operation="add paragraph",
        )


@validate_docx_write("filename")
async def add_picture(
    filename: str, image_path: str, width: float | None = None
) -> dict[str, Any]:
    """Add an image to a Word document."""
    try:
        abs_image_path = _validate_image_file(image_path)
        if width is not None and width <= 0:
            raise ValueError(f"Width must be positive, got {width}")

        with document_context(filename, mode="write") as doc:
            result = core_add_picture(doc, abs_image_path, width=width)

        return PictureResult(
            status="success",
            filename=filename,
            image_name=result["image_name"],
            image_path=abs_image_path,
            width_inches=width,
            message=f"Picture '{result['image_name']}' added to {filename}",
        ).model_dump()

    except (OSError, FileNotFoundError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add picture: {error}"),
            filename=filename,
            operation="add picture",
        )


@validate_docx_write("filename")
async def add_table(
    filename: str, rows: int, cols: int, data: list[list[Any]] | None = None
) -> dict[str, Any]:
    """Add a table to a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            core_add_table(doc, rows, cols, data)

        return TableResult(
            status="success",
            filename=filename,
            rows=rows,
            columns=cols,
            data_rows=len(data) if data else 0,
            message=f"Table with {rows} rows and {cols} columns added to {filename}",
        ).model_dump()

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add table: {str(error)}"),
            filename=filename,
            operation="add table",
        )


@validate_docx_write("filename")
async def add_table_of_contents(
    filename: str, title: str = "Table of Contents", max_level: int = 3
) -> dict[str, Any]:
    """Add a table of contents to a Word document."""
    try:
        with document_context(filename, mode="write") as doc:
            result = core_add_table_of_contents(doc, title=title, max_level=max_level)

        return TableOfContentsResult(
            status="success",
            filename=filename,
            title=title,
            entry_count=0,  # Field is created, entries depend on update
            max_level=max_level,
            message=result["message"],
        ).model_dump()

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add TOC: {str(error)}"),
            filename=filename,
            operation="add TOC",
        )


@validate_docx_write("filename")
async def add_page_break(filename: str) -> dict[str, Any]:
    """Add a page break to the document."""
    try:
        with document_context(filename, mode="write") as doc:
            core_add_page_break(doc)
            
        return OperationResult(
            status="success",
            message=f"Page break added to {filename}"
        ).model_dump()
        
    except (OSError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to add page break: {str(error)}"),
            filename=filename,
            operation="add page break",
        )


@validate_docx_write("filename")
async def delete_paragraph(filename: str, paragraph_index: int) -> dict[str, Any]:
    """Delete a paragraph from a document."""
    try:
        with document_context(filename, mode="write") as doc:
            core_delete_paragraph(doc, paragraph_index)

        return OperationResult(
            status="success",
            message=f"Paragraph at index {paragraph_index} deleted successfully."
        ).model_dump()

    except (OSError, IndexError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to delete paragraph: {str(error)}"),
            filename=filename,
            operation="delete paragraph",
        )


@validate_docx_write("filename")
async def search_and_replace(
    filename: str, find_text: str, replace_text: str
) -> dict[str, Any]:
    """Search for text and replace all occurrences in a Word document."""
    try:
        if not find_text:
            raise ValueError("Search text cannot be empty")

        with document_context(filename, mode="write") as doc:
            replacement_count = core_find_and_replace_text(doc, find_text, replace_text)
            
        if replacement_count > 0:
            return OperationResult(
                status="success",
                message=f"Replaced {replacement_count} occurrence(s) of '{find_text}' with '{replace_text}' in {filename}",
                details={"replacements_made": replacement_count}
            ).model_dump()

        return OperationResult(
            status="error",
            message=f"No occurrences of '{find_text}' found in {filename}"
        ).model_dump()

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to perform search and replace: {str(error)}"),
            filename=filename,
            operation="search and replace",
        )


def _validate_heading_level(level: int) -> int:
    """Validate and normalize heading level."""
    try:
        level_int = int(level)
        if not 1 <= level_int <= 9:
            raise ValueError(f"Heading level must be between 1 and 9, got {level_int}")
        return level_int
    except (ValueError, TypeError) as error:
        raise ValueError(
            f"Invalid heading level: {level}. Must be an integer between 1 and 9."
        ) from error


def _validate_image_file(image_path: str) -> str:
    """Validate an image file and return its absolute path."""
    abs_image_path = os.path.abspath(image_path)
    if not os.path.exists(abs_image_path):
        raise FileNotFoundError(f"Image file not found: {abs_image_path}")
    return abs_image_path
