"""Batch and composite document building tools for Office MCP Server.

This module provides high-efficiency tools to build full documents or insert
multiple content blocks (headings, paragraphs, tables, lists, breaks) in a single pass.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from word_mcp.core.batch import (
    core_build_document,
    core_create_from_markdown,
    render_element,
)
from word_mcp.core.document_context import document_context
from word_mcp.exception import (
    DocumentProcessingError,
    ExceptionTool,
    FileOperationError,
)
from word_mcp.models.response_models import BuildDocumentResult
from word_mcp.utils.security import is_path_in_allowed_directories


def _dump(model: Any) -> dict[str, Any]:
    """Dump pydantic model supporting both v1 and v2."""
    if hasattr(model, "model_dump"):
        return model.model_dump()
    return model.dict()


def _resolve_and_check_path(filename: str | None, file_path: str | None) -> tuple[str | None, dict[str, Any] | None]:
    """Resolve target path and check security access."""
    target = filename or file_path
    if not target:
        return None, {
            "status": "error",
            "message": "Missing required parameter 'filename' (or 'file_path').",
        }

    is_allowed, error_msg = is_path_in_allowed_directories(target)
    if not is_allowed:
        return None, {
            "status": "error",
            "message": f"Security restriction: {error_msg}",
        }

    directory = os.path.dirname(target)
    if directory and not os.path.exists(directory):
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as e:
            return None, ExceptionTool.handle_error(
                FileOperationError(f"Cannot create directory '{directory}': {e}"),
                filename=target,
                operation="create directory",
            )

    return target, None


async def build_document(
    filename: str | None = None,
    file_path: str | None = None,
    title: str | None = None,
    author: str | None = None,
    preset: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    orientation: str | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
    elements: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Build a complete Word document with typography, margins, header/footer and content in a single pass.

    Args:
        filename: Path to the .docx file to create or overwrite (also accepts file_path).
        file_path: Alternative alias for filename.
        title: Document title metadata.
        author: Document author metadata.
        preset: Design preset ('modern', 'formal', 'academic', 'compact').
        font_name: Global font family (e.g. 'Calibri', 'Arial', 'Times New Roman').
        font_size: Global font size in points (e.g. 11.0).
        line_spacing: Paragraph line spacing (e.g. 1.15, 1.5, 2.0).
        space_after: Spacing after paragraphs in points (e.g. 6.0).
        margins: Page margins in cm {'top': 2.5, 'bottom': 2.5, 'left': 3.0, 'right': 2.5}.
        orientation: Page orientation ('portrait' or 'landscape').
        header_text: Header text displayed at the top of pages.
        footer_text: Footer text displayed at the bottom of pages.
        page_numbers: If True, adds dynamic PAGE numbers to the footer.
        elements: Sequential list of content blocks (headings, paragraphs, tables, lists, breaks).

    Returns:
        Compact summary of the built document.
    """
    target, err = _resolve_and_check_path(filename, file_path)
    if err or not target:
        return err or {}

    try:
        with document_context(target, mode="write") as doc:
            res = core_build_document(
                doc=doc,
                title=title,
                author=author,
                preset=preset,
                font_name=font_name,
                font_size=font_size,
                line_spacing=line_spacing,
                space_after=space_after,
                margins=margins,
                orientation=orientation,
                header_text=header_text,
                footer_text=footer_text,
                page_numbers=page_numbers,
                elements=elements,
            )

        return _dump(BuildDocumentResult(
            status="success",
            filename=target,
            total_elements=res["total_elements"],
            element_counts=res["element_counts"],
            summary=res["summary"],
            message=f"Document '{target}' successfully built in a single pass ({res['summary']}).",
        ))

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to build document: {error}"),
            filename=target,
            operation="build document",
        )


async def create_from_markdown(
    markdown_content: str,
    filename: str | None = None,
    file_path: str | None = None,
    title: str | None = None,
    author: str | None = None,
    preset: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    line_spacing: float | None = None,
    space_after: float | None = None,
    margins: dict[str, float] | None = None,
    header_text: str | None = None,
    footer_text: str | None = None,
    page_numbers: bool = False,
) -> dict[str, Any]:
    """Create a styled Word document directly from Markdown text in a single pass.

    Supports markdown headers (#, ##, ###), lists (- or 1.), tables (|), blockquotes (>),
    page breaks (---), and inline **bold** and *italic* formatting.
    """
    target, err = _resolve_and_check_path(filename, file_path)
    if err or not target:
        return err or {}

    try:
        with document_context(target, mode="write") as doc:
            res = core_create_from_markdown(
                doc=doc,
                markdown_content=markdown_content,
                title=title,
                author=author,
                preset=preset,
                font_name=font_name,
                font_size=font_size,
                line_spacing=line_spacing,
                space_after=space_after,
                margins=margins,
                header_text=header_text,
                footer_text=footer_text,
                page_numbers=page_numbers,
            )

        return _dump(BuildDocumentResult(
            status="success",
            filename=target,
            total_elements=res["total_elements"],
            element_counts=res["element_counts"],
            summary=res["summary"],
            message=f"Document '{target}' created from Markdown in a single pass ({res['summary']}).",
        ))

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to create document from markdown: {error}"),
            filename=target,
            operation="create from markdown",
        )


async def batch_add_content(
    elements: list[dict[str, Any]],
    filename: str | None = None,
    file_path: str | None = None,
) -> dict[str, Any]:
    """Append multiple content blocks (headings, paragraphs, tables, lists) to an existing document in one pass."""
    target, err = _resolve_and_check_path(filename, file_path)
    if err or not target:
        return err or {}

    if not os.path.exists(target):
        return {
            "status": "error",
            "message": f"Document '{target}' does not exist. Use build_document to create a new one.",
        }

    try:
        counts: dict[str, int] = {}
        with document_context(target, mode="write") as doc:
            for elem in elements:
                if isinstance(elem, dict):
                    t = render_element(doc, elem)
                    counts[t] = counts.get(t, 0) + 1

        total = sum(counts.values())
        summary_parts = [f"{v} {k}s" for k, v in counts.items()]
        summary = ", ".join(summary_parts) if summary_parts else "No elements added"

        return _dump(BuildDocumentResult(
            status="success",
            filename=target,
            total_elements=total,
            element_counts=counts,
            summary=summary,
            message=f"Appended {total} elements to '{target}' in a single pass.",
        ))

    except (OSError, ValueError, DocumentProcessingError) as error:
        return ExceptionTool.handle_error(
            error if isinstance(error, DocumentProcessingError)
            else DocumentProcessingError(f"Failed to batch add content: {error}"),
            filename=target,
            operation="batch add content",
        )
