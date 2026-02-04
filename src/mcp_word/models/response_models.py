"""Pydantic models for structured output in MCP Word Server.

These models provide typed, documented responses for all MCP tools,
enabling better LLM understanding and programmatic processing.
"""

from typing import Any

from pydantic import BaseModel, Field


class DocumentInfo(BaseModel):
    """Information about a Word document."""

    filename: str = Field(description="Path to the Word document")
    title: str | None = Field(None, description="Document title")
    author: str | None = Field(None, description="Document author")
    subject: str | None = Field(None, description="Document subject")
    keywords: str | None = Field(None, description="Document keywords")
    comments: str | None = Field(None, description="Document comments")
    last_modified_by: str | None = Field(None, description="Last modified by")
    created: str | None = Field(None, description="Creation date")
    modified: str | None = Field(None, description="Last modification date")
    word_count: int = Field(description="Number of words in document")
    character_count: int = Field(description="Number of characters")
    character_count_with_spaces: int = Field(
        description="Number of characters including spaces"
    )
    page_count: int = Field(description="Number of pages")
    line_count: int = Field(description="Number of lines")
    paragraph_count: int = Field(description="Number of paragraphs")


class DocumentText(BaseModel):
    """Text content extracted from a Word document."""

    filename: str = Field(description="Path to the Word document")
    full_text: str = Field(description="Complete text content of the document")
    paragraph_count: int = Field(description="Number of paragraphs")
    paragraphs: list[str] = Field(
        default_factory=list, description="List of paragraphs"
    )


class HeadingInfo(BaseModel):
    """Information about a heading in a document."""

    level: int = Field(description="Heading level (1-9)")
    text: str = Field(description="Heading text")
    position: int = Field(description="Position/index of heading in document")


class TableInfo(BaseModel):
    """Information about a table in a document."""

    index: int = Field(description="Table index (0-based)")
    row_count: int = Field(description="Number of rows")
    column_count: int = Field(description="Number of columns")
    has_header_row: bool = Field(description="Whether table has a header row")


class DocumentOutline(BaseModel):
    """Structure outline of a Word document."""

    filename: str = Field(description="Path to the Word document")
    title: str | None = Field(None, description="Document title")
    headings: list[HeadingInfo] = Field(
        default_factory=list, description="List of headings"
    )
    tables: list[TableInfo] = Field(default_factory=list, description="List of tables")
    total_paragraphs: int = Field(description="Total number of paragraphs")
    total_tables: int = Field(description="Total number of tables")


class DocumentListItem(BaseModel):
    """Information about a Word document in a list."""

    name: str = Field(description="Document filename")
    path: str = Field(description="Full path to document")
    size_kb: float = Field(description="File size in kilobytes")
    source_directory: str = Field(description="Directory where document was found")


class PaginatedDocumentList(BaseModel):
    """Paginated list of Word documents."""

    status: str = Field(description="Status of the operation")
    message: str = Field(description="Human-readable status message")
    directories_searched: list[str] = Field(
        description="Directories that were searched"
    )
    total: int = Field(description="Total number of documents found")
    page: int = Field(description="Current page number")
    page_size: int = Field(description="Number of documents per page")
    has_more: bool = Field(description="Whether more results exist")
    next_offset: int | None = Field(None, description="Offset for next page")
    documents: list[DocumentListItem] = Field(
        default_factory=list, description="List of documents"
    )


class DocumentSummary(BaseModel):
    """Summary information about a document."""

    filename: str = Field(description="Path to the document")
    exists: bool = Field(description="Whether document exists")
    size_kb: float | None = Field(None, description="File size in KB")
    last_modified: str | None = Field(None, description="Last modification timestamp")


class OperationResult(BaseModel):
    """Standard operation result for successful operations."""

    status: str = Field(description="Status (always 'success')")
    message: str = Field(description="Human-readable success message")
    details: dict[str, Any] | None = Field(
        None, description="Additional operation details"
    )


class OperationError(BaseModel):
    """Structured error response for failed operations."""

    status: str = Field(description="Status (always 'error')")
    error_type: str = Field(description="Type of error")
    message: str = Field(description="Error message")
    suggestion: str | None = Field(None, description="Suggested solution")
    recoverable: bool = Field(
        default=True, description="Whether operation can be retried"
    )
    details: dict[str, Any] | None = Field(None, description="Additional error details")


class HeadingResult(BaseModel):
    """Result of adding a heading to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    heading_text: str = Field(description="Text of the heading added")
    heading_level: int = Field(description="Level of the heading")
    message: str = Field(description="Result message")


class ParagraphResult(BaseModel):
    """Result of adding a paragraph to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    paragraph_text: str = Field(description="Text of paragraph added")
    style_applied: str | None = Field(None, description="Style applied to paragraph")
    message: str = Field(description="Result message")


class TableResult(BaseModel):
    """Result of adding a table to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    rows: int = Field(description="Number of rows")
    columns: int = Field(description="Number of columns")
    data_rows: int | None = Field(None, description="Number of data rows added")
    message: str = Field(description="Result message")


class PictureResult(BaseModel):
    """Result of adding a picture to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    image_name: str = Field(description="Filename of the image added")
    image_path: str = Field(description="Full path to the image")
    width_inches: float | None = Field(None, description="Width in inches if specified")
    message: str = Field(description="Result message")


class TableOfContentsResult(BaseModel):
    """Result of adding a table of contents to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    title: str = Field(description="Title of TOC")
    entry_count: int = Field(description="Number of entries in TOC")
    max_level: int = Field(description="Maximum heading level included")
    message: str = Field(description="Result message")


class SearchReplaceResult(BaseModel):
    """Result of a search and replace operation."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    find_text: str = Field(description="Text that was searched")
    replace_text: str = Field(description="Text used as replacement")
    replacements_made: int = Field(description="Number of replacements performed")
    message: str = Field(description="Result message")


class TextFormatResult(BaseModel):
    """Result of formatting text in a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    paragraph_index: int = Field(description="Index of formatted paragraph")
    text_formatted: str = Field(description="Text that was formatted")
    formatting_applied: dict[str, Any] = Field(description="Formatting options applied")
    message: str = Field(description="Result message")


class StyleResult(BaseModel):
    """Result of creating a custom style."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    style_name: str = Field(description="Name of style created")
    style_type: str = Field(description="Type of style (paragraph, character, etc.)")
    font_properties: dict[str, Any] = Field(
        description="Font properties applied to style"
    )
    message: str = Field(description="Result message")


class TableFormatResult(BaseModel):
    """Result of formatting a table."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    table_index: int = Field(description="Index of formatted table")
    header_row_applied: bool = Field(
        description="Whether header row formatting was applied"
    )
    border_style: str | None = Field(None, description="Border style applied")
    shading_applied: bool = Field(description="Whether shading was applied")
    message: str = Field(description="Result message")


class PDFResult(BaseModel):
    """Result of converting a document to PDF."""

    status: str = Field(description="Operation status")
    success: bool = Field(description="Whether conversion succeeded")
    source_filename: str = Field(description="Path to source Word document")
    pdf_path: str | None = Field(None, description="Path to output PDF file")
    message: str = Field(description="Result message")
    hint: str | None = Field(None, description="Additional hint or suggestion")


class FootnoteResult(BaseModel):
    """Result of adding a footnote to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    footnote_id: int = Field(description="ID of footnote added")
    footnote_text: str = Field(description="Text of the footnote")
    reference_position: str = Field(description="Location of footnote reference")
    message: str = Field(description="Result message")


class EndnoteResult(BaseModel):
    """Result of adding an endnote to a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    endnote_id: int = Field(description="ID of endnote added")
    endnote_text: str = Field(description="Text of the endnote")
    reference_position: str = Field(description="Location of endnote reference")
    message: str = Field(description="Result message")


class ProtectionResult(BaseModel):
    """Result of protecting or unprotecting a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    protection_type: str = Field(description="Type of protection applied or removed")
    success: bool = Field(description="Whether operation succeeded")
    message: str = Field(description="Result message")


class ParagraphTextResult(BaseModel):
    """Result of getting text from a specific paragraph."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    paragraph_index: int = Field(description="Index of the paragraph")
    text: str = Field(description="Text content of paragraph")
    character_count: int = Field(description="Number of characters in paragraph")
    style_name: str | None = Field(None, description="Style applied to paragraph")


class TextSearchResult(BaseModel):
    """Result of searching for text in a document."""

    status: str = Field(description="Operation status")
    filename: str = Field(description="Path to the document")
    search_term: str = Field(description="Text that was searched")
    match_case: bool = Field(description="Whether case-sensitive matching was used")
    whole_word: bool = Field(description="Whether whole word matching was used")
    match_count: int = Field(description="Number of matches found")
    matches: list[dict[str, Any]] = Field(
        default_factory=list, description="List of match locations"
    )
    message: str = Field(description="Result message")
