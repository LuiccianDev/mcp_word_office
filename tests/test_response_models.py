"""Tests for response models in MCP Word Server."""

import pytest
from pydantic import ValidationError

from mcp_word.models.response_models import (
    DocumentInfo,
    DocumentListItem,
    DocumentOutline,
    HeadingInfo,
    OperationError,
    OperationResult,
    PaginatedDocumentList,
    TableInfo,
)


class TestDocumentInfo:
    """Tests for DocumentInfo model."""

    def test_create_document_info_with_all_fields(self):
        """Test creating DocumentInfo with all fields."""
        info = DocumentInfo(
            filename="/path/to/document.docx",
            title="Test Document",
            author="John Doe",
            subject="Test Subject",
            keywords="test, document",
            comments="Test comments",
            last_modified_by="Jane Doe",
            created="2024-01-01T00:00:00",
            modified="2024-01-02T00:00:00",
            word_count=100,
            character_count=500,
            character_count_with_spaces=550,
            page_count=2,
            line_count=50,
            paragraph_count=10,
        )
        assert info.filename == "/path/to/document.docx"
        assert info.title == "Test Document"
        assert info.word_count == 100
        assert info.page_count == 2

    def test_create_document_info_minimal(self):
        """Test creating DocumentInfo with minimal required fields."""
        info = DocumentInfo(
            filename="/path/to/document.docx",
            word_count=0,
            character_count=0,
            character_count_with_spaces=0,
            page_count=0,
            line_count=0,
            paragraph_count=0,
        )
        assert info.filename == "/path/to/document.docx"
        assert info.title is None
        assert info.author is None

    def test_document_info_validation(self):
        """Test that DocumentInfo validates field types."""
        with pytest.raises(ValidationError):
            DocumentInfo(
                filename="/path/to/document.docx",
                word_count="not a number",  # Should be int
                character_count=0,
                character_count_with_spaces=0,
                page_count=0,
                line_count=0,
                paragraph_count=0,
            )


class TestPaginatedDocumentList:
    """Tests for PaginatedDocumentList model."""

    def test_create_paginated_list(self):
        """Test creating PaginatedDocumentList."""
        documents = [
            DocumentListItem(
                name="doc1.docx",
                path="/path/to/doc1.docx",
                size_kb=100.5,
                source_directory="/documents",
            ),
            DocumentListItem(
                name="doc2.docx",
                path="/path/to/doc2.docx",
                size_kb=200.0,
                source_directory="/documents",
            ),
        ]

        result = PaginatedDocumentList(
            status="success",
            message="Found 2 documents",
            directories_searched=["/documents"],
            total=2,
            page=1,
            page_size=20,
            has_more=False,
            next_offset=None,
            documents=documents,
        )

        assert result.status == "success"
        assert result.total == 2
        assert len(result.documents) == 2
        assert result.has_more is False
        assert result.page == 1
        assert result.page_size == 20

    def test_paginated_list_with_more_results(self):
        """Test PaginatedDocumentList with has_more=True."""
        result = PaginatedDocumentList(
            status="success",
            message="Found 50 documents",
            directories_searched=["/docs"],
            total=50,
            page=1,
            page_size=20,
            has_more=True,
            next_offset=2,
            documents=[],
        )

        assert result.has_more is True
        assert result.next_offset == 2
        assert result.total == 50


class TestOperationResult:
    """Tests for OperationResult model."""

    def test_create_success_result(self):
        """Test creating a success OperationResult."""
        result = OperationResult(
            status="success",
            message="Document created successfully",
            details={"filename": "test.docx"},
        )
        assert result.status == "success"
        assert result.message == "Document created successfully"
        assert result.details["filename"] == "test.docx"

    def test_create_success_result_minimal(self):
        """Test creating a minimal success result."""
        result = OperationResult(
            status="success",
            message="Operation completed",
        )
        assert result.status == "success"
        assert result.details is None


class TestOperationError:
    """Tests for OperationError model."""

    def test_create_error(self):
        """Test creating an OperationError."""
        error = OperationError(
            status="error",
            error_type="FileNotFoundError",
            message="File not found",
            suggestion="Check the file path and try again",
            recoverable=True,
        )
        assert error.status == "error"
        assert error.error_type == "FileNotFoundError"
        assert error.recoverable is True

    def test_create_non_recoverable_error(self):
        """Test creating a non-recoverable error."""
        error = OperationError(
            status="error",
            error_type="CorruptedFileError",
            message="File is corrupted",
            recoverable=False,
        )
        assert error.recoverable is False

    def test_error_defaults(self):
        """Test OperationError default values."""
        error = OperationError(
            status="error",
            error_type="TestError",
            message="Test error message",
        )
        assert error.recoverable is True
        assert error.suggestion is None
        assert error.details is None


class TestDocumentOutline:
    """Tests for DocumentOutline model."""

    def test_create_document_outline(self):
        """Test creating DocumentOutline."""
        outline = DocumentOutline(
            filename="/path/to/document.docx",
            title="My Document",
            headings=[
                HeadingInfo(level=1, text="Introduction", position=0),
                HeadingInfo(level=2, text="Background", position=5),
                HeadingInfo(level=2, text="Methods", position=10),
            ],
            tables=[
                TableInfo(index=0, row_count=5, column_count=3, has_header_row=True),
                TableInfo(index=1, row_count=10, column_count=4, has_header_row=True),
            ],
            total_paragraphs=50,
            total_tables=2,
        )

        assert outline.filename == "/path/to/document.docx"
        assert len(outline.headings) == 3
        assert len(outline.tables) == 2
        assert outline.headings[0].level == 1

    def test_empty_outline(self):
        """Test creating an empty DocumentOutline."""
        outline = DocumentOutline(
            filename="/path/to/empty.docx",
            title=None,
            headings=[],
            tables=[],
            total_paragraphs=0,
            total_tables=0,
        )
        assert len(outline.headings) == 0
        assert len(outline.tables) == 0


class TestDocumentListItem:
    """Tests for DocumentListItem model."""

    def test_create_document_list_item(self):
        """Test creating a DocumentListItem."""
        item = DocumentListItem(
            name="report.docx",
            path="/documents/report.docx",
            size_kb=150.75,
            source_directory="/documents",
        )
        assert item.name == "report.docx"
        assert item.size_kb == 150.75

    def test_document_list_item_roundtrip(self):
        """Test that DocumentListItem rounds through serialization."""
        item = DocumentListItem(
            name="test.docx",
            path="/path/test.docx",
            size_kb=50.0,
            source_directory="/docs",
        )
        # Simulate JSON serialization/deserialization
        json_data = item.model_dump()
        restored = DocumentListItem(**json_data)
        assert restored.name == item.name
        assert restored.path == item.path
