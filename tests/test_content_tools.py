"""Tests for the content_tools module."""

from pathlib import Path
from typing import Generator
from unittest.mock import MagicMock, patch

import pytest
from docx import Document
from PIL import Image

# Assuming content_tools.py is in src/word_mcp/tools
from word_mcp.tools import content_tools


@pytest.fixture  # type: ignore[misc]
def temp_docx_file(tmp_path: Path) -> Generator[str, None, None]:
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("This is a test document.")
    doc.save(file_path)
    yield str(file_path)


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_heading(temp_docx_file: str) -> None:
    """Test adding a heading to a document."""
    result = await content_tools.add_heading(temp_docx_file, "Test Heading", level=1)
    assert result["status"] == "success"
    assert "Test Heading" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_paragraph(temp_docx_file: str) -> None:
    """Test adding a paragraph to a document."""
    result = await content_tools.add_paragraph(
        temp_docx_file, "This is a new paragraph."
    )
    assert result["status"] == "success"
    assert "Paragraph added" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_table(temp_docx_file: str) -> None:
    """Test adding a table to a document."""
    result = await content_tools.add_table(temp_docx_file, rows=2, cols=3)
    assert result["status"] == "success"
    assert "Table (2x3) added" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_picture_with_patch(temp_docx_file: str, tmp_path: Path) -> None:
    """Test adding a picture using patch."""

    image_path = tmp_path / "test_image.png"

    # Crear una imagen real para evitar errores de validación
    image = Image.new("RGB", (100, 100), color="red")
    image.save(image_path)

    with patch("word_mcp.tools.content_tools.Document") as MockDocument:
        mock_doc = MagicMock()
        MockDocument.return_value = mock_doc

        result = await content_tools.add_picture(temp_docx_file, str(image_path))

        # Verifica que se haya intentado agregar imagen
        mock_doc.add_picture.assert_called_once()
        mock_doc.save.assert_called_once_with(temp_docx_file)

        print("RESULTADO DEL TEST (mockeado):", result)
        assert result["status"] == "success"
        assert "Picture" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_page_break(temp_docx_file: str) -> None:
    """Test adding a page break to a document."""
    result = await content_tools.add_page_break(temp_docx_file)
    assert result["status"] == "success"
    assert "Page break added" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_add_table_of_contents(temp_docx_file: str) -> None:
    """Test adding a table of contents to a document."""
    await content_tools.add_heading(temp_docx_file, "TOC Heading", 1)
    result = await content_tools.add_table_of_contents(temp_docx_file)
    assert result["status"] == "success"
    assert "Table of contents" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_delete_paragraph(temp_docx_file: str) -> None:
    """Test deleting a paragraph from a document."""
    await content_tools.add_paragraph(temp_docx_file, "Paragraph to be deleted.")
    result = await content_tools.delete_paragraph(
        temp_docx_file, 1
    )  # Delete the second paragraph
    assert result["status"] == "success"
    assert "Paragraph at index 1 deleted" in result["message"]


@pytest.mark.asyncio  # type: ignore[misc]
async def test_search_and_replace(temp_docx_file: str) -> None:
    """Test searching and replacing text in a document."""
    result = await content_tools.search_and_replace(temp_docx_file, "test", "sample")
    assert result["status"] == "success"
    assert "Replaced" in result["message"]
