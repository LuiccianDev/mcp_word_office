"""Tests for the link_tools module."""

from collections.abc import Generator
from pathlib import Path
import pytest
from docx import Document

from mcp_word.tools import link_tools


@pytest.fixture
def temp_docx_file(tmp_path: Path) -> Generator[str, None, None]:
    """Create a temporary .docx file for testing."""
    file_path = tmp_path / "test_document.docx"
    doc = Document()
    doc.add_paragraph("First paragraph.")
    doc.save(file_path)
    yield str(file_path)


@pytest.mark.asyncio
async def test_add_bookmark(temp_docx_file: str) -> None:
    """Test adding a bookmark."""
    result = await link_tools.add_bookmark(temp_docx_file, 0, "TestMark")
    assert result["status"] == "success", result.get("message", "Unknown error")
    assert "TestMark" in result["message"]
    

@pytest.mark.asyncio
async def test_add_hyperlink_url(temp_docx_file: str) -> None:
    """Test adding an external hyperlink."""
    result = await link_tools.add_hyperlink(temp_docx_file, 0, "Google", url="https://google.com")
    assert result["status"] == "success", result.get("message", "Unknown error")


@pytest.mark.asyncio
async def test_add_hyperlink_bookmark(temp_docx_file: str) -> None:
    """Test adding an internal hyperlink."""
    await link_tools.add_bookmark(temp_docx_file, 0, "TestMark")
    result = await link_tools.add_hyperlink(temp_docx_file, 0, "Go to mark", bookmark="TestMark")
    assert result["status"] == "success", result.get("message", "Unknown error")


@pytest.mark.asyncio
async def test_add_hyperlink_invalid_args(temp_docx_file: str) -> None:
    """Test adding a hyperlink with invalid arguments (both url and bookmark)."""
    result = await link_tools.add_hyperlink(temp_docx_file, 0, "Text", url="http", bookmark="mark")
    assert result["status"] == "error"
    assert "exactly one" in result["message"]
